/**
 * @NApiVersion 2.1
 * @NScriptType ScheduledScript
 * @NModuleScope SameAccount
 *
 * Statscore JE Creator - posting script.
 *
 * Started by statscore_je_suitelet.js, never on a schedule of its own. It
 * reads the staged CSV, rebuilds the lines with the shared library and saves
 * one journal entry, writing progress back to the job file the Suitelet polls.
 *
 * The work lives here rather than in the Suitelet because a month of
 * Statscore data is 4,000-5,000 lines and the save alone runs for minutes,
 * well past what a user-interface request is allowed.
 */
define([
    'N/runtime',
    'N/record',
    'N/file',
    'N/query',
    'N/email',
    'N/url',
    'N/log',
    './statscore_je_lib'
], (runtime, record, file, query, email, url, log, lib) => {

    const JOB_PARAM = 'custscript_statscore_je_job';
    const COPY_PARAM = 'custscript_statscore_je_copyto';
    const MAX_LINE_MEMO = 999;

    function execute() {
        const started = new Date();
        const jobFileId = runtime.getCurrentScript().getParameter({ name: JOB_PARAM });

        if (!jobFileId) {
            log.error({
                title: 'No job file',
                details: 'Parameter ' + JOB_PARAM + ' is empty. This script is started by the ' +
                    'Statscore JE Creator Suitelet, not on a schedule.'
            });
            alertCopyList('Statscore JE posting script started with no job',
                'The posting script ran but parameter <b>' + esc(JOB_PARAM) + '</b> was empty, so ' +
                'there was nothing to post. This usually means the deployment was given a ' +
                'schedule; it is meant to be started only by the Suitelet.');
            return;
        }

        let job;
        try {
            job = JSON.parse(file.load({ id: jobFileId }).getContents());
        } catch (e) {
            log.error({ title: 'Could not read job file ' + jobFileId, details: e });
            alertCopyList('Statscore JE posting could not start',
                'The posting script could not read its job file (id ' + esc(jobFileId) + '), so ' +
                'nothing was posted. If the file was deleted deliberately to cancel a job, this ' +
                'is expected and no action is needed.');
            return;
        }

        // Posting creates real financial records, so only ever act on a job
        // that has not been touched. A job left at RUNNING means an earlier
        // attempt died mid-save and needs a human to check NetSuite first.
        if (job.status !== 'PENDING') {
            log.audit({
                title: 'Job ' + jobFileId + ' skipped',
                details: 'Status is ' + job.status + ', not PENDING. Nothing posted.'
            });
            alertCopyList('Statscore JE job skipped: ' + (job.memo || 'unknown period'),
                'A posting task fired for a job already marked <b>' + esc(job.status) + '</b>, so ' +
                'nothing was posted. A job stuck at RUNNING means an earlier attempt died ' +
                'mid-save: check NetSuite for a partially created entry for ' +
                esc(job.memo || 'that period') + ' before running it again.');
            return;
        }

        job.status = 'RUNNING';
        job.startedAt = started.toISOString();
        save(jobFileId, job);

        try {
            const analysis = lib.analyze(file.load({ id: job.csvFileId }).getContents());
            if (!analysis.ok) {
                throw new Error('The staged file no longer validates: ' + analysis.errors.join(' '));
            }

            if (analysis.accountIds.length) {
                const accounts = lib.validateAccounts(analysis.accountIds);
                if (accounts.invalid.length) {
                    throw new Error('Account internal ID(s) not found in NetSuite: ' +
                        accounts.invalid.slice(0, 25).join(', ') +
                        (accounts.invalid.length > 25 ? ' and more' : ''));
                }
            }

            const jeId = buildAndSave(analysis);
            job.status = 'DONE';
            job.jeId = jeId;
            job.jeTranId = tranIdOf(jeId);
            job.lineCount = analysis.lines.length;
            job.finishedAt = new Date().toISOString();
            job.error = '';
            save(jobFileId, job);

            log.audit({
                title: 'Statscore JE created',
                details: 'id=' + jeId + ' tranid=' + job.jeTranId +
                    ' lines=' + analysis.lines.length +
                    ' seconds=' + Math.round((new Date() - started) / 1000) +
                    ' remainingUsage=' + runtime.getCurrentScript().getRemainingUsage()
            });

            notify(job);

        } catch (e) {
            log.error({ title: 'Statscore JE posting failed', details: e });
            job.status = 'ERROR';
            job.error = (e.name ? e.name + ': ' : '') + (e.message || String(e));
            job.finishedAt = new Date().toISOString();
            save(jobFileId, job);
            notify(job);
        }
    }

    /**
     * Build the journal entry and save it.
     *
     * Standard (non-dynamic) mode with sourcing off: dynamic mode would
     * re-source every line and turn a few minutes into far longer.
     *
     * @param {Object} analysis result from lib.analyze
     * @returns {number} internal ID of the new journal entry
     */
    function buildAndSave(analysis) {
        const rec = record.create({ type: record.Type.JOURNAL_ENTRY, isDynamic: false });

        rec.setValue({ fieldId: 'subsidiary', value: lib.CONFIG.SUBSIDIARY });
        rec.setValue({ fieldId: 'currency', value: lib.CONFIG.CURRENCY });
        rec.setValue({ fieldId: 'trandate', value: dateOf(analysis.tranDate) });
        rec.setValue({ fieldId: 'memo', value: analysis.memo });

        // 'approved' does not exist when approval routing is switched on for
        // journals; the entry is then created pending approval instead.
        try {
            rec.setValue({ fieldId: 'approved', value: true });
        } catch (e) {
            log.audit({
                title: 'Could not set Approved',
                details: 'The journal will follow the account approval workflow. ' + e.message
            });
        }

        analysis.lines.forEach((ln, i) => {
            rec.insertLine({ sublistId: 'line', line: i });
            setLine(rec, i, 'account', Number(ln.account));
            if (ln.debit) setLine(rec, i, 'debit', ln.debit);
            if (ln.credit) setLine(rec, i, 'credit', ln.credit);
            if (ln.memo) setLine(rec, i, 'memo', ln.memo.slice(0, MAX_LINE_MEMO));
            setLine(rec, i, 'department', lib.CONFIG.DEPARTMENT);
            setLine(rec, i, 'location', lib.CONFIG.LOCATION);
            setLine(rec, i, 'cseg_location_exp', lib.CONFIG.CSEG_LOCATION_EXP);
        });

        log.audit({
            title: 'Saving journal entry',
            details: analysis.lines.length + ' lines, debit ' + lib.fmt(analysis.totalDebit) +
                ', credit ' + lib.fmt(analysis.totalCredit) + ', date ' + analysis.tranDate
        });

        return rec.save({ enableSourcing: false, ignoreMandatoryFields: false });
    }

    function setLine(rec, line, fieldId, value) {
        rec.setSublistValue({ sublistId: 'line', fieldId: fieldId, line: line, value: value });
    }

    /** 'YYYY-MM-DD' to a Date in the account's timezone. */
    function dateOf(iso) {
        const p = iso.split('-');
        return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    }

    function tranIdOf(jeId) {
        try {
            const rows = query.runSuiteQL({
                query: 'SELECT tranid FROM transaction WHERE id = ?',
                params: [jeId]
            }).asMappedResults();
            return rows.length ? String(rows[0].tranid) : '';
        } catch (e) {
            log.error({ title: 'Could not read the JE number', details: e });
            return '';
        }
    }

    /**
     * Record progress on the job file the status page polls.
     *
     * A failure here cannot stop the posting, but it must never be silent:
     * a stale "Queued" on a job that actually finished is what got a month
     * posted twice. lib.writeJson verifies the write and throws if it did
     * not take, so the execution log always shows it.
     */
    function save(jobFileId, job) {
        try {
            lib.writeJson(jobFileId, job);
        } catch (e) {
            log.error({
                title: 'Could not update job file ' + jobFileId,
                details: 'The status page will keep showing the previous state for this job. ' +
                    'The journal entry itself is unaffected. ' + (e.message || String(e))
            });
        }
    }

    // -----------------------------------------------------------------------
    // Result notification
    //
    // Added because the posting job can sit in NetSuite's shared script queue
    // for the best part of an hour. Without an email the only way to learn the
    // outcome is to keep the status page open, which is what led someone to
    // resubmit a job that had not failed and post a month twice.
    // -----------------------------------------------------------------------

    /**
     * Who is copied on every result, read from the deployment parameter.
     * Accepts employee internal IDs and email addresses, comma or semicolon
     * separated. Empty is fine: the submitter still gets their email.
     *
     * @returns {Array<number|string>}
     */
    function copyList() {
        const raw = runtime.getCurrentScript().getParameter({ name: COPY_PARAM });
        if (!raw) return [];
        return String(raw)
            .split(/[,;]/)
            .map((v) => v.trim())
            .filter((v) => v !== '')
            .map((v) => (/^\d+$/.test(v) ? Number(v) : v));
    }

    function firstEmployeeId(list) {
        for (let i = 0; i < list.length; i++) {
            if (typeof list[i] === 'number') return list[i];
        }
        return null;
    }

    /**
     * Email the outcome of a finished job to whoever submitted it, copying
     * the configured list.
     *
     * Only ever called after the job status has been written, so a mail
     * failure cannot change what was posted or what the status page reports.
     * Everything is caught: a bounced notification must not turn a good run
     * into a reported failure.
     */
    function notify(job) {
        try {
            const copies = copyList();
            const submitter = Number(job.submittedById) || null;
            const cc = copies.filter((v) => v !== submitter);   // never address one person twice
            const author = submitter || firstEmployeeId(copies) || runtime.getCurrentUser().id;

            if (!author) {
                log.audit({
                    title: 'No result email sent',
                    details: 'No submitter on the job and no employee ID in ' + COPY_PARAM + '.'
                });
                return;
            }

            const recipients = submitter ? [submitter] : cc;
            if (!recipients.length) {
                log.audit({
                    title: 'No result email sent',
                    details: 'Nobody to send to. Set ' + COPY_PARAM + ' on the deployment.'
                });
                return;
            }

            const ok = job.status === 'DONE';
            const period = pad2(job.month) + '/' + job.year;
            const options = {
                author: author,
                recipients: recipients,
                subject: ok
                    ? 'Statscore JE posted: ' + period + (job.jeTranId ? ' - ' + job.jeTranId : '')
                    : 'Statscore JE FAILED: ' + period,
                body: resultHtml(job)
            };
            if (submitter && cc.length) {
                options.cc = cc;
            }

            email.send(options);
            log.audit({
                title: 'Result emailed',
                details: 'to=' + recipients.join(',') + ' cc=' + (options.cc || []).join(',')
            });

        } catch (e) {
            log.error({ title: 'Could not send the result email', details: e });
        }
    }

    /**
     * Alert for the aborts that used to be silent: no job parameter, an
     * unreadable job file, or a job that was not PENDING. There is no known
     * submitter in those cases, so it goes to the copy list only.
     */
    function alertCopyList(subject, message) {
        try {
            const copies = copyList();
            if (!copies.length) return;                  // nobody configured; the log has it
            email.send({
                author: firstEmployeeId(copies) || runtime.getCurrentUser().id,
                recipients: copies,
                subject: subject,
                body: shell('Statscore JE Creator', '#c87f0a', 'Attention', '<p>' + message + '</p>')
            });
        } catch (e) {
            log.error({ title: 'Could not send the alert email', details: e });
        }
    }

    function resultHtml(job) {
        const ok = job.status === 'DONE';
        const period = pad2(job.month) + '/' + job.year;
        let intro;

        if (ok) {
            const link = jeLink(job.jeId);
            intro = '<p style="margin:0 0 14px 0">Journal entry <b>' +
                esc(job.jeTranId || ('internal ID ' + job.jeId)) +
                '</b> was created for <b>' + esc(period) + '</b>.</p>' +
                (link
                    ? '<p style="margin:0 0 18px 0"><a href="' + esc(link) +
                      '" style="display:inline-block;padding:9px 16px;background:#1e8449;' +
                      'color:#ffffff;text-decoration:none;border-radius:3px;font-weight:bold">' +
                      'Open the journal entry</a></p>'
                    : '');
        } else {
            intro = '<p style="margin:0 0 10px 0">Nothing was posted for <b>' + esc(period) +
                '</b>. The run failed with:</p>' +
                '<p style="margin:0 0 14px 0;padding:10px 12px;background:#fdeeec;' +
                'border-left:4px solid #c0392b;font-family:Consolas,monospace;font-size:12px">' +
                esc(job.error || 'No detail recorded.') + '</p>' +
                '<p style="margin:0 0 18px 0">The stack trace is in NetSuite under ' +
                '<b>Customization &gt; Scripting &gt; Script Execution Log</b>, filtered to ' +
                '<b>Statscore JE Creator - Post</b>. Nothing was left half-posted: the entry is ' +
                'saved in one operation, so it either exists in full or not at all.</p>';
        }

        const rows = [
            ['Period', period],
            ['Transaction date', job.tranDate],
            ['Memo', job.memo],
            ['Subsidiary', 'Statscore'],
            ['Currency', 'EUR'],
            ['Lines', String(job.lineCount)],
            ['Total debit', lib.fmt(job.totalDebit)],
            ['Total credit', lib.fmt(job.totalCredit)],
            ['Source file', job.csvName || ''],
            ['Submitted by', job.submittedBy || ''],
            ['Took', duration(job.startedAt, job.finishedAt)]
        ];
        if (job.confirmedDuplicate) {
            rows.push(['Duplicate period', 'Confirmed and posted anyway']);
        }

        return shell(ok ? 'Statscore JE posted' : 'Statscore JE failed',
            ok ? '#1e8449' : '#c0392b',
            ok ? 'Posted' : 'Failed',
            intro + table(rows));
    }

    /** Outer layout. Inline styles only: mail clients strip style blocks. */
    function shell(title, colour, badge, body) {
        return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;' +
            'color:#222222;max-width:640px">' +
            '<div style="border-left:5px solid ' + colour + ';padding-left:14px;margin-bottom:18px">' +
            '<div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:' +
            colour + ';font-weight:bold">' + esc(badge) + '</div>' +
            '<div style="font-size:19px;font-weight:bold;margin-top:2px">' + esc(title) + '</div>' +
            '</div>' +
            body +
            '<p style="margin:22px 0 0 0;padding-top:12px;border-top:1px solid #e4e7eb;' +
            'font-size:11px;color:#888888">Sent automatically by the Statscore JE Creator ' +
            'in NetSuite.</p>' +
            '</div>';
    }

    function table(rows) {
        let html = '<table cellpadding="0" cellspacing="0" ' +
            'style="border-collapse:collapse;font-size:13px">';
        rows.forEach((r) => {
            html += '<tr>' +
                '<td style="padding:4px 20px 4px 0;color:#555555;white-space:nowrap;' +
                'vertical-align:top">' + esc(r[0]) + '</td>' +
                '<td style="padding:4px 0;font-weight:bold;vertical-align:top">' +
                esc(r[1]) + '</td></tr>';
        });
        return html + '</table>';
    }

    /** Absolute link: url.resolveRecord alone gives a path, useless in email. */
    function jeLink(jeId) {
        if (!jeId) return '';
        try {
            return 'https://' + url.resolveDomain({ hostType: url.HostType.APPLICATION }) +
                url.resolveRecord({
                    recordType: 'journalentry',
                    recordId: jeId,
                    isEditMode: false
                });
        } catch (e) {
            log.error({ title: 'Could not build the journal entry link', details: e });
            return '';
        }
    }

    function duration(fromIso, toIso) {
        const a = Date.parse(fromIso);
        const b = Date.parse(toIso);
        if (isNaN(a) || isNaN(b)) return '';
        const secs = Math.round((b - a) / 1000);
        return secs < 60 ? secs + ' seconds'
            : Math.floor(secs / 60) + ' min ' + (secs % 60) + ' sec';
    }

    function pad2(n) {
        return (Number(n) < 10 ? '0' : '') + n;
    }

    function esc(s) {
        return String(s === null || s === undefined ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    return { execute: execute };
});
