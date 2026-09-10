/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * Statscore JE Creator - Suitelet (the screen the operator uses).
 *
 * Replaces the desktop Python tool. Three pages, one script:
 *   1. Upload   - pick the Statscore CSV
 *   2. Preview  - period, totals, REF balance, account validation, sample lines
 *   3. Status   - progress of the posting job, then a link to the new JE
 *
 * The JE itself is built by statscore_je_sched.js. A month of Statscore data
 * is roughly 4,000-5,000 lines, which takes longer than a Suitelet request is
 * allowed to run, so the posting is handed to a scheduled script and this
 * page polls for the result.
 */
define([
    'N/ui/serverWidget',
    'N/file',
    'N/query',
    'N/record',
    'N/task',
    'N/runtime',
    'N/url',
    'N/log',
    './statscore_je_lib'
], (serverWidget, file, query, record, task, runtime, url, log, lib) => {

    const SCHED_SCRIPT_ID = 'customscript_statscore_je_sched';
    const SCHED_DEPLOY_ID = 'customdeploy_statscore_je_sched';
    const JOB_PARAM = 'custscript_statscore_je_job';

    const WORK_FOLDER_NAME = 'Statscore JE Creator';
    const FALLBACK_FOLDER = -15;          // SuiteScripts
    const SAMPLE_LINES = 25;
    const MAX_ACCOUNTS_SHOWN = 250;
    const MAX_ERRORS_SHOWN = 30;
    const MAX_JOBS_SCANNED = 20;       // recent job files checked for an in-flight run

    // -----------------------------------------------------------------------
    // Entry point
    // -----------------------------------------------------------------------

    function onRequest(context) {
        const req = context.request;
        const res = context.response;

        try {
            if (req.method === 'GET') {
                if (req.parameters.action === 'status') {
                    return renderStatus(res, req.parameters.job);
                }
                return renderUpload(res);
            }

            if (req.parameters.custpage_action === 'create') {
                return submitJob(req, res);
            }
            return renderPreview(req, res);

        } catch (e) {
            log.error({ title: 'Statscore JE Creator failed', details: e });
            const form = serverWidget.createForm({ title: 'Statscore JE Creator' });
            addHtml(form, 'err', panel('error', 'Something went wrong',
                '<p>' + esc(e.message || String(e)) + '</p>' +
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            res.writePage(form);
        }
    }

    // -----------------------------------------------------------------------
    // Page 1: upload
    // -----------------------------------------------------------------------

    function renderUpload(res) {
        const form = serverWidget.createForm({ title: 'Statscore JE Creator' });

        addHtml(form, 'intro', styles() + panel('info', 'What this does',
            '<p>Pick the Statscore CSV. The next screen shows the period, the totals and the ' +
            'account check. Nothing is posted until you confirm it there.</p>' +
            '<ul>' +
            '<li>Every row becomes one journal entry under <b>Statscore</b> (subsidiary ' +
                lib.CONFIG.SUBSIDIARY + '), in <b>EUR</b>.</li>' +
            '<li>Rows with zero debit and zero credit are skipped.</li>' +
            '<li>A negative debit posts as a credit, a negative credit posts as a debit.</li>' +
            '<li>Every REF group must balance to within ' +
                lib.CONFIG.BALANCE_TOLERANCE.toFixed(2) + ' or the run is blocked.</li>' +
            '</ul>'));

        const csv = form.addField({
            id: 'custpage_csv',
            type: serverWidget.FieldType.FILE,
            label: 'Statscore CSV File'
        });
        csv.isMandatory = true;

        form.addSubmitButton({ label: 'Preview' });
        res.writePage(form);
    }

    // -----------------------------------------------------------------------
    // Page 2: preview
    // -----------------------------------------------------------------------

    function renderPreview(req, res) {
        const form = serverWidget.createForm({ title: 'Statscore JE Creator - Preview' });
        let html = styles();

        const uploaded = req.files ? req.files.custpage_csv : null;
        if (!uploaded) {
            addHtml(form, 'err', html + panel('error', 'No file received',
                '<p>Pick a CSV file and try again.</p>' +
                '<p><a href="' + esc(selfUrl()) + '">Back to upload</a></p>'));
            return res.writePage(form);
        }

        const csvText = readText(uploaded);
        const analysis = lib.analyze(csvText);

        // Account validation runs here so the operator sees bad IDs before the
        // long-running post rather than after it.
        let accounts = null;
        if (analysis.accountIds.length) {
            try {
                accounts = lib.validateAccounts(analysis.accountIds);
                if (accounts.invalid.length) {
                    analysis.errors.push(accounts.invalid.length +
                        ' account internal ID(s) do not exist in NetSuite.');
                    analysis.ok = false;
                }
                if (accounts.inactive.length) {
                    analysis.warnings.push(accounts.inactive.length +
                        ' account(s) are inactive. NetSuite will reject those lines.');
                }
            } catch (e) {
                log.error({ title: 'Account validation failed', details: e });
                analysis.warnings.push('Accounts could not be validated: ' + (e.message || String(e)));
            }
        }

        const duplicates = findDuplicates(analysis);

        // A job already queued for this period blocks outright. It is not a
        // warning with an override: there is no good reason to queue two.
        let inFlight = null;
        if (analysis.memo) {
            try {
                inFlight = findInFlightJob(analysis.memo, false);
            } catch (e) {
                inFlight = null;
            }
        }
        if (inFlight) {
            analysis.ok = false;
        }

        html += renderSummary(analysis, uploaded.name);
        html += renderInFlight(inFlight);
        html += renderDuplicates(duplicates);
        html += renderProblems(analysis);
        html += renderSample(analysis);
        html += renderAccounts(accounts);

        if (!analysis.ok) {
            html += panel('error', 'Blocked',
                '<p>Fix the source file and upload it again. Nothing has been posted.</p>');
            addHtml(form, 'body', html + backLink());
            return res.writePage(form);
        }

        let csvFileId;
        try {
            csvFileId = stageCsv(uploaded.name, csvText);
        } catch (e) {
            log.error({ title: 'Could not stage CSV', details: e });
            addHtml(form, 'body', html + panel('error', 'Could not stage the file',
                '<p>' + esc(e.message || String(e)) + '</p>'));
            return res.writePage(form);
        }

        html += panel('info', 'Ready to post',
            '<p>Creating the entry takes a few minutes for a file this size. It runs in the ' +
            'background, so you can leave the status page open or come back to it later.</p>');

        // The HTML goes on the form before the checkbox so the warning the
        // checkbox refers to is above it on screen.
        addHtml(form, 'body', html + backLink());

        if (duplicates.length) {
            form.addField({
                id: 'custpage_confirm_dup',
                type: serverWidget.FieldType.CHECKBOX,
                label: 'Post anyway - ' + (analysis.memo || 'this period') +
                    ' is already in NetSuite'
            }).defaultValue = 'F';
        }
        hidden(form, 'custpage_action', 'create');
        hidden(form, 'custpage_csvfile', String(csvFileId));
        hidden(form, 'custpage_csvname', uploaded.name);
        form.addSubmitButton({ label: 'Create Journal Entry' });

        res.writePage(form);
    }

    /**
     * Journal entries already posted for this period. A failed lookup is a
     * warning, not a blocker: it must not stop a legitimate first post.
     */
    function findDuplicates(analysis) {
        if (!analysis.memo) return [];
        try {
            return lib.findExistingJes(analysis.memo);
        } catch (e) {
            log.error({ title: 'Duplicate check failed', details: e });
            analysis.warnings.push('Could not check whether this period is already posted: ' +
                (e.message || String(e)));
            return [];
        }
    }

    /**
     * A job for this period that is already queued or running.
     *
     * The posted-entry check cannot see these. NetSuite runs scheduled scripts
     * from a shared queue that can hold a job for the best part of an hour,
     * and a page that only says "Queued" invites a second submission that
     * would post the month twice.
     *
     * @param {string} memo   the period memo, e.g. 'Statscore 08/2026'
     * @param {boolean} strict  throw on a lookup failure instead of returning null
     * @returns {{fileId: number, job: Object}|null}
     */
    function findInFlightJob(memo, strict) {
        let rows;
        try {
            rows = query.runSuiteQL({
                query: 'SELECT id FROM file WHERE folder = ? AND name LIKE ? ORDER BY id DESC',
                params: [workFolder(), '%.job.json']
            }).asMappedResults();
        } catch (e) {
            log.error({ title: 'Could not list job files', details: e });
            if (strict) throw e;
            return null;
        }

        for (let i = 0; i < rows.length && i < MAX_JOBS_SCANNED; i++) {
            let job;
            try {
                job = JSON.parse(file.load({ id: rows[i].id }).getContents());
            } catch (e) {
                continue;                       // deleted or unreadable, skip it
            }
            if (job.memo === memo && (job.status === 'PENDING' || job.status === 'RUNNING')) {
                return { fileId: Number(rows[i].id), job: job };
            }
        }
        return null;
    }

    function renderInFlight(inFlight) {
        if (!inFlight) return '';

        const queued = inFlight.job.status !== 'RUNNING';
        return panel('error', 'This period is already being posted',
            '<p>A job for <b>' + esc(inFlight.job.memo) + '</b> ' +
            (queued ? 'is queued and has not started yet' : 'is being posted right now') +
            ', submitted by ' + esc(inFlight.job.submittedBy || 'someone') + ' ' +
            esc(since(inFlight.job.submittedAt)) + '.</p>' +
            '<p><a href="' + esc(statusUrl(inFlight.fileId)) + '">Watch that job</a></p>' +
            '<p>Posting again would create a second journal entry for the same month, so this ' +
            'file cannot be submitted until that job finishes. If it is genuinely stuck, delete ' +
            'its <code>.job.json</code> file from the <b>' + esc(WORK_FOLDER_NAME) +
            '</b> folder in the File Cabinet and reload this page.</p>');
    }

    /** '12 minutes ago' from an ISO timestamp. */
    function since(iso) {
        if (!iso) return '';
        const then = Date.parse(iso);
        if (isNaN(then)) return '';
        const mins = Math.floor((Date.now() - then) / 60000);
        if (mins < 1) return 'less than a minute ago';
        if (mins === 1) return '1 minute ago';
        if (mins < 60) return mins + ' minutes ago';
        const hrs = Math.floor(mins / 60);
        return hrs + (hrs === 1 ? ' hour ' : ' hours ') + (mins % 60) + ' min ago';
    }

    function renderDuplicates(duplicates) {
        if (!duplicates.length) return '';

        let body = '<p>' + (duplicates.length === 1
            ? 'This period is already posted as:'
            : 'This period is already posted as ' + duplicates.length + ' entries:') + '</p><ul>';
        duplicates.forEach((d) => {
            const link = url.resolveRecord({
                recordType: 'journalentry',
                recordId: d.id,
                isEditMode: false
            });
            body += '<li><a href="' + esc(link) + '" target="_blank">' +
                esc(d.tranid || ('Internal ID ' + d.id)) + '</a> dated ' + esc(d.trandate) + '</li>';
        });
        body += '</ul><p>Posting again doubles the period. Tick the confirmation box at the bottom ' +
            'if that is genuinely what you want.</p>';

        return panel('warn', 'A journal entry for this period already exists', body);
    }

    /**
     * Entries for this period that were not already there when the job was
     * submitted, so almost certainly this job's own output.
     */
    function newlyPosted(job) {
        if (!job.memo) return [];
        const before = (job.preExistingJeIds || []).map(String);
        try {
            return lib.findExistingJes(job.memo)
                .filter((d) => before.indexOf(String(d.id)) === -1);
        } catch (e) {
            log.error({ title: 'Could not cross-check posted entries', details: e });
            return [];
        }
    }

    function jeUrl(jeId) {
        return url.resolveRecord({
            recordType: 'journalentry',
            recordId: jeId,
            isEditMode: false
        });
    }

    function backLink() {
        return '<p style="margin-top:14px"><a href="' + esc(selfUrl()) +
            '">Upload a different file</a></p>';
    }

    function renderSummary(a, fileName) {
        const diff = Math.abs(a.totalDebit - a.totalCredit);
        const balanced = diff < lib.CONFIG.BALANCE_TOLERANCE;

        let rows = '';
        rows += kv('File', esc(fileName));
        rows += kv('Rows read', String(a.dataRows));
        rows += kv('Period', a.month ? pad2(a.month) + '/' + a.year : '-');
        rows += kv('Subsidiary', 'Statscore (' + lib.CONFIG.SUBSIDIARY + ')');
        rows += kv('Currency', 'EUR (' + lib.CONFIG.CURRENCY + ')');
        rows += kv('Department', 'Statscore import (' + lib.CONFIG.DEPARTMENT + ')');
        rows += kv('Location', 'Poland (' + lib.CONFIG.LOCATION + ')');
        rows += kv('Location (Expenses)', 'Other (' + lib.CONFIG.CSEG_LOCATION_EXP + ')');
        rows += kv('Transaction date', esc(a.tranDate || '-'));
        rows += kv('Memo', esc(a.memo || '-'));
        rows += kv('Approved', 'Yes');
        rows += kv('JE lines', String(a.lines.length) +
            (a.skippedZero ? ' <span class="sc-muted">(' + a.skippedZero + ' zero-amount rows skipped)</span>' : ''));
        rows += kv('Unique accounts', String(a.accountIds.length));
        rows += kv('Total debit', lib.fmt(a.totalDebit));
        rows += kv('Total credit', lib.fmt(a.totalCredit));
        rows += kv('Balance', balanced
            ? '<span class="sc-ok">Balanced</span>'
            : '<span class="sc-bad">Out by ' + lib.fmt(diff) + '</span>');
        if (a.refGroupsOk || a.refErrors.length) {
            rows += kv('REF groups', a.refErrors.length
                ? '<span class="sc-bad">' + a.refErrors.length + ' unbalanced</span>, ' + a.refGroupsOk + ' balanced'
                : '<span class="sc-ok">All ' + a.refGroupsOk + ' balanced</span>');
        }

        return '<table class="sc-kv">' + rows + '</table>';
    }

    function renderProblems(a) {
        let html = '';

        if (a.errors.length) {
            html += panel('error', 'Blocking problems',
                '<ul><li>' + a.errors.map(esc).join('</li><li>') + '</li></ul>');
        }
        if (a.warnings.length) {
            html += panel('warn', 'Warnings',
                '<ul><li>' + a.warnings.map(esc).join('</li><li>') + '</li></ul>');
        }

        if (a.refErrors.length) {
            let t = '<table class="sc-grid"><tr><th>REF</th><th class="sc-num">Debit</th>' +
                '<th class="sc-num">Credit</th><th class="sc-num">Difference</th></tr>';
            a.refErrors.slice(0, MAX_ERRORS_SHOWN).forEach((r) => {
                t += '<tr><td>' + esc(r.ref) + '</td><td class="sc-num">' + lib.fmt(r.debit) +
                    '</td><td class="sc-num">' + lib.fmt(r.credit) +
                    '</td><td class="sc-num sc-bad">' + lib.fmt(r.diff) + '</td></tr>';
            });
            t += '</table>';
            if (a.refErrors.length > MAX_ERRORS_SHOWN) {
                t += '<p class="sc-muted">and ' + (a.refErrors.length - MAX_ERRORS_SHOWN) + ' more</p>';
            }
            html += section('Unbalanced REF groups', t);
        }

        if (a.rowErrors.length) {
            let t = '<table class="sc-grid"><tr><th>CSV line</th><th>Problem</th></tr>';
            a.rowErrors.slice(0, MAX_ERRORS_SHOWN).forEach((r) => {
                t += '<tr><td>' + r.csvLine + '</td><td>' + esc(r.message) + '</td></tr>';
            });
            t += '</table>';
            if (a.rowErrors.length > MAX_ERRORS_SHOWN) {
                t += '<p class="sc-muted">and ' + (a.rowErrors.length - MAX_ERRORS_SHOWN) + ' more</p>';
            }
            html += section('Rows that could not be converted', t);
        }

        return html;
    }

    function renderSample(a) {
        if (!a.lines.length) return '';

        let t = '<table class="sc-grid"><tr><th>Account</th><th class="sc-num">Debit</th>' +
            '<th class="sc-num">Credit</th><th>Memo</th></tr>';
        a.lines.slice(0, SAMPLE_LINES).forEach((ln) => {
            t += '<tr><td>' + esc(ln.account) + '</td><td class="sc-num">' +
                (ln.debit ? lib.fmt(ln.debit) : '') + '</td><td class="sc-num">' +
                (ln.credit ? lib.fmt(ln.credit) : '') + '</td><td>' +
                esc(ln.memo.slice(0, 80)) + '</td></tr>';
        });
        t += '</table>';
        if (a.lines.length > SAMPLE_LINES) {
            t += '<p class="sc-muted">and ' + (a.lines.length - SAMPLE_LINES) + ' more lines</p>';
        }
        return section('First ' + Math.min(SAMPLE_LINES, a.lines.length) + ' lines', t);
    }

    function renderAccounts(accounts) {
        if (!accounts) return '';

        let html = '';
        if (accounts.invalid.length) {
            html += panel('error', 'Account IDs not found in NetSuite',
                '<p>' + esc(accounts.invalid.slice(0, 100).join(', ')) +
                (accounts.invalid.length > 100 ? ' and ' + (accounts.invalid.length - 100) + ' more' : '') +
                '</p>');
        }

        const ids = Object.keys(accounts.valid).sort((x, y) => Number(x) - Number(y));
        if (!ids.length) return html;

        let t = '<table class="sc-grid"><tr><th>ID</th><th>Number</th><th>Name</th><th>Status</th></tr>';
        ids.slice(0, MAX_ACCOUNTS_SHOWN).forEach((id) => {
            const v = accounts.valid[id];
            t += '<tr><td>' + esc(id) + '</td><td>' + esc(v.acctnumber) + '</td><td>' +
                esc(v.fullname) + '</td><td>' + (v.isinactive
                    ? '<span class="sc-bad">Inactive</span>'
                    : '<span class="sc-ok">Active</span>') + '</td></tr>';
        });
        t += '</table>';
        if (ids.length > MAX_ACCOUNTS_SHOWN) {
            t += '<p class="sc-muted">and ' + (ids.length - MAX_ACCOUNTS_SHOWN) + ' more accounts</p>';
        }

        return html + section('Accounts (' + ids.length + ' matched)',
            '<div class="sc-scroll">' + t + '</div>');
    }

    // -----------------------------------------------------------------------
    // Page 3: submit the posting job
    // -----------------------------------------------------------------------

    function submitJob(req, res) {
        const form = serverWidget.createForm({ title: 'Statscore JE Creator' });
        const csvFileId = Number(req.parameters.custpage_csvfile);

        if (!csvFileId) {
            addHtml(form, 'err', styles() + panel('error', 'Lost the staged file',
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }

        // Re-check the staged file rather than trusting what the browser posted back.
        const csvText = file.load({ id: csvFileId }).getContents();
        const analysis = lib.analyze(csvText);
        if (analysis.ok && analysis.accountIds.length) {
            const accounts = lib.validateAccounts(analysis.accountIds);
            if (accounts.invalid.length) {
                analysis.errors.push(accounts.invalid.length +
                    ' account internal ID(s) do not exist in NetSuite.');
                analysis.ok = false;
            }
        }
        if (!analysis.ok) {
            addHtml(form, 'err', styles() + panel('error', 'The file no longer validates',
                '<ul><li>' + analysis.errors.map(esc).join('</li><li>') + '</li></ul>' +
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }

        // No override on this one. A job already in the queue for this period
        // means waiting, not posting a second entry.
        let inFlight;
        try {
            inFlight = analysis.memo ? findInFlightJob(analysis.memo, true) : null;
        } catch (e) {
            addHtml(form, 'err', styles() + panel('error', 'Not posted',
                '<p>NetSuite could not be checked for a job already in progress for ' +
                esc(analysis.memo) + ', so nothing was posted. Try again.</p>' +
                '<p>' + esc(e.message || String(e)) + '</p>' +
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }
        if (inFlight) {
            addHtml(form, 'err', styles() + renderInFlight(inFlight));
            return res.writePage(form);
        }

        // Strict here, unlike the preview: if the check cannot run we refuse
        // rather than risk a second entry for a month already posted.
        let duplicates;
        try {
            duplicates = analysis.memo ? lib.findExistingJes(analysis.memo) : [];
        } catch (e) {
            log.error({ title: 'Duplicate check failed at submit', details: e });
            addHtml(form, 'err', styles() + panel('error', 'Not posted',
                '<p>NetSuite could not be checked for an existing entry for ' +
                esc(analysis.memo) + ', so nothing was posted. Try again.</p>' +
                '<p>' + esc(e.message || String(e)) + '</p>' +
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }

        if (duplicates.length && req.parameters.custpage_confirm_dup !== 'T') {
            addHtml(form, 'err', styles() + renderDuplicates(duplicates) +
                panel('error', 'Not posted',
                    '<p>Tick the confirmation box on the preview screen if you really want a ' +
                    'second journal entry for ' + esc(analysis.memo) + '.</p>' +
                    '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }

        const user = runtime.getCurrentUser();
        const job = {
            status: 'PENDING',
            csvFileId: csvFileId,
            csvName: req.parameters.custpage_csvname || '',
            month: analysis.month,
            year: analysis.year,
            tranDate: analysis.tranDate,
            memo: analysis.memo,
            lineCount: analysis.lines.length,
            totalDebit: analysis.totalDebit,
            totalCredit: analysis.totalCredit,
            submittedBy: user.name,
            submittedById: user.id,
            submittedAt: new Date().toISOString(),
            confirmedDuplicate: duplicates.length > 0,
            preExistingJeIds: duplicates.map((d) => String(d.id)),
            taskId: '',
            jeId: null,
            jeTranId: '',
            error: ''
        };

        const jobFileId = writeJob(null, job, folderOf(csvFileId));

        let taskId;
        try {
            taskId = startTask(jobFileId, SCHED_DEPLOY_ID);
        } catch (e) {
            log.audit({ title: 'Named deployment busy, letting NetSuite pick one', details: e.message });
            try {
                taskId = startTask(jobFileId, null);
            } catch (e2) {
                job.status = 'ERROR';
                job.error = 'Could not start the posting script: ' + (e2.message || String(e2));
                writeJob(jobFileId, job);
                addHtml(form, 'err', styles() + panel('error', 'Could not start the posting script',
                    '<p>' + esc(job.error) + '</p>' +
                    '<p>A previous run may still be in progress. Check ' +
                    '<b>Customization &gt; Scripting &gt; Script Deployments</b>, then try again.</p>'));
                return res.writePage(form);
            }
        }

        job.taskId = taskId;
        writeJob(jobFileId, job);
        log.audit({ title: 'Statscore JE job submitted', details: 'job=' + jobFileId + ' task=' + taskId });

        res.sendRedirect({
            type: 'SUITELET',
            identifier: runtime.getCurrentScript().id,
            id: runtime.getCurrentScript().deploymentId,
            parameters: { action: 'status', job: String(jobFileId) }
        });
    }

    function startTask(jobFileId, deploymentId) {
        const params = {};
        params[JOB_PARAM] = jobFileId;
        const opts = {
            taskType: task.TaskType.SCHEDULED_SCRIPT,
            scriptId: SCHED_SCRIPT_ID,
            params: params
        };
        if (deploymentId) {
            opts.deploymentId = deploymentId;
        }
        return task.create(opts).submit();
    }

    // -----------------------------------------------------------------------
    // Page 4: status
    // -----------------------------------------------------------------------

    function renderStatus(res, jobFileId) {
        const form = serverWidget.createForm({ title: 'Statscore JE Creator - Status' });
        let html = styles();

        let job;
        try {
            job = JSON.parse(file.load({ id: jobFileId }).getContents());
        } catch (e) {
            addHtml(form, 'err', html + panel('error', 'Job not found',
                '<p><a href="' + esc(selfUrl()) + '">Start again</a></p>'));
            return res.writePage(form);
        }

        const running = job.status === 'PENDING' || job.status === 'RUNNING';

        // The scheduled script cannot record a hard failure (a timeout, say),
        // so cross-check the queue before claiming the job is still alive.
        let taskState = '';
        if (running && job.taskId) {
            try {
                taskState = task.checkStatus({ taskId: job.taskId }).status;
            } catch (e) {
                taskState = '';
            }
        }
        const stalled = running && taskState === 'FAILED';

        let rows = '';
        rows += kv('Status', statusBadge(stalled ? 'ERROR' : job.status, taskState));
        rows += kv('Source file', esc(job.csvName || ''));
        rows += kv('Period', pad2(job.month) + '/' + job.year);
        rows += kv('Transaction date', esc(job.tranDate));
        rows += kv('Memo', esc(job.memo));
        rows += kv('Lines', String(job.lineCount));
        rows += kv('Total debit', lib.fmt(job.totalDebit));
        rows += kv('Total credit', lib.fmt(job.totalCredit));
        rows += kv('Submitted by', esc(job.submittedBy || ''));
        rows += kv('Submitted at', esc(job.submittedAt || '') + (job.submittedAt
            ? ' <span class="sc-muted">(' + esc(since(job.submittedAt)) + ')</span>' : ''));
        if (job.confirmedDuplicate) {
            rows += kv('Duplicate period', '<span class="sc-bad">Confirmed and posted anyway</span>');
        }
        if (job.startedAt) rows += kv('Started at', esc(job.startedAt));
        if (job.finishedAt) rows += kv('Finished at', esc(job.finishedAt));
        html += '<table class="sc-kv">' + rows + '</table>';

        if (job.status === 'DONE' && job.jeId) {
            const link = url.resolveRecord({
                recordType: 'journalentry',
                recordId: job.jeId,
                isEditMode: false
            });
            html += panel('ok', 'Journal entry created',
                '<p><b>' + esc(job.jeTranId || ('Internal ID ' + job.jeId)) + '</b></p>' +
                '<p><a href="' + esc(link) + '" target="_blank">Open the journal entry</a></p>');
        } else if (job.status === 'ERROR') {
            html += panel('error', 'Posting failed',
                '<p>' + esc(job.error || 'No detail recorded.') + '</p>' +
                '<p>The full stack trace is in <b>Customization &gt; Scripting &gt; Script Execution Log</b>.</p>');
        } else if (stalled) {
            html += panel('error', 'The posting script stopped without finishing',
                '<p>Check the script execution log for the reason, then upload the file again.</p>');
        } else {
            // Second opinion. The job file is the primary status, but if a
            // write to it is ever lost the page would sit on "Queued" over a
            // finished job, which is how a month got posted twice.
            const appeared = newlyPosted(job);
            if (appeared.length) {
                html += panel('error', 'This period now has a journal entry',
                    '<p>The job file still reads <b>' + esc(job.status) + '</b>, but an entry for ' +
                    '<b>' + esc(job.memo) + '</b> exists that was not there when this job was ' +
                    'submitted:</p><ul>' +
                    appeared.map((d) => '<li><a href="' + esc(jeUrl(d.id)) + '" target="_blank">' +
                        esc(d.tranid || ('Internal ID ' + d.id)) + '</a> dated ' +
                        esc(d.trandate) + '</li>').join('') +
                    '</ul><p>That is almost certainly this job, with its status update lost. ' +
                    'Open the entry and check it. Do not post the file again.</p>');
            }

            const queued = job.status === 'PENDING';
            html += panel('warn',
                queued ? 'Waiting in the NetSuite script queue' : 'Building the journal entry',
                (queued
                    ? '<p>Submitted ' + esc(since(job.submittedAt)) + '. The job has not started yet.</p>' +
                      '<p>NetSuite runs scheduled scripts from a shared queue. When that queue is ' +
                      'busy a job can sit here for <b>30 to 60 minutes</b> before it starts. ' +
                      'That is normal and does not mean anything has failed.</p>'
                    : '<p>Building ' + job.lineCount + ' lines, started ' +
                      esc(since(job.startedAt)) + '. Roughly 90 seconds per 4,000 lines.</p>') +
                '<p class="sc-bad">Do not upload and post this file again while this page is open. ' +
                'A second submission creates a second journal entry.</p>' +
                '<p>This page refreshes every 10 seconds. It is safe to close and come back to.</p>');
            html += '<script>setTimeout(function(){location.reload();},10000);</script>';
        }

        html += '<p style="margin-top:14px">' +
            '<a href="' + esc(statusUrl(jobFileId)) + '">Refresh now</a> &nbsp;|&nbsp; ' +
            '<a href="' + esc(selfUrl()) + '">Upload another file</a></p>';

        addHtml(form, 'body', html);
        res.writePage(form);
    }

    function statusBadge(status, taskState) {
        const map = {
            PENDING: ['sc-wait', 'Queued'],
            RUNNING: ['sc-wait', 'Building the journal entry'],
            DONE: ['sc-ok', 'Done'],
            ERROR: ['sc-bad', 'Failed']
        };
        const m = map[status] || ['sc-muted', status];
        let text = m[1];
        if (status === 'PENDING' && taskState === 'PROCESSING') {
            text = 'Starting';
        }
        return '<span class="' + m[0] + '">' + esc(text) + '</span>';
    }

    // -----------------------------------------------------------------------
    // File Cabinet helpers
    // -----------------------------------------------------------------------

    /**
     * Read an uploaded CSV as text.
     *
     * UTF-8 is tried first because invalid UTF-8 is detectable (it decodes to
     * U+FFFD), whereas a UTF-8 file read as Windows-1252 is mangled silently.
     * The Python tool tried Windows-1252 first; the outcome is the same for
     * both file kinds, just reached from the safer direction.
     */
    function readText(f) {
        // UTF-8 first, then Windows-1252, then whatever the file claims.
        // Invalid UTF-8 is detectable (it decodes to U+FFFD); a UTF-8 file read
        // as Windows-1252 is mangled silently, so this is the safer order even
        // though the Python tool tried Windows-1252 first.
        const attempts = [file.Encoding.UTF_8, file.Encoding.WINDOWS_1252, null];
        let fallback = '';

        for (let i = 0; i < attempts.length; i++) {
            let text = '';
            try {
                if (attempts[i]) {
                    f.encoding = attempts[i];
                }
                text = f.getContents();
            } catch (e) {
                log.error({ title: 'Could not read the upload as ' + (attempts[i] || 'default'), details: e });
                continue;
            }
            if (text && text.indexOf('\uFFFD') === -1) {
                return text;
            }
            if (text && !fallback) {
                fallback = text;
            }
        }

        return fallback;
    }

    function stageCsv(name, contents) {
        const stamp = new Date().getTime();
        const f = file.create({
            name: 'statscore_je_' + stamp + '.csv',
            fileType: file.Type.CSV,
            contents: contents,
            encoding: file.Encoding.UTF_8,
            folder: workFolder(),
            description: 'Statscore JE source: ' + name
        });
        return f.save();
    }

    function writeJob(jobFileId, job, folderId) {
        if (jobFileId) {
            return lib.writeJson(jobFileId, job);
        }
        const f = file.create({
            name: 'statscore_je_' + new Date().getTime() + '.job.json',
            fileType: file.Type.JSON,
            contents: JSON.stringify(job, null, 2),
            encoding: file.Encoding.UTF_8,
            folder: folderId || workFolder()
        });
        return f.save();
    }

    function folderOf(fileId) {
        try {
            return file.load({ id: fileId }).folder;
        } catch (e) {
            return workFolder();
        }
    }

    /** Find (or create once) the working folder in the File Cabinet. */
    function workFolder() {
        try {
            const rows = query.runSuiteQL({
                query: 'SELECT id FROM mediaitemfolder WHERE name = ?',
                params: [WORK_FOLDER_NAME]
            }).asMappedResults();
            if (rows.length) {
                return Number(rows[0].id);
            }
            const folder = record.create({ type: record.Type.FOLDER });
            folder.setValue({ fieldId: 'name', value: WORK_FOLDER_NAME });
            return folder.save();
        } catch (e) {
            log.error({ title: 'Falling back to the SuiteScripts folder', details: e });
            return FALLBACK_FOLDER;
        }
    }

    // -----------------------------------------------------------------------
    // Rendering helpers
    // -----------------------------------------------------------------------

    function selfUrl() {
        return url.resolveScript({
            scriptId: runtime.getCurrentScript().id,
            deploymentId: runtime.getCurrentScript().deploymentId
        });
    }

    function statusUrl(jobFileId) {
        return url.resolveScript({
            scriptId: runtime.getCurrentScript().id,
            deploymentId: runtime.getCurrentScript().deploymentId,
            params: { action: 'status', job: String(jobFileId) }
        });
    }

    function addHtml(form, id, html) {
        form.addField({
            id: 'custpage_html_' + id,
            type: serverWidget.FieldType.INLINEHTML,
            label: ' '
        }).defaultValue = html;
    }

    function hidden(form, id, value) {
        const f = form.addField({ id: id, type: serverWidget.FieldType.TEXT, label: id });
        f.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });
        f.defaultValue = value;
    }

    function panel(kind, title, body) {
        return '<div class="sc-panel sc-panel-' + kind + '"><div class="sc-panel-t">' +
            esc(title) + '</div>' + body + '</div>';
    }

    function section(title, body) {
        return '<div class="sc-section"><div class="sc-section-t">' + esc(title) + '</div>' +
            body + '</div>';
    }

    function kv(label, value) {
        return '<tr><th>' + esc(label) + '</th><td>' + value + '</td></tr>';
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

    function styles() {
        return '<style>' +
            '.sc-kv{border-collapse:collapse;margin:10px 0 16px 0;font-size:12px}' +
            '.sc-kv th{text-align:left;padding:3px 18px 3px 0;color:#555;font-weight:normal;white-space:nowrap}' +
            '.sc-kv td{padding:3px 0;font-weight:bold}' +
            '.sc-grid{border-collapse:collapse;font-size:12px;margin:6px 0}' +
            '.sc-grid th{text-align:left;background:#f0f2f5;padding:4px 10px;border:1px solid #d8dce2;white-space:nowrap}' +
            '.sc-grid td{padding:3px 10px;border:1px solid #e4e7eb}' +
            '.sc-grid .sc-num,.sc-num{text-align:right;font-family:Consolas,monospace}' +
            '.sc-section{margin:16px 0}' +
            '.sc-section-t{font-weight:bold;font-size:13px;margin-bottom:4px}' +
            '.sc-scroll{max-height:340px;overflow:auto;border:1px solid #e4e7eb}' +
            '.sc-panel{margin:10px 0;padding:9px 12px;border-left:4px solid #999;background:#f7f8fa;font-size:12px}' +
            '.sc-panel-t{font-weight:bold;margin-bottom:3px}' +
            '.sc-panel-info{border-left-color:#2980b9}' +
            '.sc-panel-ok{border-left-color:#1e8449;background:#eefaf1}' +
            '.sc-panel-warn{border-left-color:#c87f0a;background:#fdf6e8}' +
            '.sc-panel-error{border-left-color:#c0392b;background:#fdeeec}' +
            '.sc-panel ul{margin:4px 0 0 18px;padding:0}' +
            '.sc-ok{color:#1e8449;font-weight:bold}' +
            '.sc-bad{color:#c0392b;font-weight:bold}' +
            '.sc-wait{color:#c87f0a;font-weight:bold}' +
            '.sc-muted{color:#777;font-weight:normal}' +
            '</style>';
    }

    return { onRequest: onRequest };
});
