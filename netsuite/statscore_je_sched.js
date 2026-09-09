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
    'N/log',
    './statscore_je_lib'
], (runtime, record, file, query, log, lib) => {

    const JOB_PARAM = 'custscript_statscore_je_job';
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
            return;
        }

        let job;
        try {
            job = JSON.parse(file.load({ id: jobFileId }).getContents());
        } catch (e) {
            log.error({ title: 'Could not read job file ' + jobFileId, details: e });
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

        } catch (e) {
            log.error({ title: 'Statscore JE posting failed', details: e });
            job.status = 'ERROR';
            job.error = (e.name ? e.name + ': ' : '') + (e.message || String(e));
            job.finishedAt = new Date().toISOString();
            save(jobFileId, job);
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

    function save(jobFileId, job) {
        try {
            const f = file.load({ id: jobFileId });
            f.contents = JSON.stringify(job, null, 2);
            f.save();
        } catch (e) {
            log.error({ title: 'Could not update job file ' + jobFileId, details: e });
        }
    }

    return { execute: execute };
});
