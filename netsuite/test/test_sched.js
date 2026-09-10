// Notification tests for statscore_je_sched.js.
//
//     node netsuite/test/test_sched.js
//
// execute() is driven against stubbed N/ modules so the email that reports a
// run's outcome is asserted on directly: who it goes to, what it says, and
// that a mail failure can never change what was posted.

const fs = require('fs');
const path = require('path');

const DIR = path.resolve(__dirname, '..');

function loadModule(file, deps) {
    let factory = null;
    global.define = (d, f) => { factory = f; };
    eval(fs.readFileSync(path.join(DIR, file), 'utf8'));
    return factory.apply(null, deps);
}

// ---------------------------------------------------------------------------
// Stubs
// ---------------------------------------------------------------------------

const CSV = [
    'Internal ,Account,Debit,Credit,LineMemo,REF,Date,Class,Location (Expenses),Currecny',
    '119,120003,1000.00,,Invoice A,R1,15/08/26,,Other,EUR',
    '213,121008,,1000.00,Fee B,R1,15/08/26,,Other,EUR'
].join('\n');

let FILES, SENT, SAVED_JES, PARAMS, EMAIL_THROWS, SAVE_THROWS;

function baseJob(over) {
    return Object.assign({
        status: 'PENDING', csvFileId: 500, csvName: '0 CSV.csv',
        month: 8, year: 2026, tranDate: '2026-08-31', memo: 'Statscore 08/2026',
        lineCount: 2, totalDebit: 1000, totalCredit: 1000,
        submittedBy: 'Matan Lazar', submittedById: 10,
        submittedAt: '2026-09-10T08:00:00.000Z',
        taskId: 'TASK_1', jeId: null, jeTranId: '', error: '', confirmedDuplicate: false
    }, over || {});
}

function reset(job, copyParam) {
    FILES = { 500: { contents: CSV }, 600: { contents: JSON.stringify(job || baseJob()) } };
    SENT = [];
    SAVED_JES = [];
    EMAIL_THROWS = false;
    SAVE_THROWS = false;
    PARAMS = {
        custscript_statscore_je_job: 600,
        custscript_statscore_je_copyto: copyParam === undefined ? '10, 42' : copyParam
    };
}

const runtime = {
    getCurrentScript: () => ({
        getParameter: ({ name }) => PARAMS[name],
        getRemainingUsage: () => 9870
    }),
    getCurrentUser: () => ({ id: 99, name: 'Script Owner' })
};

const fileMod = {
    load({ id }) {
        const f = FILES[Number(id)];
        if (!f) throw new Error('That file does not exist. id=' + id);
        return {
            getContents: () => f.contents,
            get contents() { return f.contents; },
            set contents(v) { f.contents = v; },
            save: () => Number(id)
        };
    }
};

const record = {
    Type: { JOURNAL_ENTRY: 'journalentry' },
    create() {
        const je = { body: {}, lines: [] };
        return {
            setValue({ fieldId, value }) { je.body[fieldId] = value; },
            insertLine({ line }) { je.lines[line] = {}; },
            setSublistValue({ fieldId, line, value }) { je.lines[line][fieldId] = value; },
            save() {
                if (SAVE_THROWS) throw new Error('INSUFFICIENT_PERMISSION: no JE permission');
                SAVED_JES.push(je);
                return 332926;
            }
        };
    }
};

const query = {
    runSuiteQL({ query: q }) {
        const rows = (() => {
            if (/FROM account/.test(q)) {
                const ids = (q.match(/IN \(([^)]*)\)/) || [, ''])[1].split(',').filter(Boolean);
                return ids.map(id => ({
                    id: id.trim(), acctnumber: 'A' + id.trim(),
                    fullname: 'Account ' + id.trim(), isinactive: 'F'
                }));
            }
            if (/SELECT tranid/.test(q)) return [{ tranid: 'JE10529' }];
            return [];
        })();
        return { asMappedResults: () => rows };
    }
};

const email = {
    send(o) {
        if (EMAIL_THROWS) throw new Error('SSS_EMAIL_SEND_FAILED');
        SENT.push(o);
    }
};

const urlMod = {
    HostType: { APPLICATION: 'APPLICATION' },
    resolveDomain: () => '11069058.app.netsuite.com',
    resolveRecord: (o) => '/app/accounting/transactions/journal.nl?id=' + o.recordId
};

const log = { error() {}, audit() {}, debug() {} };

const lib = loadModule('statscore_je_lib.js', [query]);
const sched = loadModule('statscore_je_sched.js',
    [runtime, record, fileMod, query, email, urlMod, log, lib]);

// ---------------------------------------------------------------------------

let pass = 0, fail = 0;
function check(name, cond, detail) {
    if (cond) { pass++; console.log('  ok   ' + name); }
    else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}
const job = () => JSON.parse(FILES[600].contents);

// ---------------------------------------------------------------------------
console.log('\n1. Success emails the submitter and copies the list');
reset();
sched.execute();
check('journal entry created', SAVED_JES.length === 1);
check('job marked DONE', job().status === 'DONE', 'status: ' + job().status);
check('one email sent', SENT.length === 1, 'sent: ' + SENT.length);
let m = SENT[0] || {};
check('subject names the period and the JE',
    m.subject === 'Statscore JE posted: 08/2026 - JE10529', 'subject: ' + m.subject);
check('addressed to the submitter', JSON.stringify(m.recipients) === '[10]',
    'recipients: ' + JSON.stringify(m.recipients));
check('submitter not also in cc', JSON.stringify(m.cc) === '[42]',
    'cc: ' + JSON.stringify(m.cc));
check('author is the submitter', m.author === 10);
check('body carries an absolute JE link',
    /https:\/\/11069058\.app\.netsuite\.com\/app\/accounting\/transactions\/journal\.nl\?id=332926/
        .test(m.body || ''));
check('body reports the totals', /1,000\.00/.test(m.body || ''));
check('body names the source file', /0 CSV\.csv/.test(m.body || ''));
check('no style block, inline CSS only', !/<style/i.test(m.body || ''));

// ---------------------------------------------------------------------------
console.log('\n2. Failure emails the error');
reset();
SAVE_THROWS = true;
sched.execute();
check('nothing posted', SAVED_JES.length === 0);
check('job marked ERROR', job().status === 'ERROR', 'status: ' + job().status);
check('one email sent', SENT.length === 1);
m = SENT[0] || {};
check('subject flags the failure', m.subject === 'Statscore JE FAILED: 08/2026',
    'subject: ' + m.subject);
check('body carries the error text', /no JE permission/.test(m.body || ''));
check('body points at the execution log', /Script Execution Log/.test(m.body || ''));
check('body has no JE link', !/journal\.nl/.test(m.body || ''));

// ---------------------------------------------------------------------------
console.log('\n3. Copy list handling');
reset(baseJob(), '');
sched.execute();
check('empty parameter still emails the submitter',
    SENT.length === 1 && JSON.stringify(SENT[0].recipients) === '[10]');
check('no cc key when nobody is copied', SENT[0].cc === undefined);

reset(baseJob(), '10');
sched.execute();
check('submitter alone in the list is not cc\'d to themselves', SENT[0].cc === undefined,
    'cc: ' + JSON.stringify(SENT[0].cc));

reset(baseJob(), '42; finance@example.com , 7');
sched.execute();
check('semicolons, spaces and addresses all parse',
    JSON.stringify(SENT[0].cc) === '[42,"finance@example.com",7]',
    'cc: ' + JSON.stringify(SENT[0].cc));

// ---------------------------------------------------------------------------
console.log('\n4. The aborts that used to be silent now alert');
reset();
delete FILES[600];
sched.execute();
check('nothing posted', SAVED_JES.length === 0);
check('alert sent to the copy list', SENT.length === 1 &&
    JSON.stringify(SENT[0].recipients) === '[10,42]', 'sent: ' + JSON.stringify(SENT));
check('subject says it could not start', /could not start/.test(SENT[0].subject || ''));

reset();
PARAMS.custscript_statscore_je_job = null;
sched.execute();
check('no job parameter alerts', SENT.length === 1 && /no job/i.test(SENT[0].subject || ''),
    'subject: ' + (SENT[0] || {}).subject);

reset(baseJob({ status: 'RUNNING' }));
sched.execute();
check('a job not at PENDING posts nothing', SAVED_JES.length === 0);
check('and alerts', SENT.length === 1 && /skipped/.test(SENT[0].subject || ''),
    'subject: ' + (SENT[0] || {}).subject);
check('alert warns about a half-finished earlier run', /died/.test(SENT[0].body || ''));

reset(baseJob({ status: 'RUNNING' }), '');
sched.execute();
check('no copy list means log only, no crash', SENT.length === 0);

// ---------------------------------------------------------------------------
console.log('\n5. A mail failure never changes what was posted');
reset();
EMAIL_THROWS = true;
sched.execute();
check('journal entry still created', SAVED_JES.length === 1);
check('job still marked DONE', job().status === 'DONE', 'status: ' + job().status);
check('JE number still recorded', job().jeTranId === 'JE10529');
check('execute did not throw', true);

// ---------------------------------------------------------------------------
console.log('\n6. The duplicate override is carried into the email');
reset(baseJob({ confirmedDuplicate: true }));
sched.execute();
check('override shown in the body', /Confirmed and posted anyway/.test(SENT[0].body || ''));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
