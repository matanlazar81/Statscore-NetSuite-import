// Guard tests for statscore_je_suitelet.js.
//
//     node netsuite/test/test_suitelet.js
//
// The Suitelet is driven through onRequest() against stubbed N/ modules, so
// the guards that stop a period being posted twice are exercised for real
// rather than reasoned about. Added after a duplicate August entry reached
// the ledger: the posted-entry check alone was not enough, because NetSuite's
// scheduled script queue can hold a job for the best part of an hour and the
// status page only said "Queued".

const fs = require('fs');
const path = require('path');

const DIR = path.resolve(__dirname, '..');

// ---------------------------------------------------------------------------
// Stubs
// ---------------------------------------------------------------------------

function loadModule(file, deps) {
    let factory = null;
    global.define = (d, f) => { factory = f; };
    eval(fs.readFileSync(path.join(DIR, file), 'utf8'));
    return factory.apply(null, deps);
}

/** Captures what the Suitelet builds so assertions can inspect it. */
function makeForm() {
    return {
        html: '',
        fields: [],
        submitButtons: [],
        addField(o) {
            this.fields.push(o);
            const self = this;
            const f = {
                _id: o.id,
                updateDisplayType() { return f; },
                set defaultValue(v) { self.fields.find(x => x.id === o.id).value = v; },
                get defaultValue() { return ''; },
                set isMandatory(v) {},
                get isMandatory() { return false; }
            };
            if (o.type === 'INLINEHTML') {
                Object.defineProperty(f, 'defaultValue', {
                    set(v) { self.html += v; },
                    get() { return ''; }
                });
            }
            return f;
        },
        addSubmitButton(o) { this.submitButtons.push(o.label); }
    };
}

let FORMS = [];
const serverWidget = {
    FieldType: { FILE: 'FILE', TEXT: 'TEXT', CHECKBOX: 'CHECKBOX', INLINEHTML: 'INLINEHTML' },
    FieldDisplayType: { HIDDEN: 'HIDDEN' },
    createForm(o) { const f = makeForm(); f.title = o.title; FORMS.push(f); return f; }
};

// Files the fake File Cabinet holds, keyed by id.
let FILES = {};
let NEXT_FILE_ID = 900;
const fileMod = {
    Type: { CSV: 'CSV', JSON: 'JSON' },
    Encoding: { UTF_8: 'UTF-8', WINDOWS_1252: 'windows-1252' },
    load({ id }) {
        const f = FILES[Number(id)];
        if (!f) throw new Error('That file does not exist. id=' + id);
        return {
            folder: f.folder,
            name: f.name,
            getContents: () => f.contents,
            set contents(v) { f.contents = v; },
            get contents() { return f.contents; },
            save: () => Number(id)
        };
    },
    create(o) {
        const id = ++NEXT_FILE_ID;
        return { save() { FILES[id] = { name: o.name, contents: o.contents, folder: o.folder }; return id; } };
    }
};

let POSTED_JES = [];        // rows findExistingJes should return
let SUBMITTED_TASKS = [];

const query = {
    runSuiteQL({ query: q, params }) {
        const rows = (() => {
            if (/mediaitemfolder/.test(q)) return [{ id: 199462 }];
            if (/FROM account/.test(q)) {
                const ids = (q.match(/IN \(([^)]*)\)/) || [, ''])[1].split(',').filter(Boolean);
                return ids.map(id => ({
                    id: id.trim(), acctnumber: '1200' + id.trim(),
                    fullname: 'Account ' + id.trim(), isinactive: 'F'
                }));
            }
            if (/FROM transaction/.test(q)) return POSTED_JES;
            if (/FROM file/.test(q)) {
                return Object.keys(FILES)
                    .filter(id => FILES[id].folder === params[0] && /\.job\.json$/.test(FILES[id].name))
                    .map(Number).sort((a, b) => b - a).map(id => ({ id: id }));
            }
            return [];
        })();
        return { asMappedResults: () => rows };
    }
};

const record = { Type: { FOLDER: 'folder' }, create: () => ({ setValue() {}, save: () => 199462 }) };
const task = {
    TaskType: { SCHEDULED_SCRIPT: 'SCHEDULED_SCRIPT' },
    create(o) { return { submit() { SUBMITTED_TASKS.push(o); return 'TASK_1'; } }; },
    checkStatus: () => ({ status: 'PENDING' })
};
const runtime = {
    getCurrentScript: () => ({ id: 'customscript_x', deploymentId: 'customdeploy_x' }),
    getCurrentUser: () => ({ id: 10, name: 'Matan Lazar' })
};
const urlMod = {
    resolveScript: (o) => '/suitelet?job=' + ((o.params && o.params.job) || ''),
    resolveRecord: (o) => '/je?id=' + o.recordId
};
const log = { error() {}, audit() {}, debug() {} };

const lib = loadModule('statscore_je_lib.js', [query]);
const suitelet = loadModule('statscore_je_suitelet.js',
    [serverWidget, fileMod, query, record, task, runtime, urlMod, log, lib]);

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const CSV = [
    'Internal ,Account,Debit,Credit,LineMemo,REF,Date,Class,Location (Expenses),Currecny',
    '119,120003,1000.00,,Invoice A,R1,15/08/26,,Other,EUR',
    '213,121008,,1000.00,Fee B,R1,15/08/26,,Other,EUR'
].join('\n');

const MEMO = 'Statscore 08/2026';

function uploadedFile() {
    return { name: '0 CSV.csv', set encoding(v) {}, get encoding() { return ''; }, getContents: () => CSV };
}

function run(req) {
    FORMS = [];
    let redirect = null;
    const res = { writePage() {}, sendRedirect(o) { redirect = o; } };
    suitelet.onRequest({ request: req, response: res });
    return { form: FORMS[FORMS.length - 1], redirect: redirect };
}

function reset() {
    FILES = {}; POSTED_JES = []; SUBMITTED_TASKS = []; NEXT_FILE_ID = 900;
}

function addJobFile(status, memo, submittedAt) {
    const id = ++NEXT_FILE_ID;
    FILES[id] = {
        name: 'statscore_je_' + id + '.job.json',
        folder: 199462,
        contents: JSON.stringify({
            status: status, memo: memo, month: 8, year: 2026,
            tranDate: '2026-08-31', csvName: '0 CSV.csv', lineCount: 3728,
            totalDebit: 5484042.48, totalCredit: 5484042.48,
            submittedBy: 'Matan Lazar',
            submittedAt: submittedAt || new Date(Date.now() - 39 * 60000).toISOString(),
            taskId: 'TASK_0', jeId: null, jeTranId: '', error: ''
        })
    };
    return id;
}

// ---------------------------------------------------------------------------
// Assertions
// ---------------------------------------------------------------------------

let pass = 0, fail = 0;
function check(name, cond, detail) {
    if (cond) { pass++; console.log('  ok   ' + name); }
    else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}

// ---------------------------------------------------------------------------
console.log('\n1. Clean preview offers the Create button');
reset();
let r = run({ method: 'POST', parameters: {}, files: { custpage_csv: uploadedFile() } });
check('Create button present', r.form.submitButtons.indexOf('Create Journal Entry') !== -1,
    'buttons: ' + JSON.stringify(r.form.submitButtons));
check('no in-flight panel', !/already being posted/.test(r.form.html));
check('period detected', /Statscore 08\/2026/.test(r.form.html));

// ---------------------------------------------------------------------------
console.log('\n2. A QUEUED job for the same period blocks the preview');
reset();
addJobFile('PENDING', MEMO);
r = run({ method: 'POST', parameters: {}, files: { custpage_csv: uploadedFile() } });
check('no Create button', r.form.submitButtons.indexOf('Create Journal Entry') === -1,
    'buttons: ' + JSON.stringify(r.form.submitButtons));
check('in-flight panel shown', /already being posted/.test(r.form.html));
check('says it is queued', /is queued and has not started yet/.test(r.form.html));
check('offers a link to that job', /Watch that job/.test(r.form.html));
check('explains the escape hatch', /delete/i.test(r.form.html) && /job\.json/.test(r.form.html));

// ---------------------------------------------------------------------------
console.log('\n3. A RUNNING job blocks too');
reset();
addJobFile('RUNNING', MEMO);
r = run({ method: 'POST', parameters: {}, files: { custpage_csv: uploadedFile() } });
check('no Create button', r.form.submitButtons.indexOf('Create Journal Entry') === -1);
check('says it is posting now', /is being posted right now/.test(r.form.html));

// ---------------------------------------------------------------------------
console.log('\n4. Finished jobs never block');
['DONE', 'ERROR'].forEach((st) => {
    reset();
    addJobFile(st, MEMO);
    r = run({ method: 'POST', parameters: {}, files: { custpage_csv: uploadedFile() } });
    check(st + ' does not block', r.form.submitButtons.indexOf('Create Journal Entry') !== -1);
});

console.log('\n4b. A job for a different period never blocks');
reset();
addJobFile('PENDING', 'Statscore 07/2026');
r = run({ method: 'POST', parameters: {}, files: { custpage_csv: uploadedFile() } });
check('different period does not block', r.form.submitButtons.indexOf('Create Journal Entry') !== -1);

// ---------------------------------------------------------------------------
console.log('\n5. Submit refuses while a job is in flight - the incident case');
reset();
const csvId = ++NEXT_FILE_ID;
FILES[csvId] = { name: 'staged.csv', folder: 199462, contents: CSV };
addJobFile('PENDING', MEMO);
r = run({
    method: 'POST',
    parameters: {
        custpage_action: 'create',
        custpage_csvfile: String(csvId),
        custpage_csvname: '0 CSV.csv',
        custpage_confirm_dup: 'T'          // even with the override ticked
    },
    files: {}
});
check('no task submitted', SUBMITTED_TASKS.length === 0,
    'tasks: ' + SUBMITTED_TASKS.length);
check('refusal shown', /already being posted/.test(r.form.html));
check('no redirect to a new job', r.redirect === null);

console.log('\n5b. Submit works when nothing is in flight');
reset();
const csvId2 = ++NEXT_FILE_ID;
FILES[csvId2] = { name: 'staged.csv', folder: 199462, contents: CSV };
r = run({
    method: 'POST',
    parameters: {
        custpage_action: 'create',
        custpage_csvfile: String(csvId2),
        custpage_csvname: '0 CSV.csv'
    },
    files: {}
});
check('task submitted', SUBMITTED_TASKS.length === 1);
check('job file written as PENDING', Object.keys(FILES).some(
    id => /\.job\.json$/.test(FILES[id].name) && JSON.parse(FILES[id].contents).status === 'PENDING'));
check('redirected to the status page', r.redirect && r.redirect.parameters.action === 'status');

console.log('\n5c. A posted duplicate still needs the tick');
reset();
const csvId3 = ++NEXT_FILE_ID;
FILES[csvId3] = { name: 'staged.csv', folder: 199462, contents: CSV };
POSTED_JES = [{ id: 332798, tranid: 'JE10523', trandate: '31/08/2026' }];
r = run({
    method: 'POST',
    parameters: {
        custpage_action: 'create',
        custpage_csvfile: String(csvId3),
        custpage_csvname: '0 CSV.csv'
    },
    files: {}
});
check('unticked submit refused', SUBMITTED_TASKS.length === 0);
check('refusal names the period', /Statscore 08\/2026/.test(r.form.html));

// ---------------------------------------------------------------------------
console.log('\n6. Status page tells the truth about the queue');
reset();
const jobId = addJobFile('PENDING', MEMO, new Date(Date.now() - 39 * 60000).toISOString());
r = run({ method: 'GET', parameters: { action: 'status', job: String(jobId) }, files: {} });
check('says it is waiting in the queue', /Waiting in the NetSuite script queue/.test(r.form.html));
check('shows how long', /39 minutes ago/.test(r.form.html));
check('warns against resubmitting', /Do not upload and post this file again/.test(r.form.html));
check('sets the 30-60 minute expectation', /30 to 60 minutes/.test(r.form.html));

reset();
const jobId2 = addJobFile('RUNNING', MEMO);
FILES[jobId2].contents = JSON.stringify(
    Object.assign(JSON.parse(FILES[jobId2].contents),
        { startedAt: new Date(Date.now() - 2 * 60000).toISOString() }));
r = run({ method: 'GET', parameters: { action: 'status', job: String(jobId2) }, files: {} });
check('RUNNING says building', /Building the journal entry/.test(r.form.html));
check('RUNNING shows start time', /2 minutes ago/.test(r.form.html));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
