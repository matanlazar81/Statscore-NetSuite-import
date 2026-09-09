// Parity tests for statscore_je_lib.js.
//
// The library is a SuiteScript AMD module, but analyze() is pure JavaScript
// with no NetSuite dependency, so it can be exercised with plain node:
//
//     node netsuite/test/test_lib.js
//
// Every expectation here is the behaviour of the original statscore_je_creator.py,
// except the four deliberate differences listed in netsuite/README.md.
const fs = require('fs');
const path = require('path');

const LIB = path.resolve(__dirname, '..', 'statscore_je_lib.js');

let factory = null;
global.define = (deps, f) => { factory = f; };
eval(fs.readFileSync(LIB, 'utf8'));

// N/query stub: analyze() never touches it; validateAccounts/findExistingJes do.
const queryStub = { runSuiteQL: () => { throw new Error('query should not be called by analyze()'); } };
const lib = factory(queryStub);

let pass = 0, fail = 0;
function check(name, actual, expected) {
    const a = JSON.stringify(actual), e = JSON.stringify(expected);
    if (a === e) { pass++; console.log('  ok   ' + name); }
    else { fail++; console.log('  FAIL ' + name + '\n         expected ' + e + '\n         actual   ' + a); }
}

const HEADER = 'Internal ,Account,Debit,Credit,LineMemo,REF,Date,Class,Location (Expenses),Currecny';

// ---------------------------------------------------------------------------
console.log('\n1. Happy path');
let csv = [HEADER,
    '119,120003,1000.00,,Invoice A,R1,15/07/26,,Other,EUR',
    '213,121008,,1000.00,"Fees, bank",R1,15/07/26,,Other,EUR',
    '852,122001,250.50,,Capex,R2,20/07/26,,Other,EUR',
    '523,122003,,250.50,Depreciation,R2,20/07/26,,Other,EUR'
].join('\n');
let a = lib.analyze(csv);
check('ok', a.ok, true);
check('errors', a.errors, []);
check('period', [a.month, a.year], [7, 2026]);
check('tranDate (last day of month)', a.tranDate, '2026-07-31');
check('memo', a.memo, 'Statscore 07/2026');
check('line count', a.lines.length, 4);
check('totals', [a.totalDebit, a.totalCredit], [1250.5, 1250.5]);
check('REF groups balanced', [a.refGroupsOk, a.refErrors.length], [2, 0]);
check('unique accounts', a.accountIds.sort(), ['119', '213', '523', '852']);
check('quoted memo keeps its comma', a.lines[1].memo, 'Fees, bank');

// ---------------------------------------------------------------------------
console.log('\n2. Negative amounts flip side (Python parity)');
csv = [HEADER,
    '119,120003,-500,,Reversal of a debit,R1,15/07/26,,Other,EUR',
    '213,121008,,-500,Reversal of a credit,R1,15/07/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('ok', a.ok, true);
check('negative debit becomes a credit', [a.lines[0].debit, a.lines[0].credit], [0, 500]);
check('negative credit becomes a debit', [a.lines[1].debit, a.lines[1].credit], [500, 0]);
check('totals', [a.totalDebit, a.totalCredit], [500, 500]);

// ---------------------------------------------------------------------------
console.log('\n3. Zero rows skipped, blank rows dropped');
csv = [HEADER,
    '119,120003,100,,Real,R1,15/07/26,,Other,EUR',
    '213,121008,0,0,Zero line,R1,15/07/26,,Other,EUR',
    '852,122001,,,Empty amounts,R1,15/07/26,,Other,EUR',
    ',,,,,,,,,',
    '213,121008,,100,Real,R1,15/07/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('ok', a.ok, true);
check('lines', a.lines.length, 2);
check('skippedZero counts both zero and blank amounts', a.skippedZero, 2);
check('fully blank row not counted as data', a.dataRows, 4);

// ---------------------------------------------------------------------------
console.log('\n4. Unbalanced REF blocks the run');
csv = [HEADER,
    '119,120003,100,,A,R1,15/07/26,,Other,EUR',
    '213,121008,,95,B,R1,15/07/26,,Other,EUR',
    '852,122001,50,,C,R2,15/07/26,,Other,EUR',
    '523,122003,,55,D,R2,15/07/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('blocked', a.ok, false);
check('two unbalanced REFs', a.refErrors.length, 2);
check('R1 difference', lib.round2(a.refErrors[0].diff), 5);

console.log('\n4b. A difference under the 0.02 tolerance passes');
csv = [HEADER,
    '119,120003,100.00,,A,R1,15/07/26,,Other,EUR',
    '213,121008,,99.99,B,R1,15/07/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('REF tolerated', a.refErrors.length, 0);
check('but the JE total is also within tolerance', a.ok, true);

// ---------------------------------------------------------------------------
console.log('\n5. Bad data is reported by CSV line, not crashed on');
csv = [HEADER,
    '119,120003,abc,,Bad amount,R1,15/07/26,,Other,EUR',
    'XYZ,120003,100,,Bad internal id,R1,15/07/26,,Other,EUR',
    '213,121008,100,100,Both sides,R1,15/07/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('blocked', a.ok, false);
check('three row errors', a.rowErrors.length, 3);
check('CSV line numbers', a.rowErrors.map(r => r.csvLine), [2, 3, 4]);
check('both-sides row is caught', /both a debit .* and a credit/.test(a.rowErrors[2].message), true);

// ---------------------------------------------------------------------------
console.log('\n6. Header handling');
csv = ['Internal,Debit,Credit,Line Memo,Date',
    '119,10,,Memo one,15/07/26',
    '213,,10,Memo two,15/07/26'].join('\n');
a = lib.analyze(csv);
check('"Line Memo" variant accepted', a.errors.length > 0 ? a.errors : 'no column errors', 'no column errors');
check('memo read from the "Line Memo" column', a.lines[0].memo, 'Memo one');

a = lib.analyze(['Internal,Debit,Credit,Date', '119,10,,15/07/26'].join('\n'));
check('missing LineMemo is reported', /Missing required columns: memo/.test(a.errors[0]), true);

// ---------------------------------------------------------------------------
console.log('\n7. Encoding and line endings');
a = lib.analyze('﻿' + HEADER + '\r\n119,120003,10,,A,R1,15/07/26,,Other,EUR\r\n' +
    '213,121008,,10,B,R1,15/07/26,,Other,EUR\r\n');
check('BOM + CRLF parsed', [a.ok, a.lines.length], [true, 2]);

// ---------------------------------------------------------------------------
console.log('\n8. Period detection uses the modal month');
csv = [HEADER,
    '119,120003,10,,A,R1,31/05/26,,Other,EUR',
    '213,121008,,10,B,R1,15/06/26,,Other,EUR',
    '852,122001,10,,C,R2,20/06/26,,Other,EUR',
    '523,122003,,10,D,R2,25/06/26,,Other,EUR'
].join('\n');
a = lib.analyze(csv);
check('modal month wins', [a.month, a.year], [6, 2026]);
check('date is the last day of June', a.tranDate, '2026-06-30');
check('the stray May row is flagged but kept', [a.warnings.length > 0, a.lines.length], [true, 4]);

console.log('\n9. February leap-year boundary');
a = lib.analyze([HEADER,
    '119,120003,10,,A,R1,05/02/24,,Other,EUR',
    '213,121008,,10,B,R1,06/02/24,,Other,EUR'].join('\n'));
check('2024-02-29', a.tranDate, '2024-02-29');

// ---------------------------------------------------------------------------
console.log('\n10. Amount formats');
check('1,234.56 thousands separator', lib.parseAmount('1,234.56'), { ok: true, value: 1234.56 });
check('(50.00) accounting negative', lib.parseAmount('(50.00)'), { ok: true, value: -50 });
check('blank is zero', lib.parseAmount(''), { ok: true, value: 0 });
check('text is rejected', lib.parseAmount('n/a'), { ok: false, value: 0 });
check('round2 away from zero', [lib.round2(-2.345), lib.round2(2.345)], [-2.35, 2.35]);

// ---------------------------------------------------------------------------
console.log('\n11. Two-digit year pivot matches Python %y');
check('26 -> 2026', lib.parseDate('01/01/26').year, 2026);
check('68 -> 2068', lib.parseDate('01/01/68').year, 2068);
check('69 -> 1969', lib.parseDate('01/01/69').year, 1969);
check('99 -> 1999', lib.parseDate('01/01/99').year, 1999);
check('31/02 rejected', lib.parseDate('31/02/26'), null);

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
