/**
 * @NApiVersion 2.1
 * @NModuleScope SameAccount
 *
 * Statscore JE Creator - shared library.
 *
 * Holds every piece of logic that both the Suitelet (preview) and the
 * scheduled script (posting) need, so the parsing and validation rules are
 * defined exactly once.
 *
 * Ported from statscore_je_creator.py. Behaviour kept identical:
 *   - all CSV lines become ONE Journal Entry under Statscore (subsidiary 6)
 *   - lines where Debit and Credit are both zero are skipped
 *   - a negative Debit becomes a Credit, a negative Credit becomes a Debit
 *   - every REF group must balance to within 0.02 or the run is blocked
 *   - transaction date is the last day of the detected month
 */
define(['N/query'], (query) => {

    const CONFIG = {
        SUBSIDIARY: 6,            // Statscore
        CURRENCY: 1,              // EUR
        DEPARTMENT: 24,           // Statscore import
        LOCATION: 5,              // Poland
        CSEG_LOCATION_EXP: 3,     // Location (Expenses) = Other
        BALANCE_TOLERANCE: 0.02,
        ACCOUNT_BATCH: 500        // IDs per SuiteQL account lookup
    };

    // -----------------------------------------------------------------------
    // Small helpers
    // -----------------------------------------------------------------------

    /** Round to 2 decimals, away from zero, so negatives behave like positives. */
    function round2(n) {
        const sign = n < 0 ? -1 : 1;
        return sign * Math.round((Math.abs(n) + Number.EPSILON) * 100) / 100;
    }

    /** Format a number as 1,234.56 for display. */
    function fmt(n) {
        const parts = Math.abs(n).toFixed(2).split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        return (n < 0 ? '-' : '') + parts.join('.');
    }

    function lastDayOfMonth(year, month) {
        return new Date(year, month, 0).getDate();
    }

    function pad2(n) {
        return (n < 10 ? '0' : '') + n;
    }

    // -----------------------------------------------------------------------
    // CSV parsing
    // -----------------------------------------------------------------------

    /**
     * RFC 4180 CSV parser. Handles quoted fields, embedded commas, escaped
     * quotes and CR/LF line endings.
     *
     * @param {string} text
     * @returns {Array<Array<string>>} rows of raw string cells
     */
    function parseCsv(text) {
        const rows = [];
        let row = [];
        let field = '';
        let inQuotes = false;
        let touched = false;

        // Strip a UTF-8 BOM if the file carries one.
        if (text.charCodeAt(0) === 0xFEFF) {
            text = text.slice(1);
        }

        for (let i = 0; i < text.length; i++) {
            const c = text.charAt(i);

            if (inQuotes) {
                if (c === '"') {
                    if (text.charAt(i + 1) === '"') {   // escaped quote
                        field += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field += c;
                }
                continue;
            }

            if (c === '"') {
                inQuotes = true;
                touched = true;
            } else if (c === ',') {
                row.push(field);
                field = '';
                touched = true;
            } else if (c === '\n') {
                row.push(field);
                rows.push(row);
                row = [];
                field = '';
                touched = false;
            } else if (c !== '\r') {
                field += c;
                touched = true;
            }
        }

        if (touched || field !== '') {
            row.push(field);
            rows.push(row);
        }

        return rows;
    }

    /**
     * Map the header row onto the logical columns the JE needs.
     * Matching mirrors the Python version, including the tolerated
     * 'Currecny' misspelling and the trailing spaces the export produces.
     */
    function mapColumns(headerCells) {
        const colMap = {};
        headerCells.forEach((raw, idx) => {
            const cl = String(raw || '').trim().toLowerCase();
            if (cl === 'internal') colMap.internal = idx;
            else if (cl === 'account') colMap.account = idx;
            else if (cl === 'debit') colMap.debit = idx;
            else if (cl === 'credit') colMap.credit = idx;
            else if (cl === 'linememo' || cl === 'line memo' || cl === 'line_memo') colMap.memo = idx;
            else if (cl === 'ref') colMap.ref = idx;
            else if (cl === 'date') colMap.date = idx;
            else if (cl.indexOf('class') !== -1) colMap.class = idx;
            else if (cl.indexOf('location') !== -1 && cl.indexOf('expense') !== -1) colMap.loc_exp = idx;
            else if (cl === 'currecny' || cl === 'currency') colMap.currency = idx;
        });
        return colMap;
    }

    // -----------------------------------------------------------------------
    // Cell parsing
    // -----------------------------------------------------------------------

    const THOUSANDS = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;
    const PLAIN_NUMBER = /^-?(\d+(\.\d*)?|\.\d+)$/;

    /**
     * Parse an amount cell. Blank counts as zero; anything that is not a
     * number is reported so the operator sees it instead of it silently
     * becoming zero.
     *
     * @returns {{ok: boolean, value: number}}
     */
    function parseAmount(raw) {
        let s = String(raw === null || raw === undefined ? '' : raw).trim();
        s = s.replace(/^'/, '');            // Excel text-marker apostrophe
        if (s === '' || s === '-' || s.toLowerCase() === 'nan') {
            return { ok: true, value: 0 };
        }

        let negative = false;
        if (/^\(.*\)$/.test(s)) {           // (123.45) accounting negative
            negative = true;
            s = s.slice(1, -1).trim();
        }
        if (THOUSANDS.test(s)) {
            s = s.replace(/,/g, '');
        }
        if (!PLAIN_NUMBER.test(s)) {
            return { ok: false, value: 0 };
        }

        const v = parseFloat(s);
        if (isNaN(v)) {
            return { ok: false, value: 0 };
        }
        return { ok: true, value: negative ? -v : v };
    }

    /**
     * Parse the Internal (account internal ID) cell.
     * Accepts '1234' and '1234.0', matching the Python str(int(float(x))).
     *
     * @returns {{ok: boolean, blank: boolean, value: string}}
     */
    function parseInternalId(raw) {
        const s = String(raw === null || raw === undefined ? '' : raw).trim();
        if (s === '' || s.toLowerCase() === 'nan') {
            return { ok: true, blank: true, value: '' };
        }
        if (!PLAIN_NUMBER.test(s)) {
            return { ok: false, blank: false, value: s };
        }
        const v = parseFloat(s);
        if (isNaN(v)) {
            return { ok: false, blank: false, value: s };
        }
        return { ok: true, blank: false, value: String(Math.trunc(v)) };
    }

    /**
     * Parse a dd/mm/yy date. Two digit years pivot at 68, matching Python's
     * %y. Four digit years are accepted too. Returns null when unparseable,
     * which mirrors pandas errors='coerce'.
     */
    function parseDate(raw) {
        const s = String(raw === null || raw === undefined ? '' : raw).trim();
        const m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2}|\d{4})$/.exec(s);
        if (!m) return null;

        const day = parseInt(m[1], 10);
        const month = parseInt(m[2], 10);
        let year = parseInt(m[3], 10);
        if (m[3].length === 2) {
            year = year <= 68 ? 2000 + year : 1900 + year;
        }
        if (month < 1 || month > 12 || day < 1 || day > lastDayOfMonth(year, month)) {
            return null;
        }
        return { day: day, month: month, year: year };
    }

    /** Most common value; ties resolve to the smallest, like pandas mode(). */
    function modeOf(values) {
        const counts = {};
        values.forEach((v) => { counts[v] = (counts[v] || 0) + 1; });
        let best = null;
        let bestCount = -1;
        Object.keys(counts).map(Number).sort((a, b) => a - b).forEach((k) => {
            if (counts[k] > bestCount) {
                best = k;
                bestCount = counts[k];
            }
        });
        return best;
    }

    // -----------------------------------------------------------------------
    // Main analysis
    // -----------------------------------------------------------------------

    /**
     * Parse and validate a Statscore CSV, and build the JE lines.
     *
     * Never throws on bad data: everything lands in result.errors so the
     * operator sees the full picture in one pass.
     *
     * @param {string} csvText
     * @returns {Object} analysis result
     */
    function analyze(csvText) {
        const result = {
            ok: false,
            errors: [],
            warnings: [],
            dataRows: 0,
            month: null,
            year: null,
            tranDate: '',
            memo: '',
            lines: [],
            skippedZero: 0,
            totalDebit: 0,
            totalCredit: 0,
            refGroupsOk: 0,
            refErrors: [],
            rowErrors: [],
            accountIds: []
        };

        const rows = parseCsv(csvText || '');
        if (!rows.length) {
            result.errors.push('The file is empty.');
            return result;
        }

        const colMap = mapColumns(rows[0]);
        const required = ['internal', 'debit', 'credit', 'memo', 'date'];
        const missing = required.filter((k) => colMap[k] === undefined);
        if (missing.length) {
            result.errors.push('Missing required columns: ' + missing.join(', ') +
                '. Found: ' + rows[0].map((c) => String(c).trim()).join(', '));
            return result;
        }

        // Drop rows that are entirely blank, matching df.dropna(how='all').
        const dataRows = [];
        for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (r.some((c) => String(c === null || c === undefined ? '' : c).trim() !== '')) {
                dataRows.push({ csvLine: i + 1, cells: r });
            }
        }
        result.dataRows = dataRows.length;
        if (!dataRows.length) {
            result.errors.push('The file has a header row but no data rows.');
            return result;
        }

        const cell = (r, key) => (colMap[key] === undefined ? '' : (r.cells[colMap[key]] || ''));

        // --- period ---------------------------------------------------------
        const months = [];
        const years = [];
        dataRows.forEach((r) => {
            const d = parseDate(cell(r, 'date'));
            if (d) {
                months.push(d.month);
                years.push(d.year);
            }
        });
        if (!months.length) {
            result.errors.push('No valid dates found in the Date column (expected dd/mm/yy).');
            return result;
        }
        result.month = modeOf(months);
        result.year = modeOf(years);
        result.tranDate = result.year + '-' + pad2(result.month) + '-' +
            pad2(lastDayOfMonth(result.year, result.month));
        result.memo = 'Statscore ' + pad2(result.month) + '/' + result.year;

        const mixed = months.filter((m) => m !== result.month).length;
        if (mixed) {
            result.warnings.push(mixed + ' row(s) fall outside ' +
                pad2(result.month) + '/' + result.year +
                '. They are still included; the JE posts on ' + result.tranDate + '.');
        }
        const undated = dataRows.length - months.length;
        if (undated > 0) {
            result.warnings.push(undated + ' row(s) have an unreadable date. They are still included.');
        }

        // --- per REF balance (raw column sums, like the Python) -------------
        if (colMap.ref !== undefined) {
            const refTotals = {};
            const refOrder = [];
            dataRows.forEach((r) => {
                const ref = String(cell(r, 'ref')).trim();
                if (!refTotals[ref]) {
                    refTotals[ref] = { debit: 0, credit: 0 };
                    refOrder.push(ref);
                }
                refTotals[ref].debit += parseAmount(cell(r, 'debit')).value;
                refTotals[ref].credit += parseAmount(cell(r, 'credit')).value;
            });
            refOrder.forEach((ref) => {
                const t = refTotals[ref];
                const diff = Math.abs(t.debit - t.credit);
                if (diff >= CONFIG.BALANCE_TOLERANCE) {
                    result.refErrors.push({ ref: ref, debit: t.debit, credit: t.credit, diff: diff });
                } else {
                    result.refGroupsOk++;
                }
            });
        }

        // --- build the JE lines ---------------------------------------------
        const accountSet = {};
        dataRows.forEach((r) => {
            const idCell = parseInternalId(cell(r, 'internal'));
            if (idCell.blank) {
                return;                          // Python skips NaN internal IDs
            }
            if (!idCell.ok) {
                result.rowErrors.push({
                    csvLine: r.csvLine,
                    message: 'Internal ID "' + idCell.value + '" is not a number.'
                });
                return;
            }

            const dCell = parseAmount(cell(r, 'debit'));
            const cCell = parseAmount(cell(r, 'credit'));
            if (!dCell.ok || !cCell.ok) {
                result.rowErrors.push({
                    csvLine: r.csvLine,
                    message: 'Debit "' + String(cell(r, 'debit')).trim() +
                        '" / Credit "' + String(cell(r, 'credit')).trim() + '" is not a number.'
                });
                return;
            }

            const debit = dCell.value;
            const credit = cCell.value;
            if (debit === 0 && credit === 0) {
                result.skippedZero++;
                return;
            }

            const line = {
                csvLine: r.csvLine,
                account: idCell.value,
                memo: String(cell(r, 'memo') || '').trim(),
                debit: 0,
                credit: 0
            };

            // A negative debit is a credit, a negative credit is a debit.
            if (debit > 0) line.debit = round2(debit);
            else if (debit < 0) line.credit = round2(Math.abs(debit));

            if (credit > 0) line.credit = round2(credit);
            else if (credit < 0) line.debit = round2(Math.abs(credit));

            // NetSuite rejects a line carrying both, so catch it here rather
            // than after a long-running post.
            if (line.debit > 0 && line.credit > 0) {
                result.rowErrors.push({
                    csvLine: r.csvLine,
                    message: 'Row has both a debit (' + fmt(line.debit) + ') and a credit (' +
                        fmt(line.credit) + '). NetSuite allows only one per line.'
                });
                return;
            }

            accountSet[line.account] = true;
            result.totalDebit = round2(result.totalDebit + line.debit);
            result.totalCredit = round2(result.totalCredit + line.credit);
            result.lines.push(line);
        });

        result.accountIds = Object.keys(accountSet);

        // --- blocking conditions ---------------------------------------------
        if (result.refErrors.length) {
            result.errors.push(result.refErrors.length +
                ' REF group(s) are unbalanced. Fix them in the source file before posting.');
        }
        if (result.rowErrors.length) {
            result.errors.push(result.rowErrors.length + ' row(s) could not be converted to a JE line.');
        }
        if (!result.lines.length) {
            result.errors.push('No postable lines: every row was blank, zero or invalid.');
        }
        const diff = Math.abs(result.totalDebit - result.totalCredit);
        if (result.lines.length && diff >= CONFIG.BALANCE_TOLERANCE) {
            result.errors.push('The journal entry is unbalanced by ' + fmt(diff) +
                ' (debit ' + fmt(result.totalDebit) + ' vs credit ' + fmt(result.totalCredit) + ').');
        }

        result.ok = result.errors.length === 0;
        return result;
    }

    // -----------------------------------------------------------------------
    // Account validation
    // -----------------------------------------------------------------------

    /**
     * Check every account internal ID against NetSuite.
     *
     * @param {Array<string>} ids
     * @returns {{valid: Object, invalid: Array<string>, inactive: Array<string>}}
     */
    function validateAccounts(ids) {
        const valid = {};
        const invalid = [];
        const inactive = [];

        for (let i = 0; i < ids.length; i += CONFIG.ACCOUNT_BATCH) {
            const batch = ids.slice(i, i + CONFIG.ACCOUNT_BATCH);
            // IDs are digits-only by construction in parseInternalId.
            const sql = 'SELECT id, acctnumber, fullname, isinactive FROM account WHERE id IN (' +
                batch.join(',') + ')';
            const rows = query.runSuiteQL({ query: sql }).asMappedResults();
            rows.forEach((row) => {
                const id = String(row.id);
                valid[id] = {
                    acctnumber: row.acctnumber || '',
                    fullname: row.fullname || '',
                    isinactive: String(row.isinactive) === 'T'
                };
                if (valid[id].isinactive) {
                    inactive.push(id);
                }
            });
        }

        ids.forEach((id) => {
            if (!valid[id]) invalid.push(id);
        });

        return { valid: valid, invalid: invalid, inactive: inactive };
    }

    /**
     * Journal entries already posted for this period, so nobody has to spot a
     * doubled month by eye. Voided entries do not count.
     *
     * @param {string} memo the JE memo, e.g. 'Statscore 07/2026'
     * @returns {Array<{id: string, tranid: string, trandate: string}>}
     */
    function findExistingJes(memo) {
        const sql = 'SELECT DISTINCT t.id, t.tranid, t.trandate FROM transaction t ' +
            'JOIN transactionline tl ON tl.transaction = t.id ' +
            "WHERE t.recordtype = 'journalentry' AND tl.subsidiary = " + CONFIG.SUBSIDIARY +
            " AND t.voided = 'F' AND t.memo = ?";
        return query.runSuiteQL({ query: sql, params: [memo] }).asMappedResults().map((r) => ({
            id: String(r.id),
            tranid: String(r.tranid || ''),
            trandate: String(r.trandate || '')
        }));
    }

    return {
        CONFIG: CONFIG,
        analyze: analyze,
        validateAccounts: validateAccounts,
        findExistingJes: findExistingJes,
        parseCsv: parseCsv,
        parseAmount: parseAmount,
        parseDate: parseDate,
        round2: round2,
        fmt: fmt,
        lastDayOfMonth: lastDayOfMonth
    };
});
