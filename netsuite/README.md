# Statscore JE Creator, inside NetSuite

Replaces the desktop tool (`statscore_je_creator.py` / `NetSuite_Import_Processor.exe`) with a
page inside NetSuite. Same job: take IZA's mapped CSV, check it, post one journal entry.

Nothing else changes. IZA still produces the CSV the same way. The desktop processor
(`netsuite_import_processor_all transactions.py`) is untouched.

## Why this is worth doing

| | Desktop tool | This |
| --- | --- | --- |
| Install | Python + pandas, or the packaged EXE | None, it is a NetSuite page |
| Credentials | OAuth tokens in a `.env` on `\\mainsrv` | None, runs as the logged-in user |
| Permissions | Whatever the token holds | The user's own NetSuite role |
| Audit trail | Local only | NetSuite execution log, the JE names its creator |
| Who can run it | Whoever has the EXE and the share | Any role you give the deployment to |

## How it works

```
You ──> Suitelet ──> Preview: period, totals, REF balance, accounts, duplicate check
             │
             ├─ stages the CSV in the File Cabinet
             ├─ writes a job file  {status: PENDING, ...}
             └─ starts the scheduled script ──> builds and saves ONE journal entry
                                                 └─ writes the result back to the job file
        Status page polls the job file, then links to the new JE
```

The posting sits in a scheduled script rather than in the page because a Statscore month is
around 4,000 to 5,000 lines and the save runs for minutes, longer than NetSuite lets a page
request run. The status page can be closed and reopened; the job file holds the state.

## Files

All three must live in the **same File Cabinet folder**. The scripts import the library with a
relative path (`./statscore_je_lib`), so splitting them up breaks the import.

| File | What it is |
| --- | --- |
| `statscore_je_lib.js` | Library. All parsing and validation, defined once so the page and the posting script cannot drift apart. |
| `statscore_je_suitelet.js` | The page: upload, preview, submit, status. |
| `statscore_je_sched.js` | Builds and saves the journal entry. |

## Deployment

Once, by someone with the Administrator role.

### 1. Upload the files

**Documents > Files > File Cabinet**. Create `SuiteScripts / Statscore JE Creator` and upload all
three `.js` files into it.

### 2. Create the scheduled script

**Customization > Scripting > Scripts > New**, select `statscore_je_sched.js`, then Create Script Record.

- **Name**: `Statscore JE Creator - Post`
- **ID**: `_statscore_je_sched` (NetSuite prefixes it, giving `customscript_statscore_je_sched`)

On the **Parameters** subtab add one:

| Label | ID | Type | Preference |
| --- | --- | --- | --- |
| Job File | `_statscore_je_job` (becomes `custscript_statscore_je_job`) | Integer Number | Entry Form |

Save, then **Deploy Script**:

- **ID**: `_statscore_je_sched` (becomes `customdeploy_statscore_je_sched`)
- **Status**: `Not Scheduled` (the page starts it on demand; it must never run on a timer)

### 3. Create the Suitelet

**Customization > Scripting > Scripts > New**, select `statscore_je_suitelet.js`, Create Script Record.

- **Name**: `Statscore JE Creator`
- **ID**: `_statscore_je_suitelet`

**Deploy Script**:

- **ID**: `_statscore_je_suitelet`
- **Status**: `Released`
- **Audience**: the roles that post Statscore. Leave **Available Without Login** unticked.

### 4. Bookmark it

The deployment page shows the **External URL** / **script URL**. That URL is what replaces the EXE.
To put it in the menu instead, add a Center Tab or a shortcut pointing at it.

### If NetSuite gives you different IDs

The Suitelet looks for the posting script by ID. If yours differ, edit the top of
`statscore_je_suitelet.js` and re-upload:

```js
const SCHED_SCRIPT_ID = 'customscript_statscore_je_sched';
const SCHED_DEPLOY_ID = 'customdeploy_statscore_je_sched';
const JOB_PARAM       = 'custscript_statscore_je_job';
```

## Using it

1. Open the Suitelet URL, pick IZA's CSV, click **Preview**.
2. Check the preview. It shows the period, the transaction date, the memo, every segment value,
   the line count, total debit and credit, whether each REF group balances, the first 25 lines and
   every account with its number and name.
3. Anything blocking is listed with the CSV line number. Fix the source file and upload again.
   Nothing is posted while a blocker stands.
4. If a journal entry for that period already exists, the preview names it with a link and the
   **Create Journal Entry** button will not post until you tick the confirmation box.
5. Click **Create Journal Entry**. The status page shows Queued, then Building, then Done with a
   link to the entry. Closing the page does not stop the job.

## What the posted entry looks like

Fixed for every line, matching the entries already in the account:

| | |
| --- | --- |
| Subsidiary | Statscore (6) |
| Currency | EUR (1) |
| Department | Statscore import (24) |
| Location | Poland (5) |
| Location (Expenses) | Other (3) |
| Transaction date | last day of the detected month |
| Memo | `Statscore MM/YYYY` |
| Approved | yes |

## Rules it applies

Carried over from the Python tool, verified line for line against it:

- Rows where debit and credit are both zero are skipped.
- A negative debit posts as a credit. A negative credit posts as a debit.
- Every REF group must balance to within 0.02, and so must the entry as a whole.
- The period is the most common month in the Date column. Dates are `dd/mm/yy`; a two-digit year
  of 68 or lower reads as 20xx, 69 or higher as 19xx.
- Header matching tolerates the trailing space in `Internal ` and the `Currecny` spelling.

### Where it deliberately differs

Four changes, each turning a silent or late failure into one you see on the preview screen:

1. **A non-numeric amount is reported with its CSV line number.** The Python raised an exception
   and showed a stack trace.
2. **A row carrying both a debit and a credit is caught in preview.** The Python built such a line
   and NetSuite rejected the whole entry minutes into the post.
3. **Inactive accounts are flagged.** The Python only checked that the account existed. NetSuite
   rejects lines posting to an inactive account.
4. **A row with a blank REF is included in the balance check.** pandas silently dropped those rows
   from the check, so an orphan row only surfaced later as an unbalanced entry.

One cosmetic difference: line memos are trimmed of leading and trailing spaces.

## Safety

- The preview never writes anything. The first write happens when you click Create.
- The submit handler re-reads and re-validates the staged file rather than trusting what the
  browser posted back, and re-runs the duplicate check.
- The posting script only ever acts on a job marked `PENDING`. If a run dies mid-save the job stays
  at `RUNNING` and will not be picked up again, so a month cannot be posted twice by a retry.
  Check NetSuite for a partial entry before re-running that month.

## Troubleshooting

**"Could not start the posting script"** - a previous run is still going. Check
**Customization > Scripting > Script Deployments** for `Statscore JE Creator - Post`, wait for it
to finish, then submit again.

**Status stuck on Queued** - the scheduled script queue is busy with other jobs. It will start.
The page cross-checks the queue and will say so if the script actually failed.

**Status says Failed** - the message is on the page. The full stack trace is in
**Customization > Scripting > Script Execution Log**, filtered to the posting script.

**Accented characters look wrong in memos** - the file is read as UTF-8 first, falling back to
Windows-1252. If IZA's export uses a different encoding, save it as UTF-8 before uploading.

**Housekeeping** - each run leaves the source CSV and a small job file in
`SuiteScripts / Statscore JE Creator`. A few MB a year. Clear old ones out by hand when you want to;
they are deliberately not auto-deleted, being the source record of what was posted.
