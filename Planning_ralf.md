# Planificacion

# Generating a daily Excel file from Power BI with Power Automate and Office Scripts

## 0. Problem statement

Your question has four parts:

1. **What already works:** the flow queries Power BI, parses the JSON, and builds the file name `NOA_yyyyMMdd.xlsx` in a Compose step.
2. **Goal:** generate one Excel file per run from a SharePoint template, load the data into a table, and attach the file to an email.
3. **Open questions:** where the Office Script lives, and how to make sure the Excel file is finished before it's attached.
4. **Extra request:** list possible failures and how to mitigate them.

The answer follows that order.

---

## 1. Where does the script live?

**Not inside the template, and not inside any .xlsx file.** An Office Script is a separate file (`.osts`).

- By default it's saved in your OneDrive, under `/Documents/Office Scripts/`.
- It can also be saved to a SharePoint library.
- "Associating the script with a workbook" only links it to the file so others can see and run it from Excel. The code stays in the creator's OneDrive and is not copied into the workbook.

In Power Automate, you point to the workbook and the script separately. The Excel Online (Business) connector offers two actions:

| Action | Where the script lives |
|---|---|
| **Run script** | Always in your OneDrive |
| **Run script from SharePoint library** | In a SharePoint library you specify |

**Recommendation (best practice): store the `.osts` in a team SharePoint library and use "Run script from SharePoint library".**
- The flow doesn't depend on anyone's personal OneDrive. If the owner leaves the organization, a flow tied to their OneDrive breaks.
- Inside a library, the script gets versioning and controlled permissions.
- To move it: in the Code Editor, click the script name → **Move** → pick the SharePoint folder.

One restriction: scripts stored in SharePoint can't make external calls (`fetch`). That doesn't matter here, because the data arrives as a parameter.

---

## 2. Flow architecture

```
Trigger (Recurrence, concurrency = 1)
 ├─ [existing] Power BI query → Parse JSON → Compose FileName
 ├─ Compose Rows (array of rows)
 ├─ Condition: any rows?  (decide what happens with 0 rows)
 ├─ Scope TRY
 │    1. Get file content (template)
 │    2. Create file  (NOA_yyyyMMdd.xlsx in output folder)
 │    3. Run script from SharePoint library  (on the NEW file)
 │    4. Condition: rowsWritten == length(Rows)  → otherwise Terminate Failed
 │    5. Get file content (new file)
 │    6. Send an email (V2) with attachment
 ├─ Scope CATCH  (run after: TRY failed / timed out / skipped)
 │    - Build error details + link to the run
 │    - Email/Teams alert to the owner
 │    - Terminate (Failed)
 └─ (Optional) Scope FINALLY → log to a SharePoint list
```

**Key principle: never write to the template.** First copy the template to the daily file, then run the script on that copy. The template stays untouched and every run is reproducible. If the template already has no data rows, the "clear table" step in the script is an extra safeguard.

---

## 3. Preparing the template

1. Use the **standard .xlsx** format, not "Strict Open XML". With Strict format, the Run script action returns `BadGateway`.
2. Add a table with a fixed name, e.g. `tblNOA`, whose headers match the JSON columns (details in section 4).
3. Apply number and date formats to the table columns. New rows inherit the formatting and any calculated columns.
4. Store it in a separate folder (e.g. `/Templates/`) with read-only access for users. Output goes to a different folder (e.g. `/Reports/NOA/`).

---

## 4. The Office Script

The script receives the rows as a **JSON string** rather than a complex type. This avoids schema issues in Power Automate and keeps all validation inside the script.

```typescript
interface Result {
  ok: boolean;
  rowsWritten: number;
  message: string;
}

/**
 * Clears the target table and fills it with the received rows.
 * @param rowsJson   JSON array of objects (rows from the Power BI query)
 * @param tableName  Name of the target table in the workbook
 */
function main(
  workbook: ExcelScript.Workbook,
  rowsJson: string,
  tableName: string = "tblNOA"
): Result {
  // 1) Validate target table
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' does not exist in the workbook.`);
  }

  // 2) Validate and parse the JSON
  let rows: Record<string, unknown>[];
  try {
    rows = JSON.parse(rowsJson);
  } catch (e) {
    throw new Error(`Parameter rowsJson is not valid JSON: ${e}`);
  }
  if (!Array.isArray(rows)) {
    throw new Error("rowsJson must be an array of objects.");
  }

  // 3) Clear the table if it contains data
  const existingRows = table.getRowCount();
  if (existingRows > 0) {
    table.deleteRowsAt(0, existingRows);
  }

  if (rows.length === 0) {
    return { ok: true, rowsWritten: 0, message: "No data to write." };
  }

  // 4) Map JSON keys to table headers.
  //    Power BI returns keys like "Table[Column]"; normalize them to "column".
  const normalize = (k: string): string => {
    const m = k.match(/\[([^\]]+)\]$/);
    return (m ? m[1] : k).trim().toLowerCase();
  };

  // Union of keys across ALL rows: columns with BLANK values
  // may be omitted from some rows.
  const keyMap = new Map<string, string>();
  for (const r of rows) {
    for (const k of Object.keys(r)) {
      const n = normalize(k);
      if (!keyMap.has(n)) keyMap.set(n, k);
    }
  }

  const headers = table.getHeaderRowRange().getValues()[0].map(h => String(h));
  const missing = headers.filter(h => !keyMap.has(normalize(h)));
  if (missing.length > 0) {
    throw new Error(`Table columns with no match in the JSON: ${missing.join(", ")}`);
  }

  // 5) Convert values (ISO dates → Excel serial numbers)
  const isoDate = /^(\d{4})-(\d{2})-(\d{2})(T(\d{2}):(\d{2}):(\d{2}))?/;
  const toExcel = (v: unknown): string | number | boolean => {
    if (v === null || v === undefined) return "";
    if (typeof v === "string") {
      const m = v.match(isoDate);
      if (m) {
        const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[5] ?? 0), +(m[6] ?? 0), +(m[7] ?? 0));
        return (ms - Date.UTC(1899, 11, 30)) / 86400000;
      }
      return v;
    }
    if (typeof v === "number" || typeof v === "boolean") return v;
    return String(v);
  };

  const values = rows.map(r =>
    headers.map(h => toExcel(r[keyMap.get(normalize(h))!]))
  );

  // 6) Write in batches (Excel on the web limits each request to 5 MB)
  const BATCH_SIZE = 5000;
  for (let i = 0; i < values.length; i += BATCH_SIZE) {
    table.addRows(-1, values.slice(i, i + BATCH_SIZE));
  }

  // 7) Final check
  const written = table.getRowCount();
  if (written !== rows.length) {
    throw new Error(`Expected ${rows.length} rows, but the table has ${written}.`);
  }

  return { ok: true, rowsWritten: written, message: "OK" };
}
```

Notes on the script:
- It uses `throw` on any inconsistency, so the Power Automate action **fails explicitly** instead of "succeeding" with an incomplete file.
- It writes in bulk with `addRows`, not cell by cell. This is what keeps it under the **120-second timeout**.
- If the table has calculated columns that aren't in the JSON, exclude them from `headers` (e.g. with a list of columns to ignore). Otherwise the script reports them as missing.
- Test it first in Excel on the web, using a sample JSON pasted as the default parameter value. Then connect it to the flow.

---

## 5. Configuring the flow, step by step

### 5.1 Trigger
- **Recurrence**, with an explicit time zone.
- In the trigger's *Settings*: **Concurrency control = On, Degree of parallelism = 1**. This stops two runs from writing the same file at once.

### 5.2 File name (you already have this)
Check two things:
- **Time zone:** `utcNow()` can give you the wrong day between 21:00 and midnight local time. Use, for example:
  `convertFromUtc(utcNow(), 'Montevideo Standard Time', 'yyyyMMdd')`, or whichever time zone matches the business.
- **Date meaning:** does the name represent the run date or the data date (e.g. the previous day)? Your example `NOA_20261015` suggests the previous day's data. If so, wrap it in `addDays(..., -1)`.

### 5.3 Compose "Rows"
For the *Run a query against a dataset* action:
```
body('Run_a_query_against_a_dataset')?['firstTableRows']
```
Or take it from Parse JSON, e.g. `body('Parse_JSON')?['results']?[0]?['tables']?[0]?['rows']`, depending on how your flow is built.

### 5.4 Condition: empty data
Check whether `length(outputs('Rows'))` equals 0. Then define the business rule:
- (a) send a "no new data" email without an attachment, or
- (b) generate the empty file anyway.

What matters is that this is an explicit decision, not a side effect.

### 5.5 Scope TRY

1. **Get file content** (SharePoint) → template.
2. **Create file** (SharePoint)
   - *Folder path:* output folder.
   - *File name:* `outputs('FileName')`.
   - *File content:* `body('Get_file_content')`.
   - If the daily file already exists (e.g. on a rerun), SharePoint overwrites it and creates a new version. This makes the flow **idempotent**: rerunning a day doesn't duplicate files. Confirm this behavior with a test in your tenant.
3. **Run script from SharePoint library**
   - *Workbook:* the file you just created. Use the `Id` (or path) returned by **Create file**, **never the template**.
   - *Script:* the `.osts` file in the library.
   - *rowsJson:* `string(outputs('Rows'))`.
   - *tableName:* `tblNOA`.
   - **Retry policy** (Settings): Exponential, 4 retries, interval PT10S. Right after a file is created, the Excel connector sometimes returns a transient "file not found" or "locked" error.
4. **Validation condition:**
   check that `outputs('Run_script_from_SharePoint_library')?['body/result/rowsWritten']` equals `length(outputs('Rows'))`.
   - If they don't match → **Terminate (Failed)** with a message.
5. **Get file content** (SharePoint), using the `Id` of the created file.
   - Forum users report that this step sometimes returns a version from before Excel on the web saved the changes. If you see this in testing, add a **10–15 s Delay** before it. This isn't officially documented behavior, but the mitigation is cheap.
6. **Send an email (V2)**
   - Attachment: *Name* = `outputs('FileName')`, *Content* = `body('Get_file_content_2')`.

### 5.6 How do you make sure the Excel file is finished before the email?

There are three layers:

1. **Native sequencing:** in Power Automate, each action waits for the previous one to finish, according to its *run after* settings. No manual waiting is needed.
2. **Run script is synchronous:** the action doesn't complete until the script finishes and Excel saves the changes. If the script throws, the action fails and the email is never sent.
3. **Business validation:** the script returns `rowsWritten` and the flow compares it with the expected row count. This confirms the file is *correct*, not just *finished*.

### 5.7 Scope CATCH
- Set the scope's *Configure run after* to **has failed, has timed out, is skipped** (relative to TRY).
- Error details:
  `filter(result('Try'), item()?['status'] == 'Failed')`
- Link to the run:
  ```
  concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'],
         '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])
  ```
- Alert the owner (email or Teams) with the file name, the error, and the link.
- End with **Terminate → Failed**. Otherwise the run shows "Succeeded" and monitoring misses the failure.

### 5.8 FINALLY (optional, recommended)
Log to a SharePoint list: date, status, row count, file name, and run link. This is your audit trail.

---

## 6. Failure matrix and mitigations

| # | Possible failure | Impact | Mitigation |
|---|---|---|---|
| 1 | Power BI dataset hasn't refreshed yet | Email with stale data (silent failure, the most dangerous one) | Schedule the flow well after the refresh. Validate that the max date in the data is the expected one; otherwise abort and alert |
| 2 | DAX query errors or hits limits (~100,000 rows / 1M values / 15 MB per query) | Truncated data or failure | Keep the DAX scoped. Paginate or partition the query if you get close to the limit |
| 3 | Zero rows | Empty file sent with no warning | Explicit branch (section 5.4) |
| 4 | Column renamed in the model | Mapping can't find columns | The script validates and throws, listing the missing columns |
| 5 | Template modified or deleted, or table renamed | Script fails or template gets overwritten | Template in a read-only folder with versioning on. The script validates the table. Never write to the template |
| 6 | Daily file already exists (rerun) | Duplicates or error | Create file overwrites (idempotency); SharePoint versioning keeps the previous copy |
| 7 | Concurrent runs | File corruption or lock | Trigger concurrency = 1 |
| 8 | File locked or "not found" right after creation | Transient failure | Exponential retry policy on Run script |
| 9 | 120 s Run script timeout | Incomplete file | Bulk, batched writes. If volume grows a lot, split the work into multiple script calls |
| 10 | Payload > 5 MB per Excel request / parameters > 28.6 MB | "Payload size exceeded" error | 5,000-row batches (adjustable); monitor volume |
| 11 | Limit of 1,600 Run script calls per user per day | Failure if the account is shared with other heavy flows | Dedicated service account |
| 12 | Expired connection, or flow owner leaves | Flow stops | Service-account connections, flow co-owners, script in SharePoint (not personal OneDrive) |
| 13 | Script deleted or changed | Failure or different behavior | Script library with restricted permissions and versioning |
| 14 | `Get file content` returns an earlier version | Incomplete attachment | Short delay before the step + earlier row-count validation |
| 15 | Wrong date in the file name (UTC vs. local) | File named for the wrong day | `convertFromUtc` with an explicit time zone |
| 16 | Dates written as text | Excel can't filter or calculate on them | ISO → serial conversion in the script + date format in the template |
| 17 | Email send fails (large attachment, mailbox issue) | File created but not delivered | Retry policy on Send email; CATCH alerts; file stays in SharePoint for manual resend |
| 18 | Flow turned off after continuous errors (14 days) or no activity (90 days, without a premium license) | Stops without warning | Monitor runs (log from section 5.8) and act on alerts |
| 19 | Template saved as Strict Open XML | `BadGateway` | Save as standard .xlsx |

---

## 7. Go-live checklist

- [ ] Script in a SharePoint library; flow uses **Run script from SharePoint library**.
- [ ] Service-account connections and flow co-owners.
- [ ] Standard .xlsx template with table `tblNOA`, in a read-only folder.
- [ ] Trigger concurrency = 1.
- [ ] Try / Catch with Terminate Failed and an alert that includes the run link.
- [ ] `rowsWritten` vs. expected-row validation before the email.
- [ ] Defined rules for 0 rows and for stale data.
- [ ] Tests: normal day, day with no data, same-day rerun, renamed column, template locked because a user has it open.

Sources:
- [Run Office Scripts with Power Automate – Excel connector](https://learn.microsoft.com/office/dev/scripts/develop/power-automate-integration#excel-connector)
- [Office Scripts file storage and ownership](https://learn.microsoft.com/office/dev/scripts/overview/script-storage)
- [Platform limits and requirements with Office Scripts](https://learn.microsoft.com/office/dev/scripts/testing/platform-limits)
- [Write a large dataset (Office Scripts sample)](https://learn.microsoft.com/office/dev/scripts/resources/samples/write-large-dataset)
- [Limits of automated, scheduled, and instant flows](https://learn.microsoft.com/power-automate/limits-and-config)

## Formato resaltado basado en Due
El flujo está funcionando. Quiero que la tabla de excel tenga color condicional, toda la fila, basado en el valor de una de las columnas (Due). Cuando el flujo pisa el template, se me borra el formato condicional. Hay una alternativa?

Solución Claude:
## Why the rule disappears

A conditional formatting (CF) rule is attached to a specific cell range, such as `F2:F120`. Your script calls `deleteRowsAt` to clear the table, which deletes the cells the rule points to. Once every data row is gone, the rule's range is empty and Excel drops it. Whether `addRows` later extends a surviving rule to the new rows isn't documented as guaranteed, so you shouldn't rely on the template's rule.

## Recommended approach: let the script own the formatting

After writing the data, have the script:

1. Clear any conditional formats on the target range.
2. Add the rules again.

Benefits:

- **Deterministic:** the result doesn't depend on what Excel kept or dropped.
- **Idempotent:** clearing first means reruns never pile up duplicate rules.
- **Versioned:** the rules live in code, next to the data logic, not hidden in the template.

Also remove the manual rule from the template, so there is only one source of truth.

### Code to add to the script

**1. New parameters on `main`** (they have defaults, so existing calls still work):

```typescript
function main(
  workbook: ExcelScript.Workbook,
  rowsJson: string,
  tableName: string = "tblNOA",
  dueColumnName: string = "due",
  formatScope: "cell" | "row" = "cell"   // "cell" = only the due cell, "row" = whole table row
): Result {
```

**2. Call it between step 6 (batch write) and step 7 (final check).** It sits after the early `return` for zero rows, so it never runs on an empty table:

```typescript
  // 6b) Re-apply conditional formatting (owned by the script, not the template)
  applyDueFormatting(table, dueColumnName, formatScope);
```

**3. Helper functions** (place them below `main`):

```typescript
const RED = "#C00000";
const ORANGE = "#ED7D31";

/**
 * Applies the "due" conditional formatting rules:
 *   due < 0  → bold red
 *   due = 0  → bold orange
 *   otherwise → unchanged
 * Clears previous rules on the target range first, so it is safe to rerun.
 */
function applyDueFormatting(
  table: ExcelScript.Table,
  dueColumnName: string,
  scope: "cell" | "row"
): void {
  const column = table.getColumns().find(
    c => c.getName().trim().toLowerCase() === dueColumnName.trim().toLowerCase()
  );
  if (!column) {
    throw new Error(`Column '${dueColumnName}' not found in table '${table.getName()}'.`);
  }

  const dueRange = column.getRangeBetweenHeaderAndTotal();
  const target = scope === "row" ? table.getRangeBetweenHeaderAndTotal() : dueRange;

  // Idempotency: remove any existing rules on the target range
  target.clearAllConditionalFormats();

  if (scope === "cell") {
    addCellValueRule(target, ExcelScript.ConditionalCellValueOperator.lessThan, RED);
    addCellValueRule(target, ExcelScript.ConditionalCellValueOperator.equalTo, ORANGE);
  } else {
    // Absolute column, relative row, anchored on the first due cell, e.g. "$F2"
    const firstCell = dueRange.getCell(0, 0).getAddress().split("!").pop()!; // "F2"
    const ref = "$" + firstCell;
    addFormulaRule(target, `=AND(ISNUMBER(${ref}),${ref}<0)`, RED);
    addFormulaRule(target, `=AND(ISNUMBER(${ref}),${ref}=0)`, ORANGE);
  }
}

function addCellValueRule(
  range: ExcelScript.Range,
  operator: ExcelScript.ConditionalCellValueOperator,
  color: string
): void {
  const cf = range
    .addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue)
    .getCellValue();
  cf.setRule({ formula1: "0", operator: operator });
  const font = cf.getFormat().getFont();
  font.setColor(color);
  font.setBold(true);
}

function addFormulaRule(range: ExcelScript.Range, formula: string, color: string): void {
  const cf = range
    .addConditionalFormat(ExcelScript.ConditionalFormatType.custom)
    .getCustom();
  cf.getRule().setFormula(formula);
  const font = cf.getFormat().getFont();
  font.setColor(color);
  font.setBold(true);
}
```

### Notes

- **Formula syntax:** formulas passed through the API use en-US syntax (comma separators, English function names), whatever your Excel locale is. Write `AND(...,...)`, not `Y(...;...)`.
- **Data type:** the rules compare numbers. If Power BI returns `due` as a string (e.g. `"-3"`), comparisons fail silently. Make sure the measure or column is numeric, or convert it in `toExcel`.
- **Blank cells in `cell` mode:** a "Cell value equal to 0" rule treats blank cells as 0. That's harmless here, because bold orange on an empty cell is invisible. The `row` variant guards against blanks explicitly with `ISNUMBER`.
- **Other rules get cleared too:** `clearAllConditionalFormats()` removes every rule on the target range. In `row` mode, that's the whole table body. If the template has other CF rules you want to keep, apply them in the script as well.
- **Refresh the flow action:** Power Automate caches the script's parameters. Re-select the script in the *Run script from SharePoint library* action so the new optional fields show up.

### Lighter alternative (with trade-offs)

You can define the rule in the template on the entire worksheet column (for example `$F:$F`) instead of on the table range. A whole-column range survives row deletion, so the rule persists. The downside is that it breaks silently if someone moves or inserts columns, and it applies outside the table too. The script-based approach is more robust.

Sources:
- [Office Scripts sample: Conditional formatting](https://learn.microsoft.com/office/dev/scripts/resources/samples/conditional-formatting)
- [ExcelScript.ConditionalCellValue interface](https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.conditionalcellvalue)
- [ExcelScript.Range – addConditionalFormat / clearAllConditionalFormats](https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.range)