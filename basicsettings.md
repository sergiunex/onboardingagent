---
name: onboarding-agent-1.2
description: >
  Unified Nexudus onboarding agent for setting up a new coworking space via the Nexudus CLI.
  Use this skill whenever the user is onboarding a customer, setting up a new space, or importing
  data into Nexudus — even if they only mention one part of the setup. Triggers on any of:
  creating or importing membership plans, tariffs, or memberships from a spreadsheet; creating
  bookable resources, meeting rooms, phone booths, hot desks, or event spaces; adding booking
  credits, time credits, or extra services to plans; creating resource pricing rates (hourly,
  half-day, full-day); managing floor plans, floor plan areas, zones, desks, or offices; adding
  booking products to resources; adding language translations to products or plans; listing or
  updating bookings; duplicating or cloning existing resources; moving plans between locations
  in a business network; or any general "set up the space" / "onboard this customer" request.
  Also triggers whenever a .xlsx or .csv file is provided alongside a Nexudus setup task, or
  when the user mentions floor plans, resource rates, tariffs, bookable resources, translations,
  bookings, duplicating resources, or moving plans between locations in a CLI context.
  This skill covers all ten onboarding domains — Plans, Resources, Credits, Resource Rates,
  Floor Plans, Resource Booking Products, Product Translations, Bookings, Duplicating Resources,
  and Moving Plans Between Locations — use it even when only one domain is in scope.
---

# Nexudus Onboarding Agent v1.2

> _Flags/commands last verified against **nexudus CLI v5.0.33** on **2026-10-07**. **All enum values are now confirmed** by client-side probe (see the tip below) — see [Appendix: Verified Enums](#enums) for the authoritative lists; they are no longer "based on prior empirical testing". Always re-run `<command> <subcommand> --help` when a command errors unexpectedly; the CLI changes between versions. The `resources` flags and units were re-verified by readback on **2026-09-28**._
>
> _**Tip — list an enum's valid values safely:** pass a deliberately invalid value, e.g. `nexudus resources create --system-resource-type NotARealValue`. The CLI fails to parse **client-side**, prints `Valid values are '...'`, and sends no request — so nothing is written. Works on `create` as well as `update`, and on any enum flag `--help` does not enumerate._
>
> _**Changed between v5.0.23 and v5.0.33** (verified 2026-10-07): `tariffs create --sign-up-fee` **removed**; `bookings update --billed` **removed**; `resources create/update --access-control-group-id` **removed**. New and onboarding-relevant: `extraservices --time-slots` (time-of-day scoped rates), `extraservices --fixed-cost-length`/`--fixed-cost-price`/`--default-price`/`--tariffs`/`--teams`, a large set of `resources` amenity flags, and `tariffs` usage-limit and virtual-office families. Each is documented in its own Part below._

## Overview

This skill guides you through the core onboarding tasks for a new Nexudus space:

1. **[Plans](#plans)** — Create membership plans (tariffs), attach signup fees and time passes
2. **[Resources](#resources)** — Create bookable resources (rooms, desks, phone booths)
3. **[Credits](#credits)** — Add time-based booking credits (extra services) and money booking credits (`tariffbookingcredits`) to plans
4. **[Resource Rates](#rates)** — Create hourly, half-day, and full-day pricing for resource types
5. **[Floor Plans](#floorplans)** — Manage floor plan layouts, areas, and bulk-create desk items
6. **[Resource Booking Products](#resourceproducts)** — Create store products (`products create`) and attach purchasable products to resources for the booking flow
7. **[Product Translations](#translations)** — Add language translations (e.g. French) to products/plans
8. **[Bookings Management](#bookings)** — List, filter, and update booking properties in bulk
9. **[Duplicating Resources](#duplicating)** — Clone an existing resource with all its settings and associations
10. **[Moving Plans Between Locations](#movingplans)** — Reassign tariffs to another business/location in the network

Reference: **[Appendix: Verified Enums](#enums)** — every enum flag's valid values, read from the CLI rather than guessed.

Work through whichever sections apply. The golden rules across all of them: **always show the user a summary and get approval before creating anything, and always verify writes by reading the records back — never trust exit codes (see the two shared rules below).**

---

## Shared: Reading Spreadsheet Files

**⚠️ First, check the file actually contains what its name says.** These onboarding workbooks are multi-tab Google Sheets, and exporting one tab produces a file named after the *workbook*, not the tab. Observed 2026-10-07: `D'Ventures - Plans.csv` contained the **Products** tab — byte-identical in size to `D'Ventures - Products.csv` — and no plan data at all. Confirm the header rows describe the entity you were asked about before parsing a single value, and say so immediately if they don't rather than trying to infer plans from a products grid.

**These sheets are column-oriented, not row-oriented.** Each *column* is one plan/product/resource (`MEMBERSHIP 01`…`MEMBERSHIP 30`), and each *row* is a setting. Column 0 is a category label, column 1 the setting description, and columns 2–3 are usually `EXAMPLE` columns that must be skipped — real data starts at column 4. Transpose before doing anything else.

**Most of the file is empty.** A 1,000-line export typically has 16–32 lines of content. Find them first:

```powershell
$lines = Get-Content $p
for ($i=0; $i -lt $lines.Count; $i++) { if (($lines[$i] -replace ',','').Trim().Length -gt 0) { "$($i+1)" } }
```

**⚠️ Cells contain embedded newlines, so naive line-splitting corrupts the grid.** A `Billing period` cell holding `"Monthly\nor\nEvery 3 months"` shifts every subsequent row. Use a real CSV parser:

```powershell
Add-Type -AssemblyName Microsoft.VisualBasic
$parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path)
$parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(',')
$rows = @(); while (-not $parser.EndOfData) { $rows += ,$parser.ReadFields() }
$parser.Close()
```

Prices arrive as `"HK$3,900.00"` — strip with `-replace '[^\d.]',''` before casting. Extract a leading quantity from a prose cell with a regex (`'^([\d.]+)\s*credits?'`), never by string-splitting the sentence.

For Excel files, use the PowerShell COM object:

```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open("C:\path\to\file.xlsx")
foreach ($ws in $wb.Worksheets) {
    Write-Host "=== Sheet: $($ws.Name) ==="
    $usedRange = $ws.UsedRange
    for ($r = 1; $r -le $usedRange.Rows.Count; $r++) {
        $line = @()
        for ($c = 1; $c -le $usedRange.Columns.Count; $c++) {
            $line += $usedRange.Cells($r, $c).Text
        }
        Write-Host ($line -join "`t")
    }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
```

---

## Shared: Non-ASCII Text — Write `.ps1` Files With a UTF-8 BOM

**Windows PowerShell 5.1 reads a `.ps1` file as ANSI (Windows-1252) unless it starts with a UTF-8 BOM.** A script saved as plain UTF-8 therefore turns every accented character into mojibake *before* the CLI ever sees it — `café` becomes `cafÃ©`, `Impresión` becomes `ImpresiÃ³n` — and the records are created corrupted. Exit codes are `ok=true` throughout, so nothing flags it.

This bites on every non-English account (Spanish, French, Catalan, German…). Two safe routes:

- **Run the commands inline** in a single shell invocation rather than from a file — inline command strings are passed through as UTF-8 correctly.
- **Or add a BOM** after writing the script, before running it:
  ```powershell
  $text = [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false))
  [System.IO.File]::WriteAllText($p, $text, [System.Text.UTF8Encoding]::new($true))
  ```

**Verify by readback, and grep for the damage explicitly** — console output is itself re-encoded, so a mangled console line does not prove the stored data is wrong (and a clean one does not prove it is right). Check the API:

```powershell
$bad = $all | Where-Object { $_.Name -match 'Ã|Â|�' -or $_.Description -match 'Ã|Â|�' }
```

Corrupted records can be repaired in place with `update` (no need to delete and recreate) once the encoding is fixed.

## Shared: Look Up Account IDs

Run these before any create commands:

```powershell
nexudus whoami                          # → business ID, currency ID
nexudus taxrates list --json            # → tax rate ID (empty for US accounts)
nexudus financialaccounts list --json   # → financial account ID (empty for US accounts)
```

- **US accounts** — omit `--tax-rate-id` and `--financial-account-id`
- **European/HUF accounts** — usually require both

Always apply `--caller claude` and `--intent "..."` to every command for traceability in audit logs.

---

## Shared: Verify by Readback — Never Trust Exit Codes

**The CLI's exit code and stdout are not reliable success signals. They lie in both directions.**

- A `delete` loop can return exit `0` for every call while the records are **not** deleted. (Observed: a 316-item `floorplandesks delete` reported "316 deleted, 0 failed" but a re-`list` showed 113 survived.)
- A `create` can print a parse error (e.g. `tariffextraservices create` → `ExtraServiceChargePeriod` conversion error) while the record **is** created successfully.
- An `update` can report `ok=true` per the batch but silently fail one item on a server-side constraint (e.g. a name collision returning HTTP 400).

**Rule:** after every create / update / delete — especially bulk loops — confirm the result by re-querying with `get` or `list` and comparing counts or field values against intent. Report only what the readback shows. Do **not** report "N done, 0 failed" based on `$LASTEXITCODE`. When a batch reports all-success, still spot-check by readback before telling the user it worked.

## Shared: Confirm the Active Account First

**The logged-in account can change between turns without warning.** Run `nexudus whoami` at the start of a task and again before any batch of writes, and state the business **name + ID** back to the user.

- Accounts in the same country/currency (e.g. several GBP/UK businesses) look identical except for the business name/ID — that is the only reliable tell.
- Treat an account change as a hard stop: flag it and confirm before continuing. IDs gathered under one account (resource types, tariffs, products) are meaningless under another.

## Shared: `list` Drops Collection Fields — Use `get`

**Several `list` DTOs return collection and relationship fields as `null`, empty, or default regardless of the stored value.** A `list`-based check therefore makes correctly-configured records look unconfigured — and in PowerShell `@($null).Count` is **1**, so naive emptiness checks mislead in both directions. Known cases:

| Entity | Field(s) `list` gets wrong | Use instead |
|---|---|---|
| `resources` | `LinkedResources`, `LinkedResourceIds` — always `null` | `resources get <id>` |
| `extraservices` | `TimeSlots` — omitted entirely | `extraservices get <id>` |
| `tariffbookingcredits` | `ElegibleResourceTypes` — omitted; toggle under typo'd `CaneBeUsedForBookings` | `tariffbookingcredits get <id>` |
| `products` | `AvailableAs` (always `1`), `OnlyForMembers`, `OnlyForContacts` (always `false`) | `products get <id>` |

**Rule: never conclude that a collection or relationship field is empty from `list` output.** Confirm with `get` on at least a sample before reporting anything as missing or unconfigured — and especially before telling the user that work is outstanding. Assume this applies to any collection field on any entity, not just the four above; the list is what has been caught so far, not an exhaustive one.

## Shared: Cross-Network Listings

`list` commands (`tariffs list`, `resources list`, `timepasses list`, etc.) return records from **every business in the network**, not just the active one. Product/pass/tariff IDs you reference may also live on a different business.

- Filter by **`BusinessId`**, not `BusinessName` — confirm the numeric ID matches the location you intend.
- When listing "the plans/resources in location X", always filter: `... | Where-Object { $_.BusinessId -eq <id> }`.

## Shared: Passing List / Array Fields (Repeated Flags)

> **PowerShell trap in batch loops:** never key an `[ordered]@{}` dictionary by integer IDs. `$map[1415400769]` on an OrderedDictionary is read as a **positional index**, returns `$null`, and every command in the loop goes out with an empty value. Use `@{}` with string keys (`$map["$id"]`) or an array of pairs.
>
> **⚠️ Never name a loop variable `$pid`.** `$PID` is a **read-only automatic variable** (the process ID). `foreach ($pid in $productIds)` throws `Cannot overwrite variable PID because it is read-only or constant` **once per outer iteration** and the inner loop body **never executes** — so a nested batch silently creates nothing while the outer loop's own `Write-Host` lines still print "submitted N" for every item. Observed 2026-10-07: a 135-record `resourceproducts` batch reported ten resources "submitted" with zero writes. Other reserved names to avoid for loop variables: `$input`, `$host`, `$error`, `$args`, `$matches`, `$true`/`$false`/`$null`. Use `$prodId`, `$resId`, `$tariffId` and the like.
>
> This is the exact failure mode the readback rule exists for — count the records with `list` after the loop, never trust the loop's own tally.
>
> **⚠️ Don't give a helper function a single-letter name either.** `function R($n) { $rows[$n-1] }` is shadowed by PowerShell's built-in alias `r` → `Invoke-History`, so every call fails with `Cannot locate the history for Id 21` instead of returning your row. Same hazard for `h`, `r`, `sc`, `gc`, `cd`, `cp`, `mv`, `ls`. Check with `Get-Command <name>` before defining, or just use descriptive names (`Get-Row`).

Array-valued flags (`--added-tariffs`, `--added-linked-resources`, and similar) do **not** accept a comma-separated string — that fails with `Failed to convert '...' to Int64[]`. Repeat the flag once per value instead:

```powershell
$u = @('resources','update',"$id",'--caller','claude','--intent','Attach tariffs')
foreach ($t in $tariffIds) { $u += '--added-tariffs'; $u += "$t" }
& nexudus @u --agent
```

Build the command as a PowerShell **array** and splat it (`& nexudus @args`) — this also handles names with spaces and lets you escape embedded quotes (`$val -replace '"','\"'`) for HTML/description fields.

## Shared: Windows Command-Line Length Limit (~32 KB)

A single `nexudus` invocation whose full command line exceeds ~32,767 characters fails with `The filename or extension is too long`. This bites when copying long HTML fields (`--description`, `--email-confirmation-content` can each be tens of KB).

- Keep large fields **out** of the create call: create with everything else, then set each large field in its **own isolated update** (one big field per command stays under the limit).
- `@filepath` input is **only** supported by `--time-slots` — **not** by `--description`. Passing `--description "@C:\file.txt"` stores the literal string, silently corrupting the field. Never use it for text fields.

<a name="numeric-enum-readback"></a>
## Shared: Numeric Enum Readback

Many fields accept a **name** on write but return an **integer** on `get`/`list`. Do not treat the integer as a failed write.

| Field | Write value | Reads back as |
|---|---|---|
| `PassRenewalTime` / `ServiceRenewalTime` | `Week` | `1` |
| `DayOfWeek` (time slots) | `Monday` | `1` (Sunday = 0) |
| `ChargePeriod` (extra services) | `Minutes` | `1` |
| `ItemType` (floor plan desks) | `Office` etc. | integer |
| `SystemTariffType` / `SystemResourceType` | `FullTimePrivateOffice` etc. | integer |

When verifying, compare against the integer (or accept either form) rather than expecting the name back.

---

<a name="plans"></a>
## Part 1: Plans (Tariffs)

### Step 1 — Read the Excel file

Use the shared Excel reader above.

### Step 2 — Confirm the plan list

Present a summary table and wait for approval. Include:

- Plan name, price, billing period, visibility
- Minimum term and notice/cancellation period
- Signup fee (amount + product ID)
- Time passes (pass name, qty, renewal period)
- Anything being skipped (deposits, credits, benefits) and why

Key questions before proceeding:
- If the spreadsheet says **"Monthly or Yearly"** — ask whether to create both variants
- If pass names don't exactly match what the user provided — confirm the correct ID
- If yearly variants are needed — confirm the prices separately

### Step 3 — Create tariffs

```powershell
nexudus tariffs create `
  --caller claude --intent "Create plan from Excel import" `
  --business-id <id> `
  --currency-id <id> `
  --tax-rate-id <id>           # omit if not present
  --financial-account-id <id>  # omit if not present
  --name "Plan Name" `
  --description "..." `
  --visible true `
  --price 99 `
  --invoice-every 1 `           # 1 = monthly, 12 = yearly
  --default-invoicing-day 1 `
  --prorate-day-of-month 1 `
  --prorate-cancellations false `
  --cancellation-period 30 `            # notice period in DAYS
  --cancel-member-account-after 15 `    # DAYS an invoice can be overdue before the account is suspended
  --auto-cancel-after 3 `               # BILLING CYCLES after which the contract auto-cancels (omit if none)
  --default-contract-term 2 `   # minimum term in months (omit if none)
  --disable-portal-cancellations true ` # omit if members CAN self-cancel
  --system-tariff-type PartTimeHotDesk `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

#### `--system-tariff-type` valid values

| Value | Use for |
|---|---|
| `FullTimeHotDesk` | Unlimited hot desk / 24-7 coworking |
| `PartTimeHotDesk` | Limited days/hours hot desk |
| `FullTimeDedicatedDesk` | Fixed/dedicated desk (unlimited) |
| `PartTimeDedicatedDesk` | Fixed desk (limited) |
| `FullTimePrivateOffice` | Private office (full time) |
| `PartTimePrivateOffice` | Private office (part time) |
| `Virtual` / `VirtualOffice` | Virtual mailing address plans |
| `Storage` | Storage units |
| `FullTimeOther` / `PartTimeOther` / `Other` | Everything else |

Yearly variants: same settings, `--invoice-every 12`, append `" - Annual"` to the name.

#### Plan usage limits — the flags a part-time plan actually needs (v5.0.33)

`--system-tariff-type PartTimeHotDesk` is only a *label*; it enforces nothing. A sheet saying "8 days a month" or "40 hours a month" needs real limit flags, and an earlier version of this skill documented none of them:

| Sheet wording | Flag |
|---|---|
| "N check-ins per month" | `--checkin-month-limit` |
| "N check-ins per week" | `--checkin-week-limit` |
| "N hours per month" | `--hours-month-limit` |
| "N hours per week" | `--hours-week-limit` |
| "N booking minutes per month / week" | `--booking-minute-month-limit` / `--booking-minute-week-limit` |
| Limits shared across a member's plans | `--checkin-price-plan-limit` / `--hours-price-plan-limit` |

Pair with `--use-time-passes true` when the allowance is delivered as passes instead (see Step 5). If a part-time plan has neither a usage limit nor time passes, it is functionally unlimited — flag that to the user rather than relying on the plan name.

#### Other plan flags worth knowing

- **Minimum spend:** `--minimum-price`, plus `--minimum-price-include-events` / `--minimum-price-include-extra-services` / `--minimum-price-include-time-passes` to choose what counts toward it.
- **Pausing:** `--can-be-paused`, `--pause-cycles-limit`, `--pause-yearly-limit`.
- **Invoicing cadence:** `--invoice-every-weeks` (weekly billing — distinct from `--invoice-every`, which is months), `--advance-invoice-cycles`, `--raise-invoice-every`, `--auto-raise-invoices`, `--exclude-from-invoice`, `--invoice-line-display-as`.
- **Discounts on extras:** `--discount-charges`, `--discount-extra-services`, `--discount-time-passes` — a plan giving "20% off meeting rooms" uses these, **not** a booking credit.
- **Virtual office plans** (`--is-virtual-office`): `--maximum-addresses`, `--maximum-recipients`, `--maximum-company-aliases`, `--maximum-delivery-storage-days`, the `--delivery-preferences-*` family (mail / parcels / checks / publicity / other) and the mail-handling product families (`--products-forward`, `--products-scan`, `--products-store`, `--products-shred`, `--products-collect`, `--products-recycle`, `--products-return`, `--products-deposit`, each with `--added-`/`--removed-` variants). A virtual-office sheet is mostly these flags — read `tariffs create --help` before quoting the work.
- **Identity / AML checks:** `--request-identity-check`, `--request-address-identity-check`, `--request-aml-check` and their provider/threshold/repeat-pattern companions, plus `--wait-for-identity-checks-to-activate` and `--keep-new-accounts-on-hold`. Only set when the user asks — they gate activation of new contracts.
- `--terms-and-conditions`, `--new-contract-document-url`, `--form-page-id`, `--send-onboarding-form-by-email` for signup paperwork.
- `--available-to-ai`, `--notes-for-ai`, `--price-for-ai`, `--show-price-for-ai` steer AI sales agents; `--starred` and `--display-order` affect portal presentation.

### Step 4 — Add signup fees / deposits

Signup fees **and** refundable deposits are both `tariffsignupproducts`. The difference is `--refundable`:
a **deposit** is `--refundable true`, a **joining/signup fee** is `--refundable false`. If the user calls the
product a "deposit" but it is named like a fee (or vice-versa), flag the mismatch and confirm before creating.

> **⚠️ Correction (v5.0.33):** `tariffs create --sign-up-fee` **no longer exists** — an earlier version of this skill (written against v5.0.23) recommended it for a plain flat joining amount. It is absent from v5.0.33 `--help`; passing it fails. **Always use `tariffsignupproducts`**, which is also the more expressive route (real product, refundability, deposits). `tariffsignupproducts create` does still have `--invoice-during-online-checkout` (whether the fee is charged during online checkout).

```powershell
nexudus tariffsignupproducts create `
  --caller claude --intent "Add signup fee to plan" `
  --tariff-id <tariff-id> `
  --product-id <signup-product-id> `
  --price 75 `                 # omit to inherit the product's own price
  --refundable false `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

**Before adding, check what the plan already has** — many plans already carry a signup product, and adding
another leaves the plan with **two** first-invoice charges. List first, then decide add vs replace:

```powershell
nexudus tariffsignupproducts list --json --tariff-id <tariff-id> --page-size 50 |
  ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, ProductId, ProductName, ProductPrice, Refundable
```

- To **change** refundability, use `tariffsignupproducts update <record-id> --refundable false` (no need to delete/recreate).
- To **replace** an existing fee: add the new one, then `tariffsignupproducts delete <old-record-id> --yes`.
- When swapping across many plans, delete by matching **`ProductId`** (the one you did *not* just add), so you can never
  delete the new record. Plans that had no prior fee simply get the add (nothing to delete) — handle that case.

### Step 5 — Add time passes

```powershell
nexudus tarifftimepasses create `
  --caller claude --intent "Add time pass to plan" `
  --tariff-id <tariff-id> `
  --time-pass-id <pass-id> `
  --passes-included 7 `
  --pass-renewal-time Week `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

`--pass-renewal-time` values: `Week`, `TariffMonth`, `CalendarMonth`, `Day`, `Year`

Common mappings: Day Pass × N/month → `TariffMonth` · Day Pass × N/week → `Week` · Hour passes → qty = total hours, renewal matches billing period

### Plans: Annual / upfront variants — billing vs commitment

"Annual membership" often means **monthly billing with a 12-month commitment**, *not* a yearly invoice. Confirm which the user means:

- **Monthly billing + 12-month term:** `--invoice-every 1 --default-contract-term 12` (the price in the sheet is the *monthly* price). Usually pair with `--disable-portal-cancellations true` so members can't self-cancel during the committed term.
- **True yearly invoice:** `--invoice-every 12` (price is the whole-year amount).

`--cancellation-period` is in **days**; `--default-contract-term` is in **months**. When the sheet gives a notice period in months, convert (e.g. "3 months" → `--cancellation-period 90`).

### Plans: two cancellation flags that are easy to confuse

A plans spreadsheet usually has **two separate rows** — "automatic cancellations" and "unpaid account cancellation" — that map to two different flags. Getting them the wrong way round silently sets a nonsense value, because both accept a bare integer.

| Sheet wording | Flag | Unit |
|---|---|---|
| "Automatic cancellations — when will the plan end" (e.g. "3 billing cycles") | `--auto-cancel-after` | **billing cycles** — contract is cancelled |
| "Unpaid account cancellation — days overdue before suspension" (e.g. "15 days") | `--cancel-member-account-after` | **days** — account suspended, contract left alone |
| "Notice period" (e.g. "30 days") | `--cancellation-period` | days |

An earlier version of this skill documented `--auto-cancel-after` as "days overdue before suspension" — that is **wrong**; it is billing cycles, and the overdue-suspension flag is `--cancel-member-account-after` (verified against v5.0.23 `--help`, 2026-09-28). Note `--cancellation-limit-days` also exists and is described as notice days — prefer `--cancellation-period` unless a readback shows otherwise.

**⚠️ `--cancel-member-account-after` reads back under a misspelled field: `CancelMemeberAccountAfter`** ("Memeber"). Selecting the correctly-spelled `CancelMemberAccountAfter` returns blank and looks like a failed write when the value is actually set. Add it to the known API-typo list alongside `ElegibleResourceTypes` and `CaneBeUsedForBookings` — when a readback field comes back unexpectedly empty, dump the whole record and grep the property names before concluding the write failed.

### Plans: prorating when the invoice day is not the 1st

Spreadsheets often ask for prorating on a plan whose invoicing day is mid-month (e.g. the 15th), then warn in their own notes that prorating "only works" on the 1st. Both are satisfiable — but it takes **three** flags, not two:

```powershell
--default-invoicing-day 15 `   # when the invoice is raised
--prorate-day-of-month 1 `     # align the prorate calculation to the 1st
--prorate-days-before 31 `     # how far back prorating starts — 31 covers a whole cycle
--prorate-cancellations true   # prorate the final invoice too
```

`--prorate-day-of-month` alone is **not enough**. It sets the alignment date, but `--prorate-days-before` defaults to `0`, which leaves a zero-width prorating window — so a member signing up mid-cycle is not prorated at all and the setting looks broken. Use `31` so the whole month qualifies. (Observed 2026-09-28: plans created with the prorate day set but no days-before had to be corrected afterwards.)

**Verify the whole prorate group by readback, not just the field you were thinking about.** In the same batch, one plan out of five came back with `ProrateCancellations=False` despite `true` being passed identically to the others — caught only because a later readback happened to include the column. Select every field you wrote, and diff all plans against each other; a single odd one out is the tell.
- `--use-time-passes true` is required on the plan itself for attached `tarifftimepasses` to govern check-in access; adding the pass records alone is not enough.

### Plans: Replicating an already-built template plan

The most common form of this task: *"Membership 01 is already created as a template — create the rest the same way."* **Derive the settings from the live record, not from your reading of the sheet.** The user has already resolved the sheet's ambiguities by building the template; your job is to copy its decisions, not re-make them.

1. **Read the template with `get`** and list its attachments separately (`tariffsignupproducts`, `tarifftimepasses`, `tariffextraservices`, `tariffbookingcredits` — each is its own entity, none appear on the tariff record).
2. **Check name collisions across the whole network before creating anything** (see Name uniqueness below), and identify which sheet rows already exist so you don't duplicate them. Match on trimmed, lower-cased names.
3. **Canary one plan with all its attachments**, then diff it field-by-field against the template, expecting only `Name`, `Price` and `SystemTariffType` to differ:
   ```powershell
   $fields = 'InvoiceEvery','DefaultInvoicingDay','ProrateCancellations','CancellationPeriod',
             'CancellationLimitDays','DefaultContractTerm','CancelMemeberAccountAfter','AutoCancelAfter',
             'DisablePortalCancellations','Visible','UseTimePasses','CurrencyId','BusinessId'
   foreach ($f in $fields) { if ("$($tpl.$f)" -ne "$($new.$f)") { "MISMATCH $f : $($tpl.$f) vs $($new.$f)" } }
   ```
4. **Then batch**, and verify every plan against the sheet by readback — price, each scalar field, attachment counts, and the credit value.

**Where the template contradicts the sheet, flag it and default to the template.** Observed 2026-10-07: a sheet promised a *"refundable 2-month deposit if the contract period is more than 3 months"*, while the template carried a `Deposit` product at **price 0 with `Refundable=False`** — non-refundable, zero-value, and conditional in a way the CLI cannot express (there is no conditional-deposit mechanism; `tariffsignupproducts` is unconditional). Both readings are defensible, so ask rather than silently picking. Note also that a minimum term shorter than the deposit's trigger condition means the condition can never fire — point that out.

**A zero-price deposit is often deliberate, not a mistake.** When a sheet describes a variable deposit ("2 months", "depends on term", "per negotiation"), the usual setup is a `Deposit` product attached at **price 0** as a placeholder that admins edit per contract — the amount lives on the member's contract, not the plan. So raise it once, accept "it's meant to be 0", and do not keep re-flagging it on later passes. The same reasoning applies to any plan-level amount a sheet marks as negotiable or "TBC".

**Watch for a "credits" unit system that is really money.** `tariffbookingcredits --credit` is a **currency amount**. A sheet saying *"2.5 credits per month, where 1 credit = 1 hour of function room"* is describing units, and storing `2.5` gives the member HKD 2.50 — not 2.5 hours. Surface the arithmetic (2.5 × the hourly rate) and let the user choose; do not convert silently in either direction.

**Infer `--system-tariff-type` from the plan name, and sanity-check it against the template's integer.** e.g. "Fixed desk in a shared room" + 24-hour access → `FullTimeDedicatedDesk` (reads back `3`); "Private office" → `FullTimePrivateOffice` (reads back `1`).

**Preserve names exactly as the sheet spells them**, including obvious typos (`Private office -4 pax` with a missing space). Report the typo so the user can decide — silently "fixing" a name creates a mismatch between their sheet and the portal, and renaming later is a one-line `tariffs update`.

### Plans: Name uniqueness

Plan names are unique **case-insensitively across the whole network**. `SOLO (6 months upfront)` collides with `Solo (6 months upfront)` and the second write fails with HTTP 400 `There is already a plan with this name in this network`. When bulk-renaming (e.g. stripping a `Copy of ` prefix), watch for near-duplicates that differ only in case and ask the user how to disambiguate.

### Plans: Rules

- Never add a deposit, benefits, or credits unless the user explicitly asks
- Always confirm the plan list before creating anything
- Ask about yearly variants when spreadsheet says "Monthly or Yearly"; clarify billing-cycle vs minimum-term for "annual/upfront" plans
- Use `--agent` on all commands for structured JSON output
- Verify every create/update by reading the tariff back — do not trust the exit code

### Plans: Batch pattern

```powershell
$bizId = <id>; $curId = <id>
$plans = @(
    @{ name="Plan A"; desc="..."; price=99;  visible="true";  type="PartTimeHotDesk" },
    @{ name="Plan B"; desc="..."; price=199; visible="false"; type="FullTimeHotDesk" }
)
foreach ($p in $plans) {
    $result = nexudus tariffs create --business-id $bizId --currency-id $curId `
      --name $p.name --description $p.desc --visible $p.visible --price $p.price `
      --invoice-every 1 --default-invoicing-day 1 --prorate-day-of-month 1 `
      --prorate-cancellations false --cancellation-period 30 --cancel-member-account-after 15 `
      --system-tariff-type $p.type --agent 2>&1 | ConvertFrom-Json
    Write-Host "$($p.name): $($result.ok) — $($result.summary)"
}
```

Same loop pattern works for `tariffsignupproducts` and `tarifftimepasses`.

---

<a name="resources"></a>
## Part 2: Resources

### Step 1 — Read the source file

Use the shared CSV/Excel reader above.

### Step 2 — Look up existing resource types

Resource types must already exist on the account — the CLI cannot create them.

```powershell
nexudus resourcetypes list
```

Map each resource's **Category** field to the closest resource type by name. Flag any mismatches and confirm with the user.

### Step 3 — Confirm the plan of action

Present a summary and wait for approval. Include:

- Resource name, resource type, seats/allocation, visible
- Max/min booking length, interval limit, advance booking limit, requires confirmation
- Repeat booking limits
- Amenities (Quiet Zone, Internet, Whiteboard, etc.)
- Linked resources (bidirectional pairs)
- Member/contact differentiated availability (e.g. members 24/7, contacts opening-hours) — set the flat window + a `resourceaccessrules` rule (Step 8)

### Step 4 — Create resources

Always pass `--business-id`. Omit `--allocation` entirely when spreadsheet says "n/a".

```powershell
nexudus resources create `
  --business-id <id> `
  --name "Room Name" `
  --resource-type-id <id> `
  --visible true `
  --requires-confirmation false `
  --allocation 4 `
  --allow-multiple-bookings false `
  --max-booking-length 120 `
  --min-booking-length 30 `
  --interval-limit 30 `                       # in MINUTES — must be a non-zero multiple of 15
  --book-in-advance-limit 336 `               # in HOURS (not days — 2 weeks = 336)
  --late-booking-limit 24 `                   # minimum lead time, in HOURS (decimals OK — 0.5 = 30 min)
  --repeat-booking-quantity-limit 12 `
  --repeat-booking-period-limit-in-months 3 `
  --description "..." `
  --quiet-zone true --internet true --white-board true `
  --desktop-monitor true --large-display true --display-screen true `
  --dual-display-screen true --video-conferencing true `
  --catering true --tea-and-coffee true --drinks true `
  --natural-light true --security-lock true --air-conditioning true --heating true
```

Key flag notes:
- `--allocation` — omit entirely for "n/a"; do not pass 0 or null
- `--book-in-advance-limit` — in **hours** despite help text saying "days"
- `--interval-limit` — in **minutes**, and must be a **non-zero multiple of 15** (30 = half-hour slots, 60 = hourly). The old "unit unconfirmed, possibly hours" note was wrong; confirmed by readback on v5.0.23.
- `--allow-multiple-bookings true` — for hot desks with multiple concurrent users. **It also changes what `--allocation` means:** when true, allocation is the maximum number of *concurrent bookings*; when false, allocation is informational physical capacity only and does not limit anything. A sheet that says "1 seat" but "10 customers at a time" needs `--allocation 10`, not 1 — flag the contradiction to the user.
- `--requires-confirmation true` — bookings held pending admin approval
- Member/non-member differentiated **advance limits** are not settable on the resource itself, but **are** settable as a `resourceaccessrules` rule carrying its own `--book-in-advance-limit` (see Step 8) — don't send the user to the admin UI for this
- `--system-resource-type` — distinct from `--resource-type-id`; it's the built-in category that steers AI agents and system behaviour. Valid values: `MeetingRoom`, `HotDesk`, `PrivateOffice`, `EventSpace`, `Lab`, `Kitchen`, `TreatmentRoom`, `StorageUnit`, `Machine`, `DayPass`, `PhoneBooth`, `Other` (or empty). Reads back as an integer.
- `--limit-visitors-to-allocation` — caps total visitors linked to a booking at the resource's allocation. Pass it when a sheet gives an event space a hard headcount.
- `--only-for-invoicing-business` — restricts booking to customers invoiced by this business. Distinct from `--only-for-members` / `--only-for-contacts`; relevant in a multi-location network.
- **Cancellation fees** are a five-flag family: `--charge-cancellation-fee`, `--cancellation-fee-type`, `--cancellation-fee-amount`, `--cancellation-fee-percentage`, `--cancellation-fee-product-id`. A sheet saying "50% charge if cancelled inside 24h" needs these *plus* `--late-cancellation-limit` (minutes) to define the cut-off.
- `--booking-availability-exceptions` (and `--added-`/`--removed-` variants) — closure days and holiday exceptions, separate from `resourcetimeslots`. Use this for "closed on public holidays"; don't try to model it by deleting weekday slots.
- `--google-calendar-id` / `--office365-calendar-id` — two-way calendar sync per resource. `--hide-in-calendar` keeps a resource out of the admin calendar view.
- **Don't force `--display-order`** on create; see the Part 9 note — the system reassigns it.

#### The full amenity flag set (v5.0.33)

Part 2's example above lists only the common ones. The complete set, so a spreadsheet's amenity column can be mapped without guessing:

`--quiet-zone` · `--internet` · `--white-board` · `--flip-chart` · `--projector` · `--display-screen` · `--dual-display-screen` · `--large-display` · `--desktop-monitor` · `--wireless-presentation` · `--video-conferencing` · `--conference-phone` · `--standard-phone` · `--pa-system` · `--voice-recorder` · `--wireless-charger` · `--standing-desk` · `--privacy-screen` · `--soundproof` · `--natural-light` · `--air-conditioning` · `--heating` · `--security-lock` · `--secure-storage` · `--cctv` · `--catering` · `--tea-and-coffee` · `--drinks`

All are booleans. The ones an earlier version of this skill omitted — `--flip-chart`, `--projector`, `--wireless-presentation`, `--conference-phone`, `--standard-phone`, `--pa-system`, `--voice-recorder`, `--wireless-charger`, `--standing-desk`, `--privacy-screen`, `--soundproof`, `--secure-storage`, `--cctv` — are exactly the ones event-space and meeting-room spreadsheets list most often. Check the sheet against this full list before telling the user an amenity isn't settable.
- `--shifts` — fixed bookable blocks, as a comma-separated string of `NAME: START->END` in 24-hour time, e.g. `MORNING (8.30AM - 12.30PM): 08:30->12:30, FULL DAY (9AM - 5PM): 09:00->17:00`. **The shift name must not contain a colon** or parsing breaks — write labels as `8.30AM`, never `8:30AM`. Use this for resources a spreadsheet describes as "bookable only in shifts".

#### Booking-limit units are NOT consistent — check every one

| Flag | Unit |
|---|---|
| `--max-booking-length` / `--min-booking-length` | minutes |
| `--interval-limit` | minutes (non-zero multiple of 15) |
| `--late-cancellation-limit` | minutes |
| `--no-return-policy*` | minutes |
| `--book-in-advance-limit` | **hours** |
| `--late-booking-limit` | **hours** (decimals accepted — `0.5` = 30 min, reads back `0.50000`) |

Spreadsheets state these in mixed units ("1 year", "24 hours", "30 days", "30 minutes"), so convert per-flag rather than per-sheet: 1 year → `8760`, 1 week → `168`, 30 days → `43200`, 4 hours → `240`. Confirmed against v5.0.23 `--help` and verified by readback (2026-09-28).

A resource type can live on a **different business** to the resource that uses it — `resources create` accepts a `--resource-type-id` owned by another location in the network and resolves the name correctly. Worth flagging to the user, since rates attach to the type and will therefore be defined under the other location.

### Step 5 — Link interconnected resources

Do this after both resources exist (you need both IDs). Linking is **one-directional** per record — to make a
lock mutual you must update **both** resources:

```powershell
nexudus resources update <id-A> --added-linked-resources <id-B>
nexudus resources update <id-B> --added-linked-resources <id-A>
```

- Confirmed one-directional on v5.0.23: after linking A→B, an immediate `get` on B shows `LinkedResources` still empty. (A reciprocal link may appear on B some minutes later, so **never** infer the direction from a later readback — always write both sides explicitly. Writing a link that already exists is harmless.)
- **⚠️ Verify links with `get`, never `list`.** The `resources list` DTO returns `LinkedResources` as **null** regardless of the real value; only `resources get <id>` returns the populated array. In PowerShell this is a live trap: `@($null).Count` is **1**, so a `list`-based emptiness check reports every unlinked resource as linked. Same applies to `LinkedResourceIds`, which is empty in the list DTO.
- A divisible room is modelled parent/child: parent links to **every** child, each child links back to the parent, and siblings are **not** linked to each other. When the same physical room is sold as two resources (e.g. a daytime and an "Out Of Hours" variant), link those twins to each other too — otherwise both can be booked at once. Extend the map so every resource representing a given physical space blocks every other one that overlaps it.

### Step 6 — Member / contact restrictions

Resources can be limited to members or contacts via real flags (contrary to older notes, this **is** settable via CLI):

- `--only-for-members true` — only active members with a plan can book
- `--only-for-contacts true` — only contacts (non-member customers) can book

Watch for a contradiction common in spreadsheets: a resource marked "members only" **and** given non-member (contact)
pricing. If `--only-for-members true` is set, any contact rate you create for that resource type can never fire — flag it.

### Step 7 — Set availability (opening hours / time slots)

Availability windows are a **separate entity** (`resourcetimeslots`), one record per resource per weekday — *not* a field on the resource:

```powershell
nexudus resourcetimeslots create `
  --caller claude --intent "Set Mon-Fri 9-5 availability" `
  --resource-id <resource-id> `
  --from-time "1976-01-01T09:00:00" `   # date component is always 1976-01-01 (a placeholder)
  --to-time   "1976-01-01T17:00:00" `
  --day-of-week Monday `                 # name in; stored as int (Sunday=0 … Saturday=6)
  --agent
```

- **Times are stored as UTC and displayed in the LOCATION's timezone — always offset for the location.** The value is saved with a trailing `Z` on the `1976-01-01` placeholder date, and the portal/admin renders it converted to the business's timezone using that zone's **standard-time** (winter) offset. So pass `desired local time − the location's standard UTC offset`.
  - Get the zone from `nexudus whoami` → `DefaultSimpleTimeZoneNameIana`.
  - **UTC+0** (e.g. Europe/London): store the wall-clock as-is — `07:00` displays `07:00`.
  - **UTC+1** (e.g. Europe/Berlin, CET): store `06:00` to display `07:00`. (Observed: storing `07:00` showed `08:00`.)
  - **UTC−5** (e.g. America/New_York, EST): store `13:00` to display `08:00`, `15:00` to display `10:00`. A **negative** offset means you **add** hours — and an evening end time wraps past midnight: 8pm → store `01:00` on the **next day** (`--to-time "1976-01-02T01:00:00"`). The API accepts the next-day end; keep `--day-of-week` on the local start day. (Confirmed working: NY 8am–8pm stored `1976-01-01T13:00` → `1976-01-02T01:00`, day Monday.)
  - Always **verify in the admin calendar** — that's the real proof the display is right. Canary one slot and have the user confirm before batching the rest, especially for non-UTC-0 locations.
- Loop over `Monday`…`Friday` (and Saturday/Sunday if needed) to build a full week. Verify with `resourcetimeslots list --resource-id <id>`.
- A window that genuinely crosses midnight in local time (e.g. a 19:00–02:00 night shift) needs splitting into two slots. (Distinct from the storage-offset wrap above, which is a single daytime window whose UTC end merely lands after midnight.)
- `resourcetimeslots update <id> --from-time … --to-time …` edits a slot in place — no need to delete/recreate to fix times.
- **`resources update` does NOT wipe time-slots (v5.0.23).** A scalar-field update — `--min-booking-length`, `--shifts`, amenities, any single field — leaves the resource's `resourcetimeslots` collection **intact**. Verified `slots before == after` on 2026-09-28, both as a single-field canary and across a batch of shift updates on slotted resources. Collection-add-only updates (`--added-tariffs`, `--added-linked-resources`, `--added-teams`) are equally safe. **An earlier version of this skill warned that scalar updates silently wiped the slots — that is no longer true.** There is no need to back up slots before an update, and no need to order scalar edits before slot creation. Still spot-check `slots before == after` after a bulk update, in line with the general readback rule.
- **"Always" availability = no time-slots.** To make a resource bookable 24/7, delete its `resourcetimeslots` (empty = always available) — same encoding as the access-rule "Always". Back them up first if you'll need to restore.

### Step 8 — Member/contact differentiated availability (`resourceaccessrules`)

`resourcetimeslots` on a resource is a **flat** window applied to everyone. When members and contacts need *different* availability — e.g. "open to all during opening hours, but **members can book 24/7**" — set the resource's flat time-slots to the general (opening-hours) window, then add a **resource access rule** for the extra member access. This **is** doable via CLI (`resourceaccessrules`); it is not admin-UI-only.

```powershell
nexudus resourceaccessrules create `
  --caller claude --intent "Members 24/7 access" `
  --business-id <id> `
  --name "Members 24/7 Access" `
  --active true `
  --only-for-members true `                 # or --only-for-contacts true
  --added-resources <res-id> --added-resources <res-id> `   # repeat per resource
  --agent
```

- **"Always" (24/7) = leave `--time-slots` empty.** An empty time-slots list means no time restriction; *adding* `--time-slots` would restrict the rule to those windows. So for 24/7, omit the flag entirely.
- **⚠️ There are TWO time-slot collections on `resourceaccessrules` — `--time-slots` and `--eligible-time-slots`.** The "Always = leave empty" rule above applies to `--time-slots`; if a rule behaves unexpectedly, check whether `--eligible-time-slots` needs setting too. Both are documented with their exact help text in "Access rules can do much more…" below.
- Restrict scope with `--only-for-members` / `--only-for-contacts`, and attach resources with repeated `--added-resources`.
- `--active true` turns the rule on. Verify with `resourceaccessrules get <id>` (shows `Resources`, `OnlyForMembers`, `Active`); confirm the "Always" limit in the admin UI since empty time-slots is the encoding.
- The rule layers on top of the resource's own time-slots: contacts get the opening-hours window, members get the rule's (Always) access.

#### Access rules can do much more than members-vs-contacts (v5.0.33)

`resourceaccessrules create` has 66 flags. The two time-slot collections, confirmed:

- **`--time-slots`** — *"the days and times the resources can be booked when this rule applies"* (what the rule grants).
- **`--eligible-time-slots`** — *"time slots defining when this rule applies (eligibility windows)"* (when the rule is in force). Note this one is spelled **`eligible`** correctly, unlike `tariffbookingcredits`' `--elegible-*`.

**Who the rule applies to** — beyond `--only-for-members` / `--only-for-contacts`:

| Flag | Meaning |
|---|---|
| `--allowed-tariffs` (repeated) | only customers on these plans may book when the rule applies |
| `--allowed-teams` (repeated) | only members of these teams may book |
| `--members` (repeated) | specific customers — the rule fires only for them |
| `--courses` (repeated) | only customers who completed one of these courses (induction-gated machinery, labs) |

**When the rule fires:**

| Flag | Meaning |
|---|---|
| `--apply-rule-from` / `--apply-rule-to` | date window in which the rule is evaluated at all |
| `--apply-if-within-minutes-of-start` | only for bookings made close to their start time |
| `--apply-if-more-than-minutes-from-start` | only for bookings made well in advance |
| `--apply-if-price-min` / `--apply-if-price-max` | only for bookings in a price band |

**Rule ordering matters when a resource has more than one rule:** `--evaluation-order` (lower = evaluated first) and `--stop-evaluation-if-rule-is-met` (halt after this one matches). Without these, overlapping rules on the same resource resolve unpredictably — set them explicitly whenever you create a second rule on a resource. `--reject-with-message` sets the text the member sees when the rule blocks their booking; always set it, or they get a bare rejection.

A rule can also carry its own `--book-in-advance-limit`, `--no-return-policy-all-resources` / `--no-return-policy-all-users`, and the full cancellation-fee family — so **member-vs-contact differentiated advance limits, which Part 2 says the CLI cannot set on the resource, CAN be set here** as a rule. That supersedes the "flag to user, set in admin UI" note in Part 2 Step 4.

### Temporarily relaxing limits for a booking import

A common request: relax a resource's booking limits so an import doesn't get rejected, then **put them back exactly as they were** afterwards. The limits that typically block imports are `MinBookingLength`, `IntervalLimit`, `BookInAdvanceLimit`, `LateBookingLimit`, and sometimes `RequiresConfirmation`. Availability time-slots also block bookings outside their windows (relax by making availability "Always" = deleting the slots).

Workflow:

1. **Back up two things separately** — the resource records (all fields) **and** their `resourcetimeslots` (a separate entity, *not* in the resource JSON). Save both to files.
   ```powershell
   $res | ConvertTo-Json -Depth 10 | Out-File resources-backup.json          # fields
   # per resource: resourcetimeslots list --resource-id <id>  ->  timeslots-backup.json
   ```
2. **Confirm the backup matches current state** before changing anything, so the restore target is trustworthy.
3. **Relax** the blocking limits (e.g. `--min-booking-length 0 --interval-limit <as needed>`). Ask the user which values — "relax" can still mean a specific interval (e.g. 30), not 0.
4. User runs the import.
5. **Restore**: set the fields back from the resource backup, and re-create any time-slots you deleted in step 3 from the time-slot backup. (The `resources update` itself does not disturb surviving slots — see the note above — so the slot backup only matters when you relaxed availability by deleting slots.)
6. **Verify fields *and* slot counts** against both backups by readback.

### Resources: Rules

- Always pass `--business-id` — omitting it causes a 400 error
- Always confirm the resource list before creating anything
- Omit `--allocation` entirely when seats = "n/a"
- Create linked pairs last so both IDs are available
- Member/contact differentiated availability is done with a **`resourceaccessrules`** rule (Step 8), not the admin UI — set the flat time-slots to the shared window and add a rule for the extra access

### Resources: Batch pattern

```powershell
$bizId = <id>
$resources = @(
    @{ name="Grace";  typeId=1415315728 },
    @{ name="Ubuntu"; typeId=1415315728 },
    @{ name="Equity"; typeId=1415355828 }
)
foreach ($r in $resources) {
    $result = nexudus resources create `
      --business-id $bizId --name $r.name --resource-type-id $r.typeId `
      --visible true ... --agent 2>&1 | ConvertFrom-Json
    Write-Host "$($r.name): $($result.ok)"
}
```

---

<a name="credits"></a>
## Part 3: Booking Credits

**Two different kinds of booking credit — pick the right mechanism:**

| Kind | What the sheet says | Mechanism | Command |
|---|---|---|---|
| **Time credit** | "5 hours a month", "240 minutes" | allowance of a booking-credit **extra service** (minutes) | `tariffextraservices` (Steps 1–4 below) |
| **Money credit** | "$175 in credits for the meeting rooms" | a **tariff booking credit** benefit on the plan (a dollar allowance) | `tariffbookingcredits` (see "Money credits" below) |

If a sheet's "booking credit" cell gives a **dollar amount** (e.g. "$200 in credits"), it's a **money credit** → use `tariffbookingcredits`, **not** a time-credit extra service. (Spreadsheets often mislabel this — confirm with the user when a credit is expressed in money.)

Time credits (this section) are stored in **minutes** (hours × 60). They renew on the plan's billing cycle (`TariffMonth`) or on the 1st of each calendar month (`CalendarMonth`).

### Prerequisites

1. All booking-credit extra services exist on the account
2. All plans (tariffs) exist
3. No duplicate credits exist — run `nexudus tariffextraservices list` to check
4. Wrong/old credits are deleted before recreating

### Step 1 — Gather IDs

```powershell
nexudus extraservices list --json | ConvertFrom-Json |
  Where-Object { $_.IsBookingCredit -eq $true } |
  Select-Object Id, Name

nexudus tariffs list --json
```

### Step 2 — Confirm the summary table

Show the user this table before creating anything:

| Plan | Extra Service | Hours | Minutes | Renewal |
|---|---|---|---|---|
| Coworking Basic Pass | Phone Booth Credits | 4 | 240 | TariffMonth |
| Coworking Premium - Annual | Phone Booth Credits | 4 | 240 | CalendarMonth |

Verify: monthly plans → `TariffMonth` · annual plans → `CalendarMonth` · hours correctly ×60

### Step 3 — Create via batch loop

```powershell
$entries = @(
  @{ tid=1415386327; esid=1415369978; uses=240; ren="TariffMonth" },
  @{ tid=1415386335; esid=1415369978; uses=240; ren="CalendarMonth" }
)
$ok = 0
foreach ($e in $entries) {
  $raw = nexudus tariffextraservices create `
    --caller claude --intent "Add time credits from spreadsheet" `
    --tariff-id $e.tid --extra-service-id $e.esid `
    --uses-included $e.uses --service-renewal-time $e.ren 2>&1

  # Known CLI quirk: parse error on ExtraServiceChargePeriod is COSMETIC — record is created anyway
  $ok++
  Write-Host "OK | tariff $($e.tid) | esid $($e.esid) | $($e.uses) min | $($e.ren)"
}
Write-Host "Done: $ok / $($entries.Count) submitted"
```

**Known CLI quirk:** `Error: The JSON value could not be converted to System.String. Path: $.ExtraServiceChargePeriod` — this is cosmetic. The API creates the record successfully; the CLI just fails to deserialize the response. Treat as success and verify in the admin panel.

### Step 4 — Verify in admin panel

Navigate to each plan → **Credits** or **Extra Services** tab. Confirm `UsesIncluded` (minute values) and `ServiceRenewalTime` are correct. Spot-check 3–5 plans.

### Credits: `--service-renewal-time` values

| Value | Use case |
|---|---|
| `TariffMonth` | Monthly plans — renews on billing date |
| `CalendarMonth` | Annual plans — renews on the 1st of each month |
| `Week` | Weekly credit packages |
| `Day` / `Year` | Rare |

To correct entries: delete first, then recreate — don't attempt updates.
```powershell
nexudus tariffextraservices delete <tariffextraservice-id>
```

### Credits: Rules

- Convert hours to minutes (× 60) before passing to CLI
- Monthly plans → `TariffMonth`; annual plans → `CalendarMonth`
- Always confirm the summary table first
- Ignore the CLI parse error — cosmetic bug, records are created
- Verify in admin panel after; delete then recreate for corrections

---

### Money credits (`tariffbookingcredits`)

A **money credit** — a dollar allowance a member spends on bookings (e.g. "$200/month for the meeting rooms") — is a **tariff booking credit** attached directly to the plan, created via `nexudus tariffbookingcredits create`. This is a different entity from time-credit extra services; it does **not** use `tariffextraservices` and needs no pre-existing extra service.

```powershell
nexudus tariffbookingcredits create `
  --caller claude --intent "Add money booking credit" `
  --tariff-id <tariff-id> `
  --name "Resource Credits" `                 # the benefit's display name
  --credit 200 `                              # dollar allowance
  --can-be-used-for-bookings true `
  --elegible-resource-types 1415373508 `      # repeat the flag per resource type (note API spelling "elegible")
  --elegible-resource-types 1415373509 `
  --elegible-resource-types 1415373606 `
  --service-renewal-time TariffMonth `
  --agent
```

- **Restrict to specific rooms** with repeated `--elegible-resource-types` (leave empty = all bookable resources). It's `elegible`, the API's spelling — not `eligible`. Confirmed still misspelled in v5.0.33, across every variant: `--elegible-resource-types`, `--elegible-products`, `--elegible-passes`, `--elegible-tariffs`, and their `--added-`/`--removed-` forms.
- `--can-be-used-for-bookings true` scopes it to bookings. Other scopes exist: `--can-be-used-for-events` (+`--event-categories`), and `--is-universal-credit` (products/passes/charges, restrict with `--elegible-products` / `--elegible-passes` / `--elegible-tariffs` / `--applies-to-charges`). The `--elegible-*` restrictions other than resource types apply **only when `--is-universal-credit` is true** — setting them on a bookings-only credit does nothing.
- **Verify with `get`, not `list`** — the `list` DTO omits `ElegibleResourceTypes` (and reads the toggle back under the typo'd field `CaneBeUsedForBookings`). `nexudus tariffbookingcredits get <id> --json` returns the eligible types and confirms the credit is scoped correctly.
- Each plan gets its own booking credit (it's a per-tariff benefit) — no shared service to reuse.
- **Printing credit** is yet another variant: an extra service created with `--printing-credit true` (linked to a Printing resource type). Add the allowance to a plan with `tariffextraservices` (uses = dollar amount when the service is priced $1/unit). If the printing extra service is reconfigured, remove and re-add the plan's `tariffextraservices` link so it picks up the change.

---

<a name="rates"></a>
## Part 4: Resource Rates

Resource rates are created as **extra services** linked to resource types via `nexudus extraservices create`. They are NOT products and do NOT use `resourceproducts`. Per-tariff price overrides are set separately via `nexudus extraserviceprices create`.

### Naming convention

```
[Resource Type Name] - [Rate Label]
[Resource Type Name] - [Rate Label] - Membres    # member-only
[Resource Type Name] - [Rate Label] - Contacts   # contact-only
```

Examples: `Camion - Demi-journée`, `CNC - Heure`, `Laser - Heure - Membres`

### One rate or two? (member vs contact)

Only create a **member/contact pair** when the two prices actually differ. If member price == contact price,
create a **single unrestricted rate** (no `--only-for-members` / `--only-for-contacts`) — it applies to everyone.

- Different prices → two rates: one `--only-for-members true --price <member>`, one `--only-for-contacts true --price <contact>`.
- Same price → one rate, no restriction flag.

### Rate is on the resource TYPE, not the instance

A rate links to a resource **type** (`--resource-types <type-id>`), so it applies to **every resource sharing that type**.
If two resources share a type (e.g. "Office Desk (front)" and "Office Desk (back)"), one rate prices both — you cannot
price them separately without giving one its own resource type. Flag this to the user before creating.

Resource types **can** be created via CLI (v5.0.23) — an earlier version of this skill said admin-UI only:

```powershell
nexudus resourcetypes create --caller claude --intent "Create resource type" --business-id <id> --name "Buzz OOH" --agent
nexudus resources update <resource-id> --resource-type-id <new-type-id>   # re-point; leaves slots/links/limits intact
```

Common case: a room sold as a day variant and an "Out Of Hours" variant with different pricing models (shifts vs hourly). Give the OOH variants their own `<Type> OOH` types up front so their rates never overlap. The type can live on a sibling business (see Part 2).

### Charge periods: hourly vs flat vs per-use

- **Per hour** → `--charge-period Minutes` with `--price` = the hourly amount (multiplies by booking duration; no cap).
- **Flat price per booking** → **`--charge-period Minutes` with `--price` = `--maximum-price` = the flat amount** (optionally bounded by `--min-length` / `--max-length` to make it a half-day/full-day band). This is the preferred way to build a flat price — see Steps 4–5. **Do not reach for `--charge-period Uses`** unless the user specifically asks for it.
- **`--charge-period Uses`** exists (flat per booking, `0` price valid) but is *not* the default choice for flat pricing — the Minutes + cap form composes with length bands and reads correctly in the booking UI.
- **Never monthly** — see the rules below.

#### Alternative: `--fixed-cost-length` + `--fixed-cost-price`

A genuine "first N minutes cost a flat amount, then bill hourly" structure, which the cap pattern cannot express:

```powershell
--charge-period Minutes --price 150 `   # hourly rate after the initial block
--fixed-cost-length 240 `               # first 4 hours…
--fixed-cost-price 500                  # …cost a flat 500
```

The two flags **require each other** — passing one alone fails. Use this when a sheet says "HKD 500 for the first 4 hours, then 150/hour"; use `--price` = `--maximum-price` when it says a flat session price with no overage.

#### `--min-length` / `--max-length` are in the SELECTED charge period

Help text: *"Optional minimum booking length expressed in the selected charge period."* With `--charge-period Minutes` they are minutes (so `--max-length 360` = 6 hours, as in Steps 4–5). With `Days` they are **days**, with `Weeks` **weeks**. An earlier version of this skill treated them as always-minutes — correct only because every example uses `Minutes`. Convert per charge period, not per sheet.

### ⚠️ Time-of-day scoped rates: `--time-slots`

**`extraservices create` accepts `--time-slots`** — *"the days and times this extra service price is available for booking"*, with the same `1976-01-01` placeholder date as `resourcetimeslots`. This is the mechanism for the extremely common **"Office Hours rate vs After Hours rate"** split, and an earlier version of this skill did not document it at all.

Without `--time-slots`, two rates on the same resource type both match every booking and the system picks one — so an "After Hours" rate priced higher than "Office Hours" may simply never fire, or may fire during the day. **If a sheet splits a rate by time of day, the split must be encoded with `--time-slots` on each rate; the name alone does nothing.**

- **⚠️ Verify rate time-slots with `get`, never `list`.** `extraservices list` **omits `TimeSlots` entirely**, so a correctly time-scoped rate looks completely unscoped in list output. Observed 2026-10-07: a 26-rate account was reported to the user as "not time-scoped at all" on the strength of a `list` readback; `extraservices get <id>` showed every rate properly scoped. Add this to the list-DTO omission family alongside `resources.LinkedResources` and `tariffbookingcredits.ElegibleResourceTypes` — **never conclude a collection field is empty from `list`.**
- When checking the offset, convert the stored UTC back to local before judging it: `([datetime]$slot.FromTime).AddHours(<location offset>)`. A stored `01:00Z` is correct for a 09:00 local opening at UTC+8 — it only looks wrong if you read the raw value.
- `--time-slots` also accepts `@filepath` input (it is the **only** flag that does — see the shared command-line-length rule).
- When several rates can still match the same booking, **`--default-price true`** marks which one wins: *"whether this rate is preferred when multiple valid rates match the same resource and charge period."* Set it on exactly one rate per type/period.
- Verify the split in the admin UI by pricing a test booking inside and outside the window — a readback of the slot rows proves the data, not the behaviour.

### Other rate-scoping flags (v5.0.33)

| Flag | Purpose |
|---|---|
| `--tariffs` (repeated) | restrict the rate to members on specific plans; empty = no plan restriction. Different from `extraserviceprices`, which *overrides the price* per plan rather than gating access |
| `--teams` (repeated) | restrict to specific teams |
| `--credit-price` | the price charged **in credit units** when the customer pays with booking credit; null = base price. This is how a "credits" pricing scheme is implemented — see below |
| `--apply-from` / `--apply-to` | date validity window (inclusive / exclusive) — seasonal or promotional pricing. **Dates, not times of day** — use `--time-slots` for time of day |
| `--apply-to-visitors` | multiplies the charge by the number of visitors on the booking — use for per-head event pricing |
| `--per-night-pricing` | day-based length counted by nights rather than elapsed 24h periods |
| `--invoice-display` | custom invoice line text instead of the rate name |
| `--price-factor-low-demand` / `-average-demand` / `-high-demand` | signed % demand-based adjustment (`10` = +10%, `-10` = −10%) |
| `--price-factor-last-minute` + `--last-minute-period` + `--last-minute-adjustment-type` | last-minute pricing; period in minutes before start, type `Fixed` (full factor throughout) or `Gradual` (ramps from zero) |

Demand and last-minute pricing are **not** standard onboarding — only set them when the user asks, and flag that they override the base price.

### A "credits" pricing scheme = `tariffbookingcredits` + `--credit-price` on every rate

When a space prices bookings in **credits** ("2.5 credits a month, 1 credit per hour of function room, 0.5 credit per hour of meeting room"), it takes **two** pieces that are easy to half-build:

1. The **allowance** on each plan — `tariffbookingcredits --credit <n>` scoped with `--elegible-resource-types` (Part 3).
2. The **price in credits** on each rate — `extraservices --credit-price <n>`, set on **every rate attached to those eligible resource types**.

Miss step 2 and the credit is spent at the rate's **cash** price instead, draining a 2.5-credit allowance on the first booking. Build the set from the credit's `ElegibleResourceTypes`, not from rate names:

```powershell
$credTypes = @('<type-id>','<type-id>')          # from tariffbookingcredits get
foreach ($id in $allRateIds) {
  $d = nexudus extraservices get $id --json | ConvertFrom-Json | Select-Object -ExpandProperty data
  $t = @($d.ResourceTypes) | ForEach-Object { "$_" }
  if (($t | Where-Object { $credTypes -contains $_ }).Count -gt 0) { <# set --credit-price #> }
}
```

**Verify both directions** — every eligible rate has a credit price, and no *ineligible* rate picked one up (a credit price on a type the credit cannot reach is dead config that misleads the next person).

#### ⚠️ A members-only credit cannot be spent on a contacts-only rate

`--only-for-contacts` means *"restricted to contacts without an active contract"*; `--only-for-members` means *"customers with an active contract"*. The two are mutually exclusive. Booking credits are a **plan** benefit, so only members ever hold them.

**So every credit-eligible resource type needs at least one rate a member can match** — either `--only-for-members` or unrestricted. A type whose only rate is `--only-for-contacts` is unreachable for members: their booking matches no rate at all and the credit can never be charged, no matter how correct `--credit-price` looks on that rate.

Observed 2026-10-07: two meeting-room types had contacts-only rates while 15 plans carried credits scoped to them — setting `--credit-price` on those rates changed nothing for members. The fix is a parallel `- Members` rate per type, mirroring the contacts rate and adding `--only-for-members true`. **Audit this whenever you set up credits:**

```powershell
# every eligible type must have a rate where Who != Contacts
$rows | Group-Object Type | ForEach-Object {
  if (-not ($_.Group | Where-Object { -not $_.OnlyForContacts })) { "NO MEMBER-USABLE RATE for type $($_.Name)" }
}
```

**When the sheet gives no member cash price**, that is often deliberate — members are meant to book with credits only. Setting the member rate to **0** is wrong: it lets members book free once credits run out. Default to the **same cash price as the contacts rate**, so credits are the benefit and overage falls back to the standard price, and say so.

### Step 1 — Look up required IDs

```powershell
nexudus whoami                          # → business ID, currency ID
nexudus resourcetypes list --json       # → resource type IDs
```

### Step 2 — Confirm rate plan with user

Present a summary table before creating anything:

| Rate name | Resource Type ID | Charge period | Price | Cap | Length |
|---|---|---|---|---|---|
| CNC - Heure | 1234 | Minutes | €25 | none | none |
| Camion - Demi-journée | 5678 | Minutes | €30 | €30 | ≤ 360 min |
| Camion - Journée | 5678 | Minutes | €50 | €50 | > 360 min |

Key questions: Are any rates restricted to members or contacts only? Are per-plan price overrides needed?

### Step 3 — Hourly rate (genuine per-hour billing)

Price multiplies with booking duration (e.g. €25/h → 3h = €75). No cap, no length restriction.

```powershell
nexudus extraservices create `
  --caller claude --intent "Create hourly rate" `
  --business-id <id> --currency-id <id> `
  --name "CNC - Heure" `
  --resource-types <resource-type-id> `
  --charge-period Minutes `
  --price 25 `
  --visible true `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

### Step 4 — Half-day rate (flat cost, up to 6 hours)

```powershell
nexudus extraservices create `
  --caller claude --intent "Create half-day rate" `
  --business-id <id> --currency-id <id> `
  --name "Camion - Demi-journée" `
  --resource-types <resource-type-id> `
  --charge-period Minutes `
  --price 30 `
  --maximum-price 30 `   # same as --price to enforce flat cost
  --max-length 360 `     # 6 hours in minutes
  --visible true `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

### Step 5 — Full-day rate (flat cost, over 6 hours)

```powershell
nexudus extraservices create `
  --caller claude --intent "Create full-day rate" `
  --business-id <id> --currency-id <id> `
  --name "Camion - Journée" `
  --resource-types <resource-type-id> `
  --charge-period Minutes `
  --price 50 `
  --maximum-price 50 `   # same as --price to enforce flat cost
  --min-length 361 `     # kicks in just over 6 hours
  --visible true `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

### Step 6 — Per-tariff price overrides (optional)

```powershell
nexudus extraserviceprices create `
  --caller claude --intent "Override rate price for tariff" `
  --extra-service-id <extra-service-id> `
  --tariff-id <tariff-id> `
  --price 20 `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

**New (v5.0.23):** `extraserviceprices create` also accepts `--maximum-price` — a per-tariff cap for time-based rates. Use it to give one plan a different flat-cap than the base rate (e.g. base rate is uncapped but this plan caps a half-day at €25).

### Resource Rates: Batch pattern

```powershell
$bizId = <id>; $curId = <id>
$rates = @(
    @{ name='CNC - Heure';           typeId=<id>; price=25; max=$null; maxLen=$null; minLen=$null },
    @{ name='Camion - Demi-journée'; typeId=<id>; price=30; max=30;   maxLen=360;   minLen=$null },
    @{ name='Camion - Journée';      typeId=<id>; price=50; max=50;   maxLen=$null; minLen=361   }
)
foreach ($r in $rates) {
    $cmd = @(
        'extraservices', 'create',
        '--caller', 'claude', '--intent', 'Create resource rate',
        '--business-id', $bizId, '--currency-id', $curId,
        '--name', $r.name, '--resource-types', $r.typeId,
        '--charge-period', 'Minutes', '--price', $r.price,
        '--visible', 'true', '--agent'
    )
    if ($r.max)    { $cmd += '--maximum-price'; $cmd += $r.max }
    if ($r.maxLen) { $cmd += '--max-length';    $cmd += $r.maxLen }
    if ($r.minLen) { $cmd += '--min-length';    $cmd += $r.minLen }
    $result = & nexudus @cmd 2>&1 | ConvertFrom-Json
    Write-Host "$($r.name): $($result.ok) — $($result.summary)"
}
```

### Resource Rates: Rules

- **Basically never create monthly rates.** Monthly pricing in Nexudus is done through other mechanisms (plans/tariffs, time passes, extra-service renewals) — **not** resource rates. Do not reach for `--charge-period Months` on a rate; if a user asks for "monthly" pricing on a resource, clarify and route it to a plan/pass instead. (`--charge-period Months` on a resource extra service also does not function correctly.)
- Only split into member + contact rates when the prices differ; equal prices → one unrestricted rate
- **Flat price** → `--charge-period Minutes` with `--price` = `--maximum-price`; **per hour** → `--charge-period Minutes` with the hourly price. Use `--charge-period Uses` only if the user asks for it
- **A rate split by time of day needs `--time-slots`** — "Office Hours" / "After Hours" in the name does nothing on its own
- "Every hour" in the Nexudus UI maps to `Minutes` in the API. Valid `--charge-period` values: `Minutes`, `Days`, `Weeks`, `Months`, `Uses`, `FourWeekMonths` — `Months` parses but does not function correctly, so avoid it
- Always link to the resource **type** (`--resource-types`), not the resource instance — one rate prices all resources of that type
- `ChargePeriod` reads back as an integer on `get` (`1` = Minutes, `5` = Uses) — expected numeric-enum behaviour, not a failed write
- Verify every rate by reading it back; use `--agent` for structured JSON output

---

<a name="floorplans"></a>
## Part 5: Floor Plans

### Step 1 — Check the current account

Always verify which account is active before making changes:

```powershell
nexudus.exe whoami
```

Note: **DefaultBusinessId** is the active business context for all subsequent commands.

### Step 2 — Find a floor plan layout

Floor plan **layouts** (the canvas with drawn areas) have a separate ID from floor plan **instances** (the record attached to a business).

```powershell
nexudus.exe floorplanlayouts get <layout-id>
nexudus.exe floorplans list --json   # lists instances; match by FloorPlanLayoutId
```

- Pass the floor plan **instance** ID (`FloorPlanId`) when creating desk items
- Pass the **layout** ID when querying areas, nodes, and assets

### Step 3 — List all areas (zones) on a floor plan

```powershell
nexudus.exe floorplanlayoutareas list --json --floor-plan-layout-id <layout-id> --page-size 100
```

Each area record includes:
- **Name** — zone label (e.g. "Box A2", "Sala Reunions 01")
- **Size** — computed area in m²
- **UniqueId** — used to link a desk item to its drawn shape
- **Id** — numeric area ID

If the output is too large:
```powershell
nexudus.exe floorplanlayoutareas list --json --floor-plan-layout-id <id> --page-size 100 > areas.json
$json = Get-Content areas.json -Raw | ConvertFrom-Json
```

### Step 4 — Filter areas by naming pattern

```powershell
$json = nexudus.exe floorplanlayoutareas list --json --floor-plan-layout-id <id> --page-size 100 | ConvertFrom-Json

# Letter+Number pattern (A2, B12, C7…)
$matching = $json.data | Where-Object { $_.Name -match '^[A-Z]\d+$' }

# Names starting with "Box"
$matching = $json.data | Where-Object { $_.Name -match '^Box\s' }

# Restrict to specific letter series (A and B only)
$matching = $json.data | Where-Object { $_.Name -match '^[AB]\d+$' }

$matching | Select-Object Name, Size, UniqueId | Sort-Object Name | Format-Table -AutoSize
```

Always show the filtered list to the user and confirm before creating anything.

### Step 5 — Bulk-create floor plan desk items

Each area gets one floor plan item. Use the area's **UniqueId** as `--floor-plan-layout-asset-unique-id`.

```powershell
$areas = $json.data | Where-Object { $_.Name -match '<your-pattern>' } | Sort-Object Name

foreach ($area in $areas) {
    $itemName = "Box $($area.Name)"   # adjust naming convention as instructed by the user
    $result = nexudus.exe floorplandesks create `
        --floor-plan-id <floorplan-instance-id> `
        --name $itemName `
        --item-type Office `              # Office | DedicatedDesk | HotDesk | Room | Other
        --available true `
        --available-from-time "2021-01-01T00:00:00" `
        --size $area.Size `
        --size-is-linked-to-area true `
        --floor-plan-layout-asset-unique-id $area.UniqueId 2>&1

    if ($result -match "Created") {
        Write-Host "OK: $itemName"
    } else {
        Write-Host "FAIL: $itemName`n$result"
    }
}
```

#### Key parameters

| Parameter | Values | Notes |
|---|---|---|
| `--floor-plan-id` | numeric ID | Floor plan **instance** ID, not layout ID |
| `--item-type` | `Office`, `DedicatedDesk`, `HotDesk`, `Room`, `Other` | |
| `--available` | `true` / `false` | `true` = In Service |
| `--available-from-time` | ISO 8601 UTC | e.g. `"2021-01-01T00:00:00"` |
| `--size-is-linked-to-area` | `true` | Keeps size in sync with the drawn shape |
| `--floor-plan-layout-asset-unique-id` | area UniqueId | Links item to its zone on the canvas |
| `--resource-id` | numeric ID | **New (v5.0.23):** links the desk item to a bookable `resource`, so the floor-plan item and the booking resource are the same unit. Pass it when the item should be bookable. |

Other `floorplandesks create/update` flags (v5.0.33): `--coworker-id` (assign the unit to a customer), `--capacity` (seats; feeds the AI minimum-capacity filter), `--price`, `--position-x/y/z`, `--area` (explicit size, as opposed to `--size-is-linked-to-area`), `--available-to-time` (pairs with `--available-from-time` to bound availability), `--sensor-id` (occupancy sensor binding), plus the AI-recommendation family `--available-to-ai`, `--notes-for-ai`, `--price-for-ai`, `--show-price-for-ai`.

### Step 6 — Verify created items

```powershell
nexudus.exe floorplandesks list --json --floor-plan-id <floorplan-instance-id> |
  ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Name, ItemType, Available, AvailableFromTime |
  Format-Table -AutoSize
```

### Floor Plans: Common patterns

- **"Create an office for every area named Box 1, Box 2… Box N"** — filter `^Box\s\d+$`, name items exactly as the area name, type `Office`
- **"Create offices for areas matching Letter+Number (A2, B12…), named Box A2 etc."** — filter `^[A-C]\d+$` (adjust letter range), prefix name with `"Box "`
- **"Only A, B, C series — skip other letters"** — use `^[A-C]\d+$` not `^[A-Z]\d+$`

### Floor Plans: List all desk items on a floor plan

```powershell
nexudus floorplandesks list --json --floor-plan-id <floor-plan-id> --page-size 200 2>&1 |
  ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, Name, ItemType, Available, AvailableFromTime |
  Format-Table -AutoSize
```

To check multiple floor plans at once, loop:

```powershell
$floorPlanIds = @(1414760135, 1414717706, 1414760234)
foreach ($fpId in $floorPlanIds) {
    Write-Host "=== Floor Plan $fpId ==="
    nexudus floorplandesks list --json --floor-plan-id $fpId --page-size 200 2>&1 |
      ConvertFrom-Json | Select-Object -ExpandProperty data |
      Select-Object Id, Name, ItemType, Available, AvailableFromTime |
      Format-Table -AutoSize
}
```

Note: `Name` is nested under `.data` — always use `Select-Object -ExpandProperty data` before accessing it. The raw `get` response also nests data under `.data`.

### Floor Plans: Delete desk items in bulk

The delete command prompts for `[y/n]` confirmation — pass `--yes` to skip it non-interactively:

```powershell
$ids = @(1234, 5678, ...)

$ok = 0; $fail = 0
foreach ($id in $ids) {
    $result = nexudus floorplandesks delete $id --yes 2>&1
    if ($LASTEXITCODE -eq 0) {
        $ok++; Write-Host "OK  $id"
    } else {
        $fail++; Write-Host "FAIL $id : $result"
    }
}
Write-Host "Done: $ok deleted, $fail failed"
```

Always confirm the list with the user before running — deletions are irreversible.

### Floor Plans: Rename desk items (strip name suffixes)

When desk items have names like `"4. Centelles"`, `"Box 1. Aiwealth"`, `"FD. Anna"` and you need to strip to just the unit identifier:

**Step 1 — Fetch current names** (use `get` + `Select-Object -ExpandProperty data`):

```powershell
$ids = @(1234, 5678, ...)
$items = foreach ($id in $ids) {
    $data = nexudus floorplandesks get $id --json 2>&1 | ConvertFrom-Json | Select-Object -ExpandProperty data
    [PSCustomObject]@{ Id = $id; CurrentName = $data.Name }
}
$items | Format-Table -AutoSize
```

**Step 2 — Build new names and confirm with user** before updating. Naming rules:
- `"N. Suffix"` → `"N."` (strip everything after `N.`)
- `"Box N. Suffix"` → `"Box N."` (keep `Box N.`, strip suffix)
- `"Box N Suffix"` (no period) → `"Box N."` (add period, strip suffix)
- `"FD. Suffix"` → `"FD."` (strip suffix)
- Items already matching the clean pattern → skip, no update needed

**Step 3 — Bulk update:**

```powershell
$updates = @(
    @{ Id=1234; NewName="4." },
    @{ Id=5678; NewName="Box 1." },
    @{ Id=9012; NewName="FD." }
)

$ok = 0; $fail = 0
foreach ($u in $updates) {
    $result = nexudus floorplandesks update $u.Id `
        --name $u.NewName `
        --caller claude --intent "Remove name suffix, keep unit number only" `
        --agent 2>&1 | ConvertFrom-Json
    if ($result.ok) {
        $ok++; Write-Host "OK  $($u.Id) → $($u.NewName)"
    } else {
        $fail++; Write-Host "FAIL $($u.Id): $($result.summary)"
    }
}
Write-Host "Done: $ok updated, $fail failed"
```

### Floor Plans: Bulk date updates on existing items

When the user wants to change `--available-from-time` on a list of existing floor plan desk IDs, use `floorplandesks update` in a loop. For batches of 100+ IDs, default timeout (2 min) is too short — split into 2–3 sub-batches and pass `--timeout 300000` (or increase the PowerShell timeout) between runs.

```powershell
$ids = @(1234, 5678, ...)    # full list of floorplandisk IDs
$date = "2019-01-01T00:00:00"

# Split into batches of ~50 to avoid timeouts
$batch1 = $ids[0..49]
foreach ($id in $batch1) {
    $result = nexudus floorplandesks update $id `
        --available-from-time $date 2>&1
    Write-Host "$id : $result"
}
# spot-check between batches:
nexudus floorplandesks get $ids[0] --json | ConvertFrom-Json | Select-Object AvailableFromTime

$batch2 = $ids[50..($ids.Count - 1)]
foreach ($id in $batch2) {
    $result = nexudus floorplandesks update $id `
        --available-from-time $date 2>&1
    Write-Host "$id : $result"
}
```

Verify a sample after each batch before continuing.

### Floor Plans: Link existing units to a bookable resource

To make existing floor plan units point at a booking `resource`, set `--resource-id` on each item via `floorplandesks update` (the same flag `create` accepts). Common request: "link units 4.425-001…4.425-050 to resource `<id>`".

**Finding the units:** they are `floorplandesks` items, listed per floor plan and named like `4.425-001`. Match the floor plan by `Name` (`floorplans list`), then `floorplandesks list --floor-plan-id <instance-id> --page-size 500` and filter names by regex. Confirm the target resource exists **on the same business** (`resources get <id>`) before writing.

```powershell
$deskIds = @(1415595831, 1415595832, ...)   # the 50 unit IDs, in range
$resId   = 1415389566
$ok = 0; $fail = 0
foreach ($id in $deskIds) {
    $r = nexudus floorplandesks update $id `
        --caller claude --intent "Link floor plan unit to resource" `
        --resource-id $resId --agent 2>&1 | ConvertFrom-Json
    if ($r.ok) { $ok++ } else { $fail++; Write-Host "FAIL $id : $($r.summary)" }
}
Write-Host "Done: $ok linked, $fail failed"
# verify: every unit must report the resource id
foreach ($id in $deskIds) {
    $d = nexudus floorplandesks get $id --json 2>&1 | ConvertFrom-Json | Select-Object -ExpandProperty data
    if ($d.ResourceId -ne $resId) { Write-Host "NOT LINKED $id (ResourceId=$($d.ResourceId))" }
}
```

- **`floorplandesks update` does NOT wipe unspecified fields** — a field-only desk update preserves `Name`, `FloorPlanId`, `FloorPlanLayoutAssetUniqueId` (the area link) and `ItemType`. (Confirmed: canary link of one unit left all four intact.) Still back up the desks' JSON and **canary one** before batching, then verify by readback (`ResourceId` / `ResourceName`).
- **Idempotent:** re-running the link with the same `--resource-id` is safe, so a re-run after a timeout won't corrupt anything.
- **Allocation caveat — flag to the user:** linking many units to a **single** resource does not raise that resource's capacity. If the resource has `Allocation = 1`, it still allows only **one** concurrent booking across all linked units. If each unit should be independently bookable, the resource's `Allocation` (or the resource design) needs revisiting — surface this rather than assuming the link alone makes 50 desks bookable.

---

<a name="resourceproducts"></a>
## Part 6: Resource Booking Products

Two different things, don't confuse them:
- **Create a store product** (a sellable item: day pass, refreshment, locker, etc.) → `nexudus products create` (see "Creating store products" just below).
- **Attach an existing product to a resource** so members can add it to a booking → `nexudus resourceproducts create` (the rest of this section).

### Creating store products (`products create`)

For sellable items — refreshments, day passes, lockers — priced in the store / members portal:

```powershell
nexudus products create `
  --caller claude --intent "Create product from sheet" `
  --business-id <id> --currency-id <id> `
  --name "Water" `
  --description "Water" `                  # REQUIRED — see gotcha below
  --price 1 `
  --visible true `
  --available-as OnlyOneOff `              # OnlyOneOff | OnlyRecurrent | RecurrentOrOneOff
  --system-product-type Other `            # DayPass | CreditBundle | Stationery | BookingFeature | BookingProducts | Other
  --only-for-members true `                # optional; omit for "available to everyone"
  --agent
```

- **⚠️ `--description` is REQUIRED.** Creating a product without it fails with `HTTP 400: Description: is a required field`. If the sheet leaves the description blank, fall back to the product name (and tell the user you did) rather than erroring.
- **⚠️ `--description` must be under 255 characters.** Longer values fail with `Description: must be less than 255 characters`. Catering/menu items in price lists routinely exceed this — check `$text.Length` before the call and trim, then tell the user exactly what you shortened. (Note this is the *product* description; the *resource* description has no such limit and can hold tens of KB of HTML.)
- **Product categories are `--tags`** (comma-separated, "for categorising and filtering"). There is no product-group entity in the CLI, and `--group-name` is not a `products create` flag even though `GroupName` appears in the list DTO.
- `--system-product-type` valid values: `DayPass`, `CreditBundle`, `Stationery`, `BookingFeature`, `BookingProducts`, `Lockers`, `Equipment`, `EventServices`, `AdminServices`, `FoodAndBeverage`, `Other`.
- **"Available to everyone" = pass neither `--only-for-members` nor `--only-for-contacts`.** Both default to false, which means unrestricted; there is no positive "everyone" flag to set.
- **⚠️ Verify products with `get`, never `list`.** The `products list` DTO returns **defaults** for `AvailableAs` (always `1`), `OnlyForMembers` and `OnlyForContacts` (always `false`) regardless of the stored values. A list-based check therefore shows every restriction as missing and every product as RecurrentOrOneOff — it looks exactly like a batch of silently failed writes. `products get <id>` returns the truth. `AvailableAs` integers: **1** = RecurrentOrOneOff, **2** = OnlyRecurrent, **3** = OnlyOneOff.
- **Stock quantity cannot be set via the CLI.** `products create` offers `--track-stock`, `--stock-alert-level` and `--allow-negative-stock`, but no opening-count flag, and there is no stock command. Spreadsheet stock numbers are an admin-UI step — say so rather than leaving the user to assume they landed. A stock pool shared between two products (e.g. "10 quotas across both day passes") cannot be modelled at all; stock is per-product.
- Expiration dates, access windows and booking-credit entitlements named in a products sheet are **not** product fields. A "credit bundle" product can be created, but the credit it grants is a plan benefit (`tariffbookingcredits`) — creating the product alone does not grant anything. Flag this rather than letting it look complete.

### Products: names containing a double quote

**⚠️ A `"` in a `--name` or `--description` breaks the call in Windows PowerShell 5.1.** Product names like `LED Wall 120"` or `IMAGO smart displays 65"` make the shell split the argument, and the CLI reports a nonsense error such as `Unknown command 'Wall'`. Nothing is created, so this fails loudly rather than silently — but the whole batch item is lost.

Escape the quote before building the argument array:

```powershell
$esc = $name -replace '"','\"'
$a = @('products','create','--name',$esc,'--description',$esc, ...)
```

The backslash is consumed by the shell, and the value **stores correctly** as `LED Wall 120"` — confirmed by readback (`$d.Name -ceq 'LED Wall 120"'` is true). Apply the same escape to any text field sourced from a spreadsheet; inch marks and quoted nicknames are common in AV equipment and room lists.
- **Day passes** (access for a day) → `--system-product-type DayPass`; refreshments/consumables → `Other`.
- **Expiration date and access windows are NOT settable via `products create`** — those are admin-UI only (there's no expiration/validity flag).
- `AvailableAs` / `SystemProductType` read back as integers on `get` — the usual numeric-enum readback. Observed (2026-09-29): `AvailableAs` 2 = OnlyRecurrent, 3 = OnlyOneOff; `SystemProductType` 1 = DayPass, 5 = BookingProducts, 99 = Other.
- Catering / refreshments sold through a booking → `--system-product-type BookingProducts`, `--visible false` (hidden from the store), then attach to resources with `resourceproducts`.
- **Stock count is not settable.** `--track-stock`, `--allow-negative-stock` and `--stock-alert-level` exist, but there is no flag for the starting quantity — that's an admin-UI step. Flag it when a sheet gives a stock number.
- **Other `products create` flags (v5.0.33):** `--sku` (stock code from the sheet's product-code column), `--invoice-coworker`, `--apply-pro-rating`, `--invoice-display` (invoice line text), `--visible-in-kiosk`, `--sync-nex-kiosk`, `--sync-square`, `--new-image-url` / `--clear-image-file`, `--display-order`, `--starred`, `--archived`, `--exempt-tax-rate-id` / `--reduced-tax-rate-id`, and the AI family `--available-to-ai` / `--notes-for-ai` / `--price-for-ai` / `--show-price-for-ai`.
- Restrict a product to specific plans with repeated `--tariffs <id>` (e.g. a guest pass for five membership plans).
- Verify with `products get <id>` (name, price, AvailableAs, Visible).

### Attaching products to resources (`resourceproducts`)

Resource booking products let members purchase items (breakfast, drinks, equipment hire, etc.) as part of a booking. These are distinct from resource rates — they use `nexudus resourceproducts create`, not `extraservices`.

### Key flags

| Flag | Value | Purpose |
|---|---|---|
| `--visible true` | boolean | Enables "Let customers buy this product as part of bookings" |
| `--request-quantity true` | boolean | Enables "Let customers request more than one item" |
| `--invoice-in-minutes true` | boolean | Enables **"Charge this product based on the length of the booking it is added to."** — see below |
| `--price` | amount | Per-resource price override; the product costs this on *this* resource instead of its own default price |

The first two replicate the familiar checkboxes in the Nexudus resource admin UI. All four flags exist on both `create` and `update`.

#### ⚠️ `--invoice-in-minutes` — hourly-priced products need it, and it defaults to OFF

The admin-UI label is **"Charge this product based on the length of the booking it is added to."** When `false` (the default), the product is charged as a **flat amount per booking** no matter how long the booking runs. When `true`, the price is multiplied by the booking's duration.

**Any product whose price is quoted per hour must have this set to `true`** — equipment hire, AV, staffed services. Attaching an hourly-priced product without it silently undercharges every long booking, and nothing in the create response hints at it. The price on the product record looks right either way, so this will not show up in a price readback — only in what a member is actually billed.

Decide it per product, from how the price list quotes the amount:

| Price list says | `--invoice-in-minutes` |
|---|---|
| "HKD 1,000 **per hour**" (AV, equipment, staff) | `true` |
| "HKD 50" flat, per booking (cables, flip chart, whiteboard) | `false` |
| "HKD 120 **per head**" (catering) | `false` — use `--apply-to-visitors` on a rate for per-head scaling, not this |

Fixing it afterwards is a safe field-only `resourceproducts update <record-id> --invoice-in-minutes true` — verified on 2026-10-07 that it preserves `Visible`, `RequestQuantity`, `ResourceId` and `ProductId`. Find the records to fix by listing each resource's products and filtering on `ProductId`:

```powershell
$targetProds = @('<led-wall-id>','<imago-65-id>','<imago-75-id>')
foreach ($r in $resIds) {
  @(nexudus resourceproducts list --json --resource-id $r --page-size 100 | ConvertFrom-Json |
      Select-Object -ExpandProperty data) |
    Where-Object { $targetProds -contains "$($_.ProductId)" } | ForEach-Object {
      nexudus resourceproducts update $_.Id --caller claude `
        --intent "Enable length-based charging for hourly product" `
        --invoice-in-minutes true --agent
    }
}
```

**Verify both directions:** confirm every target record is now `true` *and* that every other record is still `false` — a filter bug that flips extra records is otherwise invisible. Note `ProductId` comes back as a number, so compare it as a string (`"$($_.ProductId)"`) when matching against an array of string IDs.

### Step 1 — Gather product and resource IDs

```powershell
nexudus products list --json | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, Name | Where-Object { $_.Name -match "keyword" }

nexudus resources list --json | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, Name
```

### Step 2 — Add a single product to a resource

```powershell
nexudus resourceproducts create `
  --caller claude --intent "Add booking product to resource" `
  --resource-id <resource-id> `
  --product-id <product-id> `
  --visible true `
  --request-quantity true `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

### Step 3 — Batch: add many products to one resource

```powershell
$resourceId = <id>
$productIds = @(111, 222, 333, ...)   # all product IDs to add

$ok = 0; $skip = 0; $fail = 0
foreach ($pid in $productIds) {
    $raw = nexudus resourceproducts create `
        --caller claude --intent "Add booking product to resource" `
        --resource-id $resourceId `
        --product-id $pid `
        --visible true `
        --request-quantity true `
        --agent 2>&1

    if ($raw -match '"ok":true') {
        $ok++; Write-Host "OK  | product $pid"
    } elseif ($raw -match "already been added") {
        $skip++; Write-Host "SKIP| product $pid (already added)"
    } else {
        $fail++; Write-Host "FAIL| product $pid`n$raw"
    }
}
Write-Host "Done: $ok ok, $skip skipped, $fail failed"
```

"Already been added" (HTTP 400) is non-fatal — treat as a skip, not a failure. Match the `ok` field with `'"ok":\s*true'`, not `'"ok":true'` — `--agent` output is pretty-printed with a space after the colon, so the tighter pattern never matches and every success is miscounted as a failure.

**Do not name the inner loop variable `$pid`** — see the reserved-variable warning in the shared array-flag section; it silently skips the whole inner loop.

### Tiering products across many resources

For a space where different room classes get different add-ons, build a tier plan rather than one flat product list — it keeps the matrix reviewable before you write 100+ records:

```powershell
$tierA = @(<ids attached to every resource>)          # cables, laptop, whiteboard, per-head catering
$tierB = @(<ids for event/function rooms only>)       # chairs, tables, mics, displays, translators
$plan = @(
  @{ id=<res-id>; n='Meeting Room - 4 pax';  b=$false; led=$false; cat=$null      },
  @{ id=<res-id>; n='Event Space - 80 pax';  b=$true;  led=$true;  cat=<cat-90-id> }
)
foreach ($p in $plan) {
    $prodIds = @($tierA)
    if ($p.b)   { $prodIds += $tierB }
    if ($p.led) { $prodIds += $ledWallId }
    if ($p.cat) { $prodIds += $p.cat }      # capacity-matched catering, one per resource
    foreach ($prodId in $prodIds) { ... }
}
```

**Capacity-matched products** (catering sized 25 / 50 / 90 / 120 pax) attach **one per resource**, rounded **up** to the nearest size that covers the room — a 100-pax space gets the 120-pax tray, not the 90. Confirm the round-up with the user; it changes the price the member sees.

Also check `AvailableAs` before promising a product as a booking add-on: an `OnlyRecurrent` product **cannot** be bought one-off on a booking, so attaching it achieves nothing. Flag those and offer to change `--available-as` first.

### Step 4 — Batch: add same product list to multiple resources

```powershell
$productIds = @(111, 222, 333, ...)
$resourceIds = @(1001, 1002, 1003, ...)

foreach ($rid in $resourceIds) {
    Write-Host "=== Resource $rid ==="
    foreach ($pid in $productIds) {
        $raw = nexudus resourceproducts create `
            --caller claude --intent "Add booking products in bulk" `
            --resource-id $rid --product-id $pid `
            --visible true --request-quantity true `
            --agent 2>&1
        if ($raw -match '"ok":true') { Write-Host "  OK  $pid" }
        elseif ($raw -match "already been added") { Write-Host "  SKIP $pid" }
        else { Write-Host "  FAIL $pid`n  $raw" }
    }
}
```

### Resource Booking Products: Rules

- Always set both `--visible true` and `--request-quantity true` unless the user says otherwise
- **Set `--invoice-in-minutes true` on every per-hour product** — it defaults to false, which charges a flat amount regardless of booking length. Ask which products are hourly rather than assuming; it is invisible in a price readback
- "Already been added" is not an error — skip and continue
- Products must already exist on the account — this command links them, it does not create products
- The command links a product to a specific resource instance, not a resource type

---

<a name="translations"></a>
## Part 7: Product Translations

Nexudus stores per-language overrides for entity names (products, plans, etc.) as **language tokens**. The token `--name` field is the entity's numeric ID (as a string), and `--value` is the translated text.

### Step 1 — Find the language ID

```powershell
nexudus languages list --json | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Where-Object { $_.Name -match 'French' } |
  Select-Object Id, Name, CultureName
```

If two entries appear for French (one per business network), use the one whose ID matches the current account's network — or try both and use the one that succeeds.

### Step 2 — Create the translation token

```powershell
nexudus languagetokens create `
  --caller claude --intent "Add French translation for product" `
  --language-id <french-language-id> `
  --name "<product-id>" `       # product's numeric ID as a plain string
  --value "Translated Name" `
  --agent 2>&1
```

The `--name` field is the entity ID — this is not a human-readable name. The `--value` is what members will see in that language.

### Step 3 — Verify

In the Nexudus admin, open the product, click the language tab, and confirm the French translation is shown.

### Translations: Rules

- `--name` is the entity's numeric ID (e.g., `"1415230086"`), not a descriptive label
- `--value` is the translated display string members see
- If multiple French language IDs exist, try the second/newer one first; fall back to the first
- This pattern works for any translatable entity — products, plans, resource types, etc.
- The entity must already exist on the account before a translation can be added

---

<a name="bookings"></a>
## Part 8: Bookings Management

> **⚠️ CLI change: `--cancel-if-not-checked-in` has been REMOVED from `bookings update`.** The `CancelIfNotCheckedIn` property is no longer settable via the CLI — the flag exists on neither `bookings update`, `bookings create`, nor `resources create`. If a task needs to change "cancel if not checked in", do it in the admin UI and tell the user. The examples below use a still-valid writable property (`--tentative`) to demonstrate the list → inspect → bulk-update pattern; swap in whichever real property the task needs.
>
> **⚠️ Also removed by v5.0.33: `--billed`.** An earlier version of this skill listed it as still-valid — it is not in v5.0.33 `--help`. The writable set confirmed on **2026-10-07** is: `--resource-id`, `--floor-plan-desk-id`, `--coworker-id`, `--extra-service-id`, `--from-time`, `--to-time`, `--notes`, `--charge-now`, `--invoice-now`, `--invoice-this-coworker`, `--do-not-use-booking-credit`, `--purchase-order`, `--discount-code`, `--tentative`, `--teams-at-booking`, `--tariff-at-booking`, `--which-bookings-to-update`, `--override-price`, `--include-zoom-invite`, `--coworker-checked-in-at`, `--coworker-checked-out-at`, `--booking-products`, `--booking-visitors`.

### List bookings with date filtering

```powershell
nexudus bookings list --json `
  --from-from-time "2026-07-07" `   # earliest FromTime to include (ISO date)
  --to-from-time   "2026-09-07" `   # latest FromTime to include (omit for open-ended)
  --page-size 200 2>&1 | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, ResourceName, CoworkerFullName, FromTime, ToTime, Tentative, Billed
```

Key filter params:
- `--from-from-time` / `--to-from-time` — filter by the booking's **start time** (FromTime)
- `--page-size` — default is small; use 200 for a month-range query

Note: `CoworkerFullName` and similar member fields are redacted as `«PII:NAME:...»` in CLI output due to PII masking. This is expected behavior — don't treat it as missing data.

### Inspect a booking property across all results

```powershell
$bookings = nexudus bookings list --json `
  --from-from-time "2026-07-07" --page-size 200 2>&1 |
  ConvertFrom-Json | Select-Object -ExpandProperty data

# Show the property distribution (Tentative shown as the example property)
$bookings | Select-Object Id, Tentative | Format-Table -AutoSize

# Filter to specific values
$bookings | Where-Object { $_.Tentative -eq $true } | Select-Object Id
```

### Update a single booking property

```powershell
nexudus bookings update <booking-id> `
  --caller claude --intent "Update booking setting" `
  --tentative false `
  --agent 2>&1 | ConvertFrom-Json | Select-Object ok, summary
```

### Bulk-update a booking property

```powershell
$bookingIds = @(111, 222, 333, ...)   # IDs where the property needs changing

$ok = 0; $fail = 0
foreach ($bid in $bookingIds) {
    $result = nexudus bookings update $bid `
        --caller claude --intent "Bulk update booking property" `
        --tentative false `
        --agent 2>&1 | ConvertFrom-Json
    if ($result.ok) { $ok++ } else { $fail++; Write-Host "FAIL $bid: $($result.summary)" }
}
Write-Host "Done: $ok ok, $fail failed"
```

### Updating a repeat/recurring booking series

`bookings update` has **`--which-bookings-to-update`** to control how an edit applies to a repeat series. It is the only field that edits a repeat series as a unit — pass it when the target booking is part of a recurrence and the user wants the change applied beyond the single instance. **Accepted values (verified v5.0.33):**

| Value | Effect |
|---|---|
| *(empty)* | this booking only (default) |
| `UpdateThisBookingOnly` | this instance |
| `UpdateFutureBookingsOnly` | this instance and all later ones |
| `UpdateAllBookings` | the whole series |
| `UpdateNotChargedBookings` | only instances not yet charged |
| `DeleteAllBookings` | **deletes** the whole series |
| `DeleteBookingsAfterThis` | **deletes** this instance and all later ones |
| `DeleteNotChargedBookings` | **deletes** instances not yet charged |
| `RevertAllCharges` | reverses charges across the series |

**⚠️ Four of these nine values delete bookings or reverse charges.** Always state which value you intend and get explicit approval before passing one of the `Delete*` or `RevertAllCharges` values — they are not recoverable. Other newer write flags: `--override-price`, `--booking-products`, `--booking-visitors` (JSON array or `@filepath`).

### Known error: FOR_MEMBERS_ONLY

If a booking update returns an error containing `FOR_MEMBERS_ONLY`, the resource has a "members only" restriction set. The CLI cannot bypass this. Fix:

1. Go to the resource in the Nexudus admin UI
2. Disable the members-only booking restriction
3. Retry the update command

### Copying credits between tariffs

To copy all extra service credits from one tariff to several others:

```powershell
# 1. List credits on the source tariff
$src = nexudus tariffextraservices list --json --tariff-id <source-tariff-id> |
  ConvertFrom-Json | Select-Object -ExpandProperty data

# 2. Build entries for each target tariff
$targetTariffIds = @(111, 222, 333, ...)
$entries = foreach ($tid in $targetTariffIds) {
    foreach ($s in $src) {
        [PSCustomObject]@{ tid=$tid; esid=$s.ExtraServiceId; uses=$s.UsesIncluded; ren=$s.ServiceRenewalTime }
    }
}

# 3. Create all entries (uses same loop as Part 3 Credits)
foreach ($e in $entries) {
    $raw = nexudus tariffextraservices create `
        --caller claude --intent "Copy credits to target tariff" `
        --tariff-id $e.tid --extra-service-id $e.esid `
        --uses-included $e.uses --service-renewal-time $e.ren 2>&1
    Write-Host "tariff $($e.tid) | esid $($e.esid): submitted"
    # CLI parse error on ExtraServiceChargePeriod is cosmetic — records are created
}
```

### Bookings Management: Rules

- Always confirm which bookings will be affected before bulk-updating — show the user the filtered list first
- For `FOR_MEMBERS_ONLY` errors: ask the user to disable the resource restriction in the admin UI, then retry
- Member name/email fields are PII-masked in CLI output — this is expected, not a data issue
- Date filters use `--from-from-time` / `--to-from-time` (the booking's start time), not created-at time
- Use `--page-size 200` for any date-range query to avoid missing bookings on the default small page

---

<a name="duplicating"></a>
## Part 9: Duplicating Resources

There is **no clone command**. To duplicate a resource with its exact settings, read the source with `get` and recreate field-by-field, then replicate its associations (tariffs, linked resources, time slots, booking products) separately.

### Step 1 — Confirm what to copy

Ask the user whether "exact settings" includes the associations attached *to* the resource, or just the resource's own fields. Inspect the source first so the scope is known:

```powershell
$d     = nexudus resources get <src-id> --json | ConvertFrom-Json | Select-Object -ExpandProperty data
$slots = @(nexudus resourcetimeslots list --json --resource-id <src-id> --page-size 200 | ConvertFrom-Json | Select-Object -ExpandProperty data)
$prods = @(nexudus resourceproducts list --json --resource-id <src-id> --page-size 200 | ConvertFrom-Json | Select-Object -ExpandProperty data)
"tariffs=$(@($d.Tariffs).Count) slots=$($slots.Count) products=$($prods.Count) linked=$($d.LinkedResources -join ',')"
```

### Step 2 — Copy the resource fields

Build the create as an **array**, iterating a property→flag map and passing every populated field from the source (see the amenity/limit flags in Part 2). Then rename to the target name and set `--business-id` explicitly (copies may go to a different location than the active account). Key points:

- Attach **tariffs** and **linked resources** via **repeated flags** (`--added-tariffs X --added-tariffs Y …`), never comma-joined — see the shared array-flag rule. Do this in a follow-up `update` after create.
- **Long HTML fields** (`--description`, `--email-confirmation-content`) can blow the ~32 KB command-line limit. If the source's fields are large, create **without** them, then set each in its own isolated `update`. Never use `@filepath` for text fields — only `--time-slots` supports it.
- **`DisplayOrder` cannot be forced** — the system reassigns it (typically `100`) on create and ignores explicit updates. Expect this one field to differ from the source; it only affects portal sort order.
- **⚠️ `--access-control-group-id` does NOT exist (v5.0.33)** — on `resources create` or `resources update`. An earlier version of this skill said to copy `AccessControlGroupId` with it. If the source resource has one set, it **cannot** be copied via CLI; flag it to the user as an admin-UI step rather than silently dropping it.
- **`resources create` has 97 flags in v5.0.33** — far more than Part 2's example. Because duplication copies *every populated field*, build the property→flag map from the live `resources create --help`, not from Part 2's abbreviated list, or newer fields (the amenity set, cancellation-fee family, `--booking-availability-exceptions`, calendar-sync IDs, `--limit-visitors-to-allocation`) silently won't be copied.

### Step 3 — Replicate time slots and booking products

For each new resource, recreate the source's slots and products:

```powershell
# time slots — extract the wall-clock time, rebuild on the 1976-01-01 placeholder date
foreach ($s in $slots) {
  $ft = if ("$($s.FromTime)" -match 'T(\d{2}:\d{2}:\d{2})') { "1976-01-01T$($Matches[1])" }
  $tt = if ("$($s.ToTime)"   -match 'T(\d{2}:\d{2}:\d{2})') { "1976-01-01T$($Matches[1])" }
  $day = @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')[[int]$s.DayOfWeek]
  nexudus resourcetimeslots create --resource-id <new-id> --from-time $ft --to-time $tt --day-of-week $day --agent
}
# booking products — replicate ProductId, Visible, RequestQuantity
foreach ($p in $prods) {
  $vis = if ($p.Visible) {'true'} else {'false'}; $rq = if ($p.RequestQuantity) {'true'} else {'false'}
  nexudus resourceproducts create --resource-id <new-id> --product-id $p.ProductId --visible $vis --request-quantity $rq --agent
}
```

### Step 4 — Verify each copy against the source

Compare **counts** (tariffs, slots, products, linked) and field lengths (description/email) of each new resource against the source. Do this by readback — the create/update exit codes are not trustworthy.

### Duplicating Resources: Rules

- A rate/price lives on the resource **type**, not the instance — copies sharing a type share pricing; you cannot price them apart without separate types
- Linked resources copy one-directionally; note to the user that the partner resources are **not** updated to lock back unless asked
- Verify by readback; `DisplayOrder` is the expected exception
- Watch the command-line length limit on large HTML fields; split them into isolated updates

---

<a name="movingplans"></a>
## Part 10: Moving Plans Between Locations

A plan (tariff) is moved to another business/location by changing its `--business-id`. `list` returns tariffs across the whole network, so filter by `BusinessId` to find the right set (see the cross-network shared rule).

### Step 1 — Identify the plans and confirm the target

```powershell
# find candidates in the source location, filtered by BusinessId (not name)
nexudus tariffs list --json --page-size 500 | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Where-Object { $_.BusinessId -eq <source-business-id> -and $_.Name -match '<pattern>' } |
  Select-Object Id, Name, BusinessName

# confirm the destination really is the location the user named
nexudus businesses get <target-business-id> --json | ConvertFrom-Json | Select-Object -ExpandProperty data |
  Select-Object Id, Name, CurrencyCode, CountryName
```

Check the target's currency matches the plans' currency before moving.

### Step 2 — Move (canary first, then batch)

```powershell
nexudus tariffs update <tariff-id> `
  --caller claude --intent "Move plan to <target location>" `
  --business-id <target-business-id> --agent
```

Move one plan, read it back to confirm `BusinessId` / `BusinessName` changed, then loop the rest and re-verify all.

### Moving Plans Between Locations: Rules

- Only the **location field** changes. Attached tariff associations (resources, extra services/credits, time passes) still reference the **original** location's entities — flag this; the moved plans are not automatically self-contained in the new location
- Confirm the destination business ID resolves to the expected location name before moving
- Verify each move by readback; do not trust exit codes
- Move a canary first, confirm, then batch

---

<a name="enums"></a>
## Appendix: Verified Enums

All lists below were read directly from the CLI on **2026-10-07** against **v5.0.33**, using the invalid-value probe from the header tip. They are authoritative, not empirical guesses. Re-probe after a CLI upgrade rather than trusting this table.

| Flag | Command(s) | Valid values |
|---|---|---|
| `--charge-period` | `extraservices` | `Minutes`, `Days`, `Weeks`, `Months`, `Uses`, `FourWeekMonths` — **`Months` parses but misbehaves; avoid** |
| `--service-renewal-time` | `tariffextraservices`, `tariffbookingcredits` | `Week`, `CalendarMonth`, `TariffMonth`, `Year`, `Day` |
| `--pass-renewal-time` | `tarifftimepasses` | `Week`, `CalendarMonth`, `TariffMonth`, `Year`, `Day` |
| `--system-resource-type` | `resources` | `MeetingRoom`, `HotDesk`, `PrivateOffice`, `EventSpace`, `Lab`, `Kitchen`, `TreatmentRoom`, `StorageUnit`, `Machine`, `DayPass`, `PhoneBooth`, `Other` |
| `--system-tariff-type` | `tariffs` | `FullTimePrivateOffice`, `PartTimePrivateOffice`, `FullTimeDedicatedDesk`, `PartTimeDedicatedDesk`, `FullTimeHotDesk`, `PartTimeHotDesk`, `FullTimeOther`, `PartTimeOther`, `Storage`, `VirtualOffice`, `Virtual`, `Other` |
| `--system-product-type` | `products` | `DayPass`, `CreditBundle`, `Stationery`, `BookingFeature`, `BookingProducts`, `Lockers`, `Equipment`, `EventServices`, `AdminServices`, `FoodAndBeverage`, `Other` |
| `--available-as` | `products` | `RecurrentOrOneOff`, `OnlyRecurrent`, `OnlyOneOff` |
| `--item-type` | `floorplandesks` | `Office`, `DedicatedDesk`, `HotDesk`, `Other`, `Room` |
| `--day-of-week` | `resourcetimeslots` | `Sunday`, `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday` |
| `--which-bookings-to-update` | `bookings update` | *(empty)*, `UpdateThisBookingOnly`, `UpdateFutureBookingsOnly`, `UpdateAllBookings`, `UpdateNotChargedBookings`, `DeleteAllBookings`, `DeleteBookingsAfterThis`, `DeleteNotChargedBookings`, `RevertAllCharges` — **the last four destroy data** |
| `--last-minute-adjustment-type` | `extraservices` | `Fixed` (full factor throughout the period), `Gradual` (ramps from zero) |

**Notes on these enums**

- `--service-renewal-time` and `--pass-renewal-time` are the **same underlying type** (`eTimeSpanWeekMonth`), so the value sets are identical — the earlier guidance that they were separately-guessed lists is superseded.
- `--charge-period` **does** accept `Months` (an earlier version implied it was invalid). It is a behaviour problem, not a validation one: it parses and writes, then prices incorrectly. Keep avoiding it.
- Every one of these reads back as an **integer** on `get`/`list` — see [Shared: Numeric Enum Readback](#numeric-enum-readback). Observed mappings: `ChargePeriod` 1=Minutes, 5=Uses · `AvailableAs` 1=RecurrentOrOneOff, 2=OnlyRecurrent, 3=OnlyOneOff · `SystemProductType` 1=DayPass, 2=CreditBundle, 3=Stationery, 6=Lockers, 7=Equipment, 9=AdminServices, 10=FoodAndBeverage, 99=Other · `SystemResourceType` 1=MeetingRoom, 4=EventSpace · `SystemTariffType` 1=FullTimePrivateOffice, 3=FullTimeDedicatedDesk, 99=Other.
