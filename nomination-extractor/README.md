# Beneficiary Nomination – extraction setup

Signed PDFs dropped into a SharePoint library are read by an **Office Script** (standard licence, no Azure), validated, and written to SharePoint lists (including the PDF as a data URI for Power Apps). The original PDF stays in the library as the authentic copy.

```text
Nomination Submissions (library)          Office Script "Extract nomination"
  new PDF ──► Flow: Process nomination ──► (Excel Online: Run script)
                │  ◄──────── JSON (fields, nominees[], payees[], errors[]) ─┘
                ├─ valid   → Nomination Members (+ PdfDataUri)
                │            Nomination Nominees / Nomination Court Orders
                │            file Status = Processed | Needs review
                └─ invalid → file Status = Rejected + reason
```

| Folder | Contents |
|---|---|
| `office-script/ExtractNomination.ts` | **The extractor.** Paste into Excel on the web as an Office Script |
| `function_app/` | Alternative: the same extractor as an Azure Function (needs Premium + Azure). Not needed for the Office Script route |
| `sharepoint/Create-NominationLists.ps1` | Creates the library, three lists and all columns |
| `power-automate/parse-json-schema.json` | Paste into the flow's **Parse JSON** action |
| `power-automate/sample-response.json` | What the extractor returns for a valid form |

---

## 1. Set up the Office Script

Licence: **Run script** (Excel Online (Business)) is a *standard* connector. Office Scripts must be enabled for your organisation (Microsoft 365 admin center → Settings → Org settings → *Office Scripts*; on by default).

1. In the SharePoint site, create a library or folder for tools (e.g. *Flow Tools*) and create a blank workbook there: `Nomination Extractor.xlsx`. The workbook stays empty – it only hosts the script.
2. Open it in **Excel on the web** → **Automate** tab → **New Script**.
3. Delete the sample code, paste the **whole** of `office-script/ExtractNomination.ts`, rename the script (click its name) to `Extract nomination`, and **Save**.
4. Scripts are saved to *your* OneDrive (Documents → Office Scripts). For a team-owned setup, use the flow action **Run script from SharePoint library** instead and store the `.osts` file in the *Flow Tools* library (Automate → script → *…* → Save as / share).

Limits that matter here: script parameters up to ~28 MB (plenty for a PDF), 120 s per run (a form takes well under 1 s), 1,600 runs per user per day.

What the extractor checks – **errors** (form rejected):

- not the fillable form / fields flattened / not a PDF
- `signature_attached` is not `yes`
- member name or member number empty
- no nominees
- % Benefit doesn't total 100 (Revised column if any revised values are filled in, otherwise Current)
- a % or amount that isn't a number

**Warnings** (form accepted, marked *Needs review*): signature flag set but no image found, invalid date formats, empty declaration names, rows with details but no name.

Dates come back as ISO `yyyy-mm-dd` (or `null`) with the original text alongside; numbers come back as numbers.

### Alternative: Azure Function (only if you have Premium + an Azure subscription)

1. **Azure portal → Create → Function App**
   - Hosting: **Flex Consumption** (or Consumption) – pay per run, with a monthly free grant
   - Runtime: **Python 3.11**, region close to you (e.g. South Africa North)
2. **Deploy the code** – in VS Code with the *Azure Functions* extension: open the `function_app` folder → *Azure: Deploy to Function App…* → choose the app.
   (Or with Azure Functions Core Tools: `func azure functionapp publish <app-name>` from inside `function_app`.)
3. **Get the key** – Function App → *Functions* → `extract_nomination` → *Function Keys* → copy `default`.
4. **Test it** (PowerShell, from the `Doc management` folder):

   ```powershell
   curl.exe -X POST "https://<app-name>.azurewebsites.net/api/extract-nomination" `
     -H "x-functions-key: <key>" -H "Content-Type: application/pdf" `
     --data-binary "@Beneficiary Nomination Form - Sample Filled.pdf"
   ```

   You should get JSON with `"valid": true` and three nominees. (The sample has no signature image, so it also returns one warning about that.)

   In the flow, use an **HTTP** action (POST to that URL, header `x-functions-key`, body `body('Get_file_content')`) in place of *Run extractor*, and set Parse JSON's content to `body('Extract_nomination')`.

## 2. Create the SharePoint lists

Run `sharepoint/Create-NominationLists.ps1` (instructions at the top of the script), or create them by hand:

**Nomination Submissions** (document library): `Status` (Choice: New, Processed, Needs review, Rejected; default New), `RejectReason` (Multiple lines, plain), `MemberItemId` (Number).

**Nomination Members**: Title = member name, `MemberNumber` (indexed), `FundName`, `DateOfBirth` (date only), `BranchCode`, `CostCentre`, `HomeTelephone`, `CellTelephone`, `WorkTelephone`, `EmailAddress`, `DeclarationFullNames`, `SignatureDate` (date only), `SignatureAttachedAt`, `NomineeCount`, `TotalCurrentPct`, `TotalRevisedPct`, `FormVersion`, `Warnings` (multi-line), `SourceFileId`, `SourceFileName`, `SourceFileUrl` (multi-line), **`PdfDataUri` (Multiple lines, plain text)**.

**Nomination Nominees**: Title = full name, `Member` (lookup → Nomination Members), `MemberNumber` (indexed), `RowNumber`, `DateOfBirth`, `Relationship`, `IdentityNumber`, `BenefitCurrentPct`, `BenefitRevisedPct`.

**Nomination Court Orders**: Title = payee name, `Member` (lookup), `MemberNumber` (indexed), `RowNumber`, `Relationship`, `IdentityNumber`, `BenefitPct`, `MonthlyMaintenance`.

## 3. Build the flow "Process nomination"

Name each action exactly as shown – the expressions refer to those names (spaces become `_`).

**Trigger – SharePoint: When a file is created (properties only)**
Site + library *Nomination Submissions*. Settings → Trigger conditions: `@equals(triggerOutputs()?['body/{IsFolder}'], false)`

Add a **Scope** named `Try` and put steps 1–5 inside it.

1. **SharePoint: Get file content** – name `Get file content`
   File Identifier: `triggerOutputs()?['body/{Identifier}']`
2. **Compose** – name `PDF Data URI`
   `concat('data:application/pdf;base64,', body('Get_file_content')?['$content'])`
3. **Excel Online (Business): Run script** – name `Run extractor`
   - Location: your SharePoint site, Document Library: *Flow Tools*, File: `Nomination Extractor.xlsx`, Script: `Extract nomination`
   - pdfBase64: `body('Get_file_content')?['$content']`
4. **Parse JSON** – name `Parse nomination`
   Content: `outputs('Run_extractor')?['body/result']`, Schema: paste `power-automate/parse-json-schema.json`
5. **Condition** – name `Is valid`: `body('Parse_nomination')?['valid']` *is equal to* `true`

### If yes

**5a. SharePoint: Create item** – name `Create member`, list *Nomination Members*

| Column | Expression |
|---|---|
| Title | `body('Parse_nomination')?['member']?['member_name']` |
| MemberNumber | `body('Parse_nomination')?['member']?['member_number']` |
| FundName | `body('Parse_nomination')?['member']?['fund_name']` |
| DateOfBirth | `body('Parse_nomination')?['member']?['date_of_birth']` |
| BranchCode | `body('Parse_nomination')?['member']?['branch_code']` |
| CostCentre | `body('Parse_nomination')?['member']?['cost_centre']` |
| HomeTelephone | `body('Parse_nomination')?['member']?['home_telephone']` |
| CellTelephone | `body('Parse_nomination')?['member']?['cell_telephone']` |
| WorkTelephone | `body('Parse_nomination')?['member']?['work_telephone']` |
| EmailAddress | `body('Parse_nomination')?['member']?['email_address']` |
| DeclarationFullNames | `body('Parse_nomination')?['member']?['declaration_full_names']` |
| SignatureDate | `body('Parse_nomination')?['member']?['signature_date']` |
| SignatureAttachedAt | `body('Parse_nomination')?['member']?['signature_attached_at']` |
| NomineeCount | `body('Parse_nomination')?['totals']?['nominee_count']` |
| TotalCurrentPct | `body('Parse_nomination')?['totals']?['benefit_current_pct']` |
| TotalRevisedPct | `body('Parse_nomination')?['totals']?['benefit_revised_pct']` |
| FormVersion | `body('Parse_nomination')?['form_version']` |
| Warnings | `join(body('Parse_nomination')?['warnings'], decodeUriComponent('%0A'))` |
| SourceFileId | `triggerOutputs()?['body/ID']` |
| SourceFileName | `triggerOutputs()?['body/{FilenameWithExtension}']` |
| SourceFileUrl | `triggerOutputs()?['body/{Link}']` |
| PdfDataUri | `outputs('PDF_Data_URI')` |

**5b. Apply to each** – name `Each nominee`, over `body('Parse_nomination')?['nominees']`
→ **Create item** in *Nomination Nominees*:

| Column | Expression |
|---|---|
| Title | `items('Each_nominee')?['full_name']` |
| Member Id | `outputs('Create_member')?['body/ID']` |
| MemberNumber | `body('Parse_nomination')?['member']?['member_number']` |
| RowNumber | `items('Each_nominee')?['row']` |
| DateOfBirth | `items('Each_nominee')?['date_of_birth']` |
| Relationship | `items('Each_nominee')?['relationship']` |
| IdentityNumber | `items('Each_nominee')?['identity_number']` |
| BenefitCurrentPct | `items('Each_nominee')?['benefit_current_pct']` |
| BenefitRevisedPct | `items('Each_nominee')?['benefit_revised_pct']` |

**5c. Apply to each** – name `Each payee`, over `body('Parse_nomination')?['payees']`
→ **Create item** in *Nomination Court Orders*: Title `items('Each_payee')?['full_name']`, Member Id `outputs('Create_member')?['body/ID']`, MemberNumber, RowNumber `items('Each_payee')?['row']`, Relationship, IdentityNumber, BenefitPct `items('Each_payee')?['benefit_pct']`, MonthlyMaintenance `items('Each_payee')?['monthly_maintenance']`.

**5d. SharePoint: Get item** – name `Get member back`, list *Nomination Members*, Id `outputs('Create_member')?['body/ID']`
(Confirms the data URI was stored in full.)

**5e. Compose** – name `Data URI stored OK`
`equals(length(string(outputs('Get_member_back')?['body/PdfDataUri'])), length(outputs('PDF_Data_URI')))`

**5f. SharePoint: Update file properties** – library *Nomination Submissions*, Id `triggerOutputs()?['body/ID']`

| Column | Expression |
|---|---|
| Status | `if(and(outputs('Data_URI_stored_OK'), empty(body('Parse_nomination')?['warnings'])), 'Processed', 'Needs review')` |
| RejectReason | `concat(join(body('Parse_nomination')?['warnings'], decodeUriComponent('%0A')), if(outputs('Data_URI_stored_OK'), '', concat(decodeUriComponent('%0A'), 'PDF data URI was not stored in full.')))` |
| MemberItemId | `outputs('Create_member')?['body/ID']` |

### If no

**SharePoint: Update file properties** – Id `triggerOutputs()?['body/ID']`
Status `Rejected`, RejectReason `join(body('Parse_nomination')?['errors'], decodeUriComponent('%0A'))`

Optional: **Outlook: Send an email (V2)** to `triggerOutputs()?['body/Author/Email']` with the reasons and a request to re-sign in Adobe Acrobat Reader.

### Error handling

After `Try`, add a **Scope** named `Catch` → *Configure run after*: `Try` **has failed** / **has timed out**.
Inside it: **Update file properties** – Status `Needs review`, RejectReason `concat('Flow failed on ', utcNow(), '. See the flow run history.')`.

### Things to know

- **Licence:** every action in this flow is a standard connector (SharePoint, Excel Online (Business), Data Operations, Control, Outlook).
- **Resubmissions** create a new member item each time; in Power Apps show the newest per `MemberNumber` (see below).
- **Size:** the blank form is ~117 KB (a data URI of ~160,000 characters). Each submission adds the size of the signature image the member attached: a small scan adds little, a full-resolution phone photo can add several MB. `Data URI stored OK` flags any submission whose data URI SharePoint didn't store in full. Ask members to crop their signature photo.

## 4. Power Apps

Galleries only load the columns they use, so keep `PdfDataUri` out of gallery templates and load it for the selected item only.

```powerfx
// Members gallery – newest first
galMembers.Items = Sort('Nomination Members', Created, SortOrder.Descending)

// PDF viewer – fetches only the data URI of the selected item
pdfViewer.Document = LookUp('Nomination Members', ID = galMembers.Selected.ID, PdfDataUri)

// Nominees / court orders of the selected submission
galNominees.Items   = Sort(Filter('Nomination Nominees', Member.Id = galMembers.Selected.ID), RowNumber)
galCourtOrders.Items = Sort(Filter('Nomination Court Orders', Member.Id = galMembers.Selected.ID), RowNumber)

// Open the authentic PDF in SharePoint
btnOpenOriginal.OnSelect = Launch(galMembers.Selected.SourceFileUrl)

// Latest submission for one member
LookUp(Sort(Filter('Nomination Members', MemberNumber = txtMember.Text), Created, SortOrder.Descending), true)
```

## 5. Test plan

1. Upload `Beneficiary Nomination Form - Sample Filled.pdf` → Status **Needs review** (no signature image, by design), 1 member, 3 nominees, 1 court order, PDF opens in Power Apps.
2. Sign a copy in Adobe Acrobat Reader and upload → Status **Processed**.
3. Upload the blank `Beneficiary Nomination Form - Fillable.pdf` → **Rejected** (not signed, no member, no nominees).
4. Upload any other PDF → **Rejected** (not the fillable form).
