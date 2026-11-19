<#
Creates the SharePoint library and lists for Beneficiary Nomination processing.

Requires PnP.PowerShell 2.x+ (PowerShell 7):  Install-Module PnP.PowerShell -Scope CurrentUser
PnP now needs your own Entra ID app registration for sign-in; see
https://pnp.github.io/powershell/articles/registerapplication.html

Usage:
  ./Create-NominationLists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Pensions -ClientId <app-id>

Safe to re-run: existing lists and columns are left as they are.
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $ClientId
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

function Ensure-List($Title, $Template, $Url) {
    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "Creating $Title"
        $list = New-PnPList -Title $Title -Template $Template -Url $Url -OnQuickLaunch
    }
    return $list
}

function Ensure-Field($List, $Name, $FieldXml) {
    if (Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue) { return }
    Write-Host "  + $Name"
    Add-PnPFieldFromXml -List $List -FieldXml $FieldXml | Out-Null
    # add to the default view (except the large data URI column)
    $view = Get-PnPView -List $List | Where-Object DefaultView
    if ($Name -ne "PdfDataUri" -and $view.ViewFields -notcontains $Name) {
        Set-PnPView -List $List -Identity $view.Id -Fields (@($view.ViewFields) + $Name) | Out-Null
    }
}

function F-Text($n, $d, [switch]$Indexed) { "<Field Type='Text' Name='$n' StaticName='$n' DisplayName='$d' MaxLength='255' $(if($Indexed){"Indexed='TRUE'"}) />" }
function F-Note($n, $d)   { "<Field Type='Note' Name='$n' StaticName='$n' DisplayName='$d' NumLines='6' RichText='FALSE' AppendOnly='FALSE' />" }
function F-Num($n, $d, $dec = 2) { "<Field Type='Number' Name='$n' StaticName='$n' DisplayName='$d' Decimals='$dec' />" }
function F-Date($n, $d)   { "<Field Type='DateTime' Name='$n' StaticName='$n' DisplayName='$d' Format='DateOnly' FriendlyDisplayFormat='Disabled' />" }
function F-Lookup($n, $d, $listId) {
    "<Field Type='Lookup' Name='$n' StaticName='$n' DisplayName='$d' List='{$listId}' ShowField='Title' Indexed='TRUE' RelationshipDeleteBehavior='Cascade' />"
}

# --- 1. Library for the submitted (signed) PDFs ---------------------------------
$lib = Ensure-List "Nomination Submissions" DocumentLibrary "NominationSubmissions"
Ensure-Field $lib "Status" "<Field Type='Choice' Name='Status' StaticName='Status' DisplayName='Status' Format='Dropdown'><Default>New</Default><CHOICES><CHOICE>New</CHOICE><CHOICE>Processed</CHOICE><CHOICE>Needs review</CHOICE><CHOICE>Rejected</CHOICE></CHOICES></Field>"
Ensure-Field $lib "RejectReason" (F-Note "RejectReason" "Reject Reason")
Ensure-Field $lib "MemberItemId" (F-Num "MemberItemId" "Member Item ID" 0)

# --- 2. Members: one item per submitted form ----------------------------------
$mem = Ensure-List "Nomination Members" GenericList "Lists/NominationMembers"
Set-PnPField -List $mem -Identity "Title" -Values @{ Title = "Member Name" } | Out-Null
Ensure-Field $mem "MemberNumber"         (F-Text "MemberNumber" "Member Number" -Indexed)
Ensure-Field $mem "FundName"             (F-Text "FundName" "Fund")
Ensure-Field $mem "DateOfBirth"          (F-Date "DateOfBirth" "Date of Birth")
Ensure-Field $mem "BranchCode"           (F-Text "BranchCode" "Branch Code")
Ensure-Field $mem "CostCentre"           (F-Text "CostCentre" "Cost Centre")
Ensure-Field $mem "HomeTelephone"        (F-Text "HomeTelephone" "Home Telephone")
Ensure-Field $mem "CellTelephone"        (F-Text "CellTelephone" "Cell Telephone")
Ensure-Field $mem "WorkTelephone"        (F-Text "WorkTelephone" "Work Telephone")
Ensure-Field $mem "EmailAddress"         (F-Text "EmailAddress" "Email Address")
Ensure-Field $mem "DeclarationFullNames" (F-Text "DeclarationFullNames" "Declaration Full Names")
Ensure-Field $mem "SignatureDate"        (F-Date "SignatureDate" "Signature Date")
Ensure-Field $mem "SignatureAttachedAt"  (F-Text "SignatureAttachedAt" "Signature Attached At")
Ensure-Field $mem "NomineeCount"         (F-Num "NomineeCount" "Nominees" 0)
Ensure-Field $mem "TotalCurrentPct"      (F-Num "TotalCurrentPct" "Total % Current")
Ensure-Field $mem "TotalRevisedPct"      (F-Num "TotalRevisedPct" "Total % Revised")
Ensure-Field $mem "FormVersion"          (F-Text "FormVersion" "Form Version")
Ensure-Field $mem "Warnings"             (F-Note "Warnings" "Warnings")
Ensure-Field $mem "SourceFileId"         (F-Num "SourceFileId" "Source File ID" 0)
Ensure-Field $mem "SourceFileName"       (F-Text "SourceFileName" "Source File Name")
Ensure-Field $mem "SourceFileUrl"        (F-Note "SourceFileUrl" "Source File URL")
Ensure-Field $mem "PdfDataUri"           (F-Note "PdfDataUri" "PDF Data URI")   # kept out of the default view

# --- 3. Nominees: one item per nominee row ------------------------------------
$nom = Ensure-List "Nomination Nominees" GenericList "Lists/NominationNominees"
Set-PnPField -List $nom -Identity "Title" -Values @{ Title = "Full Name" } | Out-Null
Ensure-Field $nom "Member"            (F-Lookup "Member" "Member" $mem.Id)
Ensure-Field $nom "MemberNumber"      (F-Text "MemberNumber" "Member Number" -Indexed)
Ensure-Field $nom "RowNumber"         (F-Num "RowNumber" "Row" 0)
Ensure-Field $nom "DateOfBirth"       (F-Date "DateOfBirth" "Date of Birth")
Ensure-Field $nom "Relationship"      (F-Text "Relationship" "Relationship")
Ensure-Field $nom "IdentityNumber"    (F-Text "IdentityNumber" "Identity Number")
Ensure-Field $nom "BenefitCurrentPct" (F-Num "BenefitCurrentPct" "% Benefit (Current)")
Ensure-Field $nom "BenefitRevisedPct" (F-Num "BenefitRevisedPct" "% Benefit (Revised)")

# --- 4. Court orders: one item per payee row ----------------------------------
$pay = Ensure-List "Nomination Court Orders" GenericList "Lists/NominationCourtOrders"
Set-PnPField -List $pay -Identity "Title" -Values @{ Title = "Payee Name" } | Out-Null
Ensure-Field $pay "Member"             (F-Lookup "Member" "Member" $mem.Id)
Ensure-Field $pay "MemberNumber"       (F-Text "MemberNumber" "Member Number" -Indexed)
Ensure-Field $pay "RowNumber"          (F-Num "RowNumber" "Row" 0)
Ensure-Field $pay "Relationship"       (F-Text "Relationship" "Relationship")
Ensure-Field $pay "IdentityNumber"     (F-Text "IdentityNumber" "Identity Number")
Ensure-Field $pay "BenefitPct"         (F-Num "BenefitPct" "% of Benefit")
Ensure-Field $pay "MonthlyMaintenance" (F-Num "MonthlyMaintenance" "Monthly Maintenance")

Write-Host "Done."
