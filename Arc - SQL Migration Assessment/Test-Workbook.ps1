[CmdletBinding()]
param([string]$Path)

$ErrorActionPreference = 'Stop'
if (-not $Path) {
    $files = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.workbook' -File)
    if ($files.Count -ne 1) { throw 'Supply -Path when the folder does not contain exactly one workbook.' }
    $Path = $files[0].FullName
}
$document = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path) | ConvertFrom-Json
$script:checks = 0

function Assert-Contract([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Get-WorkbookItems($Items) {
    foreach ($item in $Items) {
        $item
        if ($item.content.items) { Get-WorkbookItems $item.content.items }
    }
}

$items = @(Get-WorkbookItems $document.items)
$byName = @{}
foreach ($item in $items) {
    Assert-Contract (-not [string]::IsNullOrWhiteSpace($item.name)) 'Item name missing.'
    Assert-Contract (-not $byName.ContainsKey($item.name)) "Duplicate item: $($item.name)"
    $byName[$item.name] = $item.content
}
Assert-Contract ($document.version -eq 'Notebook/1.0') 'Unexpected workbook version.'
$parameters = @($items | Where-Object type -EQ 9 | ForEach-Object { $_.content.parameters })
Assert-Contract (@($parameters.name | Select-Object -Unique).Count -eq $parameters.Count) 'Duplicate parameter name.'
Assert-Contract (@($parameters.id | Select-Object -Unique).Count -eq $parameters.Count) 'Duplicate parameter ID.'
$queries = @($items | Where-Object type -EQ 3 | ForEach-Object { $_.content }) + @($parameters | Where-Object query)
foreach ($query in $queries) {
    Assert-Contract ($query.queryType -eq 1 -and $query.resourceType -eq 'microsoft.resourcegraph/resources') 'Query must use ARG.'
    Assert-Contract ($query.crossComponentResources -contains '{Subscriptions}') 'Subscription scope missing.'
    foreach ($reference in [regex]::Matches($query.query, '\{([A-Za-z][A-Za-z0-9]*)\}')) {
        Assert-Contract ($parameters.name -contains $reference.Groups[1].Value) "Undefined parameter: $reference"
    }
}
$tabs = $byName['tabs-selector'].parameters[0].jsonData | ConvertFrom-Json
Assert-Contract (@($tabs).Count -eq 7) 'Expected seven tabs.'
Assert-Contract ($tabs.value -notcontains 'migration') 'The separate Assessment Health tab must be removed.'
Assert-Contract (($tabs | Where-Object value -EQ readiness).label -eq 'Assessment & Arc Readiness') 'Consolidated readiness tab label missing.'
foreach ($tab in $tabs) {
    Assert-Contract (@($items | Where-Object { $_.conditionalVisibility.parameterName -eq 'SelectedTab' -and $_.conditionalVisibility.value -eq $tab.value }).Count -eq 1) "Tab has no unique group: $($tab.value)"
}

$readiness = $byName['ready-grid']
Assert-Contract ($readiness.visualization -eq 'table' -and $readiness.showExportToExcel -and $readiness.gridSettings.filter) 'Readiness must reuse the filterable/exportable table.'
foreach ($column in @('Server Name', 'SQL Instance Name', 'Overall Readiness', 'Action Required', 'Azure Arc Connected', 'SQL Server Extension Status', 'Inventory Reporting', 'Last Successful Inventory Upload', 'Migration Assessment Enabled', 'Assessment Freshness', 'Last Migration Report Upload', 'MI Readiness', 'DB Readiness', 'VM Readiness', 'Pricing Status', 'Prerequisite Status', 'Host OS', 'Best Practices Assessment Enabled', 'Next Action', 'Migration Assessment Results', 'SQL Resource', 'Host Resource')) {
    Assert-Contract ($readiness.query.Contains("['$column'] = ")) "Readiness column missing: $column"
}
Assert-Contract ($readiness.query.Contains('join kind=fullouter')) 'SQL-detected hosts without instances must be retained.'
Assert-Contract (($parameters | Where-Object name -EQ ResourceGroup).query.Contains('properties.detectedProperties.mssqldiscovered')) 'Resource group selector must include SQL-detected hosts.'
Assert-Contract ($readiness.query.Contains('isnotempty(ResourceId) or SQLDetected == true')) 'Do not classify non-SQL hosts as SQL instances.'
Assert-Contract ($readiness.query.Contains('summarize ExtensionCount = countif(not(IsHost))')) 'Extension aggregation must count extensions only and not multiply instance rows.'
Assert-Contract ($readiness.query.Contains('ExtensionCount != 1')) 'Multiple extensions must not silently become healthy.'
Assert-Contract ($readiness.query.Contains("ScopeTags = iff(isnotempty(ResourceId), InstanceTags, HostTags)")) 'Tag scope must use instances with host-only fallback.'
Assert-Contract ($readiness.query.Contains("tobool(properties.bestPracticesAssessment.enabled)")) 'Use the documented BPA enabled flag.'
Assert-Contract ($readiness.query.Contains("InventoryTime < ago(3d)")) 'Old inventory is not current reporting.'
Assert-Contract ($readiness.query.Contains("InventoryTime > now()")) 'Future inventory must not be healthy.'
Assert-Contract ($readiness.query.Contains("ExtensionInstalled = iff(ExtensionCount > 0, 'Yes', 'Unknown')")) 'Missing extension visibility must not be labelled No.'
Assert-Contract ($readiness.query.Contains("ExtensionTime >= ago(3d) and ExtensionTime <= now() and ArcConnected == 'Yes'")) 'Healthy extension requires recent status and host connectivity.'
Assert-Contract ($readiness.query.Contains("UploadStatus =~ 'OK'")) 'Healthy extension requires successful upload evidence.'
Assert-Contract ([regex]::Matches($readiness.query, '\| join ').Count -eq 1) 'Readiness must not repeat the resources table as a right-side join.'
Assert-Contract ($readiness.query.Contains("isnull(ExtensionCount) or ExtensionCount == 0, 'Unknown'")) 'An empty extension collection must not be labelled multiple or absent.'
Assert-Contract ($readiness.query.Contains("'Attention')`n| extend Reporting = ")) 'Reporting must reference ExtensionStatus only after its extend statement.'
Assert-Contract ($readiness.query -notmatch '\| mv-expand |\| union ') 'Readiness must not multiply SQL instance rows through expansion.'
Assert-Contract ($readiness.query -notmatch 'asmt\.status|bestPracticesAssessment\.status|migration\.assessment\.status') 'Do not invent ARG assessment lifecycle fields.'
Assert-Contract ($readiness.query.Contains("not(HasInstance), 'Not applicable'")) 'Host-only assessment fields must be explicitly not applicable.'
Assert-Contract ($readiness.query.Contains("SQLMajorVersion = toint(split(CurrentVersion, '.')[0])")) 'Prerequisite checks must use the documented currentVersion build.'
Assert-Contract ($readiness.query.Contains("ReadinessSeverity = case(not(HasInstance), 1, PrerequisiteStatus startswith 'Unsupported', 2, ArcConnected == 'No', 3")) 'Readiness severity must preserve the documented triage priority.'
Assert-Contract ($readiness.query.Contains('| order by ReadinessSeverity asc, ServerName asc, SQLInstanceName asc')) 'Readiness must sort by severity, server, and instance.'
Assert-Contract ($readiness.query.IndexOf('| order by ReadinessSeverity asc') -lt $readiness.query.IndexOf("| project ['Server Name']")) 'Internal severity ordering must occur before the rank is removed from display.'
Assert-Contract (-not $readiness.gridSettings.sortBy) 'Grid sorting must not override query severity order.'
Assert-Contract ($readiness.query.Contains("not(HasInstance), 'https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/connect', PrerequisiteStatus startswith 'Unsupported', 'https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/migration-assessment?view=sql-server-ver17', ArcConnected != 'Yes'")) 'Unsupported prerequisites must have a priority guidance action.'
Assert-Contract (($readiness.gridSettings.formatters | Where-Object columnMatch -EQ 'Next Action').formatOptions.linkTarget -eq 'Url') 'Consolidated next action link missing.'
foreach ($format in $readiness.gridSettings.formatters | Where-Object formatter -EQ 18) {
    Assert-Contract (($format.formatOptions.thresholdsGrid | Where-Object operator -EQ Default).representation -eq 'yellow') 'Unverified readiness evidence must require attention.'
}
$overallFormat = $readiness.gridSettings.formatters | Where-Object columnMatch -EQ 'Overall Readiness'
$prerequisiteFormat = $readiness.gridSettings.formatters | Where-Object columnMatch -EQ 'Prerequisite Status'
foreach ($value in @('Unsupported: Linux host', 'Unsupported: SQL Server before 2014')) {
    Assert-Contract (($overallFormat.formatOptions.thresholdsGrid | Where-Object thresholdValue -EQ $value).representation -eq 'red') "Overall readiness must mark '$value' red."
    Assert-Contract (($prerequisiteFormat.formatOptions.thresholdsGrid | Where-Object thresholdValue -EQ $value).representation -eq 'red') "Prerequisite status must mark '$value' red."
}
Assert-Contract (($prerequisiteFormat.formatOptions.thresholdsGrid | Where-Object thresholdValue -EQ 'Not applicable').representation -eq 'gray') 'Host-only prerequisite status must be gray.'
foreach ($column in @('Migration Assessment Enabled', 'Best Practices Assessment Enabled')) {
    $format = $readiness.gridSettings.formatters | Where-Object columnMatch -EQ $column
    Assert-Contract (($format.formatOptions.thresholdsGrid | Where-Object thresholdValue -EQ No).representation -eq 'gray') 'Disabled assessments must be gray, not failed.'
}
Assert-Contract (($readiness.gridSettings.formatters | Where-Object columnMatch -EQ 'Migration Assessment Results').formatOptions.linkTarget -eq 'Resource') 'Use the existing resource navigation for migration results.'
foreach ($column in @('SQL Resource', 'Host Resource')) {
    Assert-Contract (($readiness.gridSettings.formatters | Where-Object columnMatch -EQ $column).formatOptions.linkTarget -eq 'Resource') "$column resource link missing."
}

$prefix = ($byName['tco-grid'].query -split '\n\| project')[0]
Assert-Contract (($byName['tco-kpi'].query -split '\n\| summarize')[0] -ceq $prefix) 'KPI and detail selection differ.'
Assert-Contract (($byName['rec-grid'].query -split '\n\| extend impacted')[0] -ceq $prefix) 'Recommendations and comparison selection differ.'
$recommendations = $byName['rec-grid'].query
$beforeExpansion = ($recommendations -split '\n\| mv-expand Targets')[0]
$projection = ($beforeExpansion -split '\n' | Where-Object { $_ -like '| project *' })
Assert-Contract (@($projection).Count -eq 1) 'Recommendations must narrow the resource row before target expansion.'
Assert-Contract ($projection.StartsWith('| project name, id, resourceGroup, subscriptionId, Currency, Strategy, PreferredTarget, ComparisonStatus, AssessmentEnabled = ')) 'Expansion must retain identity and selection fields only.'
Assert-Contract (-not $projection.Contains("'Impacted',")) 'Do not duplicate impacted-object arrays during expansion.'
Assert-Contract ([regex]::Matches($projection, "'ImpactedObjectCount',toint\(array_length\(impacted\.").Count -eq 3) 'Count impacted objects before expanding all three targets.'
$afterExpansion = ($recommendations -split '\n\| mv-expand Targets')[1]
Assert-Contract ($afterExpansion -notmatch '\b(asmt|properties|impacted|r)\.') 'Expanded query references a discarded resource field.'
foreach ($disk in @(
        @{ Column = 'DataDisk'; Array = 'dataDisks'; Count = 'dataDiskCount' },
        @{ Column = 'LogDisk'; Array = 'logDisks'; Count = 'logDiskCount' },
        @{ Column = 'TempDbDisk'; Array = 'tempDbDisks'; Count = 'tempDbDiskCount' }
    )) {
    $summary = @($recommendations -split '\n' | Where-Object { $_.StartsWith("| extend $($disk.Column) = ") })
    Assert-Contract ($summary.Count -eq 1) "Missing summary for $($disk.Column)."
    Assert-Contract ($recommendations.Contains("$($disk.Count) = array_length($($disk.Array))")) 'Disk count must use the full array.'
    Assert-Contract ($summary[0].Contains("case(isnull($($disk.Count)), 'Unknown', $($disk.Count) == 0, 'None reported'")) 'Missing and empty disk arrays must differ.'
    Assert-Contract ($summary[0].Contains("iff($($disk.Count) == 1, '1 disk: ', strcat($($disk.Count), ' disks; first: '))")) 'Multiple disks must explicitly label first-disk details.'
    foreach ($field in @('size', 'maxSizeInGib', 'maxIops', 'maxThroughputInMbps', 'caching')) {
        Assert-Contract ($summary[0].Contains("coalesce(tostring($($disk.Array)[0].$field), 'Unknown")) "Missing field fallback: $($disk.Column).$field"
    }
    Assert-Contract (-not $recommendations.Contains("$($disk.Column) = tostring($($disk.Array))")) 'Do not render disk arrays as raw JSON.'
}
Assert-Contract ([regex]::Matches($recommendations, '\| mv-expand ').Count -eq 1) 'Disk summaries must not multiply recommendation rows.'
foreach ($target in @('MI', 'DB', 'VM')) {
    Assert-Contract ($prefix -match "$($target)ReadyCost = iff\(tobool\(asmt.enabled\) == true.*?recommendationStatus\) =~ 'Ready' and $($target)Cost >= 0") "$target cost bypasses eligibility."
}
Assert-Contract ($prefix.Contains('BaselineReady = max_of(MIReadyCost, DBReadyCost, VMReadyCost)')) 'Baseline includes ineligible prices.'
Assert-Contract ($prefix.Contains("Currency == 'Unknown', 'Unknown currency'")) 'Unknown currency must be excluded.'
Assert-Contract ($byName['rec-grid'].query.Contains("IsPreferredByStrategy = ComparisonStatus == 'Comparable' and Target == PreferredTarget")) 'Preferred marker must require comparable target.'
Assert-Contract ($byName['tco-kpi'].query -match 'by Currency') 'Comparison aggregates currencies together.'
Assert-Contract (-not $byName.ContainsKey('ov-card-cost')) 'Removed overview monthly-cost item must not return.'
$titles = @($items | Where-Object { $_.content.title } | ForEach-Object { $_.content.title })
Assert-Contract ($titles -notcontains 'Monthly costs by instance and target (all rows; eligibility shown)') 'Removed monthly-cost title must not return.'
foreach ($name in @('rec-grid', 'prc-grid', 'tco-kpi', 'tco-grid')) {
    Assert-Contract ($byName[$name].query.Contains('{CurrencyFilter}')) "$name ignores the currency filter."
}
Assert-Contract ($byName['tco-grid'].query.Contains("AHBEntitlement = 'Not verified'")) 'License entitlement must not be inferred.'
Assert-Contract ($byName['tco-grid'].query.Contains('Unavailable: matched license baseline required')) 'Unsupported AHB savings claim.'

$pricing = $byName['prc-grid'].query
Assert-Contract ($pricing.Contains('dynamic([null])')) 'Empty option arrays must retain a row.'
Assert-Contract ($pricing.Contains('OptCompute >= 0 and OptStorage >= 0 and OptIops >= 0')) 'Incomplete prices must not become zero.'
foreach ($guard in @('PAYGRows == 1 and PAYGValid == 1', 'RI1Rows == 1 and RI1Valid == 1', 'RI3Rows == 1 and RI3Valid == 1')) {
    Assert-Contract ($pricing.Contains($guard)) "Missing duplicate/validity guard: $guard"
}
Assert-Contract ($pricing.Contains('Lowest = min_of(PAYG, RI_1Yr, RI_3Yr)')) 'Cheapest must compare valid options.'
Assert-Contract ($pricing.Contains('PAYG > 0 and isnotnull(RI_3Yr)')) 'Percentage must guard zero and missing prices.'
foreach ($name in @('prc-grid', 'tco-grid')) {
    foreach ($format in $byName[$name].gridSettings.formatters | Where-Object numberFormat) {
        Assert-Contract ($format.numberFormat.emptyValCustomText -eq 'Unavailable') "$name missing-price label absent."
    }
}

$instance = @($parameters | Where-Object name -EQ InstanceFilter)
Assert-Contract ($instance.Count -eq 1 -and $instance[0].query.Contains('value = tolower(id)')) 'Instance selection must use a shared full-ID parameter.'
Assert-Contract ($byName['inv-grid'].exportFieldName -eq 'ResourceId') 'Inventory must export full resource IDs.'
Assert-Contract ($byName['inv-grid'].exportParameterName -eq 'InstanceFilter') 'Inventory exports to wrong parameter.'
$inventory = $byName['inv-grid'].query
Assert-Contract ($inventory.Contains("SQLProductVersion  = coalesce(tostring(p.currentVersion), 'Unknown')")) 'Inventory must use the documented currentVersion build.'
Assert-Contract ($inventory -notmatch 'p\.productVersion') 'Inventory references undocumented productVersion.'
Assert-Contract ($inventory.Contains("indexof(tolower(id), '/extensions/')")) 'Inventory extension parent extraction must be case-insensitive.'
$database = $byName['db-grid'].query
Assert-Contract ($database.Contains("ParentId =~ '{InstanceFilter}'")) 'Database selection must compare IDs case-insensitively.'
Assert-Contract ($database.Contains('join kind=inner')) 'Database inventory must join scoped parents.'
Assert-Contract ($database.IndexOf("tags['{TagName}']") -gt $database.IndexOf('join kind=inner')) 'Tag scope belongs to parent instances.'

Assert-Contract ($byName['ov-card-coverage'].query.Contains('countif(isnull(Enabled))')) 'Unknown assessment status must be counted.'
$freshness = ($byName['ov-card-freshness'].query -split '\n' | Where-Object { $_ -like '| extend FreshnessBucket*' })
$healthFreshness = ($readiness.query -split '\n' | Where-Object { $_ -like '| extend AssessmentFreshness*' })
Assert-Contract ($healthFreshness.Count -eq 1) 'Consolidated assessment freshness definition missing.'
Assert-Contract ($freshness.IndexOf('AgeDays > Threshold') -lt $freshness.IndexOf('AgeDays <= 7')) 'Stale threshold must precede seven-day bucket.'
foreach ($state in @('Invalid threshold', 'No data', 'Invalid timestamp', 'Future timestamp', 'Stale', 'Fresh', 'Recent')) {
    Assert-Contract ($freshness.Contains("'$state'")) "Missing freshness state: $state"
    Assert-Contract ($healthFreshness.Contains("'$state'")) "Consolidated freshness state missing: $state"
}
foreach ($name in @('rec-grid')) {
    $format = $byName[$name].gridSettings.formatters | Where-Object columnMatch -EQ AssessmentEnabled
    Assert-Contract (($format.formatOptions.thresholdsGrid | Where-Object operator -EQ Default).text -eq 'Unknown') "$name formats unknown assessment status incorrectly."
}
$threshold = ($parameters | Where-Object name -EQ FreshnessThresholdDays).typeSettings.paramValidationRules[0].regExp
foreach ($valid in @('1', '3', '7', '30', '9999')) { Assert-Contract ($valid -match $threshold) "Valid threshold rejected: $valid" }
foreach ($invalid in @('', '0', '-1', '2.5', '10000', 'abc')) { Assert-Contract ($invalid -notmatch $threshold) "Invalid threshold accepted: $invalid" }
foreach ($parameter in $parameters | Where-Object name -In @('TagName', 'TagValue')) {
    $pattern = $parameter.typeSettings.paramValidationRules[0].regExp
    foreach ($valid in @('', 'Cost Center', 'prod:west')) { Assert-Contract ($valid -match $pattern) 'Valid tag rejected.' }
    foreach ($invalid in @("owner's", 'path\name', "line`nnext")) { Assert-Contract ($invalid -notmatch $pattern) 'Unsafe tag accepted.' }
}
$text = ($items | Where-Object type -EQ 1 | ForEach-Object { $_.content.json }) -join "`n"
Assert-Contract ($text -notmatch 'Azure Data Studio|sum as-is|AHBEligible|RISavings_PerMonth') 'Obsolete guidance remains.'
Write-Output "PASS: $script:checks local structure and query-contract checks. Live ARG execution and portal rendering NOT tested."