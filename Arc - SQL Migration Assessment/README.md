# Arc SQL Migration Assessment Workbook

Read-only Azure Workbook for Arc-enabled SQL Server inventory, assessment health,
recommendations, and cost comparison. Azure Resource Graph (ARG) queries do not change
resources, settings, tags, or licensing.

## Functions

The workbook has seven tabs:

- **Overview** shows assessment coverage, freshness, target readiness, and assessment
  configuration.
- **Inventory** lists SQL instances, hosts, extensions, versions, and upload times.
  SQL builds use `properties.currentVersion`; missing builds
  display `Unknown`.
- **Assessment & Arc Readiness** combines connectivity, extension and inventory health,
  migration freshness, MI/DB/VM readiness, pricing, prerequisites, and links. SQL-detected
  hosts without a visible instance remain `Not discovered / not visible`.
- **Recommendations** displays target readiness, preferred targets, SKU details, costs,
  blocker counts, and disk configuration.
- **Databases** lists database inventory and parent-instance readiness. Compatibility
  findings remain in the Arc **Migration > Assessments** report.
- **Pricing Options** compares reported PAYG, one-year RI, and three-year RI compute,
  storage, and IOPS prices. Missing, duplicate, incomplete, or negative options are
  unavailable; explicit zero prices remain valid.
- **Cost Comparison** compares Ready Azure targets by currency. It is not on-premises
  TCO or realized savings. The monthly difference is the highest-cost Ready Azure
  alternative minus the selected preferred target; 12- and 36-month values multiply
  that gap by 12 and 36 with unchanged prices and usage.

## Readiness Semantics

Inventory reporting requires a successful upload within three days. Migration freshness
uses the selected threshold. Unknown means evidence is absent or unverified, not failed.
Green indicates healthy evidence, yellow attention/unknown, red failure or unsupported,
and gray not configured/not applicable.

ARG does not expose current migration-assessment job status. Timestamps prove only that data
was uploaded. `Next Action` opens priority guidance; resource links open the Arc host or
SQL instance. Detailed reports may require additional permissions.

## Cost Semantics

For `MigrateToPaaS`, the workbook selects the cheapest Ready, priced MI or DB target
(MI wins ties). For `MigrateToIaaS`, it selects a Ready, priced VM. Disabled assessments,
unknown currencies, unsupported strategies, and unpriced targets are excluded. Stale
assessment prices remain visible and may be outdated. AHB entitlement and savings are
not inferred without a matched license-included baseline and verified licensing records.
`Not verified` means eligibility is unknown, not that the organization is ineligible.

## Use

1. Create an Azure Monitor Workbook and open **Advanced Editor**.
2. Paste the adjacent `.workbook` JSON and apply it.
3. Select subscriptions, resource groups, optional tags, freshness threshold, and currency.
4. Use table filters, Excel export, and resource links for investigation.

The SQL instance selector scopes only the Databases tab. Tag filters reject apostrophes,
backslashes, and control characters. Re-select an instance after changing scope.

## Validation

Run `./Test-Workbook.ps1`. It validates JSON structure,
parameters, query contracts, financial guards, and formatters, but not portal rendering
or tenant-specific data completeness.

Before publishing, verify all tabs with connected/disconnected hosts, missing or multiple
extensions, fresh/stale/missing timestamps, enabled/disabled assessments, Ready/NotReady
targets, currencies, host-only rows, links, exports, and result truncation. Reapply local
JSON through Advanced Editor; refreshing an imported workbook does not load local changes.

References: [Migration assessment](https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/migration-assessment?view=sql-server-ver17),
[SQL extension troubleshooting](https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/troubleshoot-extension).
