# Arc SQL Migration Assessment Workbook

Standalone, read-only Azure Workbook using Azure Resource Graph (ARG). No resources,
assessment settings, tags, or licensing configuration are changed by these queries.

## Implemented First Phase

- Missing, incomplete, duplicate, and negative pricing options are unavailable;
  explicit zero prices remain valid.
- Recommendations and Cost Comparison share the same Ready, priced target selection.
- Assessment & Arc Readiness combines connectivity, extension and inventory health,
  freshness, partial prerequisite checks, pricing availability, and MI/DB/VM readiness.
- Unknown assessment state is distinct from disabled; freshness categories are exclusive.
- Database scoping uses full parent resource IDs and parent-instance tags.
- Updated Arc portal compatibility guidance; unsupported AHB savings removed.
- Disk columns show count, SKU, capacity, IOPS, throughput, and caching instead of JSON.
  For multiple disks, details are explicitly for the first disk, not aggregate totals
  or a claim that all disks are identical. Empty arrays show `None reported`; missing
  arrays or attributes show `Unknown`.
- Inventory reads the documented `properties.currentVersion` SQL build field and shows
  `Unknown` when the build is not reported; `properties.version` remains the product family.

## Assessment & Arc Readiness

The consolidated readiness tab reuses the workbook's authenticated ARG connection, subscription scope,
resource-group/tag parameters, table filtering, resource links, and Excel export. No
API service, credentials, Azure writes, or deployment are added. Currency and the
Databases-only instance selector do not filter this non-financial fleet view.

Each visible SQL instance has one row. A full-outer host join also retains Arc machines
with SQL detected but no visible instance. These are explicitly labelled
`Not discovered / not visible`, not evidence of SQL absence. Instance resource groups
and tags control instance rows; host groups and tags control host-only rows. Hosts with
neither visible instances nor a SQL-detection flag cannot be identified as SQL servers.
Hosts and extensions are grouped into one lookup before the single full-outer join,
avoiding repeated right-table references. Extension aggregation avoids multiplying
rows; multiple extensions require attention.

The table is intentionally focused. It combines overall/actionable status, Arc
connectivity, SQL extension health, inventory reporting, assessment freshness,
MI/DB/VM readiness, pricing availability, prerequisites, BPA enablement, timestamps,
and navigation. Duplicate installation/provisioning flags, derived BPA status, combined
timestamps, and four separate action-link columns were removed. `Next Action` selects
the highest-priority guidance link.

Status evidence and limits:

- Arc connection uses the host's reported status. Missing host/extension visibility is
  Unknown, never proof of No/Not Installed. Extension Installed means an extension
  resource is visible, not that provisioning or the agent is healthy.
- Extension Healthy requires Succeeded provisioning, a Healthy agent message, OK upload,
  an extension status date within three days, and a connected host. Running describes
  Creating/Updating/Transitioning provisioning, not proof of a running SQL service.
- Reporting uses the documented last successful instance inventory upload timestamp,
  with a three-day reporting window and guards for failed uploads, connectivity, and
  missing/invalid/future evidence. This is separate from assessment freshness. Agent
  status dates are parsed from the documented message in either YYYY/MM/DD or YYYY-MM-DD
  form, at midnight UTC; unrecognized message formats remain unverified.
- Migration and BPA enabled flags are read from their respective instance properties;
  absent flags remain Unknown. An explicit false maps to Not Configured.
- **Running/Completed/Failed assessment job states are not exposed in the verified ARG
  resource schema.** The run APIs return job status separately. This read-only view
  deliberately does not invoke run APIs or invent lifecycle properties. Migration shows
  report availability/freshness, not latest-run success. BPA shows enablement and directs
  users to the portal/Log Analytics for execution history and results. Older ARG payloads
  may not expose the BPA flag at all.
- Last Successful Inventory Upload and Last Migration Report Upload remain separate;
  neither is a BPA execution time. Malformed or future values never establish healthy
  reporting or a current assessment.
- Green marks positive configuration/health evidence; yellow marks attention or unknown
  evidence; red marks explicit failure/unhealthy states; gray marks not configured.
  A green enabled flag alone does not establish successful assessment execution.

Migration Results opens the SQL resource only when a valid report-upload timestamp is
present, including historical reports retained after disabling assessment. From there,
choose **Migration > Assessments**. The SQL resource link also provides access to
**Best practices assessment**. Unverified internal portal blade routes are not guessed.
`Next Action` opens the highest-priority relevant Microsoft Learn guidance for discovery,
unsupported prerequisites, Arc or extension health, reporting, migration assessment,
or BPA. It is navigation only. Enabling BPA requires appropriate licensing, a Log
Analytics workspace, agent configuration, and permissions, and applies to all instances
on the host. Both assessments have Windows prerequisites; other prerequisites still
require checking even when connectivity is healthy.

References:

- [Migration assessment and prerequisites](https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/migration-assessment?view=sql-server-ver17)
- [Best practices assessment configuration and results](https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/assess)
- [SQL extension health and reporting signals](https://learn.microsoft.com/en-us/sql/sql-server/azure-arc/troubleshoot-extension)
- [SQL instance resource schema, API 2026-01-01](https://github.com/Azure/azure-rest-api-specs/blob/main/specification/azurearcdata/resource-manager/Microsoft.AzureArcData/AzureArcData/stable/2026-01-01/sqlServerInstances.json)

## Use

1. In Azure Monitor Workbooks, create a workbook and open Advanced Editor.
2. Replace the workbook JSON with the adjacent `.workbook` file and apply it.
3. Select subscriptions you can read. The original subscription default is retained.
4. Select resource groups, tags, and currencies. The SQL instance selector scopes only
   Databases. Inventory row selection exports its full resource ID; select the Databases
   tab manually. Automatic tab navigation is not implemented.
5. Use the resource links to open an Arc SQL instance, then Migration > Assessments
   for compatibility findings. Access to detailed reports may require telemetry permissions.

The workbook has seven tabs. Assessment Health and Arc SQL Readiness were merged into
Assessment & Arc Readiness so connectivity and assessment evidence can be triaged in
one row without duplicate identity, agent, timestamp, or status columns.

Tag filters currently reject apostrophes, backslashes, and control characters because
their values are interpolated into KQL literals. This input validation is not a security
boundary. Re-select an instance or All after changing subscription, group, or tag scope
if the prior selection no longer belongs to the scope.

## Cost Semantics

Cost Comparison is an alternative Azure target price difference, not on-premises TCO
or realized savings. The baseline is the highest Ready, non-negative, known target
price. Enabled assessments using `MigrateToPaaS` select the cheapest Ready, priced MI
or DB (MI wins ties); `MigrateToIaaS` selects a Ready, priced VM. Other strategies and
unknown currencies are explicitly excluded from totals. These are workbook rules,
not a claim about undocumented assessment-engine strategy values.

Pricing Options compares compute, storage, and IOPS for three known PAYG/RI keys,
excluding licenses. Missing components are not assumed free. Stale assessments are
retained, so prices may be outdated. Annual/three-year figures assume unchanged usage
and prices. AHB entitlement is not verified; savings require a matched license baseline.

## Verification

Run `./Test-Workbook.ps1` from this folder with PowerShell 5.1 or later. It checks JSON,
item/parameter contracts, shared query clauses, financial guards, and input validation.
It does not execute KQL, validate the complete Azure schema, or simulate Azure rendering.

Before publishing, test in the portal with authorized, representative resources:

- All seven tabs, empty scope, parameter initialization, resource links, and row export.
- Same-name instances in different groups/subscriptions; parent-only tags; scope changes.
- USD/EUR separation and currency filtering in every monetary view.
- PAYG 100, missing RI1, RI3 70: RI1 unavailable and RI3 cheapest; explicit zero valid;
  null components, duplicate keys, and empty option arrays unavailable.
- Ready MI 1000, NotReady DB 1200, Ready VM 800: baseline 1000; disabled, unknown,
  unpriced, and unsupported-strategy rows excluded. Reconcile detail and KPI by currency.
- Thresholds 3/7/30, missing/malformed/future timestamps, and unknown enabled state.
- Inventory SQL product builds match ARG `properties.currentVersion`; missing builds
  display `Unknown` while SQLVersion continues to show the product family.
- Host/extension joins and partial prerequisites against actual ARG payloads and RBAC.
- Consolidated readiness: connected/disconnected/expired hosts, missing or multiple extensions,
  failed/updating provisioning, OK/failed/missing upload status, and absent/malformed/
  future/stale inventory and extension timestamps. Test both extension date formats.
- Consolidated readiness: enabled/disabled/missing assessment flags, historical migration reports
  retained after disablement, Linux hosts, SQL-detected hosts without visible instances,
  same-name resources across subscriptions, host-only tag/group scope, and all action
  links. Confirm Unknown remains attention, not failed or not installed.
- Check the readiness join and aggregation in the portal; static
  checks do not prove executability or completeness under the caller's RBAC scope.
- Result truncation and export counts for large fleets. Grid row limits do not guarantee
  complete ARG results; narrow the scope and reconcile counts before financial decisions.

Live ARG and portal validation have not been performed. History, migration waves,
additional pricing scenarios, true TCO, and automatic navigation remain later phases.

### Recommendations Query Error

After the reported 2026-09-22 execution error, Recommendations was changed to project
only required fields before target expansion and to count impacted objects before
expanding. This restores the earlier narrow-row query shape without changing selection
rules. It is a mitigation, not a confirmed root-cause fix: tenant policy blocked the
Copilot ARG diagnostic tool, so only local contract checks could be run.

Reapply the updated workbook JSON in Advanced Editor; refreshing an already imported
workbook does not load local file changes. If the error persists, run the Recommendations
query in the workbook query editor and retain its full error details. Check whether
Cost Comparison also fails (shared selection logic) and whether selecting one target
and a smaller resource-group scope changes the outcome. Include these observations
and the timestamp/correlation ID when escalating to Azure support.
