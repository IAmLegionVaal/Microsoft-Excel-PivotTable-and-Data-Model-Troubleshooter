# Microsoft Excel PivotTable and Data Model Troubleshooter

Created by **Dewald Pretorius**.

`Troubleshooter.ps1` collects PivotTable, Power Pivot, Data Model, relationship, refresh, and calculation evidence. `Repair.ps1` adds guarded `Diagnose`, `ResetModelCaches`, and `RepairOffice` actions.

```powershell
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetModelCaches -WhatIf
.\Repair.ps1 -Action RepairOffice -Confirm
```

Excel and Mashup processes must be closed before cache repair. Existing model and query caches are preserved as timestamped backups. Microsoft 365 Apps Quick Repair may require elevation. Source-reviewed for Windows PowerShell 5.1; not runtime-tested against every workbook or data connector.
