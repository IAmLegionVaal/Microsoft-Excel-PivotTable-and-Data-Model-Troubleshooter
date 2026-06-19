#requires -Version 5.1
<# Created by Dewald Pretorius #>
param([string]$OutputPath)
if(-not $OutputPath){$OutputPath="$([Environment]::GetFolderPath('Desktop'))\Excel_Data_Model_Reports"};New-Item $OutputPath -ItemType Directory -Force|Out-Null
$events=Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddDays(-5)} -ErrorAction SilentlyContinue|Where-Object Message -match 'EXCEL|PowerPivot|Data Model|VertiPaq'|Select-Object -First 40 TimeCreated,Id,ProviderName,Message
@('EXCEL PIVOTTABLE AND DATA MODEL TROUBLESHOOTER','Created by Dewald Pretorius',"Generated: $(Get-Date)",'Check relationships, source refresh, calculated columns, measures, cache size, 32/64-bit compatibility, and workbook corruption.',($events|Format-List|Out-String -Width 220))|Set-Content (Join-Path $OutputPath 'Report.txt') -Encoding UTF8