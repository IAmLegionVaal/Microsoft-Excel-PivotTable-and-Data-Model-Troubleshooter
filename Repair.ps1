#requires -Version 5.1
<# Created by Dewald Pretorius. #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet('Diagnose','ResetModelCaches','RepairOffice')][string]$Action='Diagnose',
    [string]$OutputPath=(Join-Path ([Environment]::GetFolderPath('Desktop')) 'Excel_Data_Model_Repair')
)
$ErrorActionPreference='Stop'
$CachePaths=@("$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache","$env:LOCALAPPDATA\Microsoft\Office\16.0\PowerQuery")
New-Item -ItemType Directory -Path $OutputPath -Force|Out-Null
$Stamp=Get-Date -Format 'yyyyMMdd_HHmmss';$LogPath=Join-Path $OutputPath "Repair_$Stamp.log"
function Log([string]$Message){$Line='{0:u} {1}' -f (Get-Date),$Message;Write-Host $Line;Add-Content -LiteralPath $LogPath -Value $Line}
[ordered]@{Action=$Action;ExcelRunning=[bool](Get-Process EXCEL -ErrorAction SilentlyContinue);MashupProcesses=@(Get-Process 'Microsoft.Mashup.Container.NetFX40' -ErrorAction SilentlyContinue|Select-Object Name,Id);Caches=@($CachePaths|ForEach-Object{[pscustomobject]@{Path=$_;Exists=(Test-Path -LiteralPath $_)}})}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputPath "PreRepair_$Stamp.json") -Encoding UTF8
if($Action -eq 'Diagnose'){Log '[COMPLETE] Read-only snapshot saved.';exit 0}
try{
    if($Action -eq 'ResetModelCaches' -and $PSCmdlet.ShouldProcess('Excel model and query caches','Back up and reset')){
        if(Get-Process EXCEL,'Microsoft.Mashup.Container.NetFX40' -ErrorAction SilentlyContinue){throw 'Close Excel and wait for Mashup processes to exit.'}
        foreach($Path in $CachePaths){if(Test-Path -LiteralPath $Path){$Backup="$Path.backup-$Stamp";Move-Item -LiteralPath $Path -Destination $Backup -Force;New-Item -ItemType Directory -Path $Path -Force|Out-Null;Log "[BACKUP] $Backup"}}
    }
    elseif($Action -eq 'RepairOffice'){
        $Client=@("$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe","${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe")|Where-Object{Test-Path -LiteralPath $_}|Select-Object -First 1
        if(-not $Client){throw 'Microsoft 365 Apps repair client was not found.'}
        if($PSCmdlet.ShouldProcess('Microsoft 365 Apps','Run Quick Repair')){$Process=Start-Process -FilePath $Client -ArgumentList '/repair user displaylevel=true forceappshutdown=true' -Wait -PassThru;if($Process.ExitCode -ne 0){throw "Office repair exited with code $($Process.ExitCode)."}}
    }
}catch{Log "[FAILED] $($_.Exception.Message)";exit 5}
Log '[COMPLETE] Repair completed.'
exit 0
