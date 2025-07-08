# Copyright (C) 2025 Gordon McLellan

if($host.version.major -gt 5) {
    Write-Host "This script needs Powershell 5"
    Exit
}

# $cred = get-credential   
$adminUPN = $searchname = Read-Host -Prompt "Tenant Admin Email"
$searchemail = Read-Host -Prompt "Email address to archive"
$searchname = Read-Host -Prompt "Name for this search"

Write-host Enabling Basic Auth
$WinRMClient = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client"
$Name = "AllowBasic"
$value = "1"
IF (!(Test-Path $WinRMClient)) {
   New-Item -Path $WinRMClient -Force | Out-Null
   New-ItemProperty -Path $WinRMClient -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
} ELSE {
   New-ItemProperty -Path $WinRMClient -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}

#Check for EXO v2 module installation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if($Module.count -eq 0) {
 Write-Host Exchange Online PowerShell V2 module is not available -ForegroundColor yellow
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
 if($Confirm -match "[yY]") {
    Write-host "Installing Exchange Online PowerShell module"
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
    Import-Module ExchangeOnlineManagement
  } else {
    Write-Host EXO V2 module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
    Exit
  }
} else {
    Import-Module ExchangeOnlineManagement 
}

Write-Host "Starting EXO session"

Connect-ExchangeOnline -UserPrincipalName $adminUPN 

Write-Host "Connectiong to IPPS Session"

Connect-IPPSSession -UserPrincipalName $adminUPN 

#Connect-IPPSSession -UserPrincipalName $adminUPN -ConnectionUri https://ps.compliance.protection.office365.us/powershell-liveid/ -AzureADAuthorizationEndpointUri https://login.microsoftonline.us/common

#Connect-IPPSSession -Credential $cred

Write-Host "Starting Compliance Search"
New-ComplianceSearch -Name $searchname -ExchangeLocation $searchemail -AllowNotFoundExchangeLocationsEnabled $true 

Start-ComplianceSearch $searchname

Get-ComplianceSearch $searchname | Format-List

Write-Host -Prompt "Wait for email from Microsoft and then visit https://purview.microsoft.com to export the compliance search results."

# Write-Host "Requesting Search Result Export"

# New-ComplianceSearchAction $searchname -Export -Format Fxstream 

Write-Host "Ending EXO session"

Disconnect-ExchangeOnline -Confirm:$false
