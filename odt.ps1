<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>

using namespace System.Management.Automation.Host


function New-Menu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Question,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChoiceA,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChoiceB
    )
    
    $a = [ChoiceDescription]::new("&$( $ChoiceA )", '')
    $b = [ChoiceDescription]::new("&$( $ChoiceB )", '')

    $options = [ChoiceDescription[]]($a, $b)

    $result = $host.ui.PromptForChoice($title, $question, $options, 0)

    return $result
}

function Start-OfficeSetup {
    [CmdletBinding()]
    param (
        
        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        [Parameter(Mandatory)]
        [ValidateSet("Download", "Configure")]
        [string]
        $Type
    )

    Start-Process $Path -ArgumentList "/$( $Type ) $( $NameConfig)" -Wait

    # $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    # $pinfo.FileName = $Path
    # $pinfo.RedirectStandardError = $true
    # $pinfo.RedirectStandardOutput = $true
    # $pinfo.UseShellExecute = $false
    # $Arguments = "/$( $Type ) $( $NameConfig)"
    # $pinfo.Arguments = $Arguments
    # $p = New-Object System.Diagnostics.Process
    # $p.StartInfo = $pinfo
    # $p.Start() | Out-Null
    # $p.WaitForExit()
    # $stdout = $p.StandardOutput.ReadToEnd()
    # $stderr = $p.StandardError.ReadToEnd()
    # Write-Host "stdout start ---------------------------------"
    # Write-Host $stdout
    # Write-Host "stdout end -----------------------------------"
    # Write-Host "stderr start ---------------------------------"
    # Write-Host $stderr
    # Write-Host "stderr end -----------------------------------"
    # Write-Host
    # Write-Host "exit code: $( $p.ExitCode )"
}

$NameConfig = "config.xml"
$DownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14527-20178.exe"
$PathConfig = "$( $PSScriptRoot )\$( $NameConfig )"
$PathExePacked = "$( $PSScriptRoot )\officedeploymenttool_packed.exe"
$PathExeSetup = "$( $PSScriptRoot )\setup.exe"

if (-not (Test-Path $PathExePacked)) {
    $AllProtocols = [System.Net.SecurityProtocolType]'Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    $Client = New-Object System.Net.WebClient
    $Client.DownloadFile($DownloadUrl, $PathExePacked)
}

if (-not (Test-Path $PathExeSetup)) {
    Start-Process $PathExePacked -ArgumentList "/extract:$( $PSScriptRoot ) /passive /quiet"
}

# User Interaction
$ResultVisio = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Visio?'
switch ($ResultVisio) {
    0 {$Visio = "<Product ID=`"VisioProRetail`">
    <Language ID=`"de-DE`" />
    <ExcludeApp ID=`"Groove`" />
    <ExcludeApp ID=`"Lync`" />
    </Product>"} 
    1 {$Visio = ""}
}

$ResultPublisher = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Publisher?'
switch ($ResultPublisher) {
    0 {$Publisher = "<Product ID=`"PublisherRetail`">
    <Language ID=`"de-DE`" />
    <ExcludeApp ID=`"Groove`" />
    <ExcludeApp ID=`"Lync`" />
  </Product>"}
    1 {$Publisher = ""}
}

$ResultBit = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Office as 32-Bit version? ("No" = 64-Bit)'
switch ($ResultBit) {
    0 {$Bit = "32"}
    1 {$Bit = "64"}
}

$ResultUseAdmin = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to provide the Office 365 Administrator Credentials and automatically check for available licenses in order to choose whether to install Apps for Business or Apps for Enterprise? ("no" = choose manually)'
switch ($ResultUseAdmin) {
    0 {
        Connect-AzureAD

        $user = Get-AzureADUser | Select DisplayName, Mail, ProxyAddresses, UserPrincipalName 
        $user = $user | Out-GridView -PassThru -Title "Select the user whose license you mean to use."
        $userUPN = $user.UserPrincipalName
        $licensePlanList = Get-AzureADSubscribedSku
        $userPlanList = Get-AzureADUser -ObjectID $userUPN | Select -ExpandProperty AssignedLicenses | Select SkuID 
        $userPlanListTranlated = $licensePlanList.Where({$userPlanList.SkuID -contains $_.ObjectId.substring($_.ObjectId.length - 36, 36)})
        foreach ($userLicense in $userPlanListTranlated) {
            if ($userLicense.ServicePlans.ServicePlanName -contains "OFFICE_BUSINESS") {
                # Business Plan
                $Apps = "O365BusinessRetail"
                Write-Host "Found O365BusinessRetail."
            } elseif ($userLicense.ServicePlans.ServicePlanName -contains "OFFICESUBSCRIPTION") {
                # Enterprise Plan
                $Apps = "O365ProPlusRetail"
                Write-Host "Found O365ProPlusRetail."
            }
        }
        if (-not ($Apps)) {
            throw "Neither O365BusinessRetail nor O365ProPlusRetail found!"
        }
    } 
    1 {
        $ResultApps = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Pro Plus" -ChoiceB "Business" -Question 'Which Office do you mean to install?'
        switch ($ResultApps) {
            0 {$Apps = "O365ProPlusRetail"}
            1 {$Apps = "O365BusinessRetail"}
        }
    }
}

$ResultDisplayLevel = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to show the installation progress? ("no" = silent install)'
switch ($ResultDisplayLevel) {
    0 {$DisplayLevel = "Full"}
    1 {$DisplayLevel = "None"}
}

$ConfigFinal = "<Configuration>
  <Add OfficeClientEdition=`"$( $Bit )`" Channel=`"Current`">
    <Product ID=`"$( $Apps )`">
      <Language ID=`"de-de`" />
      <ExcludeApp ID=`"Groove`" />
      <ExcludeApp ID=`"Lync`" />
    </Product>
    $( $Visio )
    $( $Publisher )
  </Add>
  <Property Name=`"SharedComputerLicensing`" Value=`"0`" />
  <Property Name=`"FORCEAPPSHUTDOWN`" Value=`"TRUE`" />
  <Property Name=`"DeviceBasedLicensing`" Value=`"0`" />
  <Property Name=`"SCLCacheOverride`" Value=`"0`" />
  <Updates Enabled=`"TRUE`" />
  <RemoveMSI />
  <AppSettings>
    <User Key=`"software\microsoft\office\16.0\excel\options`" Name=`"defaultformat`" Value=`"51`" Type=`"REG_DWORD`" App=`"excel16`" Id=`"L_SaveExcelfilesas`" />
    <User Key=`"software\microsoft\office\16.0\powerpoint\options`" Name=`"defaultformat`" Value=`"27`" Type=`"REG_DWORD`" App=`"ppt16`" Id=`"L_SavePowerPointfilesas`" />
    <User Key=`"software\microsoft\office\16.0\word\options`" Name=`"defaultformat`" Value=`"`" Type=`"REG_SZ`" App=`"word16`" Id=`"L_SaveWordfilesas`" />
  </AppSettings>
  <Display Level=`"$( $DisplayLevel )`" AcceptEULA=`"TRUE`" />
</Configuration>"

Write-Host "Writing config file..."
Set-Content -Path $PathConfig -Value $ConfigFinal

Write-Host "Starting download..."
Start-OfficeSetup -Path $PathExeSetup -Type Download

Write-Host "Starting installation..."
Start-OfficeSetup -Path $PathExeSetup -Type Configure
