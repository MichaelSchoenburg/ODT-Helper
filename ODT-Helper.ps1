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

function Get-ODTUri {
    <#
        .SYNOPSIS
            Get Download URL of latest Office 365 Deployment Tool (ODT).
        .NOTES
            Author: Bronson Magnan
            Twitter: @cit_bronson
            Modified by: Marco Hofmann
            Twitter: @xenadmin
            Source: https://www.meinekleinefarm.net/download-and-install-latest-office-365-deployment-tool-odt/
        .LINK
            https://www.meinekleinefarm.net/
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param ()

    $url = "https://www.microsoft.com/en-us/download/details.aspx?id=49117" # Muss jedes Mal manuell aktualisiert werden
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
    }
    catch {
        Throw "Failed to connect to ODT: $url with error $_."
        Break
    }
    finally {
        $ODTUri = $response.links | Where-Object {$_.outerHTML -like '*Download*Office Deployment Tool*'} # I modified this one to work with the current website
        Write-Output $ODTUri.href
    }
}

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

    $a = New-Object System.Management.Automation.Host.ChoiceDescription "&$( $ChoiceA )", ''
    $b = New-Object System.Management.Automation.Host.ChoiceDescription "&$( $ChoiceB )", ''

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($a, $b)

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

    Write-Host "PathConfig = $( $PathConfig )"
    Start-Process $Path -ArgumentList "/$( $Type ) $( $PathConfig)" -Wait

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

$Path = "C:\TSD.CenterVision\Software\ODT"
$DownloadUrl = Get-ODTUri
$NameConfig = "config.xml"
$PathConfig = "$( $Path )\$( $NameConfig )"
$PathExePacked = "$( $Path)\officedeploymenttool_packed.exe"
$PathExeSetup = "$( $Path )\setup.exe"

if (Test-Path -Path $Path) {
    Write-Host -ForegroundColor Green "Directory for ODT exists already."
} else {
    Write-Host -ForegroundColor Gray "Creating Directory for ODT..."
    New-Item -Path $Path -ItemType Directory
}

Write-Host -ForegroundColor Gray 'Testing if ODT has already been downloaded...'
if (-not (Test-Path $PathExePacked)) {
    Write-Host 'Downloading ODT...'
    $AllProtocols = [System.Net.SecurityProtocolType]'Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    
    try {
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $PathExePacked -PassThru -UseBasicParsing
    }
    catch {
        if( $_.Exception.Response.StatusCode.Value__ -eq 404 )
        {
            throw "Can't download ODT. 404 Not Found. Maybe URL isn't up-to-date anymore?"
        }
        else {
            throw "Unknown error while downloading ODT."
        }
    }
}
Write-Host -ForegroundColor Green 'ODT downloaded.'

if (-not (Test-Path $PathExeSetup)) {
    $args = "/extract:`"$( $Path )`" /passive /quiet"
    Write-Host "Args = $( $args )"
    Start-Process $PathExePacked -ArgumentList $args
}

# $ResultBit = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Office as 32-Bit version? ("No" = 64-Bit)'
# switch ($ResultBit) {
#     0 {$Bit = "32"}
#     1 {$Bit = "64"}
# }
$Bit = "64"

$ResultUseAdmin = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to provide the Office 365 Administrator Credentials and automatically check for available licenses in order to choose whether to install Apps for Business or Apps for Enterprise? ("no" = choose manually)'
switch ($ResultUseAdmin) {
    0 {
        Write-Host -ForegroundColor Yellow "Please log in with a global admin of the tenant in which the user is located."
        try {
            if (Get-Module -Name AzureAD) {
                Write-Host -ForegroundColor Gray "Module already imported".
            } elseif (Get-Module -Name AzureAD -ListAvailable) {
                Write-Host -ForegroundColor Gray "Module already installed".
                Write-Host -ForegroundColor Gray "Importing Module..."
                Import-Module -Name AzureAD
            } else {
                Write-Host -ForegroundColor Gray "Installing Module..."
                Install-Module -Name AzureAD -Force -Scope CurrentUser
                Write-Host -ForegroundColor Gray "Importing Module..."
                Import-Module -Name AzureAD
            }
            
            Write-Host -ForegroundColor Gray "Connecting to Azure..."
            Connect-AzureAD
            $user = Get-AzureADUser | Select-Object DisplayName, Mail, ProxyAddresses, UserPrincipalName
            $user = $user | Out-GridView -Title "Select the user whose license you mean to use." -PassThru
            $userUPN = $user.UserPrincipalName
            $licensePlanList = Get-AzureADSubscribedSku
            $userPlanList = Get-AzureADUser -ObjectID $userUPN | Select-Object -ExpandProperty AssignedLicenses | Select-Object SkuID 
            $userPlanListTranslated = $licensePlanList.Where({$userPlanList.SkuID -contains $_.ObjectId.substring($_.ObjectId.length - 36, 36)})
            foreach ($userLicense in $userPlanListTranslated) {
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
                throw "This user doesn't have a license for Microsoft Apps (neither business apps nor enterprise apps)! Please assign a license containing Microsoft Apps. Aborting script."
            }
        } catch {
            $_.Exception.Message
            throw "Etwas ist schief gegangen."
        }
    }
    1 {
        $a = New-Object System.Management.Automation.Host.ChoiceDescription 'Microsoft Apps for &Enterprise (aka. "Pro Plus")', ''
        $b = New-Object System.Management.Automation.Host.ChoiceDescription 'Microsoft Apps for &Business', ''
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($a, $b)
        $ResultApps = $host.ui.PromptForChoice('Office Deployment Tool - Configuration', 'Which Office do you mean to install?', $options, 0)
        switch ($ResultApps) {
            0 {$Apps = "O365ProPlusRetail"}
            1 {$Apps = "O365BusinessRetail"}
        }
    }
}

$ResultVisio = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Visio?'
switch ($ResultVisio) {
    0 {$Visio = "<Product ID=`"VisioProRetail`">
    <Language ID=`"de-DE`" />
    <ExcludeApp ID=`"Groove`" />
    <ExcludeApp ID=`"Lync`" />
    </Product>"} 
    1 {$Visio = ""}
}

# $ResultPublisher = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to install Publisher?'
# switch ($ResultPublisher) {
#     0 {$Publisher = "<Product ID=`"PublisherRetail`">
#     <Language ID=`"de-DE`" />
#     <ExcludeApp ID=`"Groove`" />
#     <ExcludeApp ID=`"Lync`" />
#   </Product>"}
#     1 {$Publisher = ""}
# }
$Publisher = ""

# $ResultDisplayLevel = New-Menu -Title 'Office Deployment Tool - Configuration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Do you want to show the installation progress? ("no" = silent install)'
# switch ($ResultDisplayLevel) {
#     0 {$DisplayLevel = "Full"}
#     1 {$DisplayLevel = "None"}
# }
$DisplayLevel = "Full"

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
