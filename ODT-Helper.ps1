<#
.SYNOPSIS
    ODT-Helper
.DESCRIPTION
    Ein Hilfsskript zum Herunterladen und Installieren des Microsoft Office 365 Deployment Tools (ODT) und zum Installieren von Office 365 Apps.
.LINK
    GitHub: https://github.com/MichaelSchoenburg/ODT-Helper
.OUTPUTS
    ExitCode 0 = Erfolgreich
    ExitCode 1 = Fehler
    ExitCode 2 = Warnung
.NOTES
    Author: Michael Schönburg
    Version: v1.0
    Creation: 28.02.2025
    Last Edit: 28.02.2025
    
    This projects code loosely follows the PowerShell Practice and Style guide, as well as Microsofts PowerShell scripting performance considerations.
    Style guide: https://poshcode.gitbook.io/powershell-practice-and-style/
    Performance Considerations: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1
#>

#region INITIALIZATION
<# 
    Libraries, Modules, ...
#>



#endregion INITIALIZATION
#region FUNCTIONS
<# 
    Declare Functions
#>

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Protokolliert ein Ereignis in der Konsole.
    
    .DESCRIPTION
    Schreibt Text in die Konsole mit dem aktuellen Datum (US-Format) davor.
    
    .PARAMETER Text
    Ereignis/Text, der in die Konsole ausgegeben werden soll.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ aufgerufen.'
    
    Lange Form
    .EXAMPLE
    Log 'Subscript XYZ aufgerufen.'
    
    Kurze Form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    if ($Interactive) {
        # Aktuelle VerbosePreference speichern
        $VerbosePreferenceBefore = $VerbosePreference

        # Verbose-Ausgabe aktivieren
        $VerbosePreference = 'Continue'

        # Verbose-Ausgabe schreiben
        Write-Verbose "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"

        # Aktuelle VerbosePreference wiederherstellen
        $VerbosePreference = $VerbosePreferenceBefore
    } else {
        Write-Output "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"
    }
}

function Get-OdtUri {
    <#
        .SYNOPSIS
            Ruft die Download-URL des neuesten Office 365 Deployment Tools (ODT) von der Microsoft-Webseite ab.
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
    param (
        [Parameter()]
        [String]
        [ValidatePattern("https://.*")]
        $Url = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
    )

    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
    } catch {
        throw "Fehler beim Abrufen der URL $($url) für den Download des ODT mit folgendem Fehler: $( $_.Exception.Message )"
    }

    try {
        $ODTUri = $response.links | Where-Object {$_.outerHTML -like '*Download*Office Deployment Tool*'} # Ich habe dies geändert, damit es mit der aktuellen Website funktioniert
        return $ODTUri.href
    } catch {
        throw "Fehler beim Extrahieren der Download-URL von der Microsoft-Webseite für das ODT. Fehler: $( $_.Exception.Message )"
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

function Set-DenyShutdown {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.Boolean]
        $Active = $true
    )
    
    switch ($Active) {
        $true { $int = 1 }
        $false { $int = 0 }
    }

    try {
        # Setze den Registrierungsschlüssel, um das Herunterfahren zu verhindern
        $null = New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoClose" -PropertyType DWORD -Value $int -Force
        
        # Explorer neu starten, da die Änderungen sonst möglicherweise nicht angewendet werden
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        # Start-Process -FilePath explorer.exe # Wurde bei mir nicht benötigt. Explorer kam von selbst wieder hoch. Dies hat nur ein unnötiges Explorer-Fenster geöffnet.
    } catch {
        Log "Der Registrierungsschluessel zum Verhindern des Herunterfahrens konnte nicht eingestellt werden. Fehler: $( $_.Exception.Message )"
        Exit 1
    }
}

function Get-UserLicense {
    <#
    .SYNOPSIS
        Überprüfen, welche Lizenz für Microsoft Apps der Benutzer hat.
    .DESCRIPTION
        Dient dazu, die Lizenzen eines Benutzers in einem Microsoft 365-Mandanten abzurufen. 
        Diese Funktion ist besonders nützlich in einem interaktiven Modus, 
        in dem der Benutzer eine Auswahl treffen kann. 
        Sie verwendet das Microsoft Graph PowerShell-Modul, 
        um eine Verbindung zu Azure herzustellen und die Benutzerinformationen abzurufen.
    .INPUTS
        Interaktiv: Der Benutzer wird aufgefordert, einen Benutzer auszuwählen, dessen Lizenz abgefragt werden soll.
    .OUTPUTS
        O365BusinessRetail
        O365ProPlusRetail
        1 = Dieser Benutzer besitzt gar keine Lizenz. Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen.
        2 = Dieser Benutzer besitzt eine Lizenz, jedoch ohne Microsoft Apps (weder Business-Apps noch Enterprise-Apps) darin enthalten! Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen.
    #>
    
    # Die Funktion wird nur im interaktiven Modus verwendet. 
    # Im interaktiven Modus ist das Log auf verbose geschaltet, 
    # so wird die Ausgabe (return) der Funktion nicht beeinträchtigt.

    try {
        # Lade PowerShell-Modul für Microsoft Graph
        if (Get-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users) {
            Log "Modul bereits importiert."
        } elseif (Get-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users -ListAvailable) {
            Log "Modul bereits installiert."
            Log "Modul importieren..."
            Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users
        } else {
            Log "Modul installieren..."
            Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser
            Log "Modul importieren..."
            Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users
        }

        # Herstellen einer Verbindung mit Azure
        Log "Herstellen einer Verbindung mit Azure..."
        Log "Bitte melden Sie sich im folgenden mit einem globalen Administrator des Mandanten an, in dem sich der Benutzer befindet."
        $null = Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

        # Abfrage der Benutzerlizenzen
        $Properties = 'DisplayName', 'AssignedLicenses', 'DisplayName', 'Givenname', 'Surname', 'UserPrincipalName', 'OnPremisesSamAccountName'
        $users = Get-MgUser -All -Filter 'accountEnabled eq true' -Property $Properties | 
            Sort-Object -Property DisplayName | 
            Select-Object @{Name = 'Lizenzierung'; Expression = {foreach ($l in $_.AssignedLicenses) {
                    $Licenses.Where({$_.GUID -eq $l.SkuId})[0].Product_Display_Name
                }}}, 
                DisplayName, Givenname, Surname, UserPrincipalName, OnPremisesSamAccountName, AssignedLicenses
        
        # Auswahl des Benutzers durch den Anwender (interaktiv)
        $user = $users | Out-GridView -Title "Waehlen Sie den Benutzer aus, dessen Lizenz Sie verwenden moechten." -PassThru

        # Lizenzen "nachschlagen", um zu bestimmen, welche Produkte enthalten sind
        $licensePlanList = Get-MgSubscribedSku

        # Wenn der Benutzer keine Lizenz hat, wird das Skript abgebrochen
        if (-not ($user.AssignedLicenses)) {
            $return = 1
            throw "Dieser Benutzer besitzt gar keine Lizenz. Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen."
        }

        # Bestimmen, ob der Benutzer Business- oder Enterprise-Apps hat
        foreach ($license in $user.AssignedLicenses) {
            $ServicePlanNames = $licensePlanList.Where({$_.SkuId -eq $license.SkuId}).ServicePlans.ServicePlanName
            if ($ServicePlanNames -contains "OFFICE_BUSINESS") {
                # Business Plan
                $return = "O365BusinessRetail"
                Log "Found O365BusinessRetail."
            } elseif ($ServicePlanNames -contains "OFFICESUBSCRIPTION") {
                # Enterprise Plan
                $return = "O365ProPlusRetail"
                Log "Found O365ProPlusRetail."
            }
        }
        
        # Wenn der Benutzer eine Lizenz hat, jedoch keine Microsoft Apps (weder Business-Apps noch Enterprise-Apps) darin enthalten sind, wird das Skript abgebrochen
        if (-not ($return)) {
            $return = 2
            throw "Dieser Benutzer besitzt eine Lizenz, jedoch ohne Microsoft Apps (weder Business-Apps noch Enterprise-Apps) darin enthalten! Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen."
        }
    } catch {
        Log "Fehler: $( $_.Exception.Message )"
    } finally {
        $null = Disconnect-MgGraph
    }

    return $return
}

function Get-OfficeInstalled {
    $officeInstalled = $false
    
    <# 
        Überprüfen der Registrierung auf Office-Installation
    #>

    # Definieren der Registrierungspfade, um nach Office-Installationen zu suchen
    $officePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($path in $officePaths) {
        Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.GetValue("DisplayName") -match "Microsoft Office") {
                $officeInstalled = $true
            }
        }
    }

    <# 
        Überprüfen von WMI/CIM auf Office-Installation
    #>

    # Verwendung von Get-WmiObject
    $wmi = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%Office%'" 2>$null

    if ($wmi) {
        $officeInstalled = $true
    }

    # Alternativ, Verwendung von Get-CimInstance für modernere Systeme
    $cim = Get-CimInstance -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%Office%'"

    if ($cim) {
        $officeInstalled = $true
    }

    <# 
        Fazit
    #>

    if ($officeInstalled) {
        return $true
    } else {
        return $false
    }
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

    Log "PathConfig = $( $PathConfig )"
    Start-Process $Path -ArgumentList "/$( $Type ) $( $PathConfig)" -Wait
}

#endregion FUNCTIONS
#region DECLARATIONS
<#
    Declare local variables and global variables
#>

$Interactive = $true # Setze auf $false für nicht-interaktiven Modus (z.B. RMM)
$MaximumVariableCount = 8192 # Graph-Modul hat mehr als 4096 Variablen
$Path = "C:\TSD.CenterVision\Software\ODT"
$NameConfig = "config.xml"
$PathConfig = Join-Path -Path $Path -ChildPath $NameConfig
$PathExePacked = Join-Path -Path $Path -ChildPath "officedeploymenttool_packed.exe"
$PathExeSetup = Join-Path -Path $Path -ChildPath "setup.exe"
$Licenses = @(
    [PSCustomObject]@{GUID='6fd2c87f-b296-42f0-b197-1e91e994b900'; Product_Display_Name='Microsoft 365 E3'; ServicePlanName='OFFICESUBSCRIPTION'},
    [PSCustomObject]@{GUID='c42b9cae-ea4f-4ab7-9717-81576235ccac'; Product_Display_Name='Microsoft 365 Business Standard'; ServicePlanName='OFFICE_BUSINESS'},
    [PSCustomObject]@{GUID='e212cbc7-0961-4c40-9825-01117710dcb1'; Product_Display_Name='Office 365 E1'; ServicePlanName='OFFICESUBSCRIPTION'},
    [PSCustomObject]@{GUID='f245ecc8-75af-4f8e-b61f-27d8114de5f3'; Product_Display_Name='Office 365 E3'; ServicePlanName='OFFICESUBSCRIPTION'},
    [PSCustomObject]@{GUID='c1ec4a95-1f05-45b3-a911-aa3fa01094f5'; Product_Display_Name='Office 365 E5'; ServicePlanName='OFFICESUBSCRIPTION'},
    [PSCustomObject]@{GUID='9aaf7827-d63c-4b61-89c3-182f06f82e5c'; Product_Display_Name='Office 365 Business'; ServicePlanName='OFFICE_BUSINESS'},
    [PSCustomObject]@{GUID='b1188c4c-1b36-4018-b48b-ee07604f6feb'; Product_Display_Name='Office 365 Business Premium'; ServicePlanName='OFFICE_BUSINESS'}
) # Source: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

# Beispiel für die Verwendung in einem RMM (z. B. Riverbird): folgende Variablen müssen während der Laufzeit gesetzt werden:
# $Apps = "O365ProPlusRetail"
# $ResultBit = "64"
# $SharedComputerLicensing = 0
# $ResultVisio = 0
# $ResultPublisher = 0
# $ResultDisplayLevel = 0

#endregion DECLARATIONS
#region EXECUTION
<# 
    Script entry point
#>

# Überprüfen, ob Microsoft Office bereits installiert ist
if (Get-OfficeInstalled) {
    Log "Microsoft Office ist bereits installiert. Das Skript wird erfolgreich abgebrochen."
    Exit 0
} else {
    Log "Microsoft Office ist noch nicht installiert. Das Skript wird fortgesetzt."

    # Extrahiere Download-URL für das ODT von der Microsoft-Webseite
    try {
        $DownloadUrl = Get-OdtUri # Terminierender Error würde in der Funktion geworfen werden.
        Log "Download-Link für ODT gefunden = $( $DownloadUrl )"

        try {
            Log "Verhindere Herunterfahren des Computers..."
            Set-DenyShutdown -Active $true -ErrorAction Continue
        
            # Überprüfen, ob der ODT-Ordner existiert
            Log "Überprüfen, ob der ODT-Ordner existiert..."
            if (Test-Path -Path $Path) {
                Log "Ordner fuer ODT existiert bereits."
            } else {
                Log "Lege Ordner fuer ODT an..."
                New-Item -Path $Path -ItemType Directory -ErrorAction Stop # Terminierender Error, falls der Ordner nicht erstellt werden kann.
            }
        
            # Überprüfen, ob ODT bereits heruntergeladen wurde, falls nicht, herunterladen
            Log 'Teste, ob ODT bereits heruntergeladen wurde...'
            if (-not (Test-Path $PathExePacked)) {
                Log 'Lade ODT herunter...'
    
                # Setze TLS-Version auf 1.1 und 1.2 zwecks Download
                $AllProtocols = [System.Net.SecurityProtocolType]'Tls11,Tls12'
                [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
                
                try {
                    # Download des ODT
                    $response = Invoke-WebRequest -Uri $DownloadUrl -OutFile $PathExePacked -PassThru -UseBasicParsing
                    Log 'ODT erfolgreich heruntergeladen...'
                } catch {
                    if( $_.Exception.Response.StatusCode.Value__ -eq 404 ) {
                        throw "ODT kann nicht heruntergeladen werden. 404 Nicht gefunden. Vielleicht ist die URL nicht mehr aktuell oder es liegt eine Störung bei Microsoft vor (gab es bereits einmal)?"
                    } else {
                        throw "Unbekannter Fehler beim Herunterladen des ODT. Response Code: $( $_.Exception.Response.StatusCode.Value__ ) Error: $( $_.Exception.Message )"
                    }
                }
            } else {
                Log "ODT bereits heruntergeladen."
            }
        
            # Überprüfen, ob ODT bereits entpackt wurde, falls nicht, entpacken
            if (-not (Test-Path $PathExeSetup)) {
                Log "Entpacke ODT..."
                $arguments = "/extract:`"$( $Path )`" /passive /quiet"
                Log "DEBUG: arguments = $( $arguments )"
                Start-Process $PathExePacked -ArgumentList $arguments -Wait
            } else {
                Log "ODT bereits entpackt."
            }
        
            # Falls $Apps nicht durch das RMM gesetzt wurde (nicht-interaktiver Modus), frage den Benutzer (interaktiver Modus), ob er die Lizenzierung automatisch ermitteln möchte oder nicht
            if ($null -eq $Apps) {
                $ResultUseAdmin = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Moechten Sie die Office 365-Administratoranmeldeinformationen angeben und automatisch nach verfuegbaren Lizenzen suchen, um auszuwaehlen, ob Apps for Business oder Apps for Enterprise installiert werden sollen? („nein“ = manuell auswaehlen)'
                switch ($ResultUseAdmin) {
                    0 {
                        $License = Get-UserLicense
                        if (1,2 -contains $License) {
                            switch ($License) {
                                1 { $Grund =  "Dieser Benutzer besitzt gar keine Lizenz. Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen." }
                                2 { $Grund =  "Dieser Benutzer besitzt eine Lizenz, jedoch ohne Microsoft Apps (weder Business-Apps noch Enterprise-Apps) darin enthalten! Bitte weisen Sie eine Lizenz zu, die Microsoft Apps enthaelt. Skript wird abgebrochen." }
                            }
                            throw "Fehler beim Abrufen der Benutzerlizenz. Skript wird abgebrochen. $( $Grund )"
                        }
                    }

                    1 {
                        $a = New-Object System.Management.Automation.Host.ChoiceDescription 'Microsoft Apps for &Enterprise (aka. "Pro Plus")', ''
                        $b = New-Object System.Management.Automation.Host.ChoiceDescription 'Microsoft Apps for &Business', ''
                        $options = [System.Management.Automation.Host.ChoiceDescription[]]($a, $b)
                        $ResultApps = $host.ui.PromptForChoice('Office Deployment Tool - Konfiguration', 'Welches Office moechten Sie installieren?', $options, 0)
                        switch ($ResultApps) {
                            0 {$Apps = "O365ProPlusRetail"}
                            1 {$Apps = "O365BusinessRetail"}
                        }
                    }
                }
            } else {
                Log 'Variable "Apps" durch RMM bereits gesetzt.'
            }
        
            # 32-Bit oder 64-Bit?
            if ($null -eq $ResultBit) {
                $ResultBit = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Moechten Sie Office als 64-Bit-Version installieren? ("Nein" = 32-Bit)'
                switch ($ResultBit) {
                    0 {$Bit = "64"}
                    1 {$Bit = "32"}
                }
            } else {
                Log 'Variable "ResultBit" durch RMM bereits gesetzt.'
            }

            # SharedComputerLicensing (Terminal Server/RDSH) aktivieren?
            if ($null -eq $SharedComputerLicensing) {
                $ResultSharedComputerLicensing = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question "Wird der Computer von mehreren Benutzern verwendet (Terminal Server aka. Remotedesktop Session Host?"
            }

            switch ($ResultSharedComputerLicensing) {
                0 {$SharedComputerLicensing = "1"}
                1 {$SharedComputerLicensing = "0"}
            }
        
            # Visio installieren?
            if ($null -eq $ResultVisio) {
                $ResultVisio = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Moechten Sie Microsoft Visio installieren?'
            } else {
                Log 'Variable "ResultVisio" durch RMM bereits gesetzt.'
            }
            
            switch ($ResultVisio) {
                0 {$Visio = "<Product ID=`"VisioProRetail`">
                <Language ID=`"de-DE`" />
                <ExcludeApp ID=`"Groove`" />
                <ExcludeApp ID=`"Lync`" />
                </Product>"} 
                1 {$Visio = ""}
            }
        
            # Publisher installieren?
            if ($null -eq $ResultPublisher) {
                $ResultPublisher = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Moechten Sie Microsoft Publisher installieren?'
                switch ($ResultPublisher) {
                    0 {$Publisher = "<Product ID=`"PublisherRetail`">
                    <Language ID=`"de-DE`" />
                    <ExcludeApp ID=`"Groove`" />
                    <ExcludeApp ID=`"Lync`" />
                </Product>"}
                    1 {$Publisher = ""}
                }
            } else {
                Log 'Variable "ResultPublisher" durch RMM bereits gesetzt.'
            }
        
            # Installationsfortschritt anzeigen?
            if ($null -eq $ResultDisplayLevel) {
                $ResultDisplayLevel = New-Menu -Title 'Office Deployment Tool - Konfiguration' -ChoiceA "Yes" -ChoiceB "No" -Question 'Moechten Sie den Installationsfortschritt anzeigen? ("no" = silent install)'
                switch ($ResultDisplayLevel) {
                    0 {$DisplayLevel = "Full"}
                    1 {$DisplayLevel = "None"}
                }
            } else {
                Log 'Variable "ResultDisplayLevel" durch RMM bereits gesetzt.'
            }
        
            # Konfigurationsdatei zusammen bauen
            $ConfigFinal = "<Configuration>
            <Add OfficeClientEdition=`"$( $Bit )`" Channel=`"Current`">
                <Product ID=`"$( $Apps )`">
                <Language ID=`"de-de`" />
                <ExcludeApp ID=`"Lync`" />
                </Product>
                $( $Visio )
                $( $Publisher )
            </Add>
            <Property Name=`"SharedComputerLicensing`" Value=`"$( $ResultSharedComputerLicensing )`" />
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
        
            Log "Schreibe Konfigurationsdatei..."
            Set-Content -Path $PathConfig -Value $ConfigFinal
        
            Log "Starte Download..."
            Start-OfficeSetup -Path $PathExeSetup -Type Download
        
            Log "Starte Installation..."
            Start-OfficeSetup -Path $PathExeSetup -Type Configure

            # Nur wenn alles erfolgreich durchgelaufen ist, wird der ExitCode auf 0 gesetzt
            $ExitCode = 0
        } catch {
            throw # rethrow error
        } finally {
            Log "Herunterfahren wieder erlauben..."
            Set-DenyShutdown -Active $false -ErrorAction Continue
        }
    }
    catch {
        Log "Ein Fehler ist aufgetreten. Das Skript wird abgebrochen. Fehler: $( $_.Exception.Message )"
        $ExitCode = 1
    } finally {
        Exit $ExitCode
    }
}

#endregion EXECUTION