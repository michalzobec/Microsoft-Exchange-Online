#Requires -Version 5
Set-StrictMode -Version 5

<#
.SYNOPSIS
Microsoft Exchange Online Distribution Group Members Update
Update Members in the Distribution Group on Microsoft Exchange Online service

.DESCRIPTION
Microsoft Exchange Online Distribution Group Members Update
(c) 2020-2024 Michal Zobec, ZOBEC Consulting. All Rights Reserved.  
web: www.michalzobec.cz, mail: michal@zobec.net  
GitHub repository http://zob.ec/

.OUTPUTS
Only text output in console (yet).

.EXAMPLE
C:\> MS-EXO-Distribution-Group-Members.ps1

.NOTES
Twitter/X: @michalzobec
Blog   : http://www.michalzobec.cz

.LINK
About this script on my Blog in Czech
http://zob.ec/

.LINK
Documentation (ReadMe)
https://github.com/michalzobec/

.LINK
Release Notes (ChangeLog)
https://github.com/michalzobec/

#>


######
$ScriptName = "Microsoft Exchange Online Distribution Group Members Update"
$ScriptVersion = "24.09.22.085359"
######


######
# External configuration file
# $ConfigurationFileName = "Get-SystemReport-Config-Example.ps1"
######

$ScriptDir = (Split-Path $myinvocation.MyCommand.Path)

# internal logging function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False,
            HelpMessage = "Select log level.")]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True,
            HelpMessage = "Information text to logfile.")]
        [string]
        $Message,

        [Parameter(Mandatory = $False,
            HelpMessage = "Logfilename and path.")]
        [string]
        $LogFile
    )

    $Stamp = Get-Date -Format "yyyy\/MM\/dd HH:mm:ss.fff"
    $Line = "[$Stamp] [$Level] $Message"
    If ($LogFile) {
        Add-Content $LogFile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}


# Definition of the log file - save method without subdirectory
$LogDate = Get-Date -Format "yyyyMMdd"
$LogFileName = "Get-SystemReport-log"
$LogFile = $ScriptDir + "\$LogFileName-$LogDate.txt"
$CopyRightYearFrom = "2016"
$CopyRightYearTo = "2019"

# Header
Write-Host ""
Write-Host "$ScriptName, version $ScriptVersion"
Write-Host "(c) $CopyRightYearFrom-$CopyRightYearTo Michal Zobec, ZOBEC Consulting. All Rights Reserved."
Write-Host ""
Write-Host "Initializing script"

$LogFileDir = $ScriptDir + "\logs"
if (!(Test-Path $LogFileDir -pathType container)) {
    Write-Verbose "Directory $LogFileDir was not found, creating."
    Write-Log -LogFile $LogFile -Message "  Directory $LogFileDir was not found, creating."
    New-Item $LogFileDir -type directory | Out-Null
    if (!(Test-Path $LogFileDir -pathType container)) {
        Write-Verbose "Directory $LogFileDir still not exist! Exiting."
        Write-Log -LogFile $LogFile -Message "  Directory $LogFileDir still not exist! Exiting."
        exit
    }
}

# Redefinition of the log file with LogFileDir
$LogFile = $LogFileDir + "\$LogFileName-$LogDate.txt"

# $CfgFilePath = $ScriptDir + "\config\$ConfigurationFileName"
# if (!(Test-Path $CfgFilePath)) {
#     Write-Warning "File $ConfigurationFileName is required for run of this script! Exiting."
#     Write-Log -LogFile $LogFile -Message "  File $ConfigurationFileName is required for run of this script! Exiting." -Level ERROR
#     exit
# }
# Write-Host "Configuration file $ConfigurationFileName"
# . $CfgFilePath

# Custom Variables
$excelPath = $ScriptDir + "\Allowed-Senders-List\services-international.xlsx" # Cesta k Excel sešitu
$distributionGroup = "exchange-senders-infrastructure-test@zobecint.cz" # Email distribučního seznamu
$UserPrincipalName = "u7256yr@zobeccid.cz" # Tvé emailové uživatelské jméno pro přihlášení

# Check if the required modules are installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Error: Module 'ExchangeOnlineManagement' is not installed. Please install the module and try again." -ForegroundColor Red
    exit
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Error: Module 'ImportExcel' is not installed. Please install the module and try again." -ForegroundColor Red
    exit
}

# Check if the Excel file exists
if (-not (Test-Path $excelPath)) {
    Write-Host "Error: File '$excelPath' does not exist. Please check the path and try again." -ForegroundColor Red
    exit
}

# Import the Exchange Online module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online using UserPrincipalName for modern authentication
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName

# Check if the distribution group exists
if (-not (Get-DistributionGroup -Identity $distributionGroup -ErrorAction SilentlyContinue)) {
    Write-Host "Error: Distribution group '$distributionGroup' does not exist. Please check the name and try again." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Import the email addresses and display names from the Excel file
$emailList = Import-Excel -Path $excelPath

# Loop through each entry in the Excel and add missing ones to the distribution group
foreach ($entry in $emailList) {
    $email = $entry.Email
    $displayName = $entry.'Display Name'
    
    # Check if the external contact exists
    $externalContact = Get-MailContact -Identity $email -ErrorAction SilentlyContinue

    if (-not $externalContact) {
        try {
            # Create external contact if it doesn't exist
            New-MailContact -Name $displayName -ExternalEmailAddress $email
            Write-Host "Created external contact for $($displayName) with email $($email)"
        } catch {
            Write-Host "Error creating contact for '$($displayName)': $_" -ForegroundColor Red
            continue
        }
    } else {
        # Check if the Display Name matches, ignoring case and leading/trailing spaces
        if ($externalContact.DisplayName.Trim().ToLower() -ne $displayName.Trim().ToLower()) {
            try {
                # Update the Display Name to match the Excel value
                Set-MailContact -Identity $email -Name $displayName
                Write-Host "Updated Display Name for contact '$($email)' to '$($displayName)'"
            } catch {
                Write-Host "Error updating contact name for '$($email)': $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Display Name for contact '$($email)' is already correct."
        }
    }

    # Now check if the email exists in the distribution group
    if (-not (Get-DistributionGroupMember -Identity $distributionGroup | Where-Object { $_.PrimarySmtpAddress -eq $email })) {
        try {
            Add-DistributionGroupMember -Identity $distributionGroup -Member $email
            Write-Host "Added $($email) to $distributionGroup"
        } catch {
            Write-Host "Failed to add $($email): $_"
        }
    } else {
        Write-Host "$($email) is already a member of $distributionGroup"
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
