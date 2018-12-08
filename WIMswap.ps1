
##### Generated from Configuration Manager
#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows Powershell adn will connect to the site.
#
# This script was auto-generated at '12/01/2018 8:08:22 PM'.

# Uncomment the line below if running in an environment where script signing is 
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

#Site configuration
$SiteCode = 'XXX' #Replace XXX with your site code
$ProviderMachineName = "XXX.xxx.com" #Replace XXX.xxx.com with your SMS Provider machine name

# Customizations
$initParams = @{}
#initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# No not change anything below this line

# Import the ConfigurationManager.psd1 module
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams
    }

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

####### End of Generated Configuration Manager Code

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$OSVersion = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Windows 10 version `n (Example 1709 or 1803)", "Windows 10 Version", "")
    if($OSVersion.Length -lt 1)
    {
        Write-Host "No OS Value Entered"
        Exit
    }
Write-Host "$OSVersion" Selected -ForegroundColor Yellow
#=+=+=+=+=+=+=+= OPERATING SYTEM VARIABLES =+=+=+=+=+=+=+=
Write-Host Getting "$OSVersion" WIM Information -ForegroundColor Yellow

#=+=+=+=+=+=+=+= OSD WIMS =+=+=+=+=+=+=+=
$NEWWIM = (Get-CMOperatingSystemImage -Name "NEW_*$OSVersion*WIM")
$NEWWIMID = ($NEWWIM.package.ID)

$PRODWIM = (Get-CMOperatingSystemImage -Name "_PROD_*$OSVersion*WIM")
$PRODWIMID = ($PRODWIM.package.ID)

$PREV1WIM = (Get-CMOperatingSystemImage -Name "1_PREV_*$OSVersion*WIM")
$PREV1WIMID = ($PREV1WIM.package.ID)

$PREV2WIM = (Get-CMOperatingSystemImage -Name "2_PREV_*$OSVersion*WIM")
$PREV2WIMID = ($PREV2WIM.package.ID)

#=+=+=+=+=+=+=+= Upgrade Packages =+=+=+=+=+=+=+=
# Using the US-EN version as an example
$NEW_ENUS = (Get-CMOperatingSystemInstaller -Name "NEW_*$OSVersion*en-us")
$NEW_ENUSID = ($NEW_ENUS.package.ID)

$PROD_ENUS = (Get-CMOperatingSystemInstaller -Name "_PROD_*$OSVersion*en-us")
$PROD_ENUSID = ($PROD_ENUS.package.ID)

$PREV1_ENUS = (Get-CMOperatingSystemInstaller -Name "1_PREV_*$OSVersion*en-us")
$PREV1_ENUSID = ($PREV1_ENUS.package.ID)

$PREV2_ENUS = (Get-CMOperatingSystemInstaller -Name "2_PREV_*$OSVersion*en-us")
$PREV2_ENUSID = ($PREV2_ENUS.package.ID)

#=+=+=+=+=+=+=+= Rename Operating System Packages =+=+=+=+=+=+=+=

Write-host Renaming "$OSVersion" WIM Files -ForegroundColor Yellow
#=+= OSD =+=
Set-CMOperatingSystemImage -Id $PREV2WIMID -NewName "RETIRED_$OSVersion _WIM" -Description "RETIRED"
Set-CMOPeratingSystemImage -Id $PREV1WIMID -NewName "2_PREV_$OSVersion _WIM" -Description "2-Previous Month"
Set-CMOperatingSystemImage -Id $PRODWIMID -NewName "1_PREV_$OSVersion _WIM" -Description "1-Previous Month"
Set-CMOperatingSystemImage -Id $NEWWIMID -NewName "_PROD_$OSVersion _WIM" -Description "Current Month"

#=+= Upgrade =+=
Set-CMOperatingSystemInstaller -Id $PREV2_ENUSID -NewName "RETIRED_$OSVersion _en-us" -Description "RETIRED"
Set-CMOPeratingSystemInstaller -Id $PREV1_ENUSID -NewName "2_PREV_$OSVersion _en-us" -Description "2-Previous Month"
Set-CMOperatingSystemInstaller -Id $PROD_ENUSID -NewName "1_PREV_$OSVersion _en-us" -Description "1-Previous Month"
Set-CMOperatingSystemInstaller -Id $NEW_ENUSID -NewName "_PROD_$OSVersion _en-us" -Description "Current Month"

#=+=+=+=+=+=+=+= Update Task Sequences =+=+=+=+=+=+=+=
# If managing multiple versions, you can set an if ($OSVersion = XXX) also and have the following steps between {}.
#=+= OSD =+=
$PRODWIM = (Get-CMOperatingSystemImage -Name "_PROD_*$OSVersion*WIM")
Write-Host "$PRODWIM" Selected -ForegroundColor Yellow
Set-CMTSStepApplyOperatingSystem -TaskSequenceId XXXXXXXX -ImagePackage $PRODWIM -ImagePackageIndex 1 #Change XXXXXXXX to your task sequence ID.

#=+= Upgrade =+=
$PROD_ENUS = (Get-CMOperatingSystemInstaller -Name "_PROD_*$OSVersion*en-us")
Write-Host "$PROD_ENUS" Selected -ForegroundColor Yellow
Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $PROD_ENUS  #Change XXXXXXXX to your task sequence ID. 

Write-Host Task Sequences Updated