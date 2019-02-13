<#
    - Widows 10 Current Versions:
        https://technet.microsoft.com/en-us/windows/release-info

    - Download Catalog
        https://www.catalog.update.microsoft.com/Search.aspx?q=1709

    
#>

Param(

    [Parameter(Position=0, HelpMessage="Enable pause for troubleshooting. Default is False")]
    [ValidateSet('True','False')]
    [string]$EnablePause = "False",

    [Parameter(HelpMessage="Override destination source directory. Default is \\server.com\source_directory")]
    [string]$DestDirectory = "\\server.com\source_directory"

)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$OSVersion = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Windows 10 version `n (Example 1809 = 17763) `n 1709 = 16299 `n 1803 = 17134 `n 1809 = 17763", "Windows 10 Version", "16299")
    if($OSVersion.Length -lt 1)
    { 
        Write-Host "No OS Value Entered"
        Exit
    } 
Write-Host "$OSVersion" Selected -ForegroundColor Yellow

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$Language = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Language version `n (Example en-us) `n en-us `n fr-ca `n or `n pl-pl", "Language Version", "en-us")
    if($Language.Length -lt 1)
    { 
         
        Write-Host "No Production Media Value Entered"
        Exit
    } 
Write-Host "$Language Selected" -ForegroundColor Yellow 

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$CreateProdMedia = [Microsoft.VisualBasic.Interaction]::InputBox("Create Production Media? `n (Example false) `n true `n or `n false", "Create Production Media", "false")
    if($CreateProdMedia.Length -lt 1)
    { 
        
        Write-Host "No Production Media Value Entered"
        Exit
    } 
Function DownloadSource {
    
    
    Write-Host "Copying WIM File" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan
    $copysourcedir = "\\server.com\source_directory"
    $0Source = $OSVersion + ".0source_" + $Language
    $OEM = (Get-ChildItem $copysourcedir -Directory | select-string -Pattern ($0source) | Sort-Object -Descending | Select-Object -First 1)
    Write-Host "Copying $Language source files for Windows 10 $OSVersion" -ForegroundColor Yellow
    copy-item $copysourcedir\$OEM\sources\install.wim $PSScriptRoot\$OSVersion\$Language\install_$Language.wim -Verbose
        
    $TimeStamp = Get-Date
}

Function Cleanup {
    Try {
        Write-Host "Cleaning Up" -ForegroundColor Green
        if (Test-Path -path $PSScriptRoot\$OSVersion\$Language) {Remove-Item -Path $PSScriptRoot\$OSVersion\$Language -Force -ErrorAction Continue}
        & DISM /Cleanup-WIM
    }
    Catch 
    {
        Write-Warning "Cleanup failed"
        Throw $Error
    }
}

#Script Logging.  Stop Transcript if already running and begin.
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue" # or "Stop"
Start-Transcript -path "$PSScriptRoot\$OSVersion $Language updateWIM.log" -append

#Check for Windows ADK Installation
Write-Host "Checking for ADK Installation" -ForegroundColor Yellow
$ADKdir = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10"
if (!( $ADKdir))
{Write-Host "Windows ADK is not installed." -ForegroundColor DarkRed
 Exit
}

#Check for Mount Directory
$Mountdir = "$PSScriptRoot\$OSVersion\$Language\mount"
if (!(test-path -path "$PSScriptRoot\$OSVersion\$Language\mount")) {New-Item -ItemType directory -Path "$PSScriptRoot\$OSVersion\$Language\mount"
Write-Host "Creating Mount Directory" -ForegroundColor Yellow}
else
{Write-Host "Mount Directory Already Exists" -ForegroundColor Yellow}

#Copying Servicing Stack Update and Cumulative Updates
Write-Host "Copying Servicing Stack Update to Working Directory" -ForegroundColor Yellow
$PatchSource = "\\server.com\source_directory\Patches"
$SS = "SS"
$CU = "CU"
$DU = "DU"
$UPDATES = $SS,$CU,$DU
Foreach ($i in $UPDATES)
{
 if (!(test-path -path "$PSScriptRoot\$OSVersion\$Language\$i")) {New-Item -ItemType directory -Path "$PSScriptRoot\$OSVersion\$Language\$i"
Write-Host "Creating $i Directory" -ForegroundColor Yellow}
}
$SERVESTACK = "servicestack_" + $OSVersion
$PatchDir = (Get-ChildItem $PatchSource -Directory | select-string -Pattern ($SERVESTACK) | Sort-Object -Descending | Select-Object -First 1) 
Write-Host "Copying $PatchDir" -ForegroundColor Yellow
copy-item $PatchSource\$PatchDir\* $PSScriptRoot\$OSVersion\$Language\SS #-Verbose

$PATCHFOLDER = $OSVersion + "."
$PatchDir = (Get-ChildItem $PatchSource -Directory | select-string -Pattern ($PATCHFOLDER) | Sort-Object -Descending | Select-Object -First 1) 
Write-Host "Copying $PatchDir" -ForegroundColor Yellow
copy-item $PatchSource\$PatchDir\* $PSScriptRoot\$OSVersion\$Language\CU #-Verbose

$DYNAMICUP = "DynamicUpdate_" + $OSVersion
$PatchDir = (Get-ChildItem $PatchSource -Directory | select-string -Pattern ($DYNAMICUP) | Sort-Object -Descending | Select-Object -First 1) 
Write-Host "Copying $PatchDir" -ForegroundColor Yellow
copy-item $PatchSource\$PatchDir\* $PSScriptRoot\$OSVersion\$Language\DU -Recurse #-Verbose


#Download Source files
DownloadSource
Write-Host "Getting WIM and ADK information" -ForegroundColor Yellow
$WIMVersion = (Get-WindowsImage -ImagePath "$PSScriptRoot\$OSVersion\$Language\install_$Language.wim" -Name "Windows 10 Enterprise").Version
$ENTindex = (Get-WindowsImage -ImagePath "$PSScriptRoot\$OSVersion\$Language\install_$Language.wim" -Name "Windows 10 Enterprise").ImageIndex
$tempdir = Get-Location
$tempdir = $tempdir.tostring()
$dismpath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\DISM"
$appToMatch = '*Windows Assessment and Deployment Kit*'
$appVersion = $WIMVersion
$TimeStamp = Get-Date


$result = (get-item $dismpath\dism.exe).versioninfo.productversion

If ($result -lt $appVersion) {
Write-Host "Windows ADK is not correct version to modify WIM." -ForegroundColor DarkRed
Exit
 }
 else
 {Write-Host "Windows ADK $result and WIM Version $WIMVersion" -ForegroundColor Yellow

# Extract ENT WIM
Write-Host "Exporting Enterprise WIM to new file" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
Dism /Export-Image /SourceImageFile:"$PSScriptRoot\$OSVersion\$Language\install_$Language.wim" /SourceIndex:$ENTindex /DestinationImageFile:"$PSScriptRoot\$OSVersion\$Language\install-optimized_$Language.wim"
Write-Host "Export Index $ENTindex Completed" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan


# Mount Wim File
Write-Host "Mounting WIM File" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
#$ENTindex = (Get-WindowsImage -ImagePath "$PSScriptRoot\$OSVersion\$Language\install_optimized_$Language.wim" -Name "Windows 10 Enterprise").ImageIndex
dism /mount-wim /wimfile:$PSScriptRoot\$OSVersion\$Language\install-optimized_$Language.wim /index:1 /mountdir:$PSScriptRoot\$OSVersion\$Language\mount
Write-Host "Mounting Completed" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan


Write-Host "Updating WIM" -ForegroundColor Yellow
Foreach ($i in $UPDATES)
{
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
dism /image:$PSScriptRoot\$OSVersion\$Language\mount /add-package /packagepath:$PSScriptRoot\$OSVersion\$Language\$i
Write-Host "$i Updates Completed" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
}

If ($EnablePause -eq "True") {
    pause
    }

# Cleanup WIM
Write-Host "Cleaning up WIM" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
DISM /Cleanup-Image /Image=$PSScriptRoot\$OSVersion\$Language\mount /StartComponentCleanup /ResetBase
Write-Host "Cleanup Completed" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan

#Unmount WIM
Write-Host "Unmounting WIM File" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan
Dism /Unmount-Image /MountDir:"$PSScriptRoot\$OSVersion\$Language\mount" /Commit
Write-Host "Unmount Completed" -ForegroundColor Yellow
$TimeStamp = Get-Date
Write-Host "$TimeStamp" -ForegroundColor Cyan 


If ($CreateProdMedia -eq "True") {
    If ($Language -eq "en-us") {
    #Copy to Production MDT
    Write-Host "Copy WIM file to Production MDT" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan
    $MDTwimpath = "\\MDTserver\Win10x64\Operating Systems"
    $MDTOSVAR = "_b" + $OSVersion
    $MDTsourceDir = (Get-ChildItem $MDTwimpath -Directory | select-string -Pattern ($MDTOSVAR) | Sort-Object -Descending | Select-Object -First 1) 
    copy-item $PSScriptRoot\$OSVersion\$Language\install-optimized_$Language.wim $MDTwimpath\$MDTsourceDir\Sources\install.wim -Force
    Write-Host "Copy To Production MDT Completed" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan } #Copy to MDT close bracket

    #Create Upgrade Files
    Write-Host "Creating Upgrade Files" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan
    $PatchDir = (Get-ChildItem $PatchSource -Directory | select-string -Pattern ($PATCHFOLDER) | Sort-Object -Descending | Select-Object -First 1) 
    $UpdateVER = ("$PatchDir").Remove(0,14)
    Write-Host "$UpdateVER" -ForegroundColor Yellow
    $DestDirectory = "\\server\destination"
    $OS ="W10x64.ENT_" + $OSVersion + "_"
    $NewDirectory = $OS + "." + $UpdateVER + "_" + $Language
    $0Source = $OSVersion + ".0source_" + $Language
    $OEM = (Get-ChildItem $DestDirectory -Directory | select-string -Pattern ($0source) | Sort-Object -Descending | Select-Object -First 1)
    $NewDirectoryName = ("$OEM").Remove(17,19) + $UpdateVER + "_" + $Language
    Write-Host "New Directory Name is $NewDirectoryName" -ForegroundColor Yellow
    New-Item -ItemType directory -Path $DestDirectory\$NewDirectoryName
    Copy-Item "$DestDirectory\$OEM\*" -Destination "$DestDirectory\$NewDirectoryName\" -Recurse -Verbose
    Write-Host "OEM Source Files Copied" -ForegroundColor Yellow

    $DYNAMICUP = "DynamicUpdate_" + $OSVersion
    $PatchDir = (Get-ChildItem $PatchSource -Directory | select-string -Pattern ($DYNAMICUP) | Sort-Object -Descending | Select-Object -First 1) 
    copy-item $PatchSource\$PatchDir\sources*  -Destination "$DestDirectory\$NewDirectoryName\Sources\" -Recurse -Verbose
    Write-Host "Dynamic Update Source Files Copied... Now Copying Updated WIM File." -ForegroundColor Yellow

    copy-item $PSScriptRoot\$OSVersion\$Language\install-optimized_$Language.wim $DestDirectory\$NewDirectoryName\Sources\install.wim -Force
    Write-Host "UIP Files Completed" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan

    # Create Configuration Manager package
    #NOTE, can only be ran using SCCM CMDLETS
    Write-Host "Creating Configuration Manager Package" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan
    if (!(test-path -path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1")) {New-Item -ItemType directory -Path "$PSScriptRoot\mount"
        Write-Host "Configuration Manager CMDLETS not found" -ForegroundColor Yellow}
    else
    {Write-Host "Configuration Manager CMDLETS found" -ForegroundColor Yellow

    ###### Generated from Configuration Manager
    #
    # Press 'F5' to run this script. Running this script will load the ConfigurationManager
    # module for Windows PowerShell and will connect to the site.
    #
    # This script was auto-generated at '5/30/2018 8:08:22 AM'.

    # Uncomment the line below if running in an environment where script signing is 
    # required.
    #Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

    # Site configuration
    $SiteCode = "XXX" # Site code 
    $ProviderMachineName = "server.name.com" # SMS Provider machine name

    # Customizations
    $initParams = @{}
    #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
    #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

    # Do not change anything below this line

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
    $SCCMVersion = "10.0." + $UpdateVER
    $SCCMName = ("$NewDirectoryName").Remove(0,11)
    $SCCMName2 = ("$SCCMName").split('_')[0]
    $SCCMName3 = "Win10x64 " + $SCCMName2 + " " + $Language + " " + "($UpdateVER)"
    $SCCMdescription = get-date -UFormat "%B %Y"
    #New-CMOperatingSystemInstaller -Path "$DestDirectory\$NewDirectoryName" -Description "Updated with $SCCMdescription CU" -Name "$SCCMName3" -Version "$SCCMVersion"
    #$USPACKAGE = Get-CMOperatingSystemInstaller -Name "$SCCMName3"
    #$USPACKAGEID = ($USPACKAGE).Name
    #Get-CMPackage -Name $USPACKAGEID | Move-CMObject -FolderPath CAS:\OperatingSystemInstaller\IMG\Test
    #Move-CMObject -FolderPath XXX:\OperatingSystemInstaller -InputObject $USPACKAGE -Verbose

    ########### NEW ############

    if ($OSVersion -eq 16299) {
    $RETIRED = (Get-CMOperatingSystemInstaller -Name "RETIRED_*1709*$Language")    $RETIRED_ID = ($RETIRED.packageID)
    Set-CMOperatingSystemInstaller -Id $RETIRED_ID -NewName "NEW_1709 _$Language" -Description "NEW" -Path "$DestDirectory\$NewDirectoryName" -Version "$SCCMVersion"
    $NEWNAME = (Get-CMOperatingSystemInstaller -Name "NEW_*1709*$Language")
    if ($Language -eq en-us) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    if ($Language -eq fr-ca) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    if ($Language -eq pl-pl) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    }

    if ($OSVersion -eq 17134) {
    $RETIRED = (Get-CMOperatingSystemInstaller -Name "RETIRED_*1803*$Language")    $RETIRED_ID = ($RETIRED.packageID)
    Set-CMOperatingSystemInstaller -Id $RETIRED_ID -NewName "NEW_1803 _$Language" -Description "NEW" -Path "$DestDirectory\$NewDirectoryName" -Version "$SCCMVersion"
    $NEWNAME = (Get-CMOperatingSystemInstaller -Name "NEW_*1803*$Language")
    if ($Language -eq en-us) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    if ($Language -eq fr-ca) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    if ($Language -eq pl-pl) {
    Set-CMTSStepUpgradeOperatingSystem -TaskSequenceId XXXXXXXX -UpgradePackage $NEWNAME }
    }

    
    ############################

    Write-Host "Configuration Manger Package Creation Complete" -ForegroundColor Yellow
    $TimeStamp = Get-Date
    Write-Host "$TimeStamp" -ForegroundColor Cyan} #"Configuration Manager CMDLETS found Close Block
  }
} #Process Close Block

#Cleanup

Stop-Transcript