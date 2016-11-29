<#
****------------- Script Information -------------****
Brief			:	Auto copy optical drive to destination
Function		:	Copies, by file, whatever is in the optical drive to the User defined destination, then ejects the disc waiting for the next disc
Author  		:	Mike Yurick
Website			:	www.mikeyurick.com
Version			:	v0.98.3
Last Updated	:	09/16/2015
Moved to GitHub	:	11/29/2016
Prerequisites	:	Detailed below in the "Script Start Requirements" section
Options			:	Adjustable via "USER VARIABLES" below
License 		:	Free for personal use and "shareable" if you keep this "Script information" section and "commenting" in tact
Warranty		:	Absolutely none. Use at your own risk. Understand the code before running. No warranties expressed or implied.
Warnings		:	This script currently has no regard for free space and will fill the destination to capacity. This script will likely not work on most dedicated audio and video CD's, DVD's, Games, BluRay's, Special MBR's, etc.
Support			:	No support is currently available for this product
Release notes	:	Pre-release Alpha version
Bugs			:	If you select a drive letter that is not optical with an optical disc in the drive, it will attempt to copy that drive letter (even though differing from the optical drive's)
Testing	notes	:	Minor modifications were made to the code to prepare the script for publishing and are untested.
Things to add	:	Alternative (possibly automated) options to RoboCopy for performing Copy Operation
					Flash drive source support
Version History	:	v0.96 - Done - Add a minimum size variable to check against to help find audio CDs
Legal			:	"My lawyers made me put this in here." Don't sue me. Only use this script if you understand how and why it works.
****------------- End Script Information -------------****

****------------- Script Start Requirements -------------****
  1. Run on a Windows machine with PowerShell installed, patched, and updated correctly
  2. You Must have robocopy in the path!
  	    Download and install this if you haven't already
		A recommended version is currently available here:
		http://www.mikeyurick.com/blogs/media/downloads/robocopy.exe
  3. Must have ExecutionPolicy set correctly
		Run: "Get-ExecutionPolicy" to get the current policy. Likely will be set to Restricted
		Run: "Set-ExecutionPolicy RemoteSigned"
			Select Yes with the Pop-up prompt
		If that doesn't work...
			Run: "Set-ExecutionPolicy -ExecutionPolicy Unrestricted"
				Please research and understand the implications and potential security risks by running in this mode.
  4. The newly recommended ISE/IDE is to use the included PowerShell ISE or Visual Studio.
		Old 4: I also recommend having PowerGUI Script Editor installed and running
  		Old link: http://en.community.dell.com/techcenter/powergui/m/bits/20439049
****-------------  End Script Start Requirements  -------------****
#>

#------USER VARIABLES-------------

$Script:Source = 'F:\' #Make sure this is your optical drive!
$Script:BaseDestination = 'D:\OpticalCopies\' #Destination where you want the disc copied to
$NumberOfLoops = 3600 #Number of times to run. At 1 second intervals (adjustable with $TimeToWait), this will run for an hour. Longer if copy operations take place.
$TimeToWait = 500 #milliseconds to wait before next run. Also gives drive time to spin up.
$DestinationStartSerial = 14 #This will add a number to the end of the folder name in case disc volume labels are the same.
$AllowSkippingDiscs = $true #This will skip a disc if the volume label and size is the same
$SkipIfNotLargerThanBytes = 5000 #Helps to determine if audio CD.
#Future - auto eject optical on/off

#----End User Vars------------------

#-- Variable Declaration and cleanup ---

$Script:NeedToCreateFSO = $true

#$Script:TotalFolderSize = 0
#$Script:TotalFolderSizeB = 0
#$Script:TotalFolderSizeKB = 0
#$Script:TotalFolderSizeMB = 0
#$Script:TotalFolderSizeGB = 0
	
#--------------------------------------

function EjectCDRom { 
    $sa = new-object -com Shell.Application 
    #$sa.Namespace(17).ParseName("D:").InvokeVerb("Eject") #Old hardcoded Source may be needed in some cases. Comment out next line if "uncommenting this one"
	$sa.Namespace(17).ParseName($Script:Source).InvokeVerb("Eject")
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sa) 
    Remove-Variable sa 
} 

#function get-remoteregistry([string]$FolderName,[string]$FcnComputerName)
function GetFolderSize([string]$FolderName){

	$Script:TotalFolderSize = 0
	$Script:TotalFolderSizeB = 0
	$Script:TotalFolderSizeKB = 0
	$Script:TotalFolderSizeMB = 0
	$Script:TotalFolderSizeGB = 0

	if ($Script:NeedToCreateFSO -eq $true) {
		$FileSystem = new-object -com "Scripting.FileSystemObject";
		# It appears that the following line is causing False to appear on the screen...
		$Script:NeedToCreateFSO -eq $false
	}
	
	$getFolderAttr = $FileSystem.GetFolder($FolderName);
	
	$Script:TotalFolderSize = $getFolderAttr.Size #Single line folder properties. And faster!
	
	<# For what we're doing here, the above single liner should be sufficient for mosts needs
	$getSubFolderAttr = $getFolderAttr.SubFolders
	if ($getSubFolderAttr.Count -eq 0) {
		$Script:TotalFolderSize = $getFolderAttr.Size		
		Write-Host 'No Subfolders at:' $FolderName 'Using root size (' $Script:TotalFolderSize ').' -ForegroundColor Red					
	} else {				
		foreach ($SubFolders in $getSubFolderAttr) {
			$Script:TotalFolderSize += $SubFolders.Size
		}			
	}
	#>
	
	$Script:TotalFolderSizeB = "{0:N2}" -f ($Script:TotalFolderSize)
	$Script:TotalFolderSizeKB = "{0:N2}" -f ($Script:TotalFolderSize / 1024)
	$Script:TotalFolderSizeMB = "{0:N2}" -f ($Script:TotalFolderSize / 1024 / 1024)
	$Script:TotalFolderSizeGB = "{0:N2}" -f ($Script:TotalFolderSize / 1024 / 1024 / 1024)	
}

[char]13
[char]13

$TimeToWaitInSeconds = $TimeToWait / 1000
$TotalRunSeconds = $NumberOfLoops * $TimeToWaitInSeconds
$TotalRunMinutes = $TotalRunSeconds / 60
$TotalRunHours = $TotalRunMinutes / 60

Write-Host "----------------------------------------------------"
Write-Host "|      Script Started. Insert Discs when ready      |"
Write-Host '|      Wait Seconds     :' $TimeToWaitInSeconds
Write-Host '|      Total Run Minutes:' $TotalRunMinutes
Write-Host '|      Total Run Hours  :' $TotalRunHours
Write-Host "|      Hit the stop key to cancel                   |"
Write-Host "----------------------------------------------------"

for ($Loop = 1; $Loop -le $NumberOfLoops; $Loop++) {	
	#Clear variables for this loop
	$TotalSourceFolderSize = 0
	$TotalDestinationFolderSize = 0
	$SkipDiskFlag = $false
	$VolumeLabel = $null
	
	$drives = [System.IO.DriveInfo]::GetDrives()
	$cdromdrive = $drives | Where-Object { $_.DriveType -eq 'CDRom' -and $_.IsReady }
	$VolumeLabel = $cdromdrive.VolumeLabel
	
	if (!$cdromdrive) {
	    #throw 'No removable drives found.'
		Write-Host '.' -NoNewline
	} else {
		[char]13
		Write-Host 'CDRom' $VolumeLabel 'in drive'
		
		start-sleep -Milliseconds $TimeToWait #wait a little to allow spin up
		
		$DestinationStartSerial++
		$VolumeLabel = $VolumeLabel -replace(" ","")
		$VolumeLabel = $VolumeLabel -replace("&","")
		
		GetFolderSize $Script:Source
		$TotalSourceFolderSize = $Script:TotalFolderSizeB
		Write-Host 'CD Size:' $TotalSourceFolderSize

		if ($VolumeLabel.Contains("AudioCD") -and $SkipIfNotLargerThanBytes -ge $Script:TotalFolderSize) {
			Write-Host 'This is likely an Audio CD and will need to be ripped.' -ForegroundColor Red
			$SkipDiskFlag = $true
		} else {
			foreach ($DestinationSubFolders in Get-ChildItem $Script:BaseDestination) {
				if ($DestinationSubFolders.Name.Contains($VolumeLabel)) {
		
					Write-Host 'Folder:' $DestinationSubFolders.Name 'exists and contains CD name:' $VolumeLabel

					$Script:DestinationSubFoldersPath = $Script:BaseDestination + $DestinationSubFolders
					GetFolderSize $Script:DestinationSubFoldersPath
					$TotalDestinationFolderSize = $Script:TotalFolderSizeB
					Write-Host 'Folder size:' $TotalDestinationFolderSize
					
					if ($TotalSourceFolderSize -eq $TotalDestinationFolderSize) {
						Write-Host 'CD and Destination are Identically sized' -ForegroundColor Red
						$SkipDiskFlag = $true
					}
				}
			}
		}
		
		if ($SkipDiskFlag -eq $true -and $AllowSkippingDiscs -eq $true) {
			Write-Host 'Skipping disc. Change user Vars if not wanted...' -ForegroundColor Red
			[char]13
		} else {
		
			$NewFolderName = $VolumeLabel + '-' + $DestinationStartSerial
			New-Item -Path $Script:BaseDestination -Name $NewFolderName -type directory
			
			$FinalDestination = $Script:BaseDestination + $NewFolderName + '\'
			
			$CommandToRun = 'robocopy ' + $Script:Source + ' ' + $FinalDestination + ' /E /ETA /R:5 /W:10'
			
			cmd /c $CommandToRun
		}
		
		EjectCDRom
		
	}
	start-sleep -Milliseconds $TimeToWait
}

Write-Host "-----------------------------------------------------------"
Write-Host "| Script execution completed. Restart script to continue. |"
Write-Host "-----------------------------------------------------------"


