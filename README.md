###------------- Script Information -------------
* **Brief**			:	Auto copy optical drive to destination
* **Function**		:	Copies, by file, whatever is in the optical drive to the User defined destination, then ejects the disc waiting for the next disc
* **Author**  		:	Mike Yurick
* **Website**			:	www.mikeyurick.com
* **Version**			:	v0.98.3
* **Last Updated**	:	09/16/2015
* **Moved to GitHub**	:	11/29/2016
* **Prerequisites**	:	Detailed below in the "Script Start Requirements" section
* **Options**			:	Adjustable via "USER VARIABLES" below
* **License** 		:	Free for personal use and "shareable" if you keep this "Script information" section and "commenting" in tact
* **Warranty**		:	Absolutely none. Use at your own risk. Understand the code before running. No warranties expressed or implied.
* **Warnings**		:	This script currently has no regard for free space and will fill the destination to capacity. This script will likely not work on most dedicated audio and video CD's, DVD's, Games, BluRay's, Special MBR's, etc.
* **Support**			:	No support is currently available for this product
* **Release notes**	:	Alpha version
* **Bugs**			:	If you select a drive letter that is not optical with an optical disc in the drive, it will attempt to copy that drive letter (even though differing from the optical drive's)
* **Testing notes**	:	Minor modifications were made to the code to prepare the script for publishing and are untested.
* **Things to add**	:	Alternative (possibly automated) options to RoboCopy for performing Copy Operation
					Flash drive source support
* **Version History**	:	v0.96 - Done - Add a minimum size variable to check against to help find audio CDs

###------------- Script Start Requirements -------------
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
  4. The newly recommended ISE/IDE is to use the default PowerShell ISE or Visual Studio.
