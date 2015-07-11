###############################################################################
#
# Author: Omar A. Ansari
# Version: 1.2, last modified: 5/20/15
# Description: Script to automate Office 2016 Removal. To be used by
# the University of South Carolina - University Technology Services 
# Department.
###############################################################################

import subprocess

userDirectory = str(subprocess.check_output('echo $HOME', shell = True)).rstrip()

FAIL = '\033[91m'
OKGREEN = '\033[92m'
OKBLUE = '\033[94m'
ENDC = '\033[0m'

###############################################################################

###############################################################################
# Step 1: Remove the Microsoft Office 2016 Folder
###############################################################################

try:
	subprocess.call('rm -rf Microsoft\ Office\ 2016/', cwd='/Applications/',
	 shell=True)
	print OKGREEN + "REMOVED: Microsoft Office 2016 Folder" + ENDC
except:
	print FAIL + "NOTE: Could not find Microsoft Office 2016 Folder." + ENDC

###############################################################################
# Step 2: Remove preference and license files and Office folder 
###############################################################################

try:
	subprocess.call('rm -rf com.microsoft.*', cwd= userDirectory+'/Library/Preferences/',
	 shell=True)
	print OKGREEN + "REMOVED: com.microsoft.* from /Library/Preferences/" + ENDC
except:
	print FAIL + "NOTE: Could not find any com.microsoft files." + ENDC

try:
	subprocess.call('rm -rf com.microsoft*', cwd=userDirectory+'/Library/Preferences/ByHost/',
	 shell=True)
	print OKGREEN + "REMOVED: com.microsoft.* from /Library/Preferences/ByHost/" + ENDC
except:
	print FAIL + "NOTE: Could not find any com.microsoft files." + ENDC

try:
	subprocess.call('rm -rf Office', cwd=userDirectory+'/Library/Application Support/Microsoft/',
	 shell=True)
	print OKGREEN + "REMOVED: Office Folder" + ENDC
except:
	print FAIL + "NOTE: Could not find Office folder." + ENDC

try:
	subprocess.call('rm -rf com.microsoft.office.licensing.helper.plist',
		cwd = '/Library/LaunchDaemons/',
	 	shell=True)
	print OKGREEN + "REMOVED: com.microsoft.office.licensing.helper.plist" + ENDC
except:
	print FAIL + "NOTE: Could not find com.microsoft.office.licensing.helper.plist" + ENDC

try:
	subprocess.call('rm -rf com.microsoft.office.licensing.plist',
		cwd ='/Library/Preferences/',
	 	shell=True)
	print OKGREEN + "REMOVED: com.microsoft.office.licensing.plist" + ENDC
except:
	print FAIL + "NOTE: Could not find com.microsoft.office.licensing.plist" + ENDC

try:
	subprocess.call('rm -rf com.microsoft.office.licensing.helper',
		cwd='/Library/PrivilegedHelperTools/',
	 	shell=True)
	print OKGREEN + "REMOVED: com.microsoft.office.licensing.helper" + ENDC
except:
	print FAIL + "NOTE: Could not find com.microsoft.office.licensing.helper" + ENDC

###############################################################################
# Step 3: Remove Microsoft folders and Office2011 files
###############################################################################

try:
	subprocess.call('rm -rf Microsoft', cwd='/Library/Application Support/',
	 shell=True)
	print OKGREEN + "REMOVED: Microsoft folder" + ENDC
except:
	print FAIL + "NOTE: Could not find Microsoft folder." + ENDC

try:
	subprocess.call('rm -rf Microsoft', cwd='/Library/Fonts/',
	 shell=True)
	print OKGREEN + "REMOVED: Microsoft Fonts folder." + ENDC
except:
	print FAIL + "NOTE: Could not find Microsoft Fonts Folder" + ENDC

try:
	subprocess.call('rm -rf Office2011_*', cwd='/Library/Receipts/',
	 shell=True)
	print OKGREEN + "REMOVED: Office2011_* files" + ENDC
except:
	print FAIL + "NOTE: Could not find Office2011_* files" + ENDC

try:
	subprocess.call('rm -rf com.microsoft.office*', cwd='/private/var/db/receipts/',
	 shell=True)
	print OKGREEN + "REMOVED: com.microsoft.office* from /private/var/db/receipts/" + ENDC
except:
	print FAIL + "NOTE: Could not find any com.microsoft files." + ENDC


print OKBLUE + "Uninstallation complete, remember to remove the Office icons from the dock and reboot the computer. \n" + ENDC








