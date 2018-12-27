# LibCubeMacros-Test

# This is a demo to extract vba code from an excel file 

Followed steps from the web page below to automatically create .bas files from Microsoft Excel file

https://www.xltrail.com/blog/auto-export-vba-commit-hook

Requierements
 * Python 3 or greater
 * oletools for python (https://github.com/decalage2/oletools/wiki/Install)
 * the files pre-commit and pre-commit.py in the .git/hooks directory
 
 
 Steps: 
 
	1. Install python 3 or greater (https://www.python.org/downloads/)
	2. Set python and pip to PATH system variable: 
		Example:
			pip: C:\Program Files\Python\Python37-32\Scripts
			python: C:\Program Files\Python\Python37-32
	3. Install oletools for python (reference to https://github.com/decalage2/oletools/wiki/Install)
		Online:
			-Windows: In command line write the next command
				pip install -U oletools
			-Unix based: In command shell write the next command
				sudo -H pip install -U oletools
		Offline:
			First, download the oletools archive on a computer with Internet access:
			* Latest stable version: from https://pypi.org/project/oletools/ or https://github.com/decalage2/oletools/releases
			* Development version: https://github.com/decalage2/oletools/archive/master.zip
			
			Copy the archive file to the target computer.
	
			On Linux, Mac OSX, Unix, run the following command using the filename of the archive that you downloaded:
				sudo -H pip install -U oletools.zip
			On Windows:
				pip install -U oletools.zip
	4. Download the files from this repository:
		\scripts4Git\pre-commit
		\scripts4Git\pre-commit.py
		
		And copy them into your repository folder .git\hooks
	
	5. Finally you just need to copy your excel file into the repo then git add . and git commit, the pre-commit under your git\hooks will do the magic