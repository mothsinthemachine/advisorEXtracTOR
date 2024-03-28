# advisorEXtracTOR

AdvisorTrac Report Extractor (version 3)

Author: Jake Sherwood ([GitHub Profile](https://github.com/mothsinthemachine))


## Purpose

The purpose of this program is to take report text data from AdvisorTrac and 
convert it into a CSV file for use with Excel.

Values from these reports are delivered by either HTML (screen) or TEXT (.4sp). 
However, to use these values in Excel requires a trained data-entry eye and 
deep attention, making it hard to import the data. It requires extreme Excel 
know-how or manual entry.

I wrote this code to help me with this and make data reporting from AdvisorTrac 
easier. It's written for a specific department in a specific institution, so it 
may not work for all.


## Requires

- Python3 installed

	
## Process

1.	Get the data from AdvisorTrac delivered in TEXT (.4sp) form and save it to 
	the computer. Remember the location of this file.

2.	Open a command line interface. Suggestions are:

	- Command Prompt (Windows)
	- Powershell (Windows)
	- Terminal (Windows or *nix, Mac included)

3.	Use Python to run the script.

	Possible Syntax:

		python at-report-extract.py <source (required)> <destination (optional)>

	or

		python3 at-report-extract.py <source (required)> <destination (optional)>

4.	Profit.


## Revisions

03/27/24 - Renamed repo from 'at-report-extract to' to 'advisorEXtracTOR' because
			it sounds more amazing.

02/14/24 - Separating tasks into new functions. Formatting output file. Now can
			input multiple files in a directory.

01/24/24 - New version completed. Renamed old version to 'at-report-extract-v3.py'

01/23/24 - Bug discovered in version 2 code. Needed to overhaul file-writing
			design.

01/17/24 - New version completed. Renamed old version to 'at-report-extract-v1.py.'

01/16/24 - Created new version of tool. WIP.

01/10/24 - Renamed directories as A and B (as in 'how to get from A to B'). 
			Outputs by default to B directory as 'output.csv'.

01/09/24 - Initial script and commit to GitHub private repo


