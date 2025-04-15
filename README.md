# advisorEXtracTOR

AdvisorTrac Report Extractor written in Python

Authored By: [Moths in the Machine](https://github.com/mothsinthemachine)

## Purpose

The purpose of this program is to extract data from AdvisorTrac reports and convert it into a CSV file.

Data from these reports are delivered by either HTML (screen) or TEXT (.4sp). However, importing the data to spreadsheet software was impossible, and required hours of manual entry.

I wrote this CLI program to make data extraction AdvisorTrac easier. It is written for a department in a specific institution, so it may not work for all. However, the algorithm is worth perserving and can be modified.

## Requirements for use

- Python3.10 or newer installed.
	
## Process

1. Get the data from AdvisorTrac delivered in TEXT (.4sp) form and save it to the computer. Place the file into the `input` folder included in the directory where the extractor resides or store in a directory of your choosing.

2. Open a CLI. Suggestions are:
	<br>- Command Prompt or Powershell (Windows)
	<br>- Terminal (Windows or *nix, Mac included)

4.	Use Python to run the script. Possible syntaxes:

```Powershell
# Standard
python3 advisorEXtracTOR.py  <source(required)>  <destination(optional)>

# Alternate
py advisorEXtracTOR.py  <source(required)>  <destination(optional)>

# Old (please verify you have Python3 and not Python2 before using this style)
python --version
python advisorEXtracTOR.py  <source(required)>  <destination(optional)>
```

5.	Profit. Output will default to `output.csv` in same directory as extractor if no destination is specified in previous step.


## Revisions

03/09/25 - Rewritten README file.

06/28/24 - Corrections and clarifications made to Process steps in ReadMe.

06/23/24 - Made repo public. Added a discussion board. Updated author name in ReadMe.

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


