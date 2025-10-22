# advisorEXtracTOR

AdvisorTrac Report Extractor written in Python

Authored By: [Moths in the Machine](https://github.com/mothsinthemachine)

## Purpose

The purpose of this program is to extract data from AdvisorTrac reports and convert it into a CSV file.

Data from these reports are delivered by either HTML (screen) or TEXT (.4sp). However, importing the data to spreadsheet software was impossible, and required hours of manual entry.

I wrote this CLI program to make data extraction AdvisorTrac easier. It is written for a department in a specific institution, so it may not work for all. However, the algorithm is worth perserving and can be modified.

## Requirements for use

- Python3.10 or newer installed.

## Process for Version 5 (recommended and newest)

1. Create two text files (files that end with extension `.txt`) in the same folder as the program named the following:

	1. `centers.txt`: List the exact names of the centers to be extracted from the data, each on a separate line.

	2. `reasons.txt`: List the exact names of the reasons to be extracted from the data, each on a separate line.

2. Create an `input` folder in the same folder as the program. The extractor will be looking for this folder to extract the data from. Place your `.4sp` files in this folder. Name them appropriately if you wish to for the extraction to be ordered (ie. 1-9 or A-Z ordering).

3. Run the program `advisorEXtracTOR_5.py` by double-clicking it. You may be prompted to confirm overwriting any previous `output.csv` files that are in the main directory.

	- If double-clicking does not work, you can run the program using a CLI (Command Prompt or PowerShell) and entering the following command from the main directory.

```PowerShell
# Standard
python3 advisorEXtracTOR_v5.py

# Alternate if py is used as an alias of python3
py advisorEXtracTOR_v5.py

# Old (please verify you have Python3 and not Python2 before using this style)
python --version
python advisorEXtracTOR_v5.py
```

4. Profit. A new file named `output.csv` will appear in the main directory. If everything was done correctly, the data will be stored in this file.

## Process for Version 4

1. Get the data from AdvisorTrac delivered in TEXT (.4sp) form and save it to the computer. Place the file into the `input` folder included in the directory where the extractor resides or store in a directory of your choosing.

2. Open a CLI. Suggestions are:
	<br>- Command Prompt or Powershell (Windows)
	<br>- Terminal (Windows or *nix, Mac included)

4.	Use Python to run the script. Possible syntaxes:

```Powershell
# Standard
python3 advisorEXtracTOR_v4.py  <source(required)>  <destination(optional)>

# Alternate
py advisorEXtracTOR_v4.py  <source(required)>  <destination(optional)>

# Old (please verify you have Python3 and not Python2 before using this style)
python --version
python advisorEXtracTOR_v4.py  <source(required)>  <destination(optional)>
```

5.	Profit. Output will default to `output.csv` in same directory as extractor if no destination is specified in previous step.

## Known Issues

- Only works on Windows computers. The issue is the filepaths use `\\` in a few places in the program. Will need to remedy to work on Linux systems as well.

## Revisions

| Date | Revision |
| --- | --- |
| 10/22/25 | Version 5 released. Simplified workflow and more legible code. |
| 09/24/25 | Resuming development. Needed for upcoming project. |
| 05/14/25 | Housecleaning. Goal is to make this project more modular for v4. |
| 04/23/25 | Housecleaning. Moved old versions of advisorEXtracTOR to private repo. Updated README. Invited others to collaborate on v4. |
| 03/09/25 | Rewritten README file. |
| 06/28/24 | Corrections and clarifications made to Process steps in ReadMe. |
| 06/23/24 | Made repo public. Added a discussion board. Updated author name in ReadMe. |
| 03/27/24 | Renamed repo from "at-report-extract" to "advisorEXtracTOR" because it looks cool. |
| 02/14/24 | Separating tasks into new functions. Formatting output file. Now can input multiple files in a directory. |
| 01/24/24 | New version completed. Renamed old version to "at-report-extract-v3.py" | 
| 01/23/24 | Bug discovered in version 2 code. Needed to overhaul file-writing design. |
| 01/17/24 | New version completed. Renamed old version to "at-report-extract-v1.py." |
| 01/16/24 | Created new version of tool. WIP. |
| 01/10/24 | Renamed directories as A and B (as in "how to get from A to B"). Outputs by default to B directory as "output.csv". |
| 01/09/24 | Initial script and commit to GitHub private repo. |
