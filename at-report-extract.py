'''
AdvisorTrac Report Extractor (version 2)
at-report-extract2.py

Author: Jake Sherwood (MothsInTheMachine on GitHub)
Last Modified: Jan. 18, 2024

Purpose:
The purpose of this program is to take report text data from AdvisorTrac and 
convert it into a CSV file for use with Excel.

Reason:
Values from these reports are delivered by either HTML (screen) or TEXT (.4sp). 
However, to use these values in Excel requires a trained data-entry eye and 
deep attention, making it hard to import the data. It requires extreme Excel 
know-how or manual entry.

I wrote this code to help me with this and make data reporting from AdvisorTrac 
easier. It's written for a specific department in a specific institution, so it 
may not work for all.

Requires:
	- Python3 installed

Process:
1.	Get the data from AdvisorTrac delivered in TEXT (.4sp) form and save it to 
	the computer. Remember the location of this file.
2.	Open a command line interface. Suggestions are:
	- Command Prompt (Windows)
	- Powershell (Windows)
	- Terminal (Windows or *nix, Mac included)
3.	Use Python to run the script.

	Possible Syntax:

	> python at-report-extract.py <source (required)> <destination (optional)>

		or

	> python3 at-report-extract.py <source (required)> <destination (optional)>

4.	Profit.

Revisions:

01/17/24 - New version completed. Renamed old version to 'at-report-extract-v1.py.'

01/16/24 - Created new version of tool. WIP.

01/10/24 - Renamed directories as A and B (as in 'how to get from A to B'). 
	Outputs by default to B directory as 'output.csv'.

01/09/24 - Initial script and commit to GitHub private repo
'''

import sys
import os, os.path

def printHeader():
	print()
	print('/*******************************************/')
	print('/* AdvisorTrac Report Extractor in process */')
	print('/*******************************************/')
	print()

def printFooter():
	print()
	print('/*******************************************/')
	print('/*  Qutting AdvisorTrac Report Extractor.  */')
	print('/*******************************************/')
	print()

def quitProcess(reason='unknown'):
	print(f'Sorry, I could not understand that command. The reason is {reason}.')
	print('Please enter the path and file name to use extractor.')
	print('Syntax:\n\n\t> python at-report-extract.py <PATH TO 4sp FILE> [OPTIONAL OUTPUT PATH]\n')
	printFooter()
	quit()

def main():
	

	separator = ','
	defaultInputPath  = 'A\\reportOut.4sp'  # The default input file path and name from AdvisorTrac
	defaultOutputPath = 'B\\output.csv'     # The default output file path and name
	subjectLineLabel  = '<!--Group1 Dif-->' # Distinguishes the Subject line
	dateLineLabel = 'from'					# Distinguishes the date
	labLineLabels = ['DC Math Lab Total','DC Writing Lab Total','SV Math Lab Total','SV Writing Lab Total','Grand Total']
	headersLine = separator.join(['Visits','Hours','Students'])

	printHeader()

	nArgs = len(sys.argv) # Number of arguments from command.

	if (nArgs > 3):
		quitProcess('the extractor cannot process excess arguments for command')
	elif (nArgs == 1):
		inputPath  = defaultInputPath
		outputPath = defaultOutputPath
	else:
		inputPath = sys.argv[1]
		if not (os.path.exists(inputPath)):
			quitProcess(f'the extractor cannot find \"{inputPath}\"')

	if (nArgs == 3):
		outputPath = sys.argv[2]
	else:
		outputPath = defaultOutputPath
	
	# Verify overwriting of existing output file
	if (os.path.exists(outputPath)):
		inputStr = input(f'The output file \"{outputPath}\" already exists. Do you want to overwrite it? [y/n]: ')
		choice = inputStr[0].lower()
		while (choice != 'n') and (choice != 'y'):
			print("Invalid choice. Please enter Y for 'Yes' and N for 'No'...")
			inputStr = input('The output file name already exists. Do you want to overwrite it? [y/n]: ')
			choice = inputStr[0].lower()
		
		if (choice == 'n'):
			print("Cancelling process. Please run command again later.")
			printFooter()
			exit()
		
		if (choice == 'y'):
			print(f'Confirmed overwriting previous file \"{outputPath}\".')
		
	
	print(f'Extractor will process file found at \"{inputPath}\" and output to \"{outputPath}\".')
	inputFile  = open(inputPath,'r')
	outputFile = open(outputPath,'w')

	for line in inputFile:
		writeToFile = False
		if (line.find(dateLineLabel) != -1):
			found = line.find(dateLineLabel)
			outputFile.write('Dates ' + line[found:found+29] + '\n\n')
			continue
		elif (line.find(subjectLineLabel) != -1):
			cleanedList = list(filter(lambda str: str != '', line[len(subjectLineLabel):-9].split('\t')))
			writeToFile = True
		else:
			for label in labLineLabels:
				if (line.find(label) != -1):
					cleanedList = list(filter(lambda str: str != '', line[:-9].split('\t')))
					# Remove HTML tags from Grand Total line
					if (label == 'Grand Total'):
						cleanedList[1] = cleanedList[1].replace('<b>','').replace('</b>','')
						cleanedList[3] = cleanedList[3].replace('<b>','').replace('</b>','')
					writeToFile = True
		
		if (writeToFile):
			outputFile.write(cleanedList[0] + '\n')
			outputFile.write(headersLine + '\n')
			outputFile.write(separator.join(cleanedList[1:]) + '\n')

	print("Process completed successfully.")
	inputFile.close()
	outputFile.close()
	printFooter()
	exit()

# Run the script
main()