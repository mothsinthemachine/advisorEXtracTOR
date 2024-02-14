'''
AdvisorTrac Report Extractor (version 3)
at-report-extract.py

Author: Jake Sherwood (MothsInTheMachine on GitHub)
Last Modified: Jan. 24, 2024

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

02/14/24 - Separating tasks into new functions.

01/24/24 - New version completed. Renamed old version to 'at-report-extract-v3.py'

01/23/24 - Bug discovered in version 2 code. Needed to overhaul file-writing
			design.

01/17/24 - New version completed. Renamed old version to 'at-report-extract-v1.py.'

01/16/24 - Created new version of tool. WIP.

01/10/24 - Renamed directories as A and B (as in 'how to get from A to B'). 
			Outputs by default to B directory as 'output.csv'.

01/09/24 - Initial script and commit to GitHub private repo
'''

import sys
import os, os.path
from datetime import datetime


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


def bar(width,_symbol='-'):
	str = ''
	for i in range(width):
		str += _symbol
	return str


def defineKeywords(_labLabels=[]):
	keywords = []
	if (len(_labLabels) > 0):
		keywords.extend(_labLabels)
	keywords.extend(['DC Math Lab Total:','DC Writing Lab Total:','SV Math Lab Total:','SV Writing Lab Total:','Grand Total:'])
	keywords.extend(['Chemistry Total:','CIS/Programming Total:','Math Total:','Other Total:','Other Course Tutoring Total:'])
	keywords.extend(['Languages Total:','MS Office Total:','Reading Total:','Social Sciences Total:','Workshop Total:','Writing Total:'])
	return keywords


def verifyArguments(n):
	inputPath  = '.\\input'
	outputPath = '.\\output.csv'

	# Verify input and output paths
	if (n == 1):
		# Use default paths for processing
		pass
	
	if (n == 2  or  n == 3):
		# Use a different input path, but default output path
		inputPath = sys.argv[1]
		if not os.path.exists(inputPath):
			quitProcess(f'the extractor cannot find \"{inputPath}\"')
	
	if (n == 3):
		# Use a different input and output path
		outputPath = sys.argv[2]
		# Edit the formatting of the output path to CSV files
		pass
	
	if (n > 3):
		# Too many arguments. Exit extractor
		quitProcess(f'there are too many arguments to use')

	return inputPath,outputPath


def verifyOverwrite(outputPath):
	if (os.path.exists(outputPath)):
		# Verify overwriting of existing output file
		inputStr = input(f'The output file \"{outputPath}\" already exists. Do you want to overwrite it? [y/n]: ')
		choice = inputStr[0].lower()
		while (choice != 'n') and (choice != 'y'):
			print("Invalid choice. Please enter Y for 'Yes' and N for 'No'...")
			inputStr = input('The output file name already exists. Do you want to overwrite it? [y/n]: ')
			choice = inputStr[0].lower()
		
		if (choice == 'n'):
			return False
		
		if (choice == 'y'):
			print(f'Confirmed overwriting previous file \"{outputPath}\".')
			return True
	else:
		# Path does not exists
		return True


	# if not (os.path.exists(inputPath)):
	# 	quitProcess(f'the input file path \"{inputPath}\" does not exist')
	# else:
	# 	print(f'Extractor will process file found at \"{inputPath}\" and output to \"{outputPath}\".')
	

def processFiles(inputPath,outputPath,keywords,labLabels,totals):
	
	outputFile = open(outputPath,'w')
	barWidth = 36
	
	# Write the date and time of when the extractor was used
	dt = datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')
	outputFile.write(bar(barWidth,'='))
	outputFile.write(f'\nAT Report Extractor Output\nGenerated at {dt}\n')
	outputFile.write(bar(barWidth,'='))

	# Begin reading input directory
	for fileName in os.listdir(inputPath):
		filePath = inputPath + '\\' + fileName
		inputFile = open(filePath,'r')
		for line in inputFile:
			
			# Record the date of the Advisortrac report

			found = line.find('from')
			if (found != -1):
				dateStr = 'Dates ' + line[found:found+29]
				dateList = dateStr.split()
				totals.append(dateList)
				continue
				
			newList = list(filter(lambda str: str != '' and str != 'Students\n', line.split('\t')))

			# Criteria to clean up the list
			if (newList[0] == '<!--Group1 Dif-->'):
				newList.pop(0)
			if (newList[0].find('<!--# vis at beginning of report:') != -1):
				tempList = newList[0].split('>')
				newList[0] = tempList[1]
			if newList[0] not in keywords:
				continue
			if newList[0] in labLabels:
				del newList[1:]
			if (newList[0] == 'Grand Total:'):
				newList[1] = newList[1].replace('<b>','').replace('</b>','')
				newList[3] = newList[3].replace('<b>','').replace('</b>','')
			
			totals.append(newList)
		# end of for-loop
	
	# Write to file
	
	for element in totals:
		if element[0] in labLabels:
			outputFile.write('\n\n' + element[0] + ',Visits,Hours,Students')
		elif element[0] == 'Grand Total:':
			outputFile.write('\n\nGrand Total,Visits,Hours,Students')
			outputFile.write('\n' + ','.join(element))
		elif element[0] == 'Dates':
			outputFile.write('\n\n\n' + bar(barWidth))
			outputFile.write('\n' + ' '.join(element))
			outputFile.write('\n' + bar(barWidth))
		else:
			outputFile.write('\n' + ','.join(element))
	# end of for-loop
	
	print("Process completed successfully.")
	inputFile.close()
	outputFile.close()
	return True

# MAIN function
def main():
	labLabels = ['DC Math Lab','DC Writing Lab','SV Math Lab','SV Writing Lab']
	keywords = defineKeywords(labLabels)
	inputPath,outputPath = verifyArguments(len(sys.argv))

	printHeader()
	
	if not (verifyOverwrite(outputPath)):
		print("Cancelling process. Please run command again later.")
		printFooter()
		exit()

	
	processFiles(inputPath, outputPath, keywords, labLabels, [])
	
	printFooter()
	exit()

# Run the script's MAIN function
main()