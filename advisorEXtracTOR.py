'''
	advisorEXtracTOR
	AdvisorTrac Report Extractor (version 3)

	Author: Jake Sherwood (MothsInTheMachine on GitHub)
	Last Modified: Dec. 17, 2024
'''

import sys
import os, os.path
from datetime import datetime


# Define MAIN
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
# End MAIN


# FUNCTION DEFINITIONS

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
	# print('Syntax:\n\n\t> python at-report-extract.py [optional to input directory] [optional output file name]\n')
	printFooter()
	quit()


def bar(width,_symbol='-'):
	str = _symbol
	for i in range(width-1):
		str += _symbol
	return str


def defineKeywords(_labLabels=[]):
	keywords = []
	if (len(_labLabels) > 0):
		keywords.extend(_labLabels)
	keywords.extend(['DC Math Lab Total:','DC Writing Lab Total:','SV Math Lab Total:','SV Writing Lab Total:','Grand Total:'])
	keywords.extend(['Chemistry Total:','CIS/Programming Total:','Math Total:','Other Total:','Other Course Tutoring Total:','Physics Total:'])
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
		
		choice = ''
		while (choice != 'n') and (choice != 'y'):
			print(f'The output file \"{outputPath}\" already exists.')
			inputStr = input('Do you want to overwrite it? [y/n]: ')
			choice = inputStr[0].lower()
			# Gets to this line if invalid choice is made
			if (choice != 'n' and choice != 'y'):
				print("\nInvalid choice. Please enter Y for 'Yes' and N for 'No'...")
		
		if (choice == 'n'):
			return False
		
		if (choice == 'y'):
			print(f'\nConfirmed overwriting previous file \"{outputPath}\".')
			return True
		
	else:
		# Path does not exists
		return True
	

def processFiles(inputPath,outputPath,keywords,labLabels,totals):
	
	outputFile = open(outputPath,'w')
	barWidth = 36
	sep = ',' # for CSV: Comma Separated Values
	
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
				continue # Nothing else is worthwhile on this line
			
			'''
			These are very specific to our AdvisorTrac Reports and is not 
			designed to work with all Reports!
			'''

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
			outputFile.write('\n\n' + sep.join([element[0],'Visits','Hours','Students']))
		elif element[0] == 'Grand Total:':
			outputFile.write('\n\n' + sep.join(['Grand Total','Visits','Hours','Students']))
			outputFile.write('\n' + sep.join(element))
		elif element[0] == 'Dates':
			outputFile.write('\n\n\n' + bar(barWidth))
			outputFile.write('\n' + ' '.join(element))
			outputFile.write('\n' + bar(barWidth))
		else:
			outputFile.write('\n' + sep.join(element))
	# end of for-loop
	
	print(f'\nProcess completed successfully. Output file is at \"{outputPath}\".')
	inputFile.close()
	outputFile.close()
	return True


# Call MAIN
main()
