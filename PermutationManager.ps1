#
# Author: Jesse McNicoll
# Title: PermutationManager.ps1
# Creation Date: 8/10/2017
#
# Description: This script opens an input csv file and takes existing part
#  		numbers and splits them into a range of new part numbers 
#		for each size of boot corresponding a boot model.  It also adds the
#		appropriate image of the boot model to the csv file on the same line. 		   			
#
# Parameters: 
#	FileName
#		A string variable to hold the name of the new csv file.
#	InFilePath
#		A string variable to hold the name of the path and file name for the input file  
#	FolderName
#		An optional parameter that is used to make the folder to store the csv files. 
#		If not specified, "My Documents" of the current user will be used. 
#