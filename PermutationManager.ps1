#
# Author: Jesse McNicoll
# Title: PermutationManager.ps1
# Creation Date: 9/6/2017
#
# Description: This script opens an excel file that's path has been input.  It will take specified columns
#				from the excel file and create their contents (a range of sizes or colors) and create 
#				permutations of part numbers for the Epicor10 system.  It will then create a csv file of the 
#				permuted part numbers and the associated vendor part number.    		   			
#
# Parameters: 
#	InFilePath
#		A string variable to hold the name of an existing excel vendor file.
#	PartNumCol
#		An integer to specify the column number to set as the part number column
#	PermuteCol
#		An integer to specify the column number to set as the column containing data to 
#		be permuted.
#	Separator
#		A character used to separate the appended size or color information from the original part
#		number.  If not specified, this will be a dash.
#	OutFileName
#		A string variable to hold the name of the output csv file.  If not specified, "OutputFile.csv"
#		will be used.  
#	FolderName
#		An optional parameter that is used to make the folder to store the csv file. 
#		If not specified, "My Documents" of the current user will be used.
#	ActiveSheet
#		An optional parameter to get the active sheet. If not used, it is assumed
#		to be 1.

#Set up the parameters for the script to allow proper parsing of inputs.
Param(
	[String]$InFilePath, 
	[int]$PartNumCol, 
	[int]$PermuteCol,
	[char]$Separator = '-'
	[String]$OutFileName = "OutputFile.csv", #The name of the new excel file
	[String]$FolderName = "$env:SystemDrive\Users\$env:UserName\My Documents",
	[int]$ActiveSheet = 1
)

#Error-checking inputs to ensure proper execution
	#Check existence of InFilePath
		If !(Test-Path($InFilePath)){
				Write-Host "Input File Not Found.  Script is terminating."
				Exit
		}
	#Get the file name from the infile position
		$InFileName = Split-Path -Path $InFilePath -Leaf -Resolv
	#Open the excel file to check its contents
		$excel = New-Object -com "Excel.Application"
		$excel.Visible = $True
		$workbook = $excel.Workbooks.open($InFilePath)
		#Get the active sheet.
		$WorkSheet = $workbook.Sheets(ActiveSheet)
		
	#Check that PartNumCol is within the used range of the excel file sheet
		If !(PartNumCol -lt $WorkSheet.UsedRange.Columns.Count){
			Write-Host "The input part number column number is outside the used range.  Script terminating."
			Exit
	#Check that PermuteCol is within the used range of the excel file sheet
		If !(PermuteCol -lt $WorkSheet.UsedRange.Columns.Count){
			Write-Host "The input permute column number is outside the used range.  Script terminating."
			Exit
#Create pre-defined arrays of sizes 
	#Use numbers for prefixes if multiple 'extra's involved in name.  If the vendor 
	#format does not match this later in the script, it will be converted. Lowercase versions 
	#will also be converted.
	$SizeArray = '4XS','3XS','2XS','XS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL','7XL','8XL'
#Perform operations on it to get the permutated part numbers.

	#Open the file 
	
	#Loop through the part numbers row by row of the input excel file
		
		#Store the part number to a variable 
		
		#Get the permutation range from the same row.
		
		#Check the range against the pre-defined arrays to get the starting and ending indices
		# for the pre-defined arrays.
		
		#Loop from the first index to the ending index, creating the new DTI part number every time
		
			#Concatenate the vendor part num and the isolated size. 
			
			#After creating a new part number, add it to the output csv file
			
			#Move on to the next part until last index is reached.
			
		#Now that the output file has all permutated versions of the current part number, it 
		# is time to move on to the next part number in the input file
	
	#Looping through the vendor part numbers is now complete.  The input and output file can be closed.

#Save the csv file and end the script with a printed statement that declares completion.  	