#
# Author: Jesse McNicoll
# Title: PermutationManager.ps1
# Creation Date: 9/6/2017
#
# Description: This script opens an excel file that's path has been input.  This excel file will have been pre-formatted
#				to closely match the current DTI format for price lists with an excel table.  The script will take specified columns
#				from the excel file and their contents (a range of sizes or colors) and create 
#				permutations of part numbers for the Epicor10 system.  It will then create a csv file of the 
#				permuted part numbers and the associated vendor part number.    		   			
#
# Parameters: 
#	InFilePath
#		A string variable to hold the name of an existing excel vendor file.
#	PartNumCol
#		An integer to specify the column name to set as the part number column.
#		If not specified, this will be 'PartNum'
#	PermuteCol
#		An integer to specify the column name to set as the column containing data to 
#		be permuted. If not specified, this will be 'VendPartNum'
#	StartingRow
#		An integer to specify the starting row of data in the excel file.  If not specified,
#		this will be set to 2. 
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
#	RangeSeparator
#		The separator used in the existing vendor range.  Almost always assumed to be 
#		a dash.


#Define a function for converting part numbers with irregular sizing schemes
#	 to the Dooley Tackaberry convention
Function SizeConverter($InputSize){
	$InputSize = $InputSize.ToUpper()
	#If it contains XX or more, count the x's so they can be replaced with a number.
	$Converted = $InputSize
	if($InputSize.Contains("XX")){
		$StringCheck = $InputSize
		if ($StringCheck.EndsWith("L")){
			$TrimString = $StringCheck.Replace("L","")
		}
		if ($StringCheck.EndsWith("M")){
			$TrimString = $StringCheck.Replace("M","")
		}
		if ($StringCheck.EndsWith("S")){
			$TrimString = $StringCheck.Replace("S","")
		}
		$XCount = $TrimString.Length
		$Array = $InputSize.ToCharArray()
		$FinalLetter = $Array[-1]
		$Converted =  "$XCount" + "X" + "$FinalLetter"
	}
	$Converted
	return
}
		
#Set up the parameters for the script to allow proper parsing of inputs.
Param(
	[String]$InFilePath,  
	[Int]$PartNumCol = 1, 
	[string]$PermuteCol = 2,
	[int]$StartingRow = 2,
	[string]$Separator = '-',
	[string]$OutFileName = "OutputFile.csv", #The name of the new excel file
	[string]$FolderName = "$env:SystemDrive\Users\$env:UserName\My Documents",
	[int]$ActiveSheet = 1,
	[string]$RangeSeparator = '-',
    [int]$last = 1
)

#Error-checking inputs to ensure proper execution
	#Check existence of InFilePath
		If!(Test-Path($InFilePath)){
				Write-Host "Input File Not Found.  Script is terminating."
				Exit
		}
	#Get the file name from the infile position
		$InFileName = Split-Path -Path $InFilePath -Leaf -Resolve
	#Check write-ability on FolderName
		
	#Open the excel file to check its contents
		$excel = New-Object -com "Excel.Application"
		$excel.Visible = $True
		$workbook = $excel.Workbooks.open($InFilePath)
		#Get the active sheet.
		$WorkSheet = $workbook.Sheets($ActiveSheet)
		
	#Check that PartNumCol is within the used range of the excel file sheet
		If(!(PartNumCol -lt $WorkSheet.UsedRange.Columns.Count)){
			Write-Host "The input part number column number is outside the used range.  Script terminating."
			Exit
        }
	#Check that PermuteCol is within the used range of the excel file sheet
		If(!(PermuteCol -lt $WorkSheet.UsedRange.Columns.Count)){
			Write-Host "The input permute column number is outside the used range.  Script terminating."
			Exit
        }
#Create pre-defined arrays of sizes 
	#Use numbers for prefixes if multiple 'extra's involved in name.  If the vendor 
	#format does not match this later in the script, it will be converted. Lowercase versions 
	#will also be converted.
	$SizeArray = '4XS','3XS','2XS','XS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL','7XL','8XL'

#Create the path to the output file to allow writing to it later on in script.
$NewFilePath = "$FolderName\$OutFileName"

#If NewFilePath is used, create a new file path or delete the existing file to ensure secure files. 
While(Test-Path "$NewFilePath"){
    $ScreenInput = Read-Host "A file already exists with that name.  If you want to delete the file, type D.  If you want to create a new filename, type it now."
    $StringCheck = $ScreenInput.ToUpper()
    If($StringCheck -eq "D"){
        Remove-Item $NewFilePath
    }
    else{
    $NewFilePath = "$FolderName\$ScreenInput"
    }
}
#Create headers in the file to allow easy reading and understanding of the created output file
Add-Content -Path $NewFilePath 'PartNum, DTIPartNum'
	
#Perform operations on the open excel file to get the permuted part numbers.
	
	#Get name of the table in the worksheet to allow easy column referencing
	$Table = $WorkSheet.ListObjects($ActiveSheet).Name
	
	#Loop through the part numbers row by row of the input excel file
	For ($RowIndex = $StartingRow; $RowIndex -lt $WorkSheet.UsedRange.Rows.Count; $RowIndex++){	
		#Store the part number to a variable 
		$PartNum = $WorkSheet.UsedRange.Cells($RowIndex, $PartNumCol).Value
		#Get the permutation range from the same row so parsing can begin.
		$PermutationRange = $WorkSheet.UsedRange.Cells($RowIndex, $PermuteCol).Value
		#Parse the range to get the starting and ending values, allowing comparison with sizeArray
		$PermutationArray = $SizeString.Split($RangeSeparator)
        
		$FirstSize = SizeConverter($PermutationArray[0])
		$LastSize = SizeConverter($PermutationArray[-1])
		
		#Check the range against the pre-defined arrays to get the starting and ending indices
		# for the pre-defined arrays.
		$FirstIndex = [array]::indexof($SizeArray,$FirstSize)
		$LastIndex = [array]::indexof($SizeArray,$LastSize)
			
		
		#Loop from the first index to the ending index, creating the new DTI part number every time
		For($SizeIndex = $FirstIndex; $SizeIndex -le $LastIndex; $SizeIndex++){
			#Concatenate the vendor part num and the isolated size. 
			$SubSize = $SizeArray[$SizeIndex]
			$DTIPartNUM = "$PartNum" + "$Separator" +  "$SubSize"
			#After creating a new part number, add it to the output csv file
			Add-Content -Path $NewFilePath - Value "$PartNum, $DTIPartNUM" 
			#Move on to the next part until last index is reached.
		}	
		#Now that the output file has all permutated versions of the current part number, it 
		# is time to move on to the next part number in the input file
	}
	#Looping through the vendor part numbers is now complete.  The input and output file can be closed.

#Save the csv file and end the script with a printed statement that declares completion.
Write-Host "Script Completed.  Please view the output file in your documents folder or the input destination folder"  	