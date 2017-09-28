#
# Author: Jesse McNicoll
# Title: PermutationSplitter.ps1
# Creation Date: 9/13/2017
#
# Description: 
#
#   This script simply checks an input file of part numbers for sizes using regular expressions.  If a valid size is found 
#   in the part number (and is placed in such a way as to indicate a true size), it is spliced from the part number and then 
#   concatenated to the end of the part number with a dash, so as to maintain DTI conventions.  
#
#   Later, this script can be added to so to include an ability to detect size ranges, allowing ranges to be filled out with prices
#   and DTI to have unique part numbers for each vendor part.  
#
#   The input file to this script should be in the form of a price list template file, with the minimum filled columns being
#   BaseUnitPrice, PUM, and VenPartNum.
#
#   The script will create two files, one to be used for loading into supplier price, and the other for loading into supplier part.
#
# Parameters: 
#	InFilePath
#		A string variable to hold the name of an existing excel vendor file.
#   VendorID
#       A string to hold the vendorID for the input price list.
#	PartNumCol
#		An integer to specify the column name to set as the part number column.
#		If not specified, this will be 6.
#	StartingRow
#		An integer to specify the starting row of data in the excel file.  If not specified,
#		this will be set to 2. 
#   PUMCol
#       An integer to specify the column for the PUM in the input file.
#   PriceCol
#       The column from which to the vendor cost from.
#	FolderName
#		An optional parameter that is used to make the folder to store the csv file. 
#		If not specified, "My Documents" of the current user will be used.
#	ActiveSheet
#		An optional parameter to get the active sheet. If not used, it is assumed
#		to be 1.



	
#Set up the parameters for the script to allow proper parsing of inputs.
Param(
	[String]$InFilePath, 
    [string]$VendorID,
 	[int]$PartNumCol = 6, 
    [int]$StartingRow = 2,
    [int]$PUMCol = 4,
    [int]$PriceCol = 3,
	[string]$FolderName = "$env:SystemDrive\Users\$env:UserName\My Documents",
	[int]$ActiveSheet = 1,
	[int]$last = 1
)

#Include the functions file for handling size strings, including common variables.
. "c:\users\jessem\CodeProjects\Scripts\Scripts\Functions.ps1"

#Create file names for the price output file and part number mapping file.  
$PriceFileName = $VendorID + "PriceFile.csv"
$MapFileName = $VendorID + "PartMap.csv"

#Error-checking inputs to ensure proper execution
	#Check existence of InFilePath
		If(!(Test-Path($InFilePath))){
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
		$WorkSheet = $workbook.Sheets.Item($ActiveSheet)
		$Range = $WorkSheet.UsedRange.Columns.Count
        $RowRange = $WorkSheet.UsedRange.Rows.Count
    
	#Check that PartNumCol is within the used range of the excel file sheet
		If(($PartNumCol -gt $Range)){
			Write-Host "The input part number column number is outside the used range.  Script terminating."
			Exit
        }
	#Check that PermuteCol is within the used range of the excel file sheet
		If(($PermuteCol -gt $Range)){
			Write-Host "The input permute column number is outside the used range.  Script terminating."
			Exit
        }

#Create the path to the output file to allow writing to it later on in script.
$PriceFilePath = "$FolderName\$PriceFileName"
$MapFilePath = "$FolderName\$MapFileName"

$PriceFilePath = ValidatePath $PriceFilePath
$MapFilePath = ValidatePath $MapFilePath

#Create the headers for each output file, allowing data to be added 
Add-Content -Path $PriceFilePath 'Company, PartNum, BaseUnitPrice, PUM, EffectiveDate, VenPartNum, ConvFactor, ExpirationDate, DiscountPercent, VendorID'
Add-Content -Path $MapFilePath 'Company, VendorID, PartNum, VendPartNum'

#Loop through the part numbers row by row of the input excel file, allowing each part number to be checked. 
	For ($RowIndex = $StartingRow; $RowIndex -le $RowRange; $RowIndex++){
        
        [string]$PartNum = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PartNumCol).Value()
        #Obtain the part cost to group with the part number in case of size ranges
        $PartCost = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PriceCol).Value()
        #Obtain the PUM
        $PUM = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PUMCol).Value()
        $PUM = $PUM.ToUpper()

        #Trim the whitespace to prevent it from being pesky
        $PartNum = $PartNum.Trim()

        #Use the regex checking function to check for sizes or colors.  
        $NewPartNum = FindSizeAndSplit $PartNum
            #If there is a size, use a custom string function to splice
            #the size or color, convert it to a DTI standard, and insert it on the end of the part number.
        Add-Content -Path $PriceFilePath "$Company, $NewPartNum, $PartCost, $PUM, , $PartNum, , , , $VendorID "
                 
            
            
       
        
    }
Write-Host "Script Complete.  Please check the file for inconsistencies."
       

