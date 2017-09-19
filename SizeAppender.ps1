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
Add-Content -Path $PriceFilePath 'New DTI Part Numbers'
#Loop through the part numbers row by row of the input excel file, allowing each part number to be checked. 
	For ($RowIndex = $StartingRow; $RowIndex -le $RowRange; $RowIndex++){
        
        [string]$PartNum = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PartNumCol).Value()
        #Obtain the part cost to group with the part number in case of size ranges
        $PartCost = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PriceCol).Value()
        #Obtain the PUM
        $PUM = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PUMCol).Value()
        $PUM = $PUM.ToUpper()

        #Trim the whitespace!
        $PartNum = $PartNum.Trim()

        #Split the part number on the vendor appender.  This will allow the size to be accessed and stored.
        $SplitPartArray = $PartNum.Split($VendorAppender)

        #Determine if there were multiple vendor appending characters in the part number, or none.  If there are more than one, additional processing
        #will be required.  If there were none, then the vendor part number can equal the DTI part number.
        If($SplitPartArray.Length -gt $TWO_MEMBERS){
            #
            #
            #NOTE: This is an immediate patch for Protecti
            #            
            $DTISize = SizeConverter $SplitPartArray[$SecondMember]
            $DTIPartNum = $SplitPartArray[$FirstMember] + $DTISeparator + $SplitPartArray[$ThirdMember] + $DTISeparator + $DTISize
            Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNUM, $PartCost, $PUM"                                 
        }
        elseif($SplitPartArray.Length -eq $ONE_MEMBER){
            #If no appended size, the vendor part num can equal the DTI part num.  
            $DTIPartNum = $PartNum
            Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNUM, $PartCost, $PUM"    
        }
        elseif($SplitPartArray.Length -eq $TWO_MEMBERS){
            #Split the base part number on a range separator and check the last member, if it is a valid size, then the part number is a dual
            # size and should be appended as such. 
            $BasePartArray = $SplitPartArray[$FirstMember].Split($VenRangeSeparator)
            If(IsValidSize $BasePartArray[$LastMember]){
                $SecondSize = SizeConverter $SplitPartArray[$SecondMember]
                $FirstSize = SizeConverter $BasePartArray[$LastMember]
                $DualSizeAppend = $FirstSize + "/" + $SecondSize
                $DTIPartNum = $PartNum.Replace($BasePartArray[$LastMember] + $VendorAppender + $SplitPartArray[$SecondMember], "") + $DualSizeAppend

                Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNUM, $PartCost, $PUM"
                #NOTE:
                #      This does not account for dual sizes that could then be followed by a dual size range.
                #
                #      This should be addressed.
                #
            }
            #Check if the array has a range separator in the second member.  If it does, it will need to be verified as a range of sizes.  
            ElseIf($SplitPartArray[$SecondMember].Contains($VenRangeSeparator)){
                #At this point, the range may be valid, as in spanning between two sizes. Check if it is to see what further processing
                # might be required.
                $SizeRangeArray = $SplitPartArray[$SecondMember].Split($VenRangeSeparator)
                $BasePartNum = $SplitPartArray[$FirstMember]
                
                $BoolCheck1 = IsValidSize $SizeRangeArray[$FirstMember]
                $BoolCheck2 = IsValidSize $SizeRangeArray[$SecondMember]
                #
                #Note: This check does not take into account the need for number size checking.
                #
                # This is an important aspect of sizing ranges.  
                # 
                If($BoolCheck1 -and $BoolCheck2){
                   
                    #Now that it is verified as a valid range, the LoopThruPartNums functions must be called after the indices and proper array has been 
                    #discovered. 
                    $FirstSize = $SizeRangeArray[$FirstMember]
                    $FirstSize = SizeConverter $FirstSize
                    $LastSize = $SizeRangeArray[$LastMember]
                    $LastSize = SizeConverter $LastSize
                    $MasterArrayIndex = $NEGATIVE_INDEX
                    Do{
                        $MasterArrayIndex = $MasterArrayIndex + 1
                        if($MasterArrayIndex -lt $MasterArray.Length){
                            $FirstIndex = [array]::IndexOf($MasterArray[$MasterArrayIndex],$FirstSize)
                            $LastIndex = [array]::IndexOf($MasterArray[$MasterArrayIndex],$LastSize)
                        }
                    }
                    While($FirstIndex -eq $NEGATIVE_INDEX -and $LastIndex -eq $NEGATIVE_INDEX -and $MasterArrayIndex -lt $MasterArray.Length)
                    
                    if($MasterArrayIndex -ge $MasterArray.Length){
                
                        $DTIPartNum = 'Invalid'
                        $PartCost = 'Invalid'
                        $ErrorMessage = 'The appropriate sizing array was not found for this part.'
                        $ErrorBool = 'TRUE'
                        Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNum, $PartCost, $ErrorMessage, $PUM"
                    }
                    Else{
                    #Loop from the first index to the ending index, creating the new DTI part number every time.
                    #Do this by calling the LoopThruPartnums() method with the appropriate array indexed in
                    #the Master Array
                        LoopThruPartnums $MasterArray[$MasterArrayIndex] $FirstIndex $LastIndex $BasePartNum $PartCost $PartNum $PUM
                    }
                }
                Else{
                    #The last member of the part number is not a size.  Rearrange the part number to achieve the DTI part number
                    #
                    #NOTE: Specifically geared towards handling protecti.
                    #
                    $SizeAppend = $SizeRangeArray[$FirstMember]
                    $InsertionString = $SizeRangeArray[$SecondMember]
                    $DTIPartNum = $BasePartNum + $DTISeparator + $InsertionString + $DTISeparator + $SizeAppend 
                    Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNum, $PartCost, $PUM"                  
                }
            }
            Else{
                $DTISize = SizeConverter $SplitPartArray[$SecondMember]
                $DTIPartNum = $SplitPartArray[$FirstMember] + $DTISeparator + $DTISize
                Add-Content -Path $PriceFilePath -Value "$PartNum, $DTIPartNUM, $PartCost, $PUM"
            }
        }        
        else{
            #The array is empty and an error has occurred.
            Write-Host "An error has occurred in processing row $RowIndex of the input file.  There is no part number."
        }
        
    }
Write-Host "Script Complete.  Please check the file for inconsistencies."
       

