#
# Author: Jesse McNicoll
# Title: PermutationSplitter.ps1
# Creation Date: 9/13/2017
#
# Description: 
#
# Parameters: 
#	InFilePath
#		A string variable to hold the name of an existing excel vendor file.
#	OutFileName
#		A string variable to hold the name of the output csv file.  If not specified, "OutputFile.csv"
#		will be used.  
#	PartNumCol
#		An integer to specify the column name to set as the part number column.
#		If not specified, this will be 1
#	DTISeparator
#		A character used by DTI to separate the appended size or color information from the original part
#		number.  If not specified, this will be a dash.
#   VendorAppender
#       The character used by the vendor list to append the sizing information to the 
#       vendor part number. 
#	RangeSeparator
#		The separator used in the existing vendor range.  Almost always assumed to be 
#		a dash.
#	StartingRow
#		An integer to specify the starting row of data in the excel file.  If not specified,
#		this will be set to 2. 
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
    [string]$OutFileName = "OutputFile.csv", #The name of the new excel file 
	[int]$PartNumCol = 2, 
	[string]$DTISeparator = '-',
    [string]$VendorAppender = '/',
    [string]$RangeSeparator = '-',
    [int]$StartingRow = 2,
    [int]$PriceCol = 4,
	[string]$FolderName = "$env:SystemDrive\Users\$env:UserName\My Documents",
	[int]$ActiveSheet = 1,
	[int]$last = 1
)

#Static Variables
#These variables should rarely need to be changed, and when necessary the code should be altered rather than makes these values
#parameters for the script.
$Company = 'DTI01'
$FLAG = 'Erroneous instance of part suffix.  Verify this part number for viability'
$ErrorBool = 'FALSE'
$INITIAL_INDEX = 0
$NEGATIVE_INDEX = -1
$TWO_MEMBERS = 2
$ONE_MEMBER = 1
$FirstMember = 0
$SecondMember = 1
$LastMember = -1
$ThirdMember = 2

#SizeConverter()
#
#   Takes an input size string and processes it, changing it into a formatted size
#   string to meet a DTI standard.  Specifically, this function replaces multiple X's
#   with a count (XXXL -> 3XL).  It also adds the implied 'L' to sizes that have no
#   letter size (2X -> 2XL).
#   
#   Input: $InputSize, a size string that has no defined standard.
#
#   Output: $OutputSize, a size string that meets DTI standards.
Function SizeConverter($InputSize){
    
    If($InputSize.EndsWith("X")){
            $InputSize = $InputSize + "L"
    }
	$Converted = $InputSize
    #If it contains XX or more, count the x's so they can be replaced with a number.
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

#RangeConverter()
#
#   This function takes an input start or end of a range and checks it for the type of range it is.
#   If it is a simple trunk size range, it will call size converter on the size.  If it is a dual trunk size
#   range, it will perform string operations and call the size converter method on the range member.  If it 
#   is a simple number size, such as for a shoe, it will simply return the input. 
#
#   Input: InputMember, a starting or ending member of a size range (string).
#
#   Output: OutputMember, the processed string that contains the formatted range member.
#
Function RangeConverter($InputMember){
    #Remove any whitespace from the input string to prevent errors with matching strings 
    $InputMember = $InputMember.Replace(" ","")
    #Convert the input size to upper case to allow easier string handling
    $InputMember = $InputMember.ToUpper()
	
    #Check for S, M, L or X in the input size.  If these are not contained, that means this part
    # is likely a shoe size based solely on numbers.
    If((($InputMember.Contains('S')) -or ($InputMember.Contains('M')) -or ($InputMember.Contains('L')) -or ($InputMember.Contains('X'))) -and (!($InputMember.Contains('/')))){  
        #Call the sizeConverter() method on this normal size.  
        $OutputMember = SizeConverter($InputMember)
    }
    ElseIf((($InputMember.Contains('S')) -or ($InputMember.Contains('M')) -or ($InputMember.Contains('L')) -or ($InputMember.Contains('X'))) -and ($InputMember.Contains('/'))){
        #Break the string in two so that each member of the size range can be handled separately.  
        $SplitSizeArray = $InputMember.Split('/')
        $FirstSplitSize = $SplitSizeArray[$FirstMember]
        $SecondSplitSize = $SplitSizeArray[$SecondMember]
        #Throw the beginning and ending member of the size range into the sizeconverter to retrieve them with Dooley Tackaberry naming conventions
        $FirstSplitSize = SizeConverter($FirstSplitSize)
        $SecondSplitSize = SizeConverter($SecondSplitSize)
        #Combine these two together to get the correct dual size.
        $OutputMember = $FirstSplitSize + "/" + $SecondSplitSize
    }
    Else{
        #At this point, the size is likely a number.  This can simply be returned.
        $OutputMember = $InputMember
    }
	$OutputMember
	return
}

#IsValidSize()
#
# This function takes an input string and checks if it is a 
# valid size of any type.
#
# Input: String to be checked for valid size
#
# Output: bool 
#
Function IsValidSize($SizeString){
    if($SizeString -match '^[0-9]?x*[ls]{1}$'){
        $returnVal = $TRUE
    }
    elseif($SizeString -match '^[mls]{1}$'){
        $returnVal = $TRUE
    }
    elseif($SizeString -match '^[0-9]?x+$'){
        $returnVal = $TRUE
    }
    else{
        $returnVal = $FALSE
    }
    $returnVal
    return
}
    

Function LoopThruPartnums($SizeArray, $SmallIndex, $BigIndex, $PartNum, $PartCost){
    #Loop from the small index to the big index, creating the new DTI part number every time
	For($SizeIndex = $SmallIndex; $SizeIndex -le $BigIndex; $SizeIndex++){
		#Concatenate the vendor part num and the isolated size. 
		$SubSize = $SizeArray[$SizeIndex]
		$DTIPartNUM = "$PartNum" + "$DTISeparator" +  "$SubSize"
		#After creating a new part number, add it to the output csv file
		Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost" 
		#Move on to the next part until last index is reached.
	}
}

#Create pre-defined arrays of sizes 
	#Use numbers for prefixes if multiple 'extra's involved in name.  If the vendor 
	#format does not match this later in the script, it will be converted. Lowercase versions 
	#will also be converted.
	$TrunkSizeArray = '4XS','3XS','2XS','XS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL','7XL','8XL'
    $NumSizeArray = '0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25'
    $DualSizeArray = '4XS/3XS','2XS/XS','S/M','L/XL','2XL/3XL','4XL/5XL','6XL/7XL'
    $StretchDualArray = '5XS/3XS','2XS/S','M/XL','2XL/5XL','6XL/8XL'
    #Create an array of sizeArrays to allow function handling of arrays and more adaptable code.  When a new array of sizes needs to be added, it can simply be 
    # hard-coded in and then added to the array of size arrays
    $MasterArray = $TrunkSizeArray, $NumSizeArray, $DualSizeArray, $StretchDualArray
    

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
$NewFilePath = "$FolderName\$OutFileName"

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

Add-Content -Path $NewFilePath 'New DTI Part Numbers'
#Loop through the part numbers row by row of the input excel file, allowing each part number to be checked. 
	For ($RowIndex = $StartingRow; $RowIndex -le $RowRange; $RowIndex++){
        
        [string]$PartNum = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PartNumCol).Value()
        #Obtain the part cost to group with the part number in case of size ranges
        $PartCost = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PriceCol).Value()

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
            Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost"                                 
        }
        elseif($SplitPartArray.Length -eq $ONE_MEMBER){
            #If no appended size, the vendor part num can equal the DTI part num.  
            $DTIPartNum = $PartNum
            Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost"    
        }
        elseif($SplitPartArray.Length -eq $TWO_MEMBERS){
            #Split the base part number on a range separator and check the last member, if it is a valid size, then the part number is a dual
            # size and should be appended as such. 
            $BasePartArray = $SplitPartArray[$FirstMember].Split($RangeSeparator)
            If(IsValidSize $BasePartArray[$LastMember]){
                $SecondSize = SizeConverter $SplitPartArray[$SecondMember]
                $FirstSize = SizeConverter $BasePartArray[$LastMember]
                $DualSizeAppend = $FirstSize + "/" + $SecondSize
                $DTIPartNum = $PartNum.Replace($BasePartArray[$LastMember] + $VendorAppender + $SplitPartArray[$SecondMember], "") + $DualSizeAppend

                Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost"
                #NOTE:
                #      This does not account for dual sizes that could then be followed by a dual size range.
                #
                #      This should be addressed.
                #
            }
            #Check if the array has a range separator in the second member.  If it does, it will need to be verified as a range of sizes.  
            ElseIf($SplitPartArray[$SecondMember].Contains($RangeSeparator)){
                #At this point, the range may be valid, as in spanning between two sizes. Check if it is to see what further processing
                # might be required.
                $SizeRangeArray = $SplitPartArray[$SecondMember].Split($RangeSeparator)
                $BasePartNum = $SplitPartArray[$FirstMember]
                $BoolCheck = IsValidSize $SizeRangeArray[$SecondMember]
                If(IsValidSize($SizeRangeArray[$FirstMember]) -and $BoolCheck){
                    If(IsValidSize($SizeRangeArray[$SecondMember])){
                        Echo "Weird things are afoot at the circle k"
                    }
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
                        Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $PartCost, $ErrorMessage"
                    }
                    Else{
                    #Loop from the first index to the ending index, creating the new DTI part number every time.
                    #Do this by calling the LoopThruPartnums() method with the appropriate array indexed in
                    #the Master Array
                        LoopThruPartnums $MasterArray[$MasterArrayIndex] $FirstIndex $LastIndex $BasePartNum $PartCost
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
                    Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $PartCost"                  
                }
            }
            Else{
                $DTISize = SizeConverter $SplitPartArray[$SecondMember]
                $DTIPartNum = $SplitPartArray[$FirstMember] + $DTISeparator + $DTISize
                Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost"
            }
        }        
        else{
            #The array is empty and an error has occurred.
            Write-Host "An error has occurred in processing row $RowIndex of the input file.  There is no part number."
        }
        
    }
Write-Host "Script Complete.  Please check the file for consistencies."
       


















<#
        if($PartNum.EndsWith("XXXXL")){
            $DTIPartNum = $PartNum.TrimEnd("XXXXL")
            $DTIPartNum = $DTIPartNum + "-4XL"
        }
        elseif($PartNum.EndsWith("XXXL")){
            $DTIPartNum = $PartNum.TrimEnd("XXXL")
            $DTIPartNum = $DTIPartNum + "-3XL"
        }
        elseif($PartNum.EndsWith("XXL")){
            $DTIPartNum = $PartNum.TrimEnd("XXL")
            $DTIPartNum = $DTIPartNum + "-2XL"
        }
        elseif($PartNum.EndsWith("XL")){
            $DTIPartNum = $PartNum.TrimEnd("XL")
            $DTIPartNum = $DTIPartNum + "-XL"
        }
        elseif($PartNum.EndsWith("L")){
            $DTIPartNum = $PartNum.TrimEnd("L")
            $DTIPartNum = $DTIPartNum + "-L"
        }
        elseif($PartNum.EndsWith("M")){
            $DTIPartNum = $PartNum.TrimEnd("M")
            $DTIPartNum = $DTIPartNum + "-M"
        }
        elseif($PartNum.EndsWith("XXS")){
            $DTIPartNum = $PartNum.TrimEnd("XXS")
            $DTIPartNum = $DTIPartNum + "-2XS"
        }
        elseif($PartNum.EndsWith("XS")){
            $DTIPartNum = $PartNum.TrimEnd("XS")
            $DTIPartNum = $DTIPartNum + "-XS"
        }
        elseif($PartNum.EndsWith("S")){
            $DTIPartNum = $PartNum.TrimEnd("S")
            $DTIPartNum = $DTIPartNum + "-S"
        }
        else{
            $DTIPartNum = $PartNum
        }
        Add-Content -Path $NewFilePath "$DTIPartNum"
    }

Write-Host "Script complete!"
#>