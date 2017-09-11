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
#   PriceCol
#        The column that prices should be grabbed from.
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


	
#Set up the parameters for the script to allow proper parsing of inputs.
Param(
	[String]$InFilePath,  
	[int]$PartNumCol = 1, 
	[int]$PermuteCol = 2,
    [double]$PriceCol = 3,
	[int]$StartingRow = 1,
	[string]$Separator = '-',
	[string]$OutFileName = "OutputFile.csv", #The name of the new excel file
	[string]$FolderName = "$env:SystemDrive\Users\$env:UserName\My Documents",
	[int]$ActiveSheet = 1,
	[string]$RangeSeparator = '-',
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


#Define a function for converting part numbers with irregular sizing schemes
#	 to the Dooley Tackaberry convention
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
        #Break the string in two so that each split size can be handled separately.  
        $SplitSizeArray = $InputMember.Split('/')
        $FirstSplitSize = $SplitSizeArray[0]
        $SecondSplitSize = $SplitSizeArray[1]
        #Throw each split size into the sizeconverter to retrieve them with Dooley Tackaberry naming conventions
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

Function LoopThruPartnums($SizeArray, $SmallIndex, $BigIndex){
    #Loop from the small index to the big index, creating the new DTI part number every time
	For($SizeIndex = $SmallIndex; $SizeIndex -le $BigIndex; $SizeIndex++){
		#Concatenate the vendor part num and the isolated size. 
		$SubSize = $SizeArray[$SizeIndex]
		$DTIPartNUM = "$PartNum" + "$Separator" +  "$SubSize"
		#After creating a new part number, add it to the output csv file
		Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $Price" 
		#Move on to the next part until last index is reached.
	}
}

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
Add-Content -Path $NewFilePath 'Company, PartNum, BaseUnitPrice, PUM, EffectiveDate, VenPartNum, ConvFactor, ExpirationDate, DiscountPercent, VendorID'
	
#Perform operations on the open excel file to get the permuted part numbers.
	
	#Loop through the part numbers row by row of the input excel file, allowing each part number to be checked.  
	For ($RowIndex = $StartingRow; $RowIndex -le $RowRange; $RowIndex++){	
		#Store the part number to a variable 
		$PartNum = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PartNumCol).Value()
		#Get the permutation range from the same row so parsing can begin.
		$SizeString = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PermuteCol).Value()
        
        #Get the prices so it can be posted to csv file in same line as new part number
        $Price = $WorkSheet.UsedRange.Cells.Item($RowIndex, $PriceCol).Value()
        #Perform a variety of checks on the SizeString to see how the data should be handled.  This allows a decision to be made on the need for a suffix
        # to the part number.
        If($null -eq $SizeString){
            $DTIPartNum = $PartNum
            Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $Price"
        }

        #If the sizeString does not contain a range separator, it likely does not have a range of sizes and can simply be
        # made a suffix to the part.
        ElseIf(!($SizeString.Contains($RangeSeparator))){
            #Call the toUpper() string function to allow more uniform handling of the size string.
            #If there is no separator BUT contains S, M, L, or X in the size string, it is likely a single size and can be added directly to the part.
            If(($SizeString.Contains('S')) -xor ($SizeString.Contains('M')) -xor ($SizeString.Contains('L')) -xor ($SizeString.Contains('X'))){
                #Invoke the SizeConverter() method on the size string to ensure consistent style
                $SizeString = SizeConverter($SizeString)
                #Simply append the SizeString to the partnum, creating a DTI Part number that adheres to a consistent style.
                $DTIPartNum = $PartNum + "$Separator" + "$SizeString"
                #Now add that part number to the outputfile, completing this part number and allowing succession to the next.
                Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $Price"
            }
            Else{
                #If no sizes, whatever is in the size string should be appended to the part number, but due to variability
                #of whatever this string might be, a flag is added to a new fourth column.  This flag is specified in the static 
                #variables at the top of this code.  This allows easy visibility to a user and human-based editing.
                $DTIPartNum = $PartNum + "$Separator" + "$SizeString"
                Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $Price, $FLAG"
            }
        }
        Else{
            #The string contains a range separator.  This allows parsing of the range and filling out to multiple DTI part numbers.
            $PermutationArray = $SizeString.Split($RangeSeparator)
            #A 2-member array of sizes has been created.  These two sizes can be sent to the sizeconverter() method, allowing for
            # the type of sizing to be recognized and the size to be converted to DTI naming conventions.
            $FirstSize = RangeConverter($PermutationArray[$INITIAL_INDEX])
		    $LastSize = RangeConverter($PermutationArray[$NEGATIVE_INDEX])
            
            #Loop through the master array until matches are made with $FirstSize and $LastSize to mark
            # a starting position in the sizing arrays to start concatenating with the part number to create
            # DTI part numbers 
            
            #Start MasterArrayIndex at -1 to allow a post-loop check on $FirstIndex and $LastIndex
            $MasterArrayIndex = $NEGATIVE_INDEX
            Do{
                $MasterArrayIndex = $MasterArrayIndex + 1
                if($MasterArrayIndex -lt $MasterArray.Length){
                    $FirstIndex = [array]::IndexOf($MasterArray[$MasterArrayIndex],$FirstSize)
                    $LastIndex = [array]::IndexOf($MasterArray[$MasterArrayIndex],$LastSize)
                }
            }
            While($FirstIndex -eq $NEGATIVE_INDEX -and $LastIndex -eq $NEGATIVE_INDEX -and $MasterArrayIndex -lt $MasterArray.Length)

            #Check to make sure the appropriate sizing array was found.  If not, output a sizing error on that line to notify a user. 
            #Set a boolean to determine if the user should be notified to look for errors in the output file.  
            if($MasterArrayIndex -ge $MasterArray.Length){
                
                $DTIPartNum = 'Invalid'
                $Price = 'Invalid'
                $ErrorMessage = 'The appropriate sizing array was not found for this part.'
                $ErrorBool = 'TRUE'
                Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNum, $Price, $ErrorMessage"

            }
            Else{
            #Loop from the first index to the ending index, creating the new DTI part number every time.
            #Do this by calling the LoopThruPartnums() method with the appropriate array indexed in
            #the Master Array
                LoopThruPartnums $MasterArray[$MasterArrayIndex] $FirstIndex $LastIndex
            }
        }
		
		#Now that the output file has all permutated versions of the current part number, it 
		# is time to move on to the next part number in the input file
	}
	#Looping through the vendor part numbers is now complete.  The input and output file can be closed.

#Save the csv file and end the script with a printed statement that declares completion.
Write-Host "Script Completed.  Please view the output file in your documents folder or the input destination folder"
If($ErrorBool -eq "TRUE"){
    Write-Host "Some parts did not have appropriate sizing arrays.  Please look through the output file to fix.  You can search for 'not found' to locate these lines"   
}
