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

#Static Variables
#These variables should rarely need to be changed, and when necessary the code should be altered rather than makes these values
#parameters for a script.
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
$DTISeparator = '-'
$VenRangeSeparator = '-'


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
    #Use regular expressions to check for spelled out sizes
    #and convert them to one-letter equivalents
    $InputSize = $InputSize -replace "me?d", "M"
    $InputSize = $InputSize -replace "lr?g", "L"
    $InputSize = $InputSize -replace "sml?", "S"
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
    
#LoopThruPartNums()
#
#  This function takes an input size array and concatenates the sizes in the array to base 
#  base part num based on the input beginning and ending indices.  It then adds these sizes to 
#  an input file.
#
#  
#
Function LoopThruPartnums($SizeArray, $SmallIndex, $BigIndex, $BasePartNum, $PartCost, $PartNum, $PUM, $NewFilePath){
    #Loop from the small index to the big index, creating the new DTI part number every time
	For($SizeIndex = $SmallIndex; $SizeIndex -le $BigIndex; $SizeIndex++){
		#Concatenate the vendor part num and the isolated size. 
		$SubSize = $SizeArray[$SizeIndex]
		$DTIPartNUM = "$BasePartNum" + "$DTISeparator" +  "$SubSize"
		#After creating a new part number, add it to the output csv file
		Add-Content -Path $NewFilePath -Value "$PartNum, $DTIPartNUM, $PartCost, $PUM" 
		#Move on to the next part until last index is reached.
	}
}

#ValidatePath()
#
# This function tests the validity of a file output path.  If it already exists, it offers the ability to delete the file and write over it,
# or create a new file with a different name. 
#
# Input: FilePath, a string variable holding the full path to a file.  
#        FolderName, a string variable holding a path to the PriceFile folder.  
# Output: FilePath, the final path to the file.
#
Function ValidatePath($FilePath, $FolderName){
    While(Test-Path "$FilePath"){
        $ScreenInput = Read-Host "A file already exists with the name $FilePath.  If you want to delete the file, type D.  If you want to create a new filename, type it now."
        $StringCheck = $ScreenInput.ToUpper()
        If($StringCheck -eq "D"){
            Remove-Item $FilePath
        }
        else{
            $FilePath = "$FolderName\$ScreenInput"
        }
    }
    $FilePath
    return
}

#ContainsSizeOrColor()
#
# This function checks a string for a size, looking for patterns that resemble sizes or colors.
# If it finds a match, it returns true.  Else, false
#
# Input: PartNum, A part number as a string.
#
# Output: A boolean indicating whether the input part number contains a size or not.
#
Function FindSizeAndSplit($PartNumber){
    
    #Set up the garment size regex string
    #$SizeRegex = "^.+[^a-z]+X*[lsm]{1}"
    #$SizeRegex = "[^a-z]+x*[lsm]{1}"
    $SizeRegex = "((((?!\d?x*[lsm])\w*)(\W?)){1,})(\d?)(x*)(me?d|lr?g|sml?|l|s|m)(.*)"
    #Check if the part number contains a garment size
        if($PartNumber -imatch $SizeRegex){
            #Create the dooley size
            $PartSize = $Matches[5] + $Matches[6] + $Matches[7]
            $PartSize = SizeConverter $PartSize
            if($Matches[8] -eq ""){
                $DTIPartNum = $Matches[3] + $DTISeparator + $PartSize
            }
            else{
                $DTIPartNum = $Matches[3] + $DTISeparator + $Matches[8] + $DTISeparator + $PartSize
            }
        }
        else{
            #If it does not match the size expression, it may be a number size at the end.  
            #First, convert all non-word expressions to a DTI separator
            $PartNumber = $PartNumber -replace '\W', "$DTISeparator"
            if($PartNumber -match "(.*[^0-9]{1,})([0-9]{1,2})$"){
                $PartNumber = $Matches[1] + $DTISeparator + $Matches[2]
            }
            $DTIPartNum = $PartNumber
        }
    $DTIPartNum
}
     
#TestFunction
#
# This function is only used to test if DOT including is working properly.
#
Function TestFunction(){
    Echo "Hey there, the function include works!"
}

#
# Iterations of $SizeRegex
#
#
#
#$SizeRegex = "(.*)(\d*)(\W?)(\d?)(x*)([s|l|m|med|lrg|sml]{1})(.*)" 
#
# This one captures all the random characters in the beginning and does not
# account for the non-word character aspect. 
# 
# "(\w*)(\W?)(\d?)(x*)([s|l|m|med|lrg|sml]{1})(.*)"
#
# Does not account for possibility of multiple non-word characters before size.
#
# "([(\w*)(\W?)]{1,})(\d?)(x*)([s|l|m|med|lrg|sml]{1})(.*)"
#
# This expression eats up the x in a size with the first few sizes
#
# ((((?!\d?x*[lsm])\w*)(\W?)){1,})(\d?)(x*)(me?d|lr?g|sml?|l|s|m)(.*)
#
# So far, so good!
