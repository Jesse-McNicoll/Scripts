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
    

Function LoopThruPartnums($SizeArray, $SmallIndex, $BigIndex, $BasePartNum, $PartCost, $PartNum, $PUM){
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