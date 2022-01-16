# Edit these 2 before running:
$InputPath = "" #Path to xlsx file containing data, fx C:\RawData.xlsx
$OutputPath = "" #Path to where to write output, fx C:\code.zs

# Also remember to add
# import mods.efabct.EFabRecipe;
# to top of output file

#############################################
Import-module PSExcel

$recipes = new-object System.Collections.ArrayList

foreach ($recipe in (Import-XLSX -Path $InputPath -RowStart 1))
{
    $recipes.add($recipe) | out-null
}

foreach ($recipe in $recipes)
{

    #Get input items
    $RawInputFull = $recipe | Select-Object -Property input | fl | Out-String
    $RawInputFull = $RawInputFull.Trim()
    $RawInputFull = $RawInputFull.Substring(8)
    $AllInputs = $RawInputFull -split("A: ") -split (" B: ") -split (" C: ") -split (" D: ") -split (" E: ") -split (" F: ") -split (" G: ") -split (" H: ") -split (" I: ")
    $AllInputs = $AllInputs.trim()

    $RawInputA = $AllInputs[1]
    $RawInputB = $AllInputs[2]
    $RawInputC = $AllInputs[3]
    $RawInputD = $AllInputs[4]
    $RawInputE = $AllInputs[5]
    $RawInputF = $AllInputs[6]
    $RawInputG = $AllInputs[7]
    $RawInputH = $AllInputs[8]
    $RawInputI = $AllInputs[9]
    

    #Create pattern variables and puts them into positions... idfk if this works until ive tested it
    $Pattern = $recipe | Select-Object -Property pattern | fl | Out-String
    $Pattern = $pattern.Trim()
    $Pattern = $Pattern.Substring(10)

    #Sets position values

    #Pos 1
    try
    {
        if ($Pattern.Substring(0,1) -ne "-")
        {
            $PatPos1 = $Pattern.Substring(0,1)      
        }
        else
        {
            $PatPos1 = "null"
        }
    }
    catch
    {
        $PatPos1 = "null"
    }

    if ($PatPos1 -ne "null")
    {
        if ($PatPos1 -eq "A")
        {
            $inputItemA = "<"+$RawInputA+">"
        }
        if ($PatPos1 -eq "B")
        {
            $inputItemA = "<"+$RawInputB+">"
        }
        if ($PatPos1 -eq "C")
        {
            $inputItemA = "<"+$RawInputC+">"
        }
        if ($PatPos1 -eq "D")
        {
            $inputItemA = "<"+$RawInputD+">"
        }
        if ($PatPos1 -eq "E")
        {
            $inputItemA = "<"+$RawInputE+">"
        }
        if ($PatPos1 -eq "F")
        {
            $inputItemA = "<"+$RawInputF+">"
        }
        if ($PatPos1 -eq "G")
        {
            $inputItemA = "<"+$RawInputG+">"
        }
        if ($PatPos1 -eq "H")
        {
            $inputItemA = "<"+$RawInputH+">"
        }
        if ($PatPos1 -eq "I")
        {
            $inputItemA = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemA = "null"
    }

    #Pos 2
    try
    {
        if ($Pattern.Substring(1,1) -ne "-")
        {
            $PatPos2 = $Pattern.Substring(1,1)     
        }
        else
        {
            $PatPos2 = "null"
        }
    }
    catch
    {
        $PatPos2 = "null"
    }

    if ($PatPos2 -ne "null")
    {
        if ($PatPos2 -eq "A")
        {
            $inputItemB = "<"+$RawInputA+">"
        }
        if ($PatPos2 -eq "B")
        {
            $inputItemB = "<"+$RawInputB+">"
        }
        if ($PatPos2 -eq "C")
        {
            $inputItemB = "<"+$RawInputC+">"
        }
        if ($PatPos2 -eq "D")
        {
            $inputItemB = "<"+$RawInputD+">"
        }
        if ($PatPos2 -eq "E")
        {
            $inputItemB = "<"+$RawInputE+">"
        }
        if ($PatPos2 -eq "F")
        {
            $inputItemB = "<"+$RawInputF+">"
        }
        if ($PatPos2 -eq "G")
        {
            $inputItemB = "<"+$RawInputG+">"
        }
        if ($PatPos2 -eq "H")
        {
            $inputItemB = "<"+$RawInputH+">"
        }
        if ($PatPos2 -eq "I")
        {
            $inputItemB = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemB = "null"
    }

    #Pos 3
    try
    {
        if ($Pattern.Substring(2,1) -ne "-")
        {
            $PatPos3 = $Pattern.Substring(2,1)   
        }
        else
        {
            $PatPos3 = "null"
        }
    }
    catch
    {
        $PatPos3 = "null"
    }

    if ($PatPos3 -ne "null")
    {
        if ($PatPos3 -eq "A")
        {
            $inputItemC = "<"+$RawInputA+">"
        }
        if ($PatPos3 -eq "B")
        {
            $inputItemC = "<"+$RawInputB+">"
        }
        if ($PatPos3 -eq "C")
        {
            $inputItemC = "<"+$RawInputC+">"
        }
        if ($PatPos3 -eq "D")
        {
            $inputItemC = "<"+$RawInputD+">"
        }
        if ($PatPos3 -eq "E")
        {
            $inputItemC = "<"+$RawInputE+">"
        }
        if ($PatPos3 -eq "F")
        {
            $inputItemC = "<"+$RawInputF+">"
        }
        if ($PatPos3 -eq "G")
        {
            $inputItemC = "<"+$RawInputG+">"
        }
        if ($PatPos3 -eq "H")
        {
            $inputItemC = "<"+$RawInputH+">"
        }
        if ($PatPos3 -eq "I")
        {
            $inputItemC = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemC = "null"
    }

    #Pos 4
    try
    {
        if ($Pattern.Substring(3,1) -ne "-")
        {
            $PatPos4 = $Pattern.Substring(3,1)   
        }
        else
        {
            $PatPos4 = "null"
        }
    }
    catch
    {
        $PatPos4 = "null"
    }

    if ($PatPos4 -ne "null")
    {
        if ($PatPos4 -eq "A")
        {
            $inputItemD = "<"+$RawInputA+">"
        }
        if ($PatPos4 -eq "B")
        {
            $inputItemD = "<"+$RawInputB+">"
        }
        if ($PatPos4 -eq "C")
        {
            $inputItemD = "<"+$RawInputC+">"
        }
        if ($PatPos4 -eq "D")
        {
            $inputItemD = "<"+$RawInputD+">"
        }
        if ($PatPos4 -eq "E")
        {
            $inputItemD = "<"+$RawInputE+">"
        }
        if ($PatPos4 -eq "F")
        {
            $inputItemD = "<"+$RawInputF+">"
        }
        if ($PatPos4 -eq "G")
        {
            $inputItemD = "<"+$RawInputG+">"
        }
        if ($PatPos4 -eq "H")
        {
            $inputItemD = "<"+$RawInputH+">"
        }
        if ($PatPos4 -eq "I")
        {
            $inputItemD = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemD = "null"
    }

    #Pos 5
    try
    {
        if ($Pattern.Substring(4,1) -ne "-")
        {
            $PatPos5 = $Pattern.Substring(4,1)   
        }
        else
        {
            $PatPos5 = "null"
        }
    }
    catch
    {
        $PatPos5 = "null"
    }

    if ($PatPos5 -ne "null")
    {
        if ($PatPos5 -eq "A")
        {
            $inputItemE = "<"+$RawInputA+">"
        }
        if ($PatPos5 -eq "B")
        {
            $inputItemE = "<"+$RawInputB+">"
        }
        if ($PatPos5 -eq "C")
        {
            $inputItemE = "<"+$RawInputC+">"
        }
        if ($PatPos5 -eq "D")
        {
            $inputItemE = "<"+$RawInputD+">"
        }
        if ($PatPos5 -eq "E")
        {
            $inputItemE = "<"+$RawInputE+">"
        }
        if ($PatPos5 -eq "F")
        {
            $inputItemE = "<"+$RawInputF+">"
        }
        if ($PatPos5 -eq "G")
        {
            $inputItemE = "<"+$RawInputG+">"
        }
        if ($PatPos5 -eq "H")
        {
            $inputItemE = "<"+$RawInputH+">"
        }
        if ($PatPos5 -eq "I")
        {
            $inputItemE = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemE = "null"
    }

    #Pos 6
    try
    {
        if ($Pattern.Substring(5,1) -ne "-")
        {
            $PatPos6 = $Pattern.Substring(5,1)   
        }
        else
        {
            $PatPos6 = "null"
        }
    }
    catch
    {
        $PatPos6 = "null"
    }

    if ($PatPos6 -ne "null")
    {
        if ($PatPos6 -eq "A")
        {
            $inputItemF = "<"+$RawInputA+">"
        }
        if ($PatPos6 -eq "B")
        {
            $inputItemF = "<"+$RawInputB+">"
        }
        if ($PatPos6 -eq "C")
        {
            $inputItemF = "<"+$RawInputC+">"
        }
        if ($PatPos6 -eq "D")
        {
            $inputItemF = "<"+$RawInputD+">"
        }
        if ($PatPos6 -eq "E")
        {
            $inputItemF = "<"+$RawInputE+">"
        }
        if ($PatPos6 -eq "F")
        {
            $inputItemF = "<"+$RawInputF+">"
        }
        if ($PatPos6 -eq "G")
        {
            $inputItemF = "<"+$RawInputG+">"
        }
        if ($PatPos6 -eq "H")
        {
            $inputItemF = "<"+$RawInputH+">"
        }
        if ($PatPos6 -eq "I")
        {
            $inputItemF = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemF = "null"
    }

    #Pos 7
    try
    {
        if ($Pattern.Substring(6,1) -ne "-")
        {
            $PatPos7 = $Pattern.Substring(6,1)   
        }
        else
        {
            $PatPos7 = "null"
        }
    }
    catch
    {
        $PatPos7 = "null"
    }

    if ($PatPos7 -ne "null")
    {
        if ($PatPos7 -eq "A")
        {
            $inputItemG = "<"+$RawInputA+">"
        }
        if ($PatPos7 -eq "B")
        {
            $inputItemG = "<"+$RawInputB+">"
        }
        if ($PatPos7 -eq "C")
        {
            $inputItemG = "<"+$RawInputC+">"
        }
        if ($PatPos7 -eq "D")
        {
            $inputItemG = "<"+$RawInputD+">"
        }
        if ($PatPos7 -eq "E")
        {
            $inputItemG = "<"+$RawInputE+">"
        }
        if ($PatPos7 -eq "F")
        {
            $inputItemG = "<"+$RawInputF+">"
        }
        if ($PatPos7 -eq "G")
        {
            $inputItemG = "<"+$RawInputG+">"
        }
        if ($PatPos7 -eq "H")
        {
            $inputItemG = "<"+$RawInputH+">"
        }
        if ($PatPos7 -eq "I")
        {
            $inputItemG = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemG = "null"
    }

    #Pos 8
    try
    {
        if ($Pattern.Substring(7,1) -ne "-")
        {
            $PatPos8 = $Pattern.Substring(7,1)   
        }
        else
        {
            $PatPos8 = "null"
        }
    }
    catch
    {
        $PatPos8 = "null"
    }

    if ($PatPos8 -ne "null")
    {
        if ($PatPos8 -eq "A")
        {
            $inputItemH = "<"+$RawInputA+">"
        }
        if ($PatPos8 -eq "B")
        {
            $inputItemH = "<"+$RawInputB+">"
        }
        if ($PatPos8 -eq "C")
        {
            $inputItemH = "<"+$RawInputC+">"
        }
        if ($PatPos8 -eq "D")
        {
            $inputItemH = "<"+$RawInputD+">"
        }
        if ($PatPos8 -eq "E")
        {
            $inputItemH = "<"+$RawInputE+">"
        }
        if ($PatPos8 -eq "F")
        {
            $inputItemH = "<"+$RawInputF+">"
        }
        if ($PatPos8 -eq "G")
        {
            $inputItemH = "<"+$RawInputG+">"
        }
        if ($PatPos8 -eq "H")
        {
            $inputItemH = "<"+$RawInputH+">"
        }
        if ($PatPos8 -eq "I")
        {
            $inputItemH = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemH = "null"
    }

    #Pos 9
    try
    {
        if ($Pattern.Substring(8,1) -ne "-")
        {
            $PatPos9 = $Pattern.Substring(8,1)   
        }
        else
        {
            $PatPos9 = "null"
        }
    }
    catch
    {
        $PatPos9 = "null"
    }

    if ($PatPos9 -ne "null")
    {
        if ($PatPos9 -eq "A")
        {
            $inputItemI = "<"+$RawInputA+">"
        }
        if ($PatPos9 -eq "B")
        {
            $inputItemI = "<"+$RawInputB+">"
        }
        if ($PatPos9 -eq "C")
        {
            $inputItemI = "<"+$RawInputC+">"
        }
        if ($PatPos9 -eq "D")
        {
            $inputItemI = "<"+$RawInputD+">"
        }
        if ($PatPos9 -eq "E")
        {
            $inputItemI = "<"+$RawInputE+">"
        }
        if ($PatPos9 -eq "F")
        {
            $inputItemI = "<"+$RawInputF+">"
        }
        if ($PatPos9 -eq "G")
        {
            $inputItemI = "<"+$RawInputG+">"
        }
        if ($PatPos9 -eq "H")
        {
            $inputItemI = "<"+$RawInputH+">"
        }
        if ($PatPos9 -eq "I")
        {
            $inputItemI = "<"+$RawInputI+">"
        }
    }
    else
    {
        $inputItemI = "null"
    }

    #Create output item variable #WORKS
    $outputItem = $recipe | Select-Object -Property output | fl | Out-String
    $outputItem = $outputItem.Trim()
    $outputItem = $outputItem.Substring(9)

#Formats output and appends to zs file
    $codeOutput =
@"
EFabRecipe.shaped(<$outputItem>, 
[[$inputItemA, $inputItemB, $inputItemC],
 [$inputItemD, $inputItemE, $inputItemF],
 [$inputItemG, $inputItemH, $inputItemI]])
.tier('')
.time(0)
.rfPerTick(0);


"@

    $codeOutput | Out-File $OutputPath -Append -Encoding utf8
}
