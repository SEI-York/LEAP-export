REM Export LEAP Data to UNFCCC template
Option Explicit   ' Requires all variables to be DIMmed
'######################################################################################
' Define variables
' objLEAPBranchList -> cache list of branches that has tags
'######################################################################################
DIM LEAP, objExcelApp, objExcelWb, arrLEAPFuels, objLEAPBranchList
DIM intYear : intYear = 2010
' Define pollutants, Pollutant Loadings in LEAP
DIM strCO2 : strCO2 = "CO2"
DIM strCH4 : strCH4 = "CH4"
DIM strN2O : strN2O = "N2O"

'######################################################################################
' Define variables - Table 1 SECTORAL REPORT FOR ENERGY
'######################################################################################
FUNCTION GetCodesForSectoralReportForEnergy()
    GetCodesForSectoralReportForEnergy = Array("1.A.1.", _
                                "1.A.1.a.", "1.A.1.b.", "1.A.1.c.", _
                                "1.A.2.", _
                                "1.A.2.a.", "1.A.2.b.", "1.A.2.c.", "1.A.2.d.", "1.A.2.e.", "1.A.2.f.", "1.A.2.g.", _
                                "1.A.3.", _
                                "1.A.3.a.", "1.A.3.b.", "1.A.3.c.", "1.A.3.d.", "1.A.3.e.", _
                                "1.A.4.", _
                                "1.A.4.a.", "1.A.4.b.", "1.A.4.c.", _
                                "1.A.5.", _
                                "1.A.5.a.", "1.A.5.b.", _
                                "1.B.1.", _
                                "1.B.1.a.", "1.B.1.b.", "1.B.1.c.", _
                                "1.B.2.", _
                                "1.B.2.a.", "1.B.2.b.", "1.B.2.c.", "1.B.2.d.", _
                                "1.C.1.", _
                                "1.C.2.", _
                                "1.C.3.", _
                                "1.D.1.", _
                                "1.D.1.a.", "1.D.1.b.", _
                                "1.D.2.", _
                                "1.D.3.", _
                                "1.D.4.", _
                                "1.D.4.a.", "1.D.4.b." _
                                )
END Function


' Match column names in Table 1
FUNCTION GetPollutantsForSectoralReportForEnergy()
    GetPollutantsForSectoralReportForEnergy = Array(strCO2,_
                                        strCH4, _
                                        strN2O, _
                                        "NOx", _
                                        "CO", _
                                        "NMVOC", _
                                        "SO2")
END FUNCTION

FUNCTION GetFuelsForSectoralBackgroundData()
    '####################################
    ' Update fuels according to LEAP
    '####################################
    DIM dictFuels : SET dictFuels = CreateObject("Scripting.Dictionary")

    ' Define Biomass fuels
    dictFuels.Add "arrBiomassFuels", Array("Wood", "Biogas", "Charcoal", "Ethanol feedstock", "Vegetal Wastes", _
                                            "Bio diesel feedstock")

    ' Define Solid fuels (not including biomass)
    dictFuels.Add "arrSolidFuels", Array("Coal", "Bitumen")

    ' Define Liquid fuels
    dictFuels.Add "arrLiquidFuels", Array("Diesel", "Gasoline", "Gasoline premix", "Jet Kerosene", "Kerosene", _
                                        "LPG", "Residual Fuel Oil", "Crude Oil", "Ethanol", _
                                        "Biodiesel", "HFO")

    ' Define Gaseous fuels
    dictFuels.Add "arrGaseousFuels", Array("LNG", "Natural Gas")

    ' Aviation gasoline
    dictFuels.Add "arrAviationGasoline", Array("Aviation gasoline")
    ' Jet kerosene
    dictFuels.Add "arrJetKerosene", Array("Jet Kerosene", "Kerosene")
    ' Gasoline
    dictFuels.Add "arrGasoline", Array("Gasoline")
    ' Diesel oil
    dictFuels.Add "arrDieselOil", Array("Diesel")
    ' Liquefied petroleum gases (LPG)
    dictFuels.Add "arrLPG", Array("LPG")

    SET GetFuelsForSectoralBackgroundData = dictFuels
END FUNCTION


'######################################################################################
' Define variables - Table 2 Sectoral report for industrial processes and product use
'######################################################################################
FUNCTION GetCodesForSectoralReportForIndustrial()
    GetCodesForSectoralReportForIndustrial = Array("2.A.",_
                                                    "2.A.1.", "2.A.2.", "2.A.3.", "2.A.4.",_
                                                    "2.B.",_
                                                    "2.B.1.", "2.B.2.", "2.B.3.", "2.B.4.", "2.B.5.", "2.B.6.", "2.B.7.", "2.B.8.", "2.B.9.", "2.B.10.",_
                                                    "2.C.",_
                                                    "2.C.1.", "2.C.2.", "2.C.3.", "2.C.4.", "2.C.5.", "2.C.6.", "2.C.7.",_
                                                    "2.D.",_
                                                    "2.D.1.", "2.D.2.", "2.D.3.",_
                                                    "2.E.",_
                                                    "2.E.1.", "2.E.2.", "2.E.3.", "2.E.4.", "2.E.5.",_
                                                    "2.F.",_
                                                    "2.F.1.", "2.F.2.", "2.F.3.", "2.F.4.", "2.F.5.", "2.F.6.",_
                                                    "2.G.",_
                                                    "2.G.1.", "2.G.2.", "2.G.3.", "2.G.4.")
END FUNCTION

' Match column names in Table 2
FUNCTION GetPollutantsForSectoralReportForIndustrial()
    GetPollutantsForSectoralReportForIndustrial = Array(strCO2,_
                                                    strCH4, _
                                                    strN2O, _
                                                    "HFCs", _
                                                    "PFCs", _
                                                    "HFCs_Mix", _
                                                    "SF6",_
                                                    "NF3",_
                                                    "NOx",_
                                                    "CO",_
                                                    "NMVOC",_
                                                    "SOx")
END FUNCTION

'######################################################################################
' Define variables - Table 3 Sectoral report for agriculture
'######################################################################################
FUNCTION GetCodesForSectoralReportForAgriculture()
    GetCodesForSectoralReportForAgriculture = Array("3.A.", "3.A.1.",_
                                                    "3.A.1.a.", "3.A.1.b.",_
                                                    "3.A.2.", "3.A.3.", "3.A.4.",_
                                                    "3.B.", "3.B.1.",_
                                                    "3.B.1.a.", "3.B.1.b.",_
                                                    "3.B.2.", "3.B.3.", "3.B.4.", "3.B.5.",_
                                                    "3.C.",_
                                                    "3.D.", "3.D.1.",_
                                                    "3.D.1.a.", "3.D.1.b.", "3.D.1.c.", "3.D.1.d.", "3.D.1.e.", "3.D.1.f.", "3.D.1.g.",_
                                                    "3.D.2.",_
                                                    "3.E.", "3.F.", "3.G.", "3.H.", "3.I.", "3.J.")
END FUNCTION

' Match column names in Table 3
FUNCTION GetPollutantsForSectoralReportForAgriculture()
    GetPollutantsForSectoralReportForAgriculture = Array(strCO2,_
                                                        strCH4, _
                                                        strN2O, _
                                                        "NOx", _
                                                        "CO", _
                                                        "NMVOC", _
                                                        "SOx")
END FUNCTION


'######################################################################################
' Define variables - Table 4 Sectoral report for land use, land-use change and forestry
'######################################################################################
FUNCTION GetCodesForSectoralReportForLandUse()
    GetCodesForSectoralReportForLandUse = Array("4.A.", "4.A.1.", "4.A.2.",_
                                                "4.B.", "4.B.1.", "4.B.2.",_
                                                "4.C.", "4.C.1.", "4.C.2.",_
                                                "4.D.", "4.D.1.", "4.D.2.",_
                                                "4.E.", "4.E.1.", "4.E.2.",_
                                                "4.F.", "4.F.1.", "4.F.2.",_
                                                "4.G.", "4.H.")
END FUNCTION

' Match column names in Table 4
FUNCTION GetPollutantsForSectoralReportForLandUse()
    GetPollutantsForSectoralReportForLandUse = Array(strCH4, _
                                                    strN2O, _
                                                    "NOx", _
                                                    "CO", _
                                                    "NMVOC")

END FUNCTION


'######################################################################################
' Program
'######################################################################################
REM On Error Resume Next
REM Err.Clear
CALL MainScript

SUB MainScript
    On Error Resume Next
    SET LEAP = GetObject(, "LEAP.LEAPApplication")
    SET objExcelApp = CreateObject("Excel.Application")
    objExcelApp.Visible = True

    CALL PopulateLEAPFuels()
    CALL PopulateLEAPBranches()

    ' calculate the result
    LEAP.ActiveView = "Results" 'Switch result view
    LEAP.ActiveScenario = 2 'Switch to first scenario

    DIM strExcelFilePath : strExcelFilePath = GetExcelFile()
    SET objExcelWb = objExcelApp.Workbooks.Open(strExcelFilePath)

    ' Table 1
    CALL ExportSectoralReportForEnergy()

    ' Table2(I)
    CALL ExportSectoralReportForIndustrial()

    ' Table3
    CALL ExportSectoralReportForAgriculture()

    ' Table4
    CALL ExportSectoralReportForLandUse()

    PRINT "Data export finished!"

    If Err.number <> 0 Then 
        MsgBox "Error " & Err.number & ": " & Err.Description
        If Not objExcelWb Is Nothing Then objExcelWb.Close False
        If Not objExcelApp Is Nothing Then objExcelApp.Quit
    End If
END SUB

'######################################################################################
' Functions
'######################################################################################
FUNCTION GetExcelFile()
    ' COPY template file to temp folder
    DIM objFS : SET objFS = CreateObject("Scripting.FileSystemObject")
    DIM strTempSourcePath : strTempSourcePath = LEAP.DictionaryDirectory & "UNFCCC template.xlsx"
    DIM strTempTargetPath : strTempTargetPath = objFS.GetSpecialFolder(2) & "\"
    objFS.CopyFile strTempSourcePath, strTempTargetPath, True
    GetExcelFile = strTempTargetPath & "UNFCCC template.xlsx"
END FUNCTION


FUNCTION GetBranchIndex(strCode)
    ' Find the branch index if the branch name ends with the giving code
    DIM objIndexList : SET objIndexList = CreateObject("System.Collections.ArrayList")
    DIM intNum, tag
    FOR EACH intNum IN objLEAPBranchList
        FOR EACH tag IN LEAP.Branches(intNum).Tags
            IF strCode = Right(tag.Name, Len(strCode)) THEN
                objIndexList.Add intNum
                EXIT FOR
            END IF 
        NEXT
    NEXT
    ' no match found, return -1
    SET GetBranchIndex = objIndexList
END FUNCTION


FUNCTION GetRowIndex(objRange, strSearchString)
    ' The Find() function returns a Range object if successful or Nothing if not.
    DIM objFind  : SET objFind  = objRange.Find(strSearchString)
    IF objFind Is Nothing Then
        ' Not found
        ' Log message
        GetRowIndex = -1
    ELSE
        GetRowIndex = objFind.Row
    END IF
END FUNCTION

FUNCTION GetRowIndexByCode(objExcelWs, strColumn, strCode)
    ' Find the row index in the excel sheet
    DIM objRange : SET objRange = objExcelWs.Range(strColumn & ":" & strColumn)
    ' add space to end as delimiter
    GetRowIndexByCode =  GetRowIndex(objRange, strCode & " ")
END FUNCTION

FUNCTION GetRowIndexByColumnRow(objExcelWs, strSearchString, intStartRow, intEndRow, strColumn)
    DIM objRange : SET objRange = objExcelWs.Range(strColumn & intStartRow & ":" & strColumn & intEndRow)
    GetRowIndexByColumnRow =  GetRowIndex(objRange, strSearchString)
END FUNCTION

FUNCTION GetRowIndexAfterCode(objExcelWs, strColumn, strCode, strSearchString)
    ' get the row index for the code
    DIM intRowIndexForCode : intRowIndexForCode = GetRowIndexByCode(objExcelWs, strColumn, strCode)
    IF intRowIndexForCode > 0 Then
        ' get the row index for the similar code within the next 20 rows
        ' this is to avoid assign value to a wrong category
        DIM strFirstCodePart : strFirstCodePart = Mid(strCode, 1, 4)
        DIM intNextRowIndexForCode : intNextRowIndexForCode = GetRowIndexByColumnRow(objExcelWs, strFirstCodePart, intRowIndexForCode + 1, intRowIndexForCode + 20, strColumn)
        ' default search 10 rows down
        DIM intSearchEndIndex : intSearchEndIndex = intRowIndexForCode + 10
        IF intNextRowIndexForCode > 0 THEN
            intSearchEndIndex = intNextRowIndexForCode
        END IF
        ' get the row index for the search string
        GetRowIndexAfterCode = GetRowIndexByColumnRow(objExcelWs, strSearchString, intRowIndexForCode, intSearchEndIndex, strColumn)
    ELSE
        ' cannot find the code
        GetRowIndexAfterCode = -1
    END IF
END FUNCTION


FUNCTION GetWsName(strCode)
    ' get the name of excel sheet
    IF InStr(1, strCode, "1.A.1") = 1 THEN
        GetWsName = "Table1.A(a)s1"
    ELSEIF InStr(1, strCode, "1.A.2") = 1 Then
        GetWsName = "Table1.A(a)s2"
    ELSEIF InStr(1, strCode, "1.A.3") = 1 Then
        GetWsName = "Table1.A(a)s3"
    ELSEIF InStr(1, strCode, "1.A.4") = 1 OR InStr(1, strCode, "1.A.5") = 1 Then
        GetWsName = "Table1.A(a)s4"
    End IF
END FUNCTION

FUNCTION GetFuelConsumption(intBranchIndex, strUnit, strFuel)
    On Error Resume Next

    IF InArray(arrLEAPFuels, strFuel) THEN
        GetFuelConsumption = LEAP.Branches(intBranchIndex).Variable("Energy Demand Final Units").Value(intYear, strUnit, "Fuel="&strFuel)
    ELSE
        ' fuel not exist
        PRINT "Fuel '" + strFuel + "' is not defined in LEAP"
        GetFuelConsumption = 0
    END IF
    
    If Err.number <> 0 Then
        ' if there is an error finding the value, log it and output 0
        GetFuelConsumption = 0
    End If
END FUNCTION

FUNCTION GetEmission(intBranchIndex, strUnit, strPollutant, strFuel)
    ' strFuel is optional, if Null is provided, emissions from all fuels will be return
    On Error Resume Next

    DIM strFilter : strFilter = "Effect="&strPollutant
    IF NOT IsNull(strFuel) THEN
        IF InArray(arrLEAPFuels, strFuel) THEN 
            strFilter = strFilter & "|Fuel=" & strFuel
        ELSE
            ' fuel not exist
            PRINT "Fuel '" + strFuel + "' is not defined in LEAP"
            GetEmission = 0
            EXIT FUNCTION
        END IF
    END IF

    GetEmission = LEAP.Branches(intBranchIndex).Variable("Pollutant Loadings").Value(intYear, strUnit, strFilter)

    If Err.number <> 0 Then
        ' if there is an error finding the value, log it and output 0
        GetEmission = 0
    End If
END FUNCTION

' add item to array
FUNCTION ArrayAdd(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    ArrayAdd = arr
END FUNCTION

' check item in array
FUNCTION InArray(arr, val)
    InArray = False
    DIM item
    FOR EACH item IN arr
        IF item = val THEN
            InArray = True
            EXIT FOR
        END IF
    Next
END FUNCTION


'######################################################################################
' Subroutines
'######################################################################################

SUB PopulateLEAPFuels()
    arrLEAPFuels = Array()
    DIM objFuel
    FOR EACH objFuel IN LEAP.Fuels
        arrLEAPFuels = ArrayAdd(arrLEAPFuels, objFuel.Name)
    NEXT
END SUB

SUB PopulateLEAPBranches()
    SET objLEAPBranchList = CreateObject("System.Collections.ArrayList")
    DIM intNum
    FOR intNum = 1 TO LEAP.Branches.Count
        IF LEAP.Branches(intNum).Tags.Count > 0 THEN
            objLEAPBranchList.Add intNum
        END IF 
    NEXT
END SUB

' SECTORAL REPORT FOR ENERGY - Table 1
SUB ExportSectoralReportForEnergy()
    ' Define all the branches that we would like to extract data from
    DIM arrCodes : arrCodes = GetCodesForSectoralReportForEnergy()
    
    DIM arrPollutants : arrPollutants = GetPollutantsForSectoralReportForEnergy()

    DIM strCode
    FOR EACH strCode IN arrCodes
        DIM objBranchIndex
        ' search branch in LEAP
        SET objBranchIndex = GetBranchIndex(strCode)
        IF objBranchIndex.Count > 0 Then
            ' Branch found, continue 
            DIM intBranchIndex
            For Each intBranchIndex in objBranchIndex
                ' Table1 - summary sheet
                DIM objExcelWs : SET objExcelWs = objExcelWb.Worksheets("Table1")
                ' search row in EXCEL (column B)
                DIM intRow : intRow = GetRowIndexByCode(objExcelWs, "B", strCode)
                IF intRow > 0 THEN
                    ' row in excel found, output data
                    DIM strPollutant
                    ' data columns start at C in Table1 
                    DIM intColumn : intColumn = 3
                    FOR EACH strPollutant IN arrPollutants
                        ' get value from LEAP
                        DIM varEmission : varEmission = GetEmission(intBranchIndex, "Metric Tonne", strPollutant, Null)
                        ' TONNE TO KT
                        varEmission = varEmission / 1000
                        ' save it to excel file
                        CALL SetExcelCellValue(objExcelWs, intRow, intColumn, varEmission, False)
                        intColumn = intColumn + 1
                    NEXT
                END IF

                ' detailed sectoral background data
                CALL ExportSectoralBackgroundData(objExcelWb, intBranchIndex, strCode)
            NEXT
        ELSE
            ' Code not found, skip
            ' Log
        END IF
    NEXT
END SUB

' Sectoral report for industrial processes and product use Table 2
SUB ExportSectoralReportForIndustrial()
    DIM arrCodes : arrCodes = GetCodesForSectoralReportForIndustrial()

    DIM arrPollutants : arrPollutants = GetPollutantsForSectoralReportForIndustrial()

    DIM strCode
    FOR EACH strCode IN arrCodes
        DIM objBranchIndex
        ' search branch in LEAP
        SET objBranchIndex = GetBranchIndex(strCode)
        IF objBranchIndex.Count > 0 Then
            ' Branch found, continue 
            DIM intBranchIndex
            For Each intBranchIndex in objBranchIndex
                ' Table2(I)
                DIM objExcelWs : SET objExcelWs = objExcelWb.Worksheets("Table2(I)")
                ' search row in EXCEL (column B)
                DIM intRow : intRow = GetRowIndexByCode(objExcelWs, "B", strCode)
                IF intRow > 0 THEN
                    ' row in excel found, output data
                    DIM strPollutant
                    ' data columns start at C in Table1 
                    DIM intColumn : intColumn = 3
                    FOR EACH strPollutant IN arrPollutants
                        ' get value from LEAP
                        DIM varEmission : varEmission = GetEmission(intBranchIndex, "Metric Tonne", strPollutant, Null)
                        ' TONNE TO KT
                        varEmission = varEmission / 1000
                        ' save it to excel file
                        CALL SetExcelCellValue(objExcelWs, intRow, intColumn, varEmission, False)
                        intColumn = intColumn + 1
                    NEXT
                END IF
            NEXT
        ELSE
            ' Code not found, skip
            ' Log
        END IF
    NEXT
END SUB


' Sectoral report for agriculture - Table 3
SUB ExportSectoralReportForAgriculture()
    DIM arrCodes : arrCodes = GetCodesForSectoralReportForAgriculture()
    DIM arrPollutants : arrPollutants = GetPollutantsForSectoralReportForAgriculture()

    DIM strCode
    FOR EACH strCode IN arrCodes
        DIM objBranchIndex
        ' search branch in LEAP
        SET objBranchIndex = GetBranchIndex(strCode)
        IF objBranchIndex.Count > 0 Then
            ' Branch found, continue 
            DIM intBranchIndex
            For Each intBranchIndex in objBranchIndex
                ' Table2(I)
                DIM objExcelWs : SET objExcelWs = objExcelWb.Worksheets("Table3")
                ' search row in EXCEL (column B)
                DIM intRow : intRow = GetRowIndexByCode(objExcelWs, "B", strCode)
                IF intRow > 0 THEN
                    ' row in excel found, output data
                    DIM strPollutant
                    ' data columns start at C in Table1 
                    DIM intColumn : intColumn = 3
                    FOR EACH strPollutant IN arrPollutants
                        ' get value from LEAP
                        DIM varEmission : varEmission = GetEmission(intBranchIndex, "Metric Tonne", strPollutant, Null)
                        ' TONNE TO KT
                        varEmission = varEmission / 1000
                        ' save it to excel file
                        CALL SetExcelCellValue(objExcelWs, intRow, intColumn, varEmission, False)
                        intColumn = intColumn + 1
                    NEXT
                END IF
            NEXT
        ELSE
            ' Code not found, skip
            ' Log
        END IF
    NEXT

END SUB


' Sectoral report for land use, land-use change and forestry - Table 4
SUB ExportSectoralReportForLandUse()
    DIM arrCodes : arrCodes = GetCodesForSectoralReportForLandUse()
    DIM arrPollutants : arrPollutants = GetPollutantsForSectoralReportForLandUse()

    DIM strCode
    FOR EACH strCode IN arrCodes
        DIM objBranchIndex
        ' search branch in LEAP
        SET objBranchIndex = GetBranchIndex(strCode)
        IF objBranchIndex.Count > 0 Then
            ' Branch found, continue 
            DIM intBranchIndex
            For Each intBranchIndex in objBranchIndex
                ' Table2(I)
                DIM objExcelWs : SET objExcelWs = objExcelWb.Worksheets("Table4")
                ' search row in EXCEL (column B)
                DIM intRow : intRow = GetRowIndexByCode(objExcelWs, "B", strCode)
                IF intRow > 0 THEN
                    ' row in excel found, output data
                    DIM strPollutant
                    ' data columns start at D in Table1 
                    DIM intColumn : intColumn = 4
                    FOR EACH strPollutant IN arrPollutants
                        ' get value from LEAP
                        DIM varEmission : varEmission = GetEmission(intBranchIndex, "Metric Tonne", strPollutant, Null)
                        ' TONNE TO KT
                        varEmission = varEmission / 1000
                        ' save it to excel file
                        CALL SetExcelCellValue(objExcelWs, intRow, intColumn, varEmission, False)
                        intColumn = intColumn + 1
                    NEXT
                END IF
            NEXT
        ELSE
            ' Code not found, skip
            ' Log
        END IF
    NEXT
END SUB



SUB SetExcelCellValue(objExcelWs, intRow, intColumn, varValue, boolOverwrite)
    IF boolOverwrite THEN
        objExcelWs.Cells(intRow, intColumn).Formula = "=" & varValue
    ELSE:
        ' append (+) to the cell value if it is not empty
        IF objExcelWs.Cells(intRow, intColumn).Value = "" THEN
            objExcelWs.Cells(intRow, intColumn).Formula = "=" & varValue
        ELSE:
            objExcelWs.Cells(intRow, intColumn).Formula = objExcelWs.Cells(intRow, intColumn).Formula & "+" & varValue
        END IF
    END IF
END SUB

SUB ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, strFuelCategoryIdentifier, arrFuels)
    ' 1.A(a)
    ' strFuelCategoryIdentifier was used to find which row
    ' find row index
    DIM intRowIndex : intRowIndex = GetRowIndexAfterCode(objExcelWs, "B", strCode, strFuelCategoryIdentifier)
    IF intRowIndex > 0 THEN
        DIM varConsumption : varConsumption = 0
        DIM varCO2Emission : varCO2Emission = 0
        DIM varCH4Emission : varCH4Emission = 0
        DIM varN2OEmission : varN2OEmission = 0

        DIM strFuel
        FOR EACH strFuel IN arrFuels
            ' get consumption
            varConsumption = varConsumption + GetFuelConsumption(intBranchIndex, "Gigajoule", strFuel)
            ' get emissions
            varCO2Emission = varCO2Emission + GetEmission(intBranchIndex, "Metric Tonne", strCO2, strFuel)
            varCH4Emission = varCH4Emission + GetEmission(intBranchIndex, "Metric Tonne", strCH4, strFuel)
            varN2OEmission = varN2OEmission + GetEmission(intBranchIndex, "Metric Tonne", strN2O, strFuel)
        NEXT

        ' output to excel
        ' consumption (Column C) unit: TJ
        CALL SetExcelCellValue(objExcelWs, intRowIndex, 3, varConsumption / 1000, False)

        ' emissions
        ' CO2 (Column H) unit: kt
        ' CH4 (Column I) unit: kt
        ' N2O (Column J) unit: kt
        CALL SetExcelCellValue(objExcelWs, intRowIndex, 8, varCO2Emission / 1000, False)
        CALL SetExcelCellValue(objExcelWs, intRowIndex, 9, varCH4Emission / 1000, False)
        CALL SetExcelCellValue(objExcelWs, intRowIndex, 10, varN2OEmission / 1000, False)
    ELSE
        ' not found
        ' log it

    END IF
END SUB

SUB ExportSectoralBackgroundData(objExcelWb, intBranchIndex, strCode)
    DIM dictFuels : SET dictFuels = GetFuelsForSectoralBackgroundData()
    ' Define Biomass fuels
    DIM arrBiomassFuels : arrBiomassFuels = dictFuels.Item("arrBiomassFuels")

    ' Define Solid fuels (not including biomass)
    DIM arrSolidFuels : arrSolidFuels = dictFuels.Item("arrSolidFuels")

    ' Define Liquid fuels
    DIM arrLiquidFuels : arrLiquidFuels = dictFuels.Item("arrLiquidFuels")

    ' Define Gaseous fuels
    DIM arrGaseousFuels : arrGaseousFuels = dictFuels.Item("arrGaseousFuels")

    ' Aviation gasoline
    DIM arrAviationGasoline : arrAviationGasoline = dictFuels.Item("arrAviationGasoline")
    ' Jet kerosene
    DIM arrJetKerosene : arrJetKerosene = dictFuels.Item("arrJetKerosene")
    ' Gasoline
    DIM arrGasoline : arrGasoline = dictFuels.Item("arrGasoline")
    ' Diesel oil
    DIM arrDieselOil : arrDieselOil = dictFuels.Item("arrDieselOil")
    ' Liquefied petroleum gases (LPG)
    DIM arrLPG : arrLPG = dictFuels.Item("arrLPG")


    ' get worksheet by code
    DIM strSheetName : strSheetName = GetWsName(strCode)

    DIM objExcelWs : set objExcelWs = objExcelWb.Worksheets(strSheetName)

    ' loop through fuel type
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Liquid fuels", arrLiquidFuels)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Solid fuels", arrSolidFuels)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Gaseous fuels", arrGaseousFuels)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Biomass", arrBiomassFuels)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Aviation gasoline", arrAviationGasoline)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Jet kerosene", arrJetKerosene)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Gasoline", arrGasoline)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Diesel oil", arrDieselOil)
    CALL ExportSectoralBackgroundDataByFuel(objExcelWs, intBranchIndex, strCode, "Liquefied petroleum gases (LPG)", arrLPG)

END SUB
