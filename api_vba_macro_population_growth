'*****************************************************************************************************
' Script Purpose: Fetch data from the Data USA API to analyze changes in the United States population.
'
' Input:          Number of years (default is 10).
'
' Output:         Adds the following data to the first sheet:
'                 - Column A: Year
'                 - Column B: Population
'                 - Column C: Percent Increase
'                 Overwrites existing data in these columns.
'
' Requirements:   1. Add JsonConverter to the current project's Modules.
'                 - Source: https://github.com/VBA-tools/VBA-JSON
'                 2. Enable the following references from Tools -> References:
'                  - Microsoft WinHTTP Services
'                  - Microsoft Scripting Runtime
'
' Dependencies:   Active internet connection (to call the Data USA API).
'
' Author:         Kadi Jairus
' Date:           January 1, 2025
' Version:        1.0
'*****************************************************************************************************

Sub Population_macro()
    
    'Set worksheet
    Dim wsheet As Worksheet
    Set wsheet = ActiveSheet

    'Set url
    Dim apiUrl As String, parameters As String
    apiUrl = "https://datausa.io/api/data"
    parameters = "?drilldowns=Nation&measures=Population"
    
    'Send Request
    Dim request As Object
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "Get", apiUrl & parameters
    request.Send
    
    If request.Status <> 200 Then
        MsgBox "Päringu viga: " & request.ResponseText
        Exit Sub
    End If
    
    'Get JSON and parse it
    Dim json As Dictionary
    Set json = JsonConverter.ParseJson(request.ResponseText)
    
    'Get the data
    Dim years As Collection
    Set years = json("data")
    
    'Get needed year range
    Dim yearRange As Long

    If Not IsNumeric(wsheet.Cells(2, 8).Value) Then
        MsgBox "Soovitud aastate arv peab olema number!", vbOKOnly + vbExclamation, "VIGA"
        Exit Sub
    End If
    
    yearRange = Int(wsheet.Cells(2, 8).Value)
    
    If IsEmpty(yearRange) Then
        MsgBox "Palun täpsusta soovitud aastate arv!", vbOKOnly + vbExclamation, "VIGA"
        Exit Sub
    ElseIf yearRange > years.Count Or yearRange < 1 Then
        MsgBox "Sobimatu aastate arv: " & yearRange & "." & vbCrLf _
        & "Näitan viimase " & years.Count & " aasta infot.", vbOKOnly + vbExclamation, "VIGA"
        yearRange = years.Count
    End If
    
    'Create header, clear current data in case yearRange is smaller than before
    Range("A1").Value = "Aasta"
    Range("B1").Value = "Rahvaarv"
    Range("C1").Value = "Kasv (Protsent)"
    
    ' Clear previous data if needed
    lastRow = wsheet.Cells(wsheet.Rows.Count, 2).End(xlUp).Row
    If lastRow > yearRange Then
        wsheet.Range(wsheet.Cells(2, 1), wsheet.Cells(lastRow, 3)).ClearContents
    End If
    
    'Get data until yearRange is reached. Fill worksheet upwards.
    Dim Year As Dictionary
    Dim i As Long
    Dim nextYearsPopulation As Long
    Dim changeInPopulation As Long
    Dim processedYears As Long
    
    processedYears = 0
    i = yearRange + 1
    For Each Year In years
        If processedYears >= yearRange Then Exit For
        
        wsheet.Cells(i, 1) = Year("ID Year")
        wsheet.Cells(i, 2) = Year("Population")
        
        If i <> yearRange + 1 Then
            changeInPopulation = nextYearsPopulation - Year("Population")
            wsheet.Cells(i + 1, 3) = FormatPercent(changeInPopulation / Year("Population"), 2)
        End If
        
        nextYearsPopulation = Year("Population")
        i = i - 1
        processedYears = processedYears + 1
    Next Year
    
    'Adjust column width and right align numbers
    wsheet.Columns(1).Resize(, 3).EntireColumn.AutoFit
    wsheet.Range(wsheet.Cells(2, 1), wsheet.Cells(yearRange + 1, 2)).HorizontalAlignment = xlRight
    
End Sub
