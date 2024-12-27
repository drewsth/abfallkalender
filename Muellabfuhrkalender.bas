Attribute VB_Name = "Muellabfuhrkalender"
' Common definitions
Const g_sUniqueIds As String = "Eindeutige ID's"
Const g_sId As String = "ID"
Const g_sStreetNames As String = "Straßenname"
Const g_nMaxArraySize As Integer = 50
Const g_sConfig As String = "Config"
Const g_sMain As String = "Straßenindex"
Const g_sRest As String = "Restmüll"
Const g_sBio As String = "Biomüll"
Const g_sGS As String = "GelberSack"
Const g_sGarden As String = "Garten"

' Initialize global values for Config processing
Const g_nColumnYear = 2
Const g_nRowYear = 1

' Initialize global values for Restmüll processing
Const g_nRmColumnStart As Integer = 3
Const g_nRmColumnStop As Integer = 6
' Note that end row will be detected automatically
Const g_nRmRowStart As Integer = 4

' Initialize global values for Biomüll processing
Const g_nBioColumnStart As Integer = 2
Const g_nBioColumnStop As Integer = 4
' Note that end row will be detected automatically
Const g_nBioRowStart As Integer = 4

' Initialize global values for Garden processing
Const g_nColumnGarden = 2
Const g_nRowGardenSpring = 2
Const g_nRowGardenAutmn = 3


Sub Korrektur_Restmuell()
  Call Korrektur(g_sRest, g_nRmColumnStart, g_nRmColumnStop, g_nRmRowStart)
End Sub

Sub Reset_Restmuell()
  Call Reset(g_sRest, g_nRmColumnStart, g_nRmColumnStop, g_nRmRowStart)
End Sub

Sub Korrektur_Biomüll()
  Call Korrektur(g_sBio, g_nBioColumnStart, g_nBioColumnStop, g_nBioRowStart)
End Sub

Sub Reset_Biomüll()
  Call Reset(g_sBio, g_nBioColumnStart, g_nBioColumnStop, g_nBioRowStart)
End Sub

Sub Reset(sTabName, nColumnStart, nColumnStop, nRowStart)
  Dim nCount As Integer
  
  nCount = 0
  nRow = nRowStart
  nColumn = nColumnStop + 1
  Application.Goto ActiveWorkbook.Sheets(sTabName).Cells(nRow, nColumn)
  ActDate = ActiveCell.Value
  Do While IsDate(ActDate)
    nCount = nCount + 1
    nRow = nRow + 1
    Application.Goto ActiveWorkbook.Sheets(sTabName).Cells(nRow, nColumn)
    ActDate = ActiveCell.Value
  Loop
    
  Application.Goto ActiveWorkbook.Sheets(sTabName).Range(Cells(nRowStart, nColumnStop + 1), Cells(nRowStart + nCount - 1, nColumnStop + nColumnStop - nColumnStart + 1))
  Selection.ClearContents
End Sub

Sub Korrektur(sTabName, nColumnStart, nColumnStop, nRowStart)
  Dim ActDate, CorrFromDate, CorrToDate As Date
  Dim nRow, nCorrRow As Integer


  For nColumn = nColumnStart To nColumnStop
    nRow = nRowStart
    Application.Goto ActiveWorkbook.Sheets(sTabName).Cells(nRow, nColumn)
    ActDate = ActiveCell.Value
    Do While IsDate(ActDate)
      nCorrRow = 2
      CorrToDate = ActDate
      Do
        Application.Goto ActiveWorkbook.Sheets("Korrektur").Cells(nCorrRow, 1)
        CorrFromDate = ActiveCell.Value
        If ActDate = CorrFromDate Then
          Application.Goto ActiveWorkbook.Sheets("Korrektur").Cells(nCorrRow, 2)
          CorrToDate = ActiveCell.Value
          Exit Do
        End If
        nCorrRow = nCorrRow + 1
      Loop Until CorrFromDate = ""
      
      'nColumnStop nColumnStop - nColumnStart + 1
      Application.Goto ActiveWorkbook.Sheets(sTabName).Cells(nRow, nColumn + nColumnStop - nColumnStart + 1)
      ActiveCell.Value = CorrToDate
    
      nRow = nRow + 1
      Application.Goto ActiveWorkbook.Sheets(sTabName).Cells(nRow, nColumn)
      ActDate = ActiveCell.Value
    Loop
  Next
End Sub

Sub Export_Data()
  Dim nColumn, nRow, nTopRow As Integer
  Dim nColumnUniqueIds As Integer
  Dim vCellContent As Variant
  Dim nActId As Integer
  Dim bRet As Boolean
  
  nColumnUniqueIds = FindColumnByString(g_sUniqueIds, 1)
  If nColumnUniqueIds < 0 Then Exit Sub

  ' Get to the top row that starts with unique ID's
  nColumn = nColumnUniqueIds
  nRow = 1
  vCellContent = Cells(nRow, nColumn).Value
  Do While Not IsNumeric(vCellContent) Or IsEmpty(vCellContent)
    nRow = nRow + 1
    vCellContent = Cells(nRow, nColumn).Value
  Loop
  nActId = vCellContent
  nTopRow = nRow
  
  ' Loop over all unique ID's to write one file for each ID
  Do
    If nActId > 4 Then
      bRet = IdWriteFile(nActId, nTopRow)
    End If
    nRow = nRow + 1
    Worksheets(g_sMain).Activate
    If IsNumeric(Cells(nRow, nColumn).Value) Then
      nActId = Cells(nRow, nColumn).Value
    Else
      Exit Do
    End If
  Loop Until nActId <= 0
  
  MsgBox ("Alle Dateien erfolgreich exportiert.")
End Sub

Public Function IdWriteFile(nId, nStartRow) As Boolean
  Dim nColumnUniqueIds, nColumnId, nColumnStreetNames As Integer
  Dim nColumnRest, nColumnBio, nColumnGS As Integer
  Dim nColumn, nRow, nStartSearchRow As Integer
  Dim nActId, nActYear As Integer
  Dim vCellContent As Variant
  Dim arsStreetList(g_nMaxArraySize) As String
  Dim nIdx As Integer
  Dim sRestDay, sBioDay, sGSTour As String
  Dim nGSTour As Integer
  Dim ardRest(g_nMaxArraySize), ardBio(g_nMaxArraySize), ardGS(g_nMaxArraySize) As Variant
  Dim sFileName As String
    
  ' Find the column that includes street names
  nColumnStreetNames = FindColumnByString(g_sStreetNames, 1)
  If nColumnStreetNames < 0 Then Exit Function
  
  ' Find the column that includes ID assigned values
  nColumnId = FindColumnByString(g_sId, 1)
  If nColumnId < 0 Then Exit Function
  
  ' Generate list of street names which belong to nActId
  ' and extract search string for date tables
  nActId = nId
  nColumn = nColumnId
  nRow = nStartRow
  nIdx = 0
  sRestDay = ""
  sBioDay = ""
  nGSTour = -1
  Do
    vCellContent = Cells(nRow, nColumn).Value
    If vCellContent = nActId Then
      arsStreetList(nIdx) = Cells(nRow, nColumnStreetNames).Value
      nIdx = nIdx + 1
      ' Extract search string for Restmüll
      If sRestDay = "" And nActId > 7 Then
        sRestDay = Cells(nRow, nColumnStreetNames + 1).Value
        If sRestDay <> "" Then
          sRestDay = sRestDay & " (1/k)"
        Else
          sRestDay = Cells(nRow, nColumnStreetNames + 2).Value
          If sRestDay <> "" Then
            sRestDay = sRestDay & " (2/k)"
          End If
        End If
      End If
      ' Extract search string for Biomüll
      If sBioDay = "" And nActId > 7 Then
        sBioDay = Cells(nRow, nColumnStreetNames + 3).Value
        If sBioDay <> "" Then
          sBioDay = sBioDay & " (1/k)"
        End If
      End If
      ' Extract search string for GelberSack Tour Nummer
      If nGSTour = -1 Then
        nGSTour = Cells(nRow, nColumnStreetNames + 4).Value
        sGSTour = "Tour " & nGSTour
      End If
    End If
    nRow = nRow + 1
  Loop Until Not IsNumeric(vCellContent) Or IsEmpty(vCellContent)
  
  nColumnRest = 0
  nColumnBio = 0
  nColumnGS = 0
  
  ' Generate Restmüll date table
  If sRestDay <> "" Then
    nStartSearchRow = 3
    nIdx = 0
    Worksheets(g_sRest).Activate
    nColumnRest = FindColumnByString(sRestDay, nStartSearchRow)
    If nColumnRest < 0 Then Exit Function
    nRow = nStartSearchRow + 1
    Do
      vCellContent = Cells(nRow, nColumnRest).Value
      If IsDate(vCellContent) Then
        ardRest(nIdx) = vCellContent
        nRow = nRow + 1
        nIdx = nIdx + 1
      End If
    Loop Until IsEmpty(vCellContent)
  End If
  
  ' Generate Biomüll date table
  If sBioDay <> "" Then
    nStartSearchRow = 3
    nIdx = 0
    Worksheets(g_sBio).Activate
    nColumnBio = FindColumnByString(sBioDay, nStartSearchRow)
    If nColumnBio < 0 Then Exit Function
    nRow = nStartSearchRow + 1
    Do
      vCellContent = Cells(nRow, nColumnBio).Value
      If IsDate(vCellContent) Then
        ardBio(nIdx) = vCellContent
        nRow = nRow + 1
        nIdx = nIdx + 1
      End If
    Loop Until IsEmpty(vCellContent)
  End If
  
  ' Generate GelberSack date table
  If nGSTour > 0 Then
    nStartSearchRow = 1
    nIdx = 0
    Worksheets(g_sGS).Activate
    nColumnGS = FindColumnByString(sGSTour, nStartSearchRow)
    If nColumnGS < 0 Then Exit Function
    nRow = nStartSearchRow + 1
    Do
      vCellContent = Cells(nRow, nColumnGS).Value
      If IsDate(vCellContent) Then
        ardGS(nIdx) = vCellContent
        nRow = nRow + 1
        nIdx = nIdx + 1
      End If
    Loop Until IsEmpty(vCellContent)
  End If
  
  ' Write collected data to file
  Application.Goto ActiveWorkbook.Sheets(g_sConfig).Cells(g_nRowYear, g_nColumnYear)
  nActYear = ActiveCell.Value
  
  nIdx = 0
  sFileName = "Abfallkalender_" & nActYear & "_ID-" & nId & ".txt"
  Open sFileName For Output As #1
  
  Print #1, "ID:"; nActId
  Print #1, ""
  
  Print #1, "Straßen-Namen:"
  Do
    Print #1, arsStreetList(nIdx);
    If arsStreetList(nIdx + 1) <> "" Then
      Print #1, ", ";
    End If
    nIdx = nIdx + 1
  Loop Until arsStreetList(nIdx) = ""
  
  Print #1, ""
  Print #1, ""

  If sRestDay <> "" Then
    nIdx = 0
    Print #1, "Restmüll:"
    Do
      Print #1, ardRest(nIdx);
      If ardRest(nIdx + 1) <> "" Then
        Print #1, ", ";
      End If
      nIdx = nIdx + 1
    Loop Until ardRest(nIdx) = ""

    Print #1, ""
    Print #1, ""
  End If

  If sBioDay <> "" Then
    nIdx = 0
    Print #1, "Biomüll:"
    Do
      Print #1, ardBio(nIdx);
      If ardBio(nIdx + 1) <> "" Then
        Print #1, ", ";
      End If
      nIdx = nIdx + 1
    Loop Until ardBio(nIdx) = ""
  
    Print #1, ""
    Print #1, ""
  End If

  If nGSTour > 0 Then
    nIdx = 0
    Print #1, "Gelber Sack:"
    Do
      Print #1, ardGS(nIdx);
      If IsDate(ardGS(nIdx + 1)) Then
        Print #1, ", ";
      End If
      nIdx = nIdx + 1
    Loop Until Not IsDate(ardGS(nIdx))
  
    Print #1, ""
    Print #1, ""
  End If
  
  Print #1, "Gartenabfälle:"
  
  Worksheets(g_sGarden).Activate
  vCellContent = Cells(g_nRowGardenSpring, g_nColumnGarden).Value

  Print #1, vCellContent & " , ";
  
  vCellContent = Cells(g_nRowGardenAutmn, g_nColumnGarden).Value
  
  Print #1, vCellContent;
  
  Close #1
  'arsStreetList(g_nMaxArraySize)
  'ardRest(g_nMaxArraySize), ardBio(g_nMaxArraySize), ardGS(g_nMaxArraySize)
End Function

Public Function FindColumnByString(sString, nStartRow) As Integer
  Dim nColumn, nRow, nMaxColumn As Integer
  Dim vCellContent As Variant
  Dim sMsg As String
  
  nMaxColumn = 15
  nColumn = 1
  nRow = nStartRow
  vCellContent = Cells(nRow, nColumn).Value
  Do While StrComp(vCellContent, sString) And nColumn < nMaxColumn
    nColumn = nColumn + 1
    vCellContent = Cells(nRow, nColumn).Value
  Loop
  
  If nColumn = nMaxColumn Then
    sMsg = "Search String <" & sString & "> not found in row " & nRow & "!"
    MsgBox (sMsg)
    FindColumnByString = -1
  Else
    FindColumnByString = nColumn
  End If
  
End Function

