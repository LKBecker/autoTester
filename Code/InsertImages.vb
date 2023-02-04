Option Explicit
Const BasePath As String = "W:\Pathology\Biochem\LKBecker\Projects\autoTester"

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Sub Amend_Template()
    Dim WS As Worksheet
    Dim Search, Replacement As String
    
    Set WS = ThisWorkbook.Worksheets("Template")
    
    Replacement = InputBox("What analyte is this test script for?", "Analyte script")
    WS.Cells(2, 1).Value = Replacement
        
    Replacement = InputBox("Who is writing the script?", "Analyte script")
    WS.Cells(6, 2).Value = Replacement
    
    Replacement = InputBox("What date are you writing the script?", "Analyte script")
    WS.Cells(5, 2).Value = Replacement

End Sub

Sub Dupe()
    Dim curSheet, scenarioSheet As Worksheet
    Dim NumDupes, D, curScenario, curPTCount As Integer
    Dim TestName, curPtOverride, curScenarioFull, tmpStr As String
    Dim TestLog() As Variant
    
    Const ColumnSampleID As Integer = 10
    Const ColumnSampleTakenDT As Integer = 11
    Const ColumnScenario As Integer = 12
    Const ColumnAuthorisationQueue As Integer = 13
    Const ColumnScenarioID As Integer = 14
    
    Application.ScreenUpdating = False
    
    TestLog = Get2DArray(ThisWorkbook.Worksheets("TestLog").Cells(2, 1))
    Debug.Print "TestLog array size is " & LBound(TestLog) & ":" & UBound(TestLog)
    
    Set scenarioSheet = ThisWorkbook.Worksheets("Scenarios")
    TestName = scenarioSheet.Cells(1, 1).Value
    'Debug.Print TestName
        
    NumDupes = scenarioSheet.Range("R9")
    
    curPtOverride = CStr(scenarioSheet.Cells(4, 6).Value)
    curPTCount = 0
        
    For D = 1 To NumDupes
        curScenario = scenarioSheet.Cells(4 + (D - 1), 1).Value
        Debug.Print "Processing scenario " & curScenario
        If (StrComp(curPtOverride, scenarioSheet.Cells(4 + (D - 1), 6).Value, vbTextCompare) <> 0) Then
            curPtOverride = CStr(scenarioSheet.Cells(4 + (D - 1), 6).Value)
            curPTCount = 1
        Else
            curPTCount = curPTCount + 1
        End If
        
        If (WorksheetExists(TestName & "_" & curScenario, ThisWorkbook)) Then
            GoTo NextIterationD
        End If
        
        Sheets("Template").Copy After:=Sheets(Sheets.Count)
        Set curSheet = Sheets(Sheets.Count)
        curScenario = scenarioSheet.Cells(4 + (D - 1), 1).Value
        curSheet.Name = TestName & "_" & curScenario
        curSheet.Cells(2, 7).Value = CStr(curScenario)
        curSheet.Cells(2, 1).Value = TestName
        curSheet.Cells(15, 2).Value = scenarioSheet.Cells(4 + (D - 1), 3).Value   'Transpose Scenario
        curSheet.Cells(15, 6).Value = scenarioSheet.Cells(4 + (D - 1), 13).Value 'Transpose Expected Queue
        curSheet.Cells(14, 6).Value = scenarioSheet.Cells(4 + (D - 1), 14).Value 'Transpose Expected Autocomment
        curSheet.Cells(18, 6).Value = scenarioSheet.Cells(4 + (D - 1), 15).Value 'Transpose Expected Flags to Validation Behaviour
        curScenarioFull = TestName & "-" & Format(CInt(curPtOverride), "000") & "_" & Format(curPTCount, "000")
        curSheet.Cells(17, 2).Value = curScenarioFull 'Transpose current scenario in Worklist format
        
        'curSheet.Cells(7, 6).Value = curSheet.Cells(7, 6).Value 'Replace Scenario(s) w/ value
        curSheet.Cells(7, 6).Value = GetMatchingValuesAsString(TestLog, ColumnScenarioID, Array(ColumnScenario, ColumnSampleID, ColumnSampleTakenDT), curScenarioFull)
        
        'curSheet.Cells(15, 8).Value = curSheet.Cells(15, 8).Value 'Replace formula for queue(s) w/ value
        curSheet.Cells(15, 8).Value = GetMatchingValuesAsString(TestLog, ColumnScenarioID, Array(ColumnAuthorisationQueue, ColumnSampleID), curScenarioFull)
                
        'curSheet.Cells(16, 2).Value = curSheet.Cells(16, 2).Value 'Replace formula for Sample Number(s) w/ value
        curSheet.Cells(16, 2).Value = GetMatchingValuesAsString(TestLog, ColumnScenarioID, Array(ColumnSampleID), curScenarioFull)
        
NextIterationD:
    Next D
    
    Application.ScreenUpdating = True
        
End Sub

Sub AddImage(ByRef currentSheet As Worksheet, ByVal ImgPath As String, ByVal imgTop As Integer, ByVal imgLeft As Integer)
    ImgPath = Replace(ImgPath, "/", Application.PathSeparator)
    'Debug.Print (ImgPath)
    If Dir(ImgPath) <> "" Then
        currentSheet.Shapes.AddPicture Filename:=ImgPath, LinkToFile:=False, SaveWithDocument:=True, Left:=imgLeft, Top:=imgTop, Width:=-1, Height:=-1
    End If
End Sub

Sub Imageify()
    Dim PicString As String, Analyte As String, PicPath As String
    
    Analyte = ThisWorkbook.Worksheets("Scenarios").Cells(1, 1).Value
    PicPath = BasePath & "/" & Analyte & " Screenshots/"
          
    Dim sheetCounter As Integer
    Dim currentSheet As Worksheet
    For sheetCounter = 1 To ActiveWorkbook.Worksheets.Count
        Set currentSheet = ActiveWorkbook.Worksheets(sheetCounter)
        currentSheet.Activate
        If Left(currentSheet.Name, Len(Analyte)) = Analyte Then
            'Debug.Print ("Sheet " & currentSheet.Name & " matched. Proceeding...")
            Dim DivisorPos, CurrentSheetNumber As Integer
            DivisorPos = InStr(1, currentSheet.Name, "_")
            If DivisorPos <> 0 Then
                'Debug.Print ("Divisor located.")
                CurrentSheetNumber = CInt(Right(currentSheet.Name, Len(currentSheet.Name) - DivisorPos))
                                                
                PicString = PicPath & currentSheet.Cells(17, 2).Value & "_1.png"
                AddImage currentSheet, PicString, currentSheet.Cells(21, 1).Top, currentSheet.Cells(21, 1).Left
                
                PicString = PicPath & currentSheet.Cells(17, 2).Value & "_2.png"
                AddImage currentSheet, PicString, currentSheet.Cells(41, 1).Top, currentSheet.Cells(41, 1).Left
                
                PicString = PicPath & currentSheet.Cells(17, 2).Value & "_1_AuthLogic_1.png"
                AddImage currentSheet, PicString, currentSheet.Cells(4, 12).Top, currentSheet.Cells(4, 12).Left
                
                PicString = PicPath & currentSheet.Cells(17, 2).Value & "_2_AuthLogic_1.png"
                AddImage currentSheet, PicString, currentSheet.Cells(7, 12).Top, currentSheet.Cells(7, 12).Left
                'TODO: LOOP for multi-auth...
            End If
        End If
    Next sheetCounter
End Sub

Sub RemoveAllImages()
    Dim sheetCounter As Integer
    Dim currentSheet As Worksheet
    Dim Analyte As String
    Analyte = ThisWorkbook.Worksheets("Scenarios").Cells(1, 1).Value
    For sheetCounter = 1 To ActiveWorkbook.Worksheets.Count
        Set currentSheet = ActiveWorkbook.Worksheets(sheetCounter)
        If Left(currentSheet.Name, Len(Analyte)) = Analyte Then
            Dim DivisorPos, CurrentSheetNumber As Integer
            DivisorPos = InStr(1, currentSheet.Name, "_")
            If DivisorPos <> 0 Then
                currentSheet.Pictures.Delete
            End If
        End If
    Next sheetCounter
End Sub

'Turns a table (usually a pivot table copy/paste) into an array for faster processing.
'Truncates anything over 10000 rows or columns. But if you have that many values, God help you.
Function Get2DArray(ByVal StartCell As Range) As Variant
    Dim LastLocation As Range, StartRow As Integer, StartColumn As Integer, EndRow As Integer, EndColumn As Integer
    'Search from start of range to end of range -> No interruptions, or from the left?
    StartRow = StartCell.Cells(1).Row
    StartColumn = StartCell.Cells(1).Column
    EndColumn = StartCell.Parent.Cells(StartRow, 10000).End(xlToLeft).Column
    EndRow = StartCell.Parent.Cells(10000, StartColumn).End(xlUp).Row
    Set LastLocation = Range(StartCell.Parent.Cells(StartRow, StartColumn), StartCell.Parent.Cells(EndRow, EndColumn))
    'Convert Range To Array and Return
    If (EndRow = StartRow) Then
        Get2DArray = Array(LastLocation.Value)
        Exit Function
    End If
    Get2DArray = LastLocation.Value
End Function

Function GetMatchingValuesAsString(ByRef InputArray As Variant, ByVal MatchColumn As Integer, ByRef ValueColumns As Variant, _
    ByVal MatchVal As Variant, Optional ByVal Separator As String = vbNewLine) As Variant
    Dim Output As String, EntryCounter As Integer, ValueColumn As Integer
    Output = ""
        
    For EntryCounter = LBound(InputArray) To UBound(InputArray)
        If UBound(ValueColumns) > 0 Then
            If InputArray(EntryCounter, MatchColumn) = MatchVal Then
                For ValueColumn = LBound(ValueColumns) To UBound(ValueColumns)
                    If ValueColumn = LBound(ValueColumns) Then
                         Output = Output & InputArray(EntryCounter, ValueColumns(ValueColumn))
                    ElseIf ValueColumn = UBound(ValueColumns) Then
                        Output = Output & " (" & InputArray(EntryCounter, ValueColumns(ValueColumn)) & ")" & Separator
                    Else
                        Output = Output & " (" & InputArray(EntryCounter, ValueColumns(ValueColumn)) & ")"
                    End If
                Next ValueColumn
            End If
        Else
            If InputArray(EntryCounter, MatchColumn) = MatchVal Then: Output = Output & Separator & InputArray(EntryCounter, ValueColumns(0))
        End If
        'TODO: Shouldn't this be a string match? maybe coerce both to str first? or binary match??
    Next EntryCounter
    If UBound(ValueColumns) = 0 Then
        Output = Right(Output, Len(Output) - Len(Separator)) 'Remove first separator
    End If
GetMatchingValuesAsString = Output
End Function
