Attribute VB_Name = "merge_sheets"
Option Explicit

Dim i As Integer
Sub merge_sheets()

    info
    break_links

End Sub
Sub info()

    Dim anserw As Integer
    Dim Cell As Range
    Dim selectedCell As Range
    
    Do
        anserw = MsgBox("Do you want to merge all open workbooks by this sheet: " & (ActiveSheet.Name), vbQuestion + vbYesNo + vbDefaultButton2, "Choose Sheet Name")
        
        If anserw = vbYes Then
            merge
            Exit Do
        ElseIf anserw = vbNo Then
            Set selectedCell = Application.InputBox("Select a cell on sheet to merge by", Type:=8)
            If selectedCell Is Nothing Then
                MsgBox "No cell selected. Try again"
            Else
            
            End If
        Else
            MsgBox "Invalid selection. Try again"
        End If
        
    Loop While anserw = vbNo
    
    
End Sub
Sub merge()
    
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim LastRow_1 As Long
    Dim FirstRow_1 As Long
    Dim Data_Sheet As Worksheet
    Dim Final_dest As Worksheet
    Dim Copy_Range As Range
    Dim Copy_Headers As Range
    Dim Name As String
    
    On Error Resume Next
    
    Name = ActiveSheet.Name
    
    Worksheets("Merged").Delete
    Sheets.Add Before:=ActiveSheet
    ActiveSheet.Name = "Merged"
    Application.DisplayAlerts = True
    
    Final_dest.Cells.Clear
    
    For i = 1 To Application.Workbooks.Count
        Set Data_Sheet = Workbooks(i).Worksheets(Name)
        Set Final_dest = ActiveWorkbook.Worksheets("Merged")
        
        With Data_Sheet
            LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
            Set Copy_Range = .Cells(2, 1).Resize(LastRow, LastColumn)
            Set Copy_Headers = .Cells(1, 1).Resize(1, LastColumn)
        End With
        
        With Final_dest
            LastRow_1 = .Cells(.Rows.Count, 1).End(xlUp).Row
            FirstRow_1 = .Cells(1, 1).End(xlToLeft).Column
            
            Copy_Headers.Copy Destination:=.Cells(FirstRow_1, 1)
            Copy_Range.Copy Destination:=.Cells(LastRow_1 + 1, 1)
        End With

        Set Copy_Range = Nothing
        Set Final_dest = Nothing
        Set Data_Sheet = Nothing
        Set Copy_Headers = Nothing
    Next i
    
    With Worksheets("Merged")
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        .Cells(1, 1).Select
        .Cells(1, 1).AutoFilter
    End With
    
End Sub
Sub break_links()

    Dim External_Links As Variant
    Dim x As Long

    External_Links = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks)
    
    If Not (IsEmpty(External_Links)) Then
        For x = 1 To UBound(External_Links)
            ActiveWorkbook.BreakLink External_Links(x), xlLinkTypeExcelLinks
        Next x
    End If
    
End Sub
