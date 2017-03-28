Attribute VB_Name = "Module11"
Public Sub TestFind()
    Application.DisplayAlerts = False
    Err.Clear
    On Error Resume Next
    If Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "RefList" Then Worksheets("RefList").Delete
'Delete reflist if exists
    

Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "RefList"

Set ws = Worksheets("RefList")

ws.Cells(1, 1).Value = "References"
ws.Cells(1, 2).Value = "Sheet name"
ws.Cells(1, 3).Value = "Cell Format"

ws.Range("A2").ListNames

Dim sht As Worksheet
Dim LastRow As Long
Dim TheType As String



Set sht = ws
LastRow = sht.Cells(sht.Rows.Count, "B").End(xlUp).Row
LastRow = LastRow - 1




 
    For k = 1 To LastRow
        Range("B" & k).Value = Right(Range("B" & k).Value, Len(Range("B" & k).Value) - 1)
        TheType = ""
        TheType = Range(Range("A" & k).Value).NumberFormat
        
        
        If TheType = "General" Then
            TheType = "General/Character"
        ElseIf TheType = "0" Or TheType = "#,##0.00" Or TheType = "#,##0" Then
            TheType = "Number"
        ElseIf TheType = "m/d/yyyy" Then
            TheType = "Date"
        ElseIf TheType = "0.00%" Then
            TheType = "Percentage"
        ElseIf TheType = "@" Then
            TheType = "Text"
        End If
        
    Range("C" & k).Value = TheType
        
    Next k
    
    
    
End Sub

