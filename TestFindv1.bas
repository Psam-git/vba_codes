Attribute VB_Name = "Module11"
Sub TestFind()

For k = 1 To Worksheets.Count
    If Worksheets(k).Name = "reflist" Then Worksheets(k).Delete
Next k

Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "reflist"


Set ws = Worksheets("reflist")

ws.Cells(1, 1).Value = "Reference"
ws.Cells(1, 2).Value = "Pagenumber"
ws.Cells(1, 3).Value = "Number_format"

ws.Range("A2").ListNames

Dim sht As Worksheet
Dim LastRow As Long
Dim TheType As String



Set sht = ws
LastRow = sht.Cells(sht.Rows.Count, "B").End(xlUp).Row
LastRow = LastRow

 
    For k = 2 To LastRow
        Range("B" & k).Value = Right(Range("B" & k).Value, Len(Range("B" & k).Value) - 1)
        TheType = ""
        TheType = Range(Range("A" & k).Value).NumberFormat
        
        
        If TheType = "General" Then
            TheType = "General/Character"
        ElseIf TheType = "0" Or TheType = "#,##0.00" Or TheType = "" Then
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

