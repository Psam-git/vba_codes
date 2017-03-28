Attribute VB_Name = "Module1"
Public Sub TestFind()

Set ws = Worksheets.Add(after:=Worksheets(Worksheets.Count))


ws.Range("A1").ListNames

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
        ElseIf TheType = "0" Then
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

