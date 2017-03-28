Attribute VB_Name = "Module5"
Sub UpdateCellReference()

    Dim mycell As Range
    Dim MyStr   As String

    For Each mycell In Selection
        MyStr = MyStr & mycell.Address(False, False) & "SUM"
    Next mycell

    AllocateNamedRange ThisWorkbook, MyStr, "='" & Selection.Parent.Name & "'!" & Selection.Address, "A1"
End Sub

Public Sub AllocateNamedRange(Book As Workbook, sName As String, sRefersTo As String, Optional ReferType = "R1C1")
    With Book
        If NamedRangeExists(Book, sName) Then .Names(sName).Delete
            If ReferType = "R1C1" Then
                .Names.Add Name:=sName, RefersToR1C1:=sRefersTo
        ElseIf ReferType = "A1" Then
                .Names.Add Name:=sName, RefersTo:=sRefersTo
        End If
    End With
End Sub

Public Function NamedRangeExists(Book As Workbook, sName As String) As Boolean
    On Error Resume Next
        NamedRangeExists = Book.Names(sName).Index <> (Err.Number = 0)
    On Error GoTo 0
End Function
