Attribute VB_Name = "Module2"
Sub Button9_Click()
Attribute Button9_Click.VB_ProcData.VB_Invoke_Func = "d\n14"
For Each Mycell In Selection
    Mycell.Value = "Test Pass"
Next Mycell
End Sub
Sub Button10_Click()
Attribute Button10_Click.VB_ProcData.VB_Invoke_Func = "g\n14"
For Each Mycell In Selection
    Mycell.Value = "Test Fail"
Next Mycell
End Sub
Sub Button11_Click()
Attribute Button11_Click.VB_ProcData.VB_Invoke_Func = "f\n14"
For Each Mycell In Selection
    Mycell.Value = "Test Skip"
Next Mycell
End Sub

