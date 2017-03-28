Attribute VB_Name = "Module4"
Sub blue()
For Each mycell In Selection
 mycell.Interior.Color = RGB(153, 204, 255)
Next mycell
End Sub
