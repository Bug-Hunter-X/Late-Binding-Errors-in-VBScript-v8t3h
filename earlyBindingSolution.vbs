Early Binding:
```vbscript
On Error GoTo ErrorHandler

Dim objExcel As Object
Set objExcel = CreateObject("Excel.Application")
MsgBox objExcel.Version

Exit Sub

ErrorHandler:
MsgBox "Error accessing Excel: " & Err.Description
End Sub
```
This improved version uses early binding (declaring the variable type) and includes error handling.  It checks for the existence of Excel before accessing its methods.  Remember to add a reference to the Microsoft Excel Object Library in your VBScript environment for early binding to work correctly.