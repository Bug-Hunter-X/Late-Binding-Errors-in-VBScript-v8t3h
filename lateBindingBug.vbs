Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where the availability of components isn't guaranteed.

Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
'This line might fail if Excel is not installed
MsgBox objExcel.Version
```

Early Binding: While early binding is preferable for performance and error detection, it requires explicit type declarations and references, which can be complex in larger projects.