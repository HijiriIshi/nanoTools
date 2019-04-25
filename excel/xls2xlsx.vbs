'*******************************************************************************
' xls to xlsx Convert
'******************************************************************************* 

Const xlOpenXMLWorkbook = 51 ' .xlsx

' args
set args = WScript.Arguments
fileList = ""
for each arg in args
  fileList = fileList & vbNewLine & arg
next

set fobj = CreateObject("Scripting.FileSystemObject")
for each arg in args
    ext = fobj.GetextensionName(arg)
    if ext <> "xls" then
        Wscript.echo WScript.StdOut.Write("invalid Extension") & vbNewLine & fileList
        WScript.Quit
    end if
next

' convert
set objXlsApp = CreateObject("Excel.Application")
for each path in args
    objXlsApp.Application.Visible = false
    set book = objXlsApp.Application.Workbooks.Open(path)
    book.SaveAs Replace(path, ".xls", ".xlsx"), xlOpenXMLWorkbook
    book.Close
next
objXlsApp.Quit
set objXlsApp = nothing

WScript.StdOut.Write("Finish")
