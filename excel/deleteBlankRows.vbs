'*******************************************************************************
' delete blank rows
'   EXCELの余計な空白行を削除.1行目と1列目にアンカーをつけて使用.
'******************************************************************************* 

Const xlUp = -4162
Const xlDown = -4121
Const xlToLeft = -4159
Const xlToRight = -4161
Const xlShiftUp = -4162
Const xlShiftToLeft = -4159
Const upperRight = "XFD1"
Const lowerLeft = "A1048576"

' args
set args = WScript.Arguments
fileList = ""
for each arg in args
  fileList = fileList & vbNewLine & arg
next

' 引数のチェック。対象ファイル以外が混ざっている場合終了。
set fobj = CreateObject("Scripting.FileSystemObject")
for each arg in args
    ext = fobj.GetextensionName(arg)
    if ext <> "xlsx" then
        WScript.StdOut.Write(WScript.StdOut.Write("invalid Extension") & vbNewLine & fileList)
        WScript.Quit
    end if
next

' delete
set oXlsApp = CreateObject("Excel.Application")
for each path in args
    oXlsApp.Application.Visible = false
    set book = oXlsApp.Application.Workbooks.Open(path)
    For i = 1 To book.Sheets.Count
        Set objWorksheet = book.Sheets(i)
        If objWorksheet.Visible Then
            objWorksheet.Range(upperRight).Select
            objWorksheet.Range(oXlsApp.Selection, oXlsApp.Selection.End(xlToLeft)).Select
            objWorksheet.Range(oXlsApp.Selection, oXlsApp.Selection.End(xlDown)).Select
            oXlsApp.Selection.Delete xlShiftToLeft
            
            objWorksheet.Range(lowerLeft).Select
            objWorksheet.Range(oXlsApp.Selection, oXlsApp.Selection.End(xlUp)).Select
            objWorksheet.Range(oXlsApp.Selection, oXlsApp.Selection.End(xlToRight)).Select
            oXlsApp.Selection.Delete xlShiftUp
            
            oXlsApp.Range("A1").Select
        End If
    Next
    book.Save
    book.Close
next
oXlsApp.Quit
set oXlsApp = nothing

WScript.StdOut.Write("Finish")
