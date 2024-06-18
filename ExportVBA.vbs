'Excelを起動
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.ScreenUpdating = False
'Excelファイルを開く
Set objWorkbook = objExcel.Workbooks.Open(WScript.Arguments(0))

'VBProject内の各ファイルを取得
For Each objVBComponent in objWorkbook.VBProject.VBComponents
    'ファイル名を取得
    strFileName = objVBComponent.Name
    '対象拡張子決定
    strExtension = ".cls"
    If Left(objVBComponent.Name, 3) = "mod" Then
        strExtension = ".bas"
    End If
    'ファイルの内容を取得
    If objVBComponent.CodeModule.CountOfLines > 0 Then
        strFileContent = objVBComponent.CodeModule.Lines(1, objVBComponent.CodeModule.CountOfLines)
        'テキストファイルに出力
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.CreateTextFile(WScript.Arguments(1) & strFileName & strExtension, True)
        objFile.Write(strFileContent)
        objFile.Close
    End If
Next

'Excelファイルを閉じる
objWorkbook.Close False
Set objWorkbook = Nothing
'Excelを終了する
objExcel.ScreenUpdating = True
objExcel.DisplayAlerts = True
objExcel.Quit
Set objExcel = Nothing
