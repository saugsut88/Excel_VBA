Private strImportFilePath As String
Private strImportFileName As String
Private strCurrentfileName As String
Private strCurrentFilePath As String
Private app As New Excel.Application
Private bkImport As Excel.Workbook


Property Get CurrentFileName() As String
    CurrentFileName = strCurrentfileName
End Property

Private Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function browseFilePathOpen() As Boolean
    'open widows file explorer to import raw data file'
    
    On Error GoTo err
    Dim fileExplorer As FileDialog
   
    MsgBox "Select Raw Data File", vbOKOnly, "Data Import"
    
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)
          
    With fileExplorer
        .AllowMultiSelect = False
        '.Filters "Excel Files (*.xlsx, *.xls, *.xlsxm)", "*.xlsx, *.xls, *.xlsxm"
        If .Show = -1 Then 'Any file is selected
            strImportFilePath = .SelectedItems.item(1)
            strImportFileName = GetFilenameFromPath(strImportFilePath)
            app.Visible = False
            Set bkImport = app.Workbooks.Add(strImportFilePath)
            'Workbooks.Open strFilePath
            browseFilePathOpen = True
        Else
            MsgBox "No File Selected"
            'fileChosen = Null ' when cancelled set blank as file path.
            browseFilePathOpen = False
        End If
    End With
err:
    Exit Function
End Function



Sub copyActiveSheet()
    'Copy open file into Financial analyzer and rename raw data'
    Dim fileCheck As New fileChecker
    
    app.Workbooks(1).Worksheets(1).Activate
    app.ActiveSheet.Cells.Select
    app.Selection.copy
    
    Windows(strCurrentfileName).Activate
    ActiveWorkbook.Sheets(1).Name = "Raw Data"
    ActiveWorkbook.Sheets("Raw Data").Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    app.DisplayAlerts = False
    app.Quit
    Set app = Nothing

    

End Sub

Function browseFilePathSaveAs() As Boolean
    'choose where to report'
    'saves but not correct file format
    
    Dim Spreadsheet_Name As Variant
    
    Spreadsheet_Name = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsm), *.xlsm", Title:="Save Financial Record")
    
    If Spreadsheet_Name <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=Spreadsheet_Name
        strCurrentFilePath = Spreadsheet_Name
        strCurrentfileName = GetFilenameFromPath(strCurrentFilePath)
        Application.DisplayAlerts = True
        browseFilePathSaveAs = True
    Else
        browseFilePathSaveAs = False
    End If

End Function
