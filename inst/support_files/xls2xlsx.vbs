Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = False                              'Do not have screen updating (i.e. not see MS Excel window)
objExcel.DisplayAlerts = False                        'Do not display warnings/errors

strDirPath = WScript.Arguments(0)                     'Read in the directory from R script

Set fso = CreateObject("Scripting.FileSystemObject")  'FileSystemObject to be used as list of files at directory

For Each f In fso.GetFolder(strDirPath).Files         'Loop though all "f" files in the specified directory
  If LCase(fso.GetExtensionName(f)) = "xls" Then      'IF the file is extension ".xls"... continue
    Set objWorkbook = objExcel.Workbooks.Open(f.Path) 'set the looped object as the "f"-file path
    objWorkbook.SaveAs f.Path&"x", 51                 'save "file.xls" as "file.xlsx", file format = 51
    objWorkbook.Close True                            'close the workbook
  End if                            
Next

Set objExcel = Nothing