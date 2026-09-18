#' Convert xls file to xlsx at a specified directory
#'
#' @param dir_convertXLS A directory where the AWWA FWAS reports are located (as .xls files) and it is desired to convert them to .xlsx files.
#'                       This must be formatted in standard R syntax using forward slashes (/) instead of backslashes (\).
#'                       An example may be "C:/Users/---InsertUserName---/Documents/..."
#'
#' @returns This function requires a directory at which there are .xls files (presumably AWWA FWAS) which need to be converted to .xlsx prior to running through the AWWA FWAS "compileR".
#'          This function uses system commands to (1) open the directory in windows explorer, and (2) run a visual basic script (vbs) to convert all .xls files within the directory.
#'          The .vbs script is housed in the installation files of this R-package, but is included below for transparency:
#'
#' ```r
#'  Set objExcel = CreateObject("Excel.Application")
#'
#'  objExcel.Visible = False                              'Do not have screen updating (i.e. not see MS Excel window)
#'  objExcel.DisplayAlerts = False                        'Do not display warnings/errors
#'
#'  strDirPath = WScript.Arguments(0)                     'Read in the directory from R script
#'
#'  Set fso = CreateObject("Scripting.FileSystemObject")  'FileSystemObject to be used as list of files at directory
#'
#'  For Each f In fso.GetFolder(strDirPath).Files         'Loop though all "f" files in the specified directory
#'    If LCase(fso.GetExtensionName(f)) = "xls" Then      'IF the file is extension ".xls"... continue
#'      Set objWorkbook = objExcel.Workbooks.Open(f.Path) 'set the looped object as the "f"-file path
#'      objWorkbook.SaveAs f.Path&"x", 51                 'save "file.xls" as "file.xlsx", file format = 51
#'      objWorkbook.Close True                            'close the workbook
#'    End if
#'  Next
#'
#'  Set objExcel = Nothing
#' ```
#'
#'
#' @export
#'
#' @examples
#' # Reference the directory of example water audits in this R package, which are .xls files
#' dir_ex_xls_audits <- system.file("example_audits_xls", package = "awwafwas")
#'
#' # Run the function which will:
#' #    (1) open the above location in windows explorer and
#' #    (2) convert the xls files to xlsx
#' FWAS_xls2xlsx(dir_ex_xls_audits)
#'
#'
#'


FWAS_xls2xlsx <- function(dir_convertXLS){

  ### Open an explorer window at the specified dir where .xls files are
  browseURL(dir_convertXLS)

  ### Reformat the directory such that it can be passed to cmd
  dir_convertXLS <- gsub("/","\\\\",dir_convertXLS)

  ### Get the directory of the .vbs script to do the converting
  dir_vbs_script <- system.file("support_files", package = "awwafwas")  # Get directory where it is stored
  vbs_script <- paste0(dir_vbs_script,"/xls2xlsx.vbs")        # Manually create full file path using directory info
  vbs_script <- gsub("/","\\\\",vbs_script)

  ### Format and run the system command
  system_command <- paste0("WScript",' "',vbs_script,'" "',dir_convertXLS,'"')
  system(command=system_command, wait=TRUE)

}


