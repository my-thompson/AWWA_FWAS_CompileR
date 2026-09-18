#' AWWA Free Water Audit Software (FWAS) Report Compiler
#'
#' @param path_reports A directory where the AWWA FWAS reports are located. This must be formatted in standard R syntax using forward slashes (/) instead of backslashes (\)
#' @param path_save Optional. The directory where you wish to save the compiled data as a MS Excel workbook.
#' @param filename Optional. The name of the file should you wish to save one.
#'
#' @returns This function returns a data frame containing the summary data extracted from any version of the AWWA Free Water Audit Software (i.e. 6.1/6.0, 5.0, 4.2/4.1).
#'          Optionally, the function can export the generated data frame as a Microsoft Excel file by specifying a directory using the "path_save" argument.
#'          Unless specified using the "filename" argument, the saved Microsoft Excel file will be named with using the following convention "AWWA_FWAS_compiled_(timestamp).xslx".
#'
#' @export
#'
#' @examples
#' # Reference the directory of example water audits in this R package
#' example_audits <- system.file("example_audits_xlsx", package = "awwafwas")
#'
#' # Run that list of audits through the compiler function
#' # Leaving the "path_save" as NULL, the function does not save the output file.
#' # Specifying the "path_save" as a directory will save the output file at that location (MS Excel).L
#' example_audits_compiled <- FWAS_compileR(path_reports = example_audits,
#'                                          path_save    = NULL,
#'                                          filename     = NULL)
#'
#'

FWAS_compileR <- function(path_reports, path_save = NULL, filename = NULL){


  { ####### CHECK THAT SPECIFIED DIRECTORIES EXIST ----------------------------+

    if(dir.exists(path_reports)==FALSE){
      stop('While using FWAS_compileR(): Directory for "path_report" does not exist or has incorrect syntax.')
    }
    if(!is.null(path_save)){
      if(dir.exists(path_save)==FALSE){
        stop('While using FWAS_compileR(): Directory for "path_save" does not exist or has incorrect syntax.')
      }
    }

  } ### CHECK THAT SPECIFIED DIRECTORIES EXIST-----------------------------+
  { ####### LOAD THE INDEXING FILES -------------------------------------------+

    ### Specify the index file name
    nam_indexfile <- "DRBC_FWAS_CompilerIndex_V6.1_250425.xlsx"

    ### Grab the directory location of where the files were installed with the package
    dir_indexfile <- system.file("support_files", package = "awwafwas")

    ### Read in the necessary data from the
    compiler_data_index      <- openxlsx::read.xlsx(xlsxFile = paste0(dir_indexfile, "/", nam_indexfile), sheet="Index", sep.names = " ")       # Contains the list of data to grab, and where it is located in the FWAS
    compiler_data_IDGranges  <- openxlsx::read.xlsx(xlsxFile = paste0(dir_indexfile, "/", nam_indexfile), sheet="IDG_Ranges", sep.names = " ")  # Contains a list of ranges for IDG answers to find limiting questions


  } ### LOAD THE INDEXING FILES -------------------------------------------+
  { ####### CREATE A TEMPLATE TO LOAD DATA INTO (BASED ON v6) -----------------+

    compiled_data_template           <- data.frame(matrix(ncol=length(compiler_data_index$Item), nrow=1, data=""))      # Create a template data_frame the size of the headers list
    colnames(compiled_data_template) <- compiler_data_index$VariableName                                                # Name each variables based on header name
    compiled_data_template[]         <- lapply(compiled_data_template, as.character)                                    # Initially force everything to a data type = "character"

    ### Create a list of MS Excel column names to be used as data.frame col names
    MSE_cols <- c(LETTERS, paste0("A",LETTERS), paste0("B",LETTERS))

  } ### CREATE A TEMPLATE TO LOAD DATA INTO (BASED ON v6) -----------------+
  { ####### NAVIGATE TO THE AUDITS IN A FOLDER --------------------------------+

    file_loc <- path_reports                                # A folder containing all the audits to compile
    file_list <- list.files(file_loc, full.names = TRUE)    # List all the files in the directory
    file_list <- file_list[grepl("xlsx", file_list)]        # Drop all non-xlsx from list

  } ### NAVIGATE TO THE AUDITS IN A FOLDER --------------------------------+
  { ####### COMPILER
    ### i-Loop: Cycling down the rows of data (i.e. different audit files)
    for(i in 1:length(file_list)){
      tryCatch(
        {
          ### Create a blank row of data for each new loop of "i"
          compiled_data_template[i,] <- ""

          ### Version-specific steps
          { ####### Figure out what version of the AWWA FWAS is being referenced.
            #
            wbk_sheets <- readxl::excel_sheets(path=file_list[i])
            if("Instructions"%in%wbk_sheets){
              version <- readxl::read_excel(path=file_list[i], sheet="Instructions", na="", range="B2", col_names = "ver")[[1]]
            }else{
              version <- readxl::read_excel(path=file_list[i], sheet="Start Page", na="", range="B2", col_names = "ver")[[1]]
            }
            version_simple <- substr(version, nchar(version)-2, nchar(version))
            #
          } ### Figure out what version of the AWWA FWAS is being referenced
          { ####### VERSION 6.1 COMPILER
            if(version_simple=="6.1"){
              ### Choose which list of fields & locations to use based on FWAS version
              ### Rename the columns to be standard for any FWAS version
              index_loop <- compiler_data_index[,c(1:8)]
              colnames(index_loop) <- c("Item","VariableName","ColType","Worksheet","Range","Data_Frame","Row","Column")
              index_loop <- index_loop[!is.na(index_loop$Range),]

              ### Read in the data from each worksheet into individual dataframes
              ws_StartPage               <- readxl::read_excel(path=file_list[i], sheet="Start Page",               na="", range="A1:AI100", col_names = MSE_cols[1:35])
              ws_Worksheet               <- readxl::read_excel(path=file_list[i], sheet="Worksheet",                na="", range="A1:AB100", col_names = MSE_cols[1:28])
              ws_InteractiveDataGrading  <- readxl::read_excel(path=file_list[i], sheet="Interactive Data Grading", na="", range="A1:E1000", col_names = MSE_cols[1:5])
              ws_Dashboard               <- readxl::read_excel(path=file_list[i], sheet="Dashboard",                na="", range="A1:BS100", col_names = MSE_cols[1:71])
              ws_Notes                   <- readxl::read_excel(path=file_list[i], sheet="Notes",                    na="", range="A1:H100",  col_names = MSE_cols[1:8])
              ws_CarbonCalculations      <- readxl::read_excel(path=file_list[i], sheet="Carbon Calculations",      na="", range="A1:H100",  col_names = MSE_cols[1:8])

              ### j-Loop: Run a loop through the list of compiler parameters on "index_loop"
              ### Reference data using (1) Data_frame (worksheet name), (2) Row, and (3) Column
              for(j in 1:length(index_loop$Item)){
                compiled_data_template[i,index_loop$Item[j]] <- get(index_loop$Data_Frame[j])[index_loop$Row[j],index_loop$Column[j]]
              }

              ### k-Loop: Run a loop to get all the "limiting" variables from the IDG
              ### Need to subset the possible range of data, find rows with "limiting" flag,
              ### and concatenate the respective question reference IDs into a single value
              for(k in 1:19){
                test <- ws_InteractiveDataGrading[compiler_data_IDGranges$Row_Start[k]:compiler_data_IDGranges$Row_End[k],]
                compiled_data_template[i,which(colnames(compiled_data_template) %in% compiler_data_IDGranges$VariableName[k])] <- paste0(test$B[test$E%in%"Limiting"], collapse=", ")
              }

              ### Concatenate the date range for the audit from the start & end dates
              Audit_Period_Start_Date <- format(as.Date(as.numeric(ws_StartPage[20,28]), origin = "1899-12-30"), "%m/%d/%y")
              Audit_Period_End_Date   <-  format(as.Date(as.numeric(ws_StartPage[21,28]), origin = "1899-12-30"), "%m/%d/%y")
              compiled_data_template$`Reporting Period`[i] <- paste0(Audit_Period_Start_Date," - ",Audit_Period_End_Date)


            }
          } ### VERSION 6.1 COMPILER
          { ####### VERSION 6   COMPILER
            if(version_simple=="6.0"){
              ### Choose which list of fields & locations to use based on FWAS version
              ### Rename the columns to be standard for any FWAS version
              index_loop <- compiler_data_index[,c(1:3,9:13)]
              colnames(index_loop) <- c("Item","VariableName","ColType","Worksheet","Range","Data_Frame","Row","Column")
              index_loop <- index_loop[!is.na(index_loop$Range),]

              ### Read in the data from each worksheet into individual dataframes
              ws_StartPage               <- readxl::read_excel(path=file_list[i], sheet="Start Page",               na="", range="A1:AI100", col_names = MSE_cols[1:35])
              ws_Worksheet               <- readxl::read_excel(path=file_list[i], sheet="Worksheet",                na="", range="A1:AB100", col_names = MSE_cols[1:28])
              ws_InteractiveDataGrading  <- readxl::read_excel(path=file_list[i], sheet="Interactive Data Grading", na="", range="A1:E1000", col_names = MSE_cols[1:5])
              ws_Dashboard               <- readxl::read_excel(path=file_list[i], sheet="Dashboard",                na="", range="A1:BS100", col_names = MSE_cols[1:71])
              ws_Notes                   <- readxl::read_excel(path=file_list[i], sheet="Notes",                    na="", range="A1:H100",  col_names = MSE_cols[1:8])

              ### j-Loop: Run a loop through the list of compiler parameters on "index_loop"
              ### Reference data using (1) Data_frame (worksheet name), (2) Row, and (3) Column
              for(j in 1:length(index_loop$Item)){
                compiled_data_template[i,index_loop$Item[j]] <- get(index_loop$Data_Frame[j])[index_loop$Row[j],index_loop$Column[j]]
              }

              ### k-Loop: Run a loop to get all the "limiting" variables from the IDG
              ### Need to subset the possible range of data, find rows with "limiting" flag,
              ### and concatenate the respective question reference IDs into a single value
              for(k in 1:19){
                test <- ws_InteractiveDataGrading[compiler_data_IDGranges$Row_Start[k]:compiler_data_IDGranges$Row_End[k],]
                compiled_data_template[i,which(colnames(compiled_data_template) %in% compiler_data_IDGranges$VariableName[k])] <- paste0(test$B[test$E%in%"Limiting"], collapse=", ")
              }

              ### Concatenate the date range for the audit from the start & end dates
              Audit_Period_Start_Date <- format(as.Date(as.numeric(ws_StartPage[20,28]), origin = "1899-12-30"), "%m/%d/%y")
              Audit_Period_End_Date   <-  format(as.Date(as.numeric(ws_StartPage[21,28]), origin = "1899-12-30"), "%m/%d/%y")
              compiled_data_template$`Reporting Period`[i] <- paste0(Audit_Period_Start_Date," - ",Audit_Period_End_Date)


            }
          } ### VERSION 6   COMPILER
          { ####### VERSION 5   COMPILER
            if(version_simple=="5.0"){
              ### Choose which list of fields & locations to use based on FWAS version
              ### Rename the columns to be standard for any FWAS version
              index_loop <- compiler_data_index[,c(1:3,14:18)]
              colnames(index_loop) <- c("Item","VariableName","ColType","Worksheet","Range","Data_Frame","Row","Column")
              index_loop <- index_loop[!is.na(index_loop$Range),]

              ### Read in the data from each worksheet into individual dataframes
              ws_Instructions          <- readxl::read_excel(path=file_list[i], sheet="Instructions",           na="", range="A1:AI100", col_names = MSE_cols[1:35])
              ws_ReportingWorksheet    <- readxl::read_excel(path=file_list[i], sheet="Reporting Worksheet",    na="", range="A1:AB100", col_names = MSE_cols[1:28])
              ws_PerformanceIndicators <- readxl::read_excel(path=file_list[i], sheet="Performance Indicators", na="", range="A1:Q47",   col_names = MSE_cols[1:17])
              ws_Dashboard             <- readxl::read_excel(path=file_list[i], sheet="Dashboard",              na="", range="A1:Q38",   col_names = MSE_cols[1:17])
              ws_Comments              <- readxl::read_excel(path=file_list[i], sheet="Comments",               na="", range="A1:H100",  col_names = MSE_cols[1:8])

              ### j-Loop: Run a loop through the list of compiler parameters on "index_loop"
              ### Reference data using (1) Data_frame (worksheet name), (2) Row, and (3) Column
              for(j in 1:length(index_loop$Item)){
                compiled_data_template[i,index_loop$Item[j]] <- get(index_loop$Data_Frame[j])[index_loop$Row[j],index_loop$Column[j]]
              }
            }
          } ### VERSION 5   COMPILER
          { ####### VERSION 4   COMPILER
            if(version_simple%in%c("4.1","4.2")){
              ### Choose which list of fields & locations to use based on FWAS version
              ### Rename the columns to be standard for any FWAS version
              index_loop <- compiler_data_index[,c(1:3,19:23)]
              colnames(index_loop) <- c("Item","VariableName","ColType","Worksheet","Range","Data_Frame","Row","Column")
              index_loop <- index_loop[!is.na(index_loop$Range),]

              ### Read in the data from each worksheet into individual dataframes
              ws_Instructions          <- readxl::read_excel(path=file_list[i], sheet="Instructions",           na="", range="A1:AI100", col_names = MSE_cols[1:35])
              ws_ReportingWorksheet    <- readxl::read_excel(path=file_list[i], sheet="Reporting Worksheet",    na="", range="A1:AB200", col_names = MSE_cols[1:28])
              ws_GradingAssessment     <- readxl::read_excel(path=file_list[i], sheet="GradingAssessment",      na="", range="A1:G22",  col_names = MSE_cols[1:7])

              ### j-Loop: Run a loop through the list of compiler parameters on "index_loop"
              ### Reference data using (1) Data_frame (worksheet name), (2) Row, and (3) Column
              for(j in 1:length(index_loop$Item)){
                compiled_data_template[i,index_loop$Item[j]] <- get(index_loop$Data_Frame[j])[index_loop$Row[j],index_loop$Column[j]]
              }
            }
          } ### VERSION 4   COMPILER

          ### Non-version specific compiled information (applies to all versions of the FWAS)
          compiled_data_template$`Original Version`[i]      <- version_simple                                                                                    # Specify simple version number
          compiled_data_template$FileName[i]                <- file_list[i]                                                                                      # Add the file name
          compiled_data_template$`Compiled Date / Time`[i]  <- paste0(substr(as.POSIXct(Sys.time(), format = "%Y-%m-%d %H:%M:%OS", tz = "UTC"), 1,22), " UTC")   # Add the date and time-stamp
          # compiled_data_template$Compiler[i]                <- Sys.getenv("USERNAME")                                                                            # The person using the compiler to compile

          ### Check warnings
          compiled_data_template$`Warning (no score)`[i] <- is.na(as.numeric(compiled_data_template$`Water Audit Data Validity Score`[i]))  # Test if the data for DVS is numeric
          compiled_data_template$`Warning (neg loss)`[i] <- as.numeric(compiled_data_template$`Water Losses`[i]) < 0                        # Test if the data for water loss is < 0

          ### Print loop number
          print(paste0("File ",i,": Complete (",file_list[i],")"))
        },
        ### Error Handler (create a list)
        error = function(e){
          Compiler_errors <<- paste0("LOOP: ", i, "; FILE: ", file_list[i],"; ERROR: ", conditionMessage(e))  # Retain the errors in a data.frame
          print(paste0("File ", i, ": Error (",file_list[i],")"))                                             # Print loop number (error message)
        }
      )
    }

    ### Make sure that numeric fields are stored as numbers
    ### This also usually creates many warnings of NA introduced by coercion" - so just suppress those
    suppressWarnings({
    compiled_data_template[,compiler_data_index$Item[compiler_data_index$ColType%in%"numeric"]] <-
      lapply(compiled_data_template[,compiler_data_index$Item[compiler_data_index$ColType%in%"numeric"]], as.numeric)
    })

    ### Adjust the way that warnings are stored (e.g. no "FALSE", just leave blank)
    compiled_data_template$`Warning (no score)`[compiled_data_template$`Warning (no score)`%in%FALSE] <- "--"
    compiled_data_template$`Warning (neg loss)`[compiled_data_template$`Warning (neg loss)`%in%FALSE] <- "--"


  } ### COMPILER ----------------------------------------------------------+
  { ####### SAVE FILE (optional) ----------------------------------------------+
    if(!is.null(path_save)){

      ### If saving, open that location in windows exploere
      browseURL(path_save)

      ### If specifying a filename, use that, otherwise default
      if(!is.null(filename)){
        full_path_save <- paste0(path_save,"/",filename)
      }else{
        full_path_save <- paste0(path_save,"/AWWA_FWAS_compiled_",format(Sys.time(), "%Y%m%d%H%M%OS"),".xlsx")
      }

      ### Create header style and save the file using {openxslx}
      hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize = 12, fontName = "Arial Narrow", fgFill = "#4F80BD")
      openxlsx::write.xlsx(compiled_data_template, file=full_path_save, rowNames=FALSE, headerStyle = hs, firstActiveCol=1, firstActiveRow=1, withFilter = TRUE)
    }
  } ### SAVE FILE (optional) ----------------------------------------------+


  ### Return the summary data frame
  return(compiled_data_template)


}





