#' Create Excel Sheets for Semi-Automated Datastream Download
#'
#' @param .data 
#' A Dataframe with company Identifier used by Datastream\cr
#' Notes:\cr
#' - The names of the columns do not matter.\cr
#' - The Dataframe can contain any number of additional columns
#' @param .col_index The Index of the column that contain Datstream Identifiers (1,2,3,...)
#' @param .formula 
#' The whole formula Datastream uses to get your query, the first arguments MUST be exchanged by "{range}" and the formula MUST be enquoted in single-quotes\cr
#' Example:\cr
#' '=DSGRID({range};"VO;UVO;RI";"2004-01-01";"2009-12-31";"D";"RowHeader=true;ColHeader=true;Transpose=true;Code=true;DispSeriesDescription=true;YearlyTSFormat=false;QuarterlyTSFormat=false")'
#' @param .dir Full Path of the folder where the excel sheets should be stored
#' @param .name Prefix for the Excel Files
#' @param .split Numbe rof Companies in one Excel Sheet
#' @param .progress Show Progress?
#'
#' @return Excel Sheets
#' @export
ds_create_excel <- function(.data, .col_index, .formula, .dir, .name, .split = 1000, .progress = TRUE) {
  
  lst_split <- split(.data, ceiling(1:nrow(.data) / .split))
  
  # Prepare Formula ---------------------------------------------------------
  f_prep_form <- function(.data, .formula, .col_index) {
    col_name <- LETTERS[.col_index]
    form  <- gsub("\n|\\s", "", .formula)
    range <- paste0("companies!$", col_name, "$2:$", col_name, "$", nrow(.data) + 1)
    form  <- glue::glue(form)
    return(form)
  }
  lst_form <- purrr::map(lst_split, f_prep_form, .formula, .col_index)
  
  
  if (!dir.exists(.dir)) dir.create(.dir)
  int_pad <-  nchar(length(lst_split))
  
  if (.progress) pb <- progress::progress_bar$new(total = length(lst_split))
  for (i in 1:length(lst_split)) {
    if (.progress) pb$tick()
    
    fil <- file.path(.dir, paste0(.name, "_", stringi::stri_pad_left(i, int_pad, "0"), ".xlsx"))
    tab_data_sheet <- tibble::tibble(tmp = lst_form[[i]])
    tab_comp_sheet <- lst_split[[i]]
    
    openxlsx::write.xlsx(tab_data_sheet, fil, sheetName = "data", col.names = FALSE)
    wb <- openxlsx::loadWorkbook(fil)
    openxlsx::addWorksheet(wb, "companies")
    openxlsx::writeData(wb, "companies", tab_comp_sheet)
    openxlsx::saveWorkbook(wb, fil, TRUE)
  }
  
  
}