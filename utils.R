# Get codelist info

convert_date <- function(date) {
  # Transform to timestamp
  ts <- as.POSIXct(date / 1000, origin = "1970-01-01")
  # Round date
  ts_round <- ceiling_date(ts, unit = "day")
  # Change the format
  date_formated <- format(ts_round, "%d/%b/%Y")
  # Return
  return(date_formated)
}

get_codelist_info <- function(codelist_id){
  # /!\ Initialize client before using /!\
  
  # Get codelist info
  codes_info <- getCodelistInfo(codelist_id)
  
  # Retrieve the codes form the codes info
  codes_info_codes <- data.table(codes_info$codes)
  
  # Unlist the children
  codes_info_codes[, children := paste(unlist(children), collapse = ", "), by = id]
  
  # Rename id/code column
  setnames(codes_info_codes, "id", "code")
  
  # Transform the start and end dates
  codes_info_codes[, start_date := convert_date(as.numeric(start_date))]
  codes_info_codes[, end_date := convert_date(as.numeric(end_date))]
  
  # Return
  return(codes_info_codes)
}
