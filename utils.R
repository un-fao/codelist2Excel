# Get codelist info

convert_date <- function(date) {
  # Transform to timestamp
  ts <- as.POSIXct(date / 1000, origin = "1970-01-01")
  # Round date
  ts_round <- ceiling_date(ts, unit = "day")
  # Change the format
  date_formated <- format(ts_round, "%Y/%m/%d")
  # Return
  return(date_formated)
}


make_excel <- function(codelist_data, codelist_name){
  # Add codelist name (scheme name)
  codelist_data[, codelist := codelist_name]
  # Find the top concepts
  codelist_children <- unique(trimws(unlist(strsplit(unlist(codelist_data$children), ","))))
  top_concept_codes <- setdiff(codelist_data$code, codelist_children)
  # Add the top concept column
  codelist_data[code %in% top_concept_codes, top := code]
  # Add the code's parent
  parent_children_data <- codelist_data[, .(code, children)]
  parent_children_data <- parent_children_data[, .(children = unlist(strsplit(as.character(children), ", "))), by = code]
  setnames(parent_children_data, c("code", "children"), c("parent", "code"))
  codelist_data <- merge(codelist_data, parent_children_data, by = "code", all.x = TRUE)
  # Remove children column
  codelist_data[, children := NULL]
  # Reorder columns
  wanted_cols <- c("code", "parent", "label_en", "label_fr", "label_es",
                   "label_ru", "label_zh", "label_ar", "description",
                   "start_date", "end_date", "order", "unit", "virtual", "type",
                   "codelist", "top")
  
  existing_cols <- intersect(wanted_cols, names(codelist_data))

  codelist_data <- codelist_data[, ..existing_cols]

  # Return
  return(codelist_data)
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
  
  # Transform for the Excel
  codes_info_codes <- make_excel(codes_info_codes, codelist_id)
  
  # Return
  return(codes_info_codes)
}
