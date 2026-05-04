# Get codelist info

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
  
  # Return
  return(codes_info_codes)
}
