#' Convert markdown file to Whitehall Publisher (GOV.UK) govspeak markdown
#' format
#' @param file_path Name and location of the *.md file you want to extract from sharepoint
#' @param site Name of sharepoint site you want to save a file to. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @param drive Name of the sharepoint drive you want to save the file to Can be found in the URL of the sharepoint location e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail/RailStats", the drive name is "RailStats". Can also be found using the list_sharepoint_drives() function in this package.
#' @param remove_blocks bool; decision to remove block elements from *.md file.
#'   This includes code blocks and warnings
#' @param sub_pattern bool or vector; decision to increase hashed headers by
#' one level in *.md file.
#'   TRUE will substitute all, FALSE will substitute none, while a vector of
#'   the desired substitution levels allows individual headings to be modified
#'   as required.
#'   e.g. "#" will modify only first level headers.
#' @param page_break string; chooses what page breaks are converted to on
#' Whitehall.
#' If "line", page breaks are replaced with a horizontal rule. If "none"
#' they are replaced with a line break.
#' If "unchanged" they are not removed.
#' @export
#' @import govspeakr
#' @import Microsoft365R
#' @importFrom tools file_ext
#' @name convert_sharepoint_rmd
#' @title Convert standard markdown file to govspeak
convert_sharepoint_rmd <- function(file_path,
                                   site,
                                   drive,
                                   remove_blocks = FALSE,
                                   sub_pattern = TRUE,
                                   page_break = "line") {

  ##Check the path is a .md format
  if(!grep("[.]md$", file_path)){
    stop("Please provide a .md format file.
         Note that this function does not work directly on the .Rmd rmarkdown document.")
  }

  # Get SharePoint location
  site_loc <- get_sharepoint_site(site)
  drive_loc <- site_loc$get_drive(drive)

  # Create temporary file for unconverted output
  tmp_file <- tempfile(fileext = paste0(".", tools::file_ext(file_path)))

  # download file into temporary location
  drive_loc$download_file(file_path, dest = tmp_file, overwrite = TRUE)

  #Read in the raw lines from the md
  md_file <- readLines(tmp_file)

  #Nicely format the image references
  md_file <- generate_image_references(md_file)

  ##Turn md file into plain text
  govspeak_file <- paste(md_file, collapse = "\n")

  govspeak_file <- remove_header(govspeak_file)

  govspeak_file <- convert_callouts(govspeak_file)

  #Remove long strings of hashes/page break indicators
  #and move all headers down one level
  if (page_break == "line") {
    govspeak_file <- gsub("#####|##### <!-- Page break -->",
                          "-----",
                          govspeak_file)
  }else if (page_break == "none") {
    govspeak_file <- gsub("#####|##### <!-- Page break -->",
                          "\n",
                          govspeak_file)
  }else if (page_break == "unchanged") {
    govspeak_file <- govspeak_file
  }

  ##Substitute hashes according to pattern
  govspeak_file <- hash_sub(govspeak_file, sub_type = sub_pattern)

  ##Remove YAML block if requested
  if (remove_blocks) {
    govspeak_file <- remove_rmd_blocks(govspeak_file)
  }

  # Create temporary file for converted output
  tmp_conv_file <- tempfile(fileext = paste0(".", tools::file_ext(file_path)))

  ##Write output as converted file to temporary location
  write(govspeak_file, tmp_conv_file)

  ## upload file to path
  drive_loc$upload_file(src = tmp_conv_file,
                        dest = gsub("\\.md", "_converted\\.md", file_path))

  if(file.exists(gsub("\\.md", "_converted\\.md", file_path))){
    message("File converted successfully")
  }

  # Clean up temp file
  unlink(c(tmp_file, tmp_conv_file))

}
