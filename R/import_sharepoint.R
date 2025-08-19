#' Import a file from a specified Sharepoint location
#' @name import_sharepoint
#' @param file_path Name and location of the file you want to extract from sharepoint. Includes the extension and any folders after the drive name
#' @param site Name of sharepoint site you want to take a file from. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @param drive Name of the sharepoint drive the file sits on. Can be found in the URL of the sharepoint location e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail/RailStats", the drive name is "RailStats". Can also be found using the list_sharepoint_drives() function in this package.
#' @param destination Folder and filename location to download the file to. Defaults to NULL, which saves the file with it's current filename in your current working directory.
#' @export
#' @import Microsoft365R
#' @importFrom readxl read_excel
#' @importFrom data.table fread

import_sharepoint <- function(site, drive, file_path, destination = NULL) {

  site <- get_sharepoint_site(site)
  drive <- site$get_drive(drive)

  # Get file metadata to check size
  item <- drive$get_item(file_path)
  file_size_mb <- as.numeric(item$properties$size) / 1e6

  if (file_size_mb > 500) {
    stop(sprintf("File is %.1f MB, which exceeds the limit of 500MB. Please use a GCP bucket to store larger files", file_size_mb))
  }

  # Use filename only if destination not provided
  if (is.null(destination)) {
    destination <- basename(file_path)
  }

  # Download the file
  drive$download_file(src = file_path, dest = destination)
  message(sprintf("File downloaded to '%s' (%.1f MB)", destination, file_size_mb))
}


#' Read a CSV or Excel file directly from SharePoint
#' @param site SharePoint site name
#' @param drive Drive name within SharePoint
#' @param file_path Path to the file on SharePoint (including filename and extension)
#' @param use_openxlsx Logical. If TRUE, uses the openxlsx package to read Excel files. Defaults to FALSE, which uses readxl.
#' @param ... Additional arguments passed to read.csv() or read_excel() functions
#' @return A data.frame or tibble
#' @import Microsoft365R
#' @importFrom readxl read_excel
#' @importFrom openxlsx read.xlsx
#' @importFrom data.table fread
#' @importFrom tools file_ext
#' @export
read_sharepoint_file <- function(site, drive, file_path, use_openxlsx = FALSE, ...) {
  # Get SharePoint location
  site_loc <- get_sharepoint_site(site)
  drive_loc <- site_loc$get_drive(drive)
  ##Get file extension
  ext <- tolower(tools::file_ext(file_path))
  # Download file to a temp location
  tmp_file <- tempfile(fileext = paste0(".", ext))
  drive_loc$download_file(file_path, dest = tmp_file, overwrite = TRUE)

  # Decide how to read based on file extension
  if (ext == "csv") {

    df <- data.table::fread(tmp_file, ...)

  } else if (ext == "xlsx") {

    if(use_openxlsx){

    df <- openxlsx::read.xlsx(tmp_file, ...)

    } else {

    df <- readxl::read_excel(tmp_file, ...)
  }

  } else {

    stop("Unsupported file type: ", ext)

  }

  # Clean up temp file
  unlink(tmp_file)

  return(df)
}
