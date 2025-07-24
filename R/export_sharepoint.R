
#' Export an R object as a CSV or XLSX file to a specified Sharepoint location
#' @name export_sharepoint
#' @param x data object in R
#' @param site Name of sharepoint site you want to save a file to. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @param drive Name of the sharepoint drive you want to save the file to Can be found in the URL of the sharepoint location e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail/RailStats", the drive name is "RailStats". Can also be found using the list_sharepoint_drives() function in this package.
#' @param dest_path Name and location of the file path you want to save to in sharepoint. Includes the extension and any folders before the drive name
#' @export
#' @importFrom data.table fwrite
#' @importFrom writexl write_xlsx
#' @importFrom tools file_ext
#' @import Microsoft365R

export_sharepoint <- function(x, site, drive, dest_path){

  site_loc <- get_sharepoint_site(site)
  drive_loc <- site_loc$get_drive(drive)

  ##Check your filename has a usable extension
  if(!grepl("[.](xlsx|csv)$", dest_path)){
    stop("Please make sure your file name includes the file extension, and is either xlsx or csv")
  }

  ##Check your folder exists
  if(grepl("[/]", dest_path)){
    ##Folder without file name
    folder <- gsub("(.*[/]).*$", "\\1", dest_path)
    drive_exists <- try(drive_loc$get_item(folder), silent = TRUE)

    ##Check if folder exists, offer to make it if not
    if("try-error" %in% class(drive_exists) ) {
      message("Folder location ", folder, " not found")
      response <- tolower(readline(prompt = "Do you want to create a new folder? (yes/no): "))

      if (response %in% c("yes", "y")) {
        message("Creating folder: ", folder)

        # Create the folder — assuming drive is defined
        drive_loc$create_folder(folder)
        message("Folder created successfully.")
      } else {
        stop(folder, " not found and not created")
      }
    }
  }
  ##Check file extension
  ext <- tolower(tools::file_ext(dest_path))

  if(ext == "xlsx"){

    #Turn your R object into a file
    temp <- tempfile(fileext = ext)
    writexl::write_xlsx(x, temp)

  } else if(ext == "csv"){

    #Turn your R object into a file
    temp <- tempfile(fileext = ".csv")
    data.table::fwrite(x, temp)

  } else{
    stop("Please make sure your file name includes the file extension, and is either xlsx or csv")
  }

  #Save that file on sharepoint
  drive_loc$upload_file(src = temp,
                        dest = dest_path)

  message("File '", dest_path, "' uploaded successfully")

  # Clean up temporary file
  unlink(temp)

}

#' Export any file to a specified Sharepoint location
#' @name export_sharepoint_file
#' @param x Name and location of a local file.
#' @param site Name of sharepoint site you want to save a file to. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @param drive Name of the sharepoint drive you want to save the file to Can be found in the URL of the sharepoint location e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail/RailStats", the drive name is "RailStats". Can also be found using the list_sharepoint_drives() function in this package.
#' @param dest_path Name and location on sharepoint you want the file to be saved to. Includes the extension and any folders after the drive name
#' @export
#' @import Microsoft365R

export_sharepoint_file <- function(x, site, drive, dest_path){

  site_loc <- get_sharepoint_site(site)

  drive_loc <- site_loc$get_drive(drive)

  ##Check your file exists
  if(!file.exists(x)){
    stop("File ", x, " not found")
  }

  ##Check your folder exists
  ##Skip if no folder
  if(grepl("[/]", dest_path)){
    ##Folder without file name
    folder <- gsub("(.*[/]).*$", "\\1", dest_path)
    drive_exists <- try(drive_loc$get_item(folder), silent = TRUE)

    ##Check if folder exists, offer to make it if not
    if("try-error" %in% class(drive_exists) ) {
      message("Folder location ", folder, " not found")
      response <- tolower(readline(prompt = "Do you want to create a new folder? (yes/no): "))

      if (response %in% c("yes", "y")) {
        message("Creating folder: ", folder)

        # Create the folder — assuming drive is defined
        drive_loc$create_folder(folder)
        message("Folder created successfully.")
      } else {
      stop(folder, " not found and not created")
      }
    }
  }

  #Save that file on sharepoint
  tryCatch({
    drive_loc$upload_file(src = x, dest = dest_path)
    message("File '", dest_path, "' uploaded successfully")
    },
      error = function(e) {
    stop("Failed to upload file '", dest_path, "': ", e$message)
  })

}


#' Upload a local folder to SharePoint
#' @name export_sharepoint_folder
#' @param local_folder Path to the local folder to upload
#' @param site SharePoint site name (e.g. "Rail")
#' @param drive SharePoint drive name (e.g. "RailStats")
#' @param dest_path Destination path in SharePoint drive (e.g. "Reports/July2025"). Defaults to NULL, which will put the folder in the same location as it is locally.
#' @export
#' @import Microsoft365R

export_sharepoint_folder <- function(local_folder, site, drive, dest_path = NULL) {

  #If folder is NULL, set to same as local
  if(is.null(dest_path)){
    dest_path <- local_folder
  }
  # Check local folder exists
  if (!dir.exists(local_folder)) {
    stop("Local folder '", local_folder, "' does not exist.")
  }

  # Get SharePoint location and drive
  site_loc <- get_sharepoint_site(site)
  drive_loc <- site_loc$get_drive(drive)

  # Upload folder to destination
  tryCatch({
    drive_loc$upload_folder(src = local_folder, dest = dest_path, recursive = TRUE)
    message("Folder '", local_folder, "' uploaded successfully")
  }, error = function(e) {
    stop("Failed to upload folder: ", e$message)
  })
}
