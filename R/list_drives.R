#' List the sharepoint drives available in the specified site, useful for finding their names
#' @name list_sharepoint_drives
#' @param site Name of sharepoint site you want to take a file from. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @export
#' @importFrom Microsoft365R get_sharepoint_site
list_sharepoint_drives <- function(site) {

  site_list <- get_sharepoint_site(site)

  ##Get the name and url for each drive
  all_drives <- site_list$list_drives()

  list_drives <- lapply(all_drives, function(item) {
    list(
      name = item$properties$name,
      webUrl = item$properties$webUrl
    )
  })

  names(list_drives) <- sapply(list_drives, function(x) x$name)

  return(list_drives)

}


#' List the sharepoint objects available in the specified drive, useful for finding the full paths for objects
#' @name list_sharepoint_objects
#' @param site Name of sharepoint site. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @param drive Name of sharepoint drive you want to list objects associated with. This is case sensitive.
#' @param path By default, this function will list the objects at the top level of the drive. If you want to view lower-level items (e.g. inside a folder), you can specify the folder paths here. Default is blank
#' @export
#' @importFrom Microsoft365R get_sharepoint_site
list_sharepoint_objects <- function(site, drive, path = "") {

  ##Get site
  site_list <- get_sharepoint_site(site)

  ##Get the specified drive
  drive_list <- site_list$get_drive(drive)

  ##List objects in it
  drive_list$list_files(path)

}



