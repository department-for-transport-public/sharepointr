#' List the sharepoint drives available in the specified site, useful for finding their names
#' @name list_sharepoint_drives
#' @param site Name of sharepoint site you want to take a file from. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @export
#' @importFrom Microsoft365R get_sharepoint_site
list_sharepoint_drives <- function(site) {

  site_list <- tryCatch({
    # Attempt to get the list of drives on site
    site_obj <- Microsoft365R::get_sharepoint_site(
      site_url = paste0("https://departmentfortransportuk.sharepoint.com/sites/", site),
      auth_type = "device_code"
    )
  },
  error = function(e) {
    # Check if the error is a 404 (site not found)
    if (grepl("404", e$message)) {
      stop("Error: SharePoint site: ", site, " not found. Either you have spelled the site name wrong or do not have permission to view it")
    } else {
      stop("Error accessing SharePoint site: ", e$message)
    }
  })

  ##Get the name and url for each drive
  all_drives <- site_list$list_drives()

  list_drives <- lapply(all_drives, function(item) {
    list(
      name = item$properties$name,
      formatted_name = gsub("^.*/", "", item$properties$webUrl),
      webUrl = item$properties$webUrl
    )
  })

  names(list_drives) <- sapply(list_drives, function(x) x$name)

  return(list_drives)

}
