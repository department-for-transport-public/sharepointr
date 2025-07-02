#' Get the sharepoint site object
#' @name get_sharepoint_site
#' @param site Name of sharepoint site. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @importFrom Microsoft365R get_sharepoint_site

get_sharepoint_site <- function(site){
  site_obj <- tryCatch({
    # Attempt to get the list of drives on site
   Microsoft365R::get_sharepoint_site(
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

  return(site_obj)
}
