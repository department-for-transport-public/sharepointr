#' List the sharepoint drives available in the specified site, useful for finding their names
#' @name list_sharepoint_drives
#' @param site Name of sharepoint site you want to take a file from. Can be found in the URL of the Sharepoint location, e.g. for "departmentfortransportuk.sharepoint.com/sites/Rail", the site name would be Rail
#' @export
#' @importFrom Microsoft365R get_sharepoint_site
list_sharepoint_drives <- function(site) {

  site <- Microsoft365R::get_sharepoint_site(
    site_url =
      paste0("https://departmentfortransportuk.sharepoint.com/sites/", site),
    auth_type = "device_code")

  site$list_drives()
}
