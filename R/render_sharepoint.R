#' Render an RMarkdown file and upload output to SharePoint
#'
#' @param input_path Path to the `.Rmd` file.
#' @param site SharePoint site name (e.g. "Rail").
#' @param drive SharePoint drive name (e.g. "Reports").
#' @param dest_folder Path inside the SharePoint drive to upload to.
#' @param ... Additional arguments passed to `rmarkdown::render()`.
#' @export
#' @importFrom rmarkdown render

render_to_sharepoint <- function(input_path,
                                 site,
                                 drive,
                                 dest_path,
                                 ...) {
  # Check input Rmd exists
  if (!file.exists(input_path)) {
    stop("Input file '", input_path, "' not found.")
  }

  # Create temp output folder
  tmp_dir <- tempdir()
  tmp_subfolder <- file.path(tmp_dir, "render_output")


  message("Rendering '", input_path)


  tryCatch({
    suppressWarnings(
      suppressMessages(
        rmarkdown::render(
          input = input_path,
          output_dir = tmp_subfolder,
          ...
          )
        )
    )
    message("Successfully rendered ", input_path)
  }, error = function(e) {
    stop("Rendering failed: ", e$message)
  })

  # Connect to SharePoint
  site_loc <- get_sharepoint_site(site)
  drive_loc <- site_loc$get_drive(drive)

  tryCatch({

    drive_loc$upload_folder(src = tmp_subfolder, dest = dest_path)

    message("Sharepoint render complete!")
  }, error = function(e) {
    stop("Failed to upload rendered document: ", e$message)
  })

  ##Delete the folder I made
  unlink(tmp_subfolder, recursive = TRUE)

}
