% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/export_sharepoint.R
\name{export_sharepoint_folder}
\alias{export_sharepoint_folder}
\title{Upload a local folder to SharePoint}
\usage{
export_sharepoint_folder(local_folder, site, drive, dest_path = NULL)
}
\arguments{
\item{local_folder}{Path to the local folder to upload}

\item{site}{SharePoint site name (e.g. "Rail")}

\item{drive}{SharePoint drive name (e.g. "RailStats")}

\item{dest_path}{Destination path in SharePoint drive (e.g. "Reports/July2025"). Defaults to NULL, which will put the folder in the same location as it is locally.}
}
\description{
Upload a local folder to SharePoint
}
