% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/import_sharepoint.R
\name{read_sharepoint_file}
\alias{read_sharepoint_file}
\title{Read a CSV or Excel file directly from SharePoint}
\usage{
read_sharepoint_file(site, drive, file_path, use_openxlsx = FALSE, ...)
}
\arguments{
\item{site}{SharePoint site name}

\item{drive}{Drive name within SharePoint}

\item{file_path}{Path to the file on SharePoint (including filename and extension)}

\item{use_openxlsx}{Logical. If TRUE, uses the openxlsx package to read Excel files. Defaults to FALSE, which uses readxl.}

\item{...}{Additional arguments passed to read.csv() or read_excel() functions}
}
\value{
A data.frame or tibble
}
\description{
Read a CSV or Excel file directly from SharePoint
}
