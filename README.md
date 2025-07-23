# Sharepointr

This package is designed to simplify the process of connecting to and interacting with Microsoft SharePoint sites using the Microsoft Graph API. It provides functions list sites and files already present in Sharepoint, and perform basic operations such as uploading and downloading files to SharePoint.

## Installation

You can install the package from Github using:

```R
remotes::install_github("department-for-transport-public/sharepointr")
```

## Usage

The first time you use the package, you will need to authenticate with your Microsoft account. You will be prompted to do this the first time you run any of the functions in the package. You will receive a notification in the console which provides a link to go to, and a code to enter at this link.

You may also be asked to enter your Microsoft credentials; it's fine to do this.

Once you've authenticated the first time, your credentials will be saved, and you won't need to log in again unless your credentials expire or you clear them.

## How Sharepoint is structured

To make use of the functions in this package, you will need to know how the structure of Sharepoint folders and urls works. A typical SharePoint URL looks like this:

```
https://yourtenant.sharepoint.com/sites/yoursite/Shared_Documents/yourfolder/yourfile.txt
```

In this example:

* `yourtenant` is the name of your SharePoint tenant. This is set automatically in this package and you don't need to worry about it.
* `yoursite` is the name of the SharePoint **site** you are accessing. This is the site where your files are stored.
* `Shared_Documents` is the name of the **drive** where your files are stored.
* `yourfolder/yourfile.txt` is the full **file path** of the location of your file. This includes any folders below the drive, and the name of the file itself, as well as the file extension (e.g. docx or xlsx).

For all of the functions in this package, you will be asked to provide the `site`, `drive` and `path` parameters. These correspond to the `yoursite`, `Shared_Documents` and `yourfolder/yourfile.txt` parts of the URL above.

## Functions

### Helper listing functions

* `list_drives()`: Lists all drives (document libraries) in the SharePoint site you are accessing. This makes it easy to find the drive you want to work with, and troubleshoot if you're struggling.

* `list_sharepoint_objects()`: Lists all files and folders in a specified Sharepoint drive. This is useful for navigating the structure of your Sharepoint site and finding the files you want to work with. By default, it lists the contents of the root folder of the specified drive, but you can specify a subfolder by providing the `path` parameter. For example:

```
list_sharepoint_objects(site = "yoursite", drive = "Shared_Documents", path = "yourfolder")

list_sharepoint_objects(site = "yoursite", drive = "Shared_Documents", path = "yourfolder/subfolder")

```

The first example will list all files and folders in the `yourfolder` subfolder of the `Shared_Documents` drive in the `yoursite` SharePoint site. The second example will list all files and folders in the `subfolder` folder of the `yourfolder` folder in the same drive.

### File imports

These functions allow you to import data and files stored in sharepoint. There are two different approaches available; if you are reading in data from a file, you can use the `read_sharepoint_file()` function, which will automatically read the file into R using the appropriate function based on the file type (e.g. `read.csv()` for CSV files, `read.xlsx()` for xlsx files). This creates and deletes a temporary file, so you don't need to worry about file management yourself.

If you want to download a file to your local machine, for example a template file, you can use the `import_sharepoint()` function, which will save the file to a specified location on your computer. This won't work for large files (> 500MB), when you will get a message asking you to use alternative storage instead.

* `read_sharepoint_file()`: Reads a file from SharePoint and returns the data in R. This function automatically detects the file type and uses the appropriate function to read it for common data file types. You can specify the `site`, `drive`, and `file_path` parameters to locate the file you want to read. For some xlsx files, openxlsx gives different results to readxl; you may specify `use_openxlsx = TRUE` if you would like to use openxlsx specifically.

* `import_sharepoint()`: Downloads a file from SharePoint to your local machine. This is useful for downloading files that you want to keep on your computer, such as templates or reports. You can specify the `site`, `drive`, and `path` parameters to locate the file you want to download, and the `file_path` parameter to specify where to save the file on your local drive.

### File exports

These functions allow you to export data and files to SharePoint. If you have a data frame that you want to save as a file in SharePoint, you can use the `export_sharepoint()` function, which will automatically write the data frame to a file in SharePoint using the appropriate function based on the file type. This creates and deletes a temporary file, so you don't need to worry about file management yourself.

If you want to upload a file from your local machine to SharePoint, you can use the `export_sharepoint_file()` function, which will save the file to a specified location in SharePoint.

You can also upload an entire folder of files to SharePoint using the `export_sharepoint_folder()` function, which will recursively upload all files in the specified folder to SharePoint.

* `export_sharepoint()`: Exports a data frame to SharePoint as a file. This function automatically detects the file type and uses the appropriate function to write it for common data file types. You can specify the `site`, `drive`, and `dest_path` parameters to locate where you want to save the file in SharePoint.

* `export_sharepoint_file()`: Uploads a file from your local machine to SharePoint. This is useful for uploading files that you have created or modified on your computer, such as reports or unusual file types. You can specify the `site`, `drive`, and `dest_path` parameters to locate where you want to save the file in SharePoint, and the `file` parameter to specify the path to the file on your local machine.

* `export_sharepoint_folder()`: Uploads an entire folder of files to SharePoint. This is useful for uploading multiple files at once, such as a folder of reports or data files. You can specify the `site`, `drive`, and `dest_path` parameters to locate where you want to save the files in SharePoint, and the `dest_path` parameter to specify the path to the folder on your local machine. If a destination is not specified, the function will copy the files to the root of the specified drive in SharePoint, in the same folder structure as you have locally.

## Rendering Rmarkdown documents directly to Sharepoint

You can also render Rmarkdown documents directly to SharePoint using the `render_sharepoint()` function. This function will render an Rmarkdown document and save the output file directly to SharePoint, without needing to save it locally first.

* `render_sharepoint()`: Renders an Rmarkdown document and saves the output file directly to SharePoint. You can specify the `site`, `drive`, and `path` parameters to locate where you want to save the file in SharePoint, and the `input` parameter to specify the path to the Rmarkdown file on your local machine. The function will automatically detect the output format (e.g. HTML, PDF, Word) and save the file in the appropriate format. Additional parameters can be passed to the `rmarkdown::render()` function using the `...` argument, allowing you to customize the rendering process as needed.

## Troubleshooting

### I'm having to repeatedly login

Usually this is accompanied by an error message saying that a file is missing; this is due to corruption of your credentials file. You can reset this by running `rm(~/.local/share/AzureR/graph_logins.json)`; you will need to log in one final time, but then your credentials should be remembered.
