# netbox-cmn-import-script

## Introduction

This repository tracks a Python script used to generate a series of JSON files from the CMN Configuration Data Spreadsheets to be imported into NetBox.
The files generated follow the data structure required to import a JSON file into NetBox, they include a device, interface and IP address JSON.
There are a numer of prerequisites which must be fulfilled within NetBox before the files generated may be imported, these are documeneted below under the NetBox Prerequisites heading.

## Pip Dependencies

This script requires the argparse package to pass an arguement from the command line to the script when running the script.
The openpyxl library is required to access data from the configuration spreadsheet.
The json library is required to write the dictionaries created in a JSON file.
The re library is required to run regex searches on strings within the spreadsheet

## Before running

The path.py file stores a single variable 'path' which must be editted before running the script to include the directory where generated JSON files may be written to.

The name of the spreadsheets must also be updated to include the site where the CloudVision instances are located. This could be drawn from the spreadsheet itself based on the hostname, however the CloudVision instances in the the SCO has devices onboarded form the outside of the SCO, hence confusion may arise from this method.
The file names must have either 'sco_' or 'nco' appended to the start of the file name.

## NetBox Prerequisites

The following prerequisites **MUST** be fulfilled before the JSON files may be imported into NetBox, else the import will fail.
These include;
- Creation of the sites where devices are located
- Creation of all device types covered in the spreadsheets
- Creation of the Aggregated and Address prefixes
- Creation of all custom fields & custom filed choices
- Creation of all TAGs assigned to devices, interfaces and addresses

## Running the script

To run this script from the CLI, run the following command from the directory where the script is stored;
```bash
py ./generate-import-files.py -f <path to config files>
```



