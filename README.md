# convertSAStoCSV
Uses SAS Enterprise Guide automation to convert between SAS binary (sas7bdat) and CSV formats

## Features
* Either import from CSV to SAS or export from SAS to CSV
* Set SAS log file location
* If a SAS log file is not explicitly noted, then the log is saved in the same directory as the executing script
* Utilize a where clause on the SAS dataset
* Protects against accidental replacing of files by prompting the user t    o replace a file if it already exists
* Set the EG profile name using a configuration file
* Check command line arguments

## Options
* Required:
  * `/conv:`    ==> conversion type - either an import or an export
  * `/sas:`     ==> location of SAS file, requires sas7bdat extension
  * `/csv:`     ==> location of CSV file, requires csv extension

* Optional:
  * `/log:`     ==> location of log file
  * `/config:`  ==> location of config file; formatted as an ini file, does not require ini extension; used to set the EG profile and any other config options
  * `/where:`   ==> where clause to be applied to SAS file
  * `/repl`     ==> if argument is used, always replace output

* Help:
  * `/help`		  ==> print argument options and usage

## Example Usage
  * `cscript convertSAStoCSV.vbs /help`

## Requirements
Requires matching version of cscript.exe with SAS Enterprise Guide.

For 32-bit EG on 64-bit Windows, use the counterintuitive version c:\Windows\SysWOW64\cscript.exe

## Inspiration
#### Reading
[Doing More with EG Automation - PDF](http://support.sas.com/documentation/onlinedoc/guide/examples/SASGF2012/Hemedinger_298-2012.pdf)

[Doing More with EG Automation - SAS Community](http://www.sascommunity.org/wiki/Not_Just_for_Scheduling:_Doing_More_with_SAS_Enterprise_Guide_Automation)

#### Code
[Run SAS program "like a batch job"](http://support.sas.com/documentation/onlinedoc/guide/examples/SASGF2012/BatchProject.vbs.txt)
