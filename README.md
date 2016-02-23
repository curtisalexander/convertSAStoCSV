# convertSAStoCSV
Uses SAS Enterprise Guide automation to convert between SAS binary (sas7bdat) and CSV formats

## Features
* Either import from CSV to SAS or export from SAS to CSV
* Set SAS log file location
* If a SAS log file is not explicitly noted, then the log is saved in the same directory as the executing script
* Utilize a where clause on the SAS dataset
* Protects against accidental replacing of files by prompting the user to replace a file if it already exists
* Set the EG profile name and version using a configuration file
* Check command line arguments

## Options
```
Required:
    /conv:      ==> conversion type - either an import or an export
    /sas:       ==> location of SAS file, requires sas7bdat extension
    /csv:       ==> location of CSV file, requires csv extension

Optional:
    /log:       ==> location of log file
    /config:    ==> location of config file
                    formatted as an ini file, does not require ini extension
                    used to set the EG profile, version, and any other config options
    /where:     ==> where clause to be applied to SAS file
    /repl       ==> if argument is used, always replace output

Help:
    /help       ==> print argument options and usage
```

## Examples
* Minimum Required
    * Export from SAS to CSV
        * `cscript convertSAStoCSV.vbs /conv:export /sas:"\\server\SAS Files\myfile.sas7bdat" /csv:"\\server\CSV Files\myfile.csv"`
    * Import from CSV to SAS
        * `cscript convertSAStoCSV.vbs /conv:import /sas:"\\server\SAS Files\myfile.sas7bdat" /csv:"\\server\CSV Files\myfile.csv"`
<br><br>
* All Options
    * Export from SAS to CSV
        * `cscript convertSAStoCSV.vbs /conv:export /sas:"\\server\SAS Files\myfile.sas7bdat" /csv:"\\server\CSV Files\myfile.csv" /log:"\\server\Log Files\myfile.log" /config:"\\server\Config Files\myconfigfile" /where:"myvariable < 10" /repl`
    * Import from CSV to SAS
        * `cscript convertSAStoCSV.vbs /conv:import /sas:"\\server\SAS Files\myfile.sas7bdat" /csv:"\\server\CSV Files\myfile.csv" /log:"\\server\Log Files\myfile.log" /config:"\\server\Config Files\myconfigfile" /where:"myvariable < 10" /repl`

## Requirements
The script `echoSASProfile.vbs` assumes Enterprise Guide version 5.1.  Update accordingly if using a different version.

Requires matching version of cscript.exe with SAS Enterprise Guide.

For 32-bit EG on 64-bit Windows, use the counterintuitive version c:\Windows\SysWOW64\cscript.exe

## Configuration File
Within the repo is a sample configuration file, `.sasrc`.  The value needed for the `[EGProfile]` option can be obtained by running the script `echoSASProfile.vbs`.

## [Standing on the shoulders](https://en.wikipedia.org/wiki/Standing_on_the_shoulders_of_giants)
Thanks to [Chris Hemedinger](https://github.com/cjdinger) for his Enterprise Guide Automation writings.

#### Reading
[Doing More with EG Automation - PDF](http://support.sas.com/documentation/onlinedoc/guide/examples/SASGF2012/Hemedinger_298-2012.pdf)

[Doing More with EG Automation - SAS Community](http://www.sascommunity.org/wiki/Not_Just_for_Scheduling:_Doing_More_with_SAS_Enterprise_Guide_Automation)

#### Code
[Run a SAS program "like a batch job"](http://support.sas.com/documentation/onlinedoc/guide/examples/SASGF2012/BatchProject.vbs.txt)
