' ------------------------------------------------
' convertSAStoCSV.vbs
'
' Description:
'   Uses SAS Enterprise Guide automation to convert between SAS binary (sas7bdat) and CSV formats
'
' Features:
'   - Either import from CSV to SAS or export from SAS to CSV
'   - Set SAS log file location
'   - If a SAS log file is not explicitly noted, then the log is saved in the same directory as the executing script
'   - Utilize a where clause on the SAS dataset
'   - Protects against accidental replacing of files by prompting the user to replace a file if it already exists
'   - Set the EG profile name using a configuration file
'   - Check command line arguments
'
' Options:
'   Required:
'       /conv:      ==> conversion type - either an import or an export
'       /sas:       ==> location of SAS file, requires sas7bdat extension
'       /csv:       ==> location of CSV file, requires csv extension
'
'   Optional:
'       /log:       ==> location of log file
'       /config:    ==> location of config file
'                       formatted as an ini file, does not require ini extension
'                       used to set the EG profile and any other config options
'       /where:     ==> where clause to be applied to SAS file
'       /repl       ==> if argument is used, always replace output
'
'   Help:
'       /help       ==> print argument options and usage
'
' Example Usage:
'   cscript convertSAStoCSV.vbs /help
'
' Requirements:
'   Requires matching version of cscript.exe with SAS Enterprise Guide
'   For 32-bit EG on 64-bit Windows, use the counterintuitive version
'       c:\Windows\SysWOW64\cscript.exe
'------------------------------------------------

' force declaration of variables in VB Script
Option Explicit

' ----------------
' Argument Parsing and Checking
' ----------------
  
' print help
If WScript.Arguments.Named.Exists("help") Then
  Call echoAndQuit("", "help")
End If

' parse command line arguments
Dim sasBinaryFile : sasBinaryFile = WScript.Arguments.Named.Item("sas")
Dim csvFile : csvFile = WScript.Arguments.Named.Item("csv")
Dim convType : convType = WScript.Arguments.Named.Item("conv")
Dim sasLog : sasLog = WScript.Arguments.Named.Item("log")
Dim whereClause : whereClause = WScript.Arguments.Named.Item("where")
Dim replExists : replExists = WScript.Arguments.Named.Exists("repl")
Dim configFile : configFile = WScript.Arguments.Named.Item("config")

' create other variables from command line arguments
Dim sasFileDir : sasFileDir = getBasePath(sasBinaryFile)
Dim sasDatasetName : sasDatasetName = getFileNoExt(sasBinaryFile)
Dim sasFileExtension : sasFileExtension = getExtension(sasBinaryFile)
Dim csvFileName : csvFileName = getFileNoExt(csvFile)
Dim csvFileExtension : csvFileExtension = getExtension(csvFile)

' check required command line arguments
If sasBinaryFile = "" OR _
   csvFile = "" OR _
   convType = "" Then
  Call echoAndQuit("Script requires three arguments - /sas: /csv: /conv:", "err")
End If

If InStr(sasBinaryFile, "sas7bdat") = 0 Then
  Call echoAndQuit("The named argument '/sas:' should be the full path name to the SAS binary file, including sas7bdat extension", "err")
End If

If InStr(csvFile, "csv") = 0 Then
  Call echoAndQuit("The named argument '/csv:' should be the full path name to the CSV file, including csv extension", "err")
End If

If convType <> "import" AND _
   convType <> "export" Then
  Call echoAndQuit("The named argument '/conv:' must be either import or export", "err")
End If

' check if config file exists
Dim sasRCFile
If configFile = "" Then
   sasRCFile = getUserProfile() & "\.sasrc"
  If Not dataFileExists(sasRCFile) Then
    Dim errorMsg : errorMsg = "A configuration file does not exist.  " _
                              & "The default file assumed is " _
                              & vbNewLine _
                              & "  " & getUserProfile() & "\.sasrc" _
                              & vbNewLine _
                              & "Alternatively, the configuration file may " _
                              & "also be set with the /config: argument"
    Call echoAndQuit(errorMsg, "err")
  End If
Else
  sasRCFile = configFile
End If

' check if SAS and CSV files exist
If convType = "import" Then
  If dataFileExists(sasBinaryFile) AND Not replExists Then
    ' run for side effects
    Call fileWriteExists(sasDatasetName, sasFileExtension)
  End If
  If Not dataFileExists(csvFile) Then
    ' run for side effects
    Call fileReadNotExists(csvFileName, csvFileExtension)
  End If
' convType = "export"
Else
  If dataFileExists(csvFile) AND Not replExists Then
    ' run for side effects
    Call fileWriteExists(csvFileName, csvFileExtension)
  End If
  If Not dataFileExists(sasBinaryFile) Then
    ' run for side effects
    Call fileReadNotExists(sasDatasetName, sasFileExtension)
  End If
End If

' optional
WScript.Echo ""
WScript.Echo "########"
WScript.Echo "# Note #"
WScript.Echo "########"
WScript.Echo ""
WScript.Echo "  The following arguments are optional"
WScript.Echo ""
WScript.Echo "    log file ==> " & Replace(WScript.ScriptFullName, ".vbs", ".log")
WScript.Echo "        if the log file already exists, it is ALWAYS replaced"
If configFile <> "" Then
  WScript.Echo "    config file ==> " & configFile
Else
  WScript.Echo "    config file ==> " & getUserProfile() & "\.sasrc"
End If
If whereClause <> "" Then
  WScript.Echo "    where clause ==> " & Q(whereClause)
Else
  WScript.Echo "    where clause ==> None"
End If
If replExists Then
  WScript.Echo "    repl ==> Yes"
Else
  WScript.Echo "    repl ==> No"
End If
WScript.Echo "        no value is need for the /repl argument"
WScript.Echo "        if present, replace files without prompting"

' ----------------
' Conversion
' ----------------

' Application ==> Project ==> Code Collection (Program)

' create a new SAS Enterprise Guide automation session
Dim Application
Set Application = WScript.CreateObject("SASEGObjectModel.Application.5.1")

' the argument to the method SetActiveProfile should be the profile name within SAS Enterprise Guide
' the profile name is set within a config file
' the config file should be of the form
'   [EGProfile]
'   profile=myprofilename
' script echoSASProfile.vbs can be used to select from possible profiles
'   in order to create the needed config file

Dim egProfile : egProfile = getConfigVars(sasRCFile, "EGProfile", "profile")
Application.SetActiveProfile(egProfile)

' New project
Dim Project : Set Project = Application.New

' New SAS program within the project just created
Dim SASProgram : Set SASProgram = Project.CodeCollection.Add

' Override application defaults
SASProgram.UseApplicationOptions = False
SASProgram.GenHTML = False
SASProgram.GenListing = False
SASProgram.GenPDF = False
SASProgram.GenRTF = False
SASProgram.GenSASReport = False

' Set the SAS application server
Dim sasAppServer : sasAppServer = getConfigVars(sasRCFile, "ApplicationServer", "server")
SASProgram.Server = sasAppServer

' Write the SAS program and run
' Always create a space at the end of the SAS snippets
'		so that they can be chained together
Dim sasLibname : sasLibname = "libname sasdir " & Q(sasFileDir) & "; "

Dim procExportWhere
Dim procImportWhere
If whereClause <> "" Then
  procExportWhere = "proc export data=sasdir." & sasDatasetName _
                    & "(where=(" & whereClause & "))"

  procImportWhere = " out=sasdir." & sasDatasetName _
                    & "(where=(" & whereClause & "))" & ";"
Else
  procExportWhere = "proc export data=sasdir." & sasDatasetName
  procImportWhere = "out=sasdir." & sasDatasetName & ";"
End If

Dim procExport
procExport =  procExportWhere _
              & " outfile=" & Q(csvFile) _
              & " dbms=csv" _
              & " replace;" _
              & " run; "

Dim procImport
procImport =  "proc import datafile=" & Q(csvFile) _
              & " dbms=csv" _
              & " replace" _
              & procImportWhere _
              & " guessingrows=5000;" _
              & " run; "

If convType = "export" Then
  SASProgram.Text = sasLibname & procExport
Else
  SASProgram.Text = sasLibname & procImport
End If

SASProgram.Run

' Save the log file to disk
If sasLog = "" Then
  SASProgram.Log.SaveAs Replace(WScript.ScriptFullName, ".vbs", ".log")
Else
  SASProgram.Log.SaveAs sasLog
End If

Application.Quit

' ----------------
' helper functions
' ----------------

' to return a value from a function, assign the value to the function name

Function getUserProfile()
  Dim oshell : Set oshell = CreateObject("WScript.Shell")
  getUserProfile = oshell.ExpandEnvironmentStrings("%userprofile%")
End Function

' return \\server\folder1 when given
' 	the file path \\server\folder1\myfile.csv
Function getBasePath(fullPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  getBasePath = fso.GetParentFolderName(fullPath)
End Function

' return myfile when given
'		the file path \\server\folder1\myfile.csv
Function getFileNoExt(fullPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  getFileNoExt = fso.GetBaseName(fullPath)
End Function

' return csv when given
'		the file path \\server\folder1\myfile.csv
Function getExtension(fullPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  getExtension = fso.GetExtensionName(fullPath)
End Function

' quote a string
Function Q(s)
  Q = chr(34) & s & chr(34)
End Function

' file exists
Function dataFileExists(fullPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  dataFileExists = fso.FileExists(fullPath)
End Function

' prompt for and get back result from stdin
Function getStdIn(prompt)
  WScript.Echo prompt & " "
  WScript.StdIn.Read(0)
  getStdIn = WScript.StdIn.ReadLine()
End Function

' file to write already exists
Function fileWriteExists(dsName, dsExtension)
  WScript.Echo ""
  Dim replacePrompt
  replacePrompt = "The file " & dsName & "." & dsExtension _
                  & " already exists.  Would you like to replace " _
                  & dsName & "." & dsExtension & "? [Y/N] "
  Dim userReply
  userReply = getStdIn(replacePrompt)
  while userReply <> "Y" AND _
        userReply <> "y" AND _
        userReply <> "N" AND _
        userReply <> "n"
    WScript.Echo ""
    WScript.Echo "Expecting either 'Y' or 'N'"
    userReply = getStdIn(replacePrompt)
  Wend
  If userReply = "N" OR userReply = "n" Then
    WScript.Quit -1
  End If
End Function

' file to read does not exist
Function fileReadNotExists(dsName, dsExtension)
  WScript.Echo ""
  WScript.Echo "The file " & dsName & "." & dsExtension _
               & " does not exist."
  WScript.Echo ""
  WScript.Quit -1
End Function

' read config file
' return a dictionary within a dictionary
' Example:
'   If I have a config file in ini format
'     [section1]
'     a = 1
'     b = 2
'     [section2]
'     a = 3
'     c = 4
'   Then the JSON equivalent returned is
'   {
'     "section1": {
'       "a": "1",
'       "b": "2"
'     },
'     "section2": {
'       "a": "4",
'       "c": "5"
'     }
'   }
'   In this example, 'section1' is the key pointing to the
'   dictionary object with keys 'a' and 'b'
Function readConfigFile(configFile)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim iniDict : Set iniDict = CreateObject("Scripting.Dictionary")
  If Not dataFileExists(configFile) Then
    Dim errorMsg : errorMsg = "The configuration file " _
                              & configFile & " does not exist." _
                              & vbNewLine & vbNewLine _
                              & "  Please use the default file " _
                              & getUserProfile() & "\.sasrc" _
                              & vbNewLine _
                              & "    or ensure " & configFile & " exists"
    Call echoAndQuit(errorMsg, "err")
  Else
    Dim iniFile : Set iniFile = fso.OpenTextFile(configFile)
  End If
  Dim iniLine, iniSection, sectionKVArray
  iniSection = ""
  Do Until iniFile.AtEndOfStream
    iniLine = Trim(iniFile.ReadLine)
    If "[" = Left(iniLine, 1) Then
      iniSection = Mid(iniLine, 2, Len(iniLine) - 2)
      Set iniDict(iniSection) = CreateObject("Scripting.Dictionary")
    ElseIf iniLine <> "" AND iniSection <> "" Then
      sectionKVArray = Split(iniLine, "=")
      If UBound(sectionKVArray) = 1 Then
        iniDict(iniSection)(Trim(sectionKVArray(0))) = Trim(sectionKVArray(1))
      End If
    End If
  Loop
  iniFile.Close
  Set readConfigFile = iniDict
End Function

' get a variable value from a config file
Function getConfigVars(configFile, wantedSection, wantedKey)
  Dim iniDict : Set iniDict = readConfigFile(configFile)
  Dim iniSection, finalSection, _
      sectionKey, finalItem, _
      errorMsg
  finalSection = ""
  finalItem = ""
  For Each iniSection In iniDict.Keys()
    If LCase(iniSection) = LCase(wantedSection) Then
      finalSection = wantedSection
      For Each sectionKey in iniDict(iniSection).Keys()
        If LCase(sectionKey) = LCase(wantedKey)  Then
          finalItem = iniDict(iniSection)(sectionKey)
          Exit For
        End If
      Next
      If finalItem = "" Then
        errorMsg =  "There is not a " & wantedKey & " variable " _
                    & "within the config file " & configFile _
                    & vbNewLine & vbNewLine _
                    & "  The config file should be formatted as below" _
                    & vbNewLine & vbNewLine _
                    & "    [EGProfile]" _
                    & vbNewLine _
                    & "    profile=myEGProfile" _
                    & vbNewLine & vbNewLine _
                    & "    [ApplicationServer]" _
                    & vbNewLine _
                    & "    server=SASApp"
        Call echoAndQuit(errorMsg, "err")
      End If
    End If
  Next
  If finalSection = "" Then
    errorMsg =  "There is not {a, an} " & wantedSection & " section within " _
                & "the config file " & configFile
    Call echoAndQuit(errorMsg, "err")
  End If
  getConfigVars = finalItem
End Function

Function echoAndQuit(msgStatement, msgType)
  If msgType = "err" Then
    WScript.Echo ""
    WScript.Echo "#########"
    WScript.Echo "# Error #"
    WScript.Echo "#########"
    WScript.Echo ""
    WScript.Echo "  " & msgStatement
    WScript.Echo ""
  Else
    WScript.Echo "#############"
    WScript.Echo "# Arguments #"
    WScript.Echo "#############"
    WScript.Echo ""
    WScript.Echo "  Required:"
    WScript.Echo ""
    WScript.Echo "    /conv:    ==> conversion type - either an import or an export"
    WScript.Echo "    /sas:     ==> location of SAS file, requires sas7bdat extension"
    WScript.Echo "    /csv:     ==> location of CSV file, requires csv extension"
    WScript.Echo ""
    WScript.Echo "  Optional:"
    WScript.Echo ""
    WScript.Echo "    /log:     ==> location of log file"
    WScript.Echo "    /config:  ==> location of config file" & vbNewLine & "                  formatted as an ini file, does not require ini extension" & vbNewLine & "                  used to set the EG profile and any other config options"
    WScript.Echo "    /where:   ==> where clause to be applied to SAS file"
    WScript.Echo "    /repl     ==> if argument is used, always replace output"
    WScript.Echo "    /help     ==> print argument options and usage"
  End If
  WScript.Echo ""
  WScript.Echo "#################"
  WScript.Echo "# Example Usage #"
  WScript.Echo "#################"
  WScript.Echo ""
  WScript.Echo "  Required:"
  WScript.Echo ""
  WScript.Echo "    Export from SAS to CSV"
  WScript.Echo ""
  WScript.Echo "      cscript convertSAStoCSV.vbs /conv:export /sas:" & Q("\\server\SAS Files\myfile.sas7bdat") & " /csv:" & Q("\\server\CSV Files\myfile.csv")
  WScript.Echo ""
  WScript.Echo "    Import from CSV to SAS"
  WScript.Echo ""
  WScript.Echo "      cscript convertSAStoCSV.vbs /conv:import /sas:" & Q("\\server\SAS Files\myfile.sas7bdat") & " /csv:" & Q("\\server\CSV Files\myfile.csv")
  WScript.Echo ""
  WScript.Echo ""
  WScript.Echo "  All Options:"
  WScript.Echo ""
  WScript.Echo "    Export from SAS to CSV"
  WScript.Echo ""
  WScript.Echo "      cscript convertSAStoCSV.vbs /conv:export /sas:" & Q("\\server\SAS Files\myfile.sas7bdat") & " /csv:" & Q("\\server\CSV Files\myfile.csv") & " /log:" & Q("\\server\Log Files\myfile.log") & " /config:" & Q("\\server\Config Files\myconfigfile") & " /where:" & Q("myvariable < 10") & " /repl"
  WScript.Echo ""
  WScript.Echo "    Import from CSV to SAS"
  WScript.Echo ""
  WScript.Echo "      cscript convertSAStoCSV.vbs /conv:import /sas:" & Q("\\server\SAS Files\myfile.sas7bdat") & " /csv:" & Q("\\server\CSV Files\myfile.csv") & " /log:" & Q("\\server\Log Files\myfile.log") & " /config:" & Q("\\server\Config Files\myconfigfile") & " /where:" & Q("myvariable < 10") & " /repl"
  WScript.Quit -1
End Function
