'------------------------------------------------
' echoSASProfile.vbs
'
' Description:
'   Print out the SAS Enteprise Guide profile
'   Used to set the active profile, Application.SetActiveProfile("myprofile")
'
' Options:
'   None
'
' Example Usage:
'   cscript echoSASProfile.vbs
'
' Requirements:
'   Requires matching version of cscript.exe with SAS Enterprise Guide 
'   For 32-bit EG on 64-bit Windows, use the counterintuitive version
'     c:\Windows\SysWOW64\cscript.exe
' ------------------------------------------------

' force declaration of variables in VB Script
Option Explicit

' create a new SAS Enterprise Guide automation session
Dim Application
Set Application = WScript.CreateObject("SASEGObjectModel.Application.5.1")
WScript.Echo Application.Name & ", Version: " & Application.Version

' Echo the available profiles that are defined for the current user
Dim i
Dim oShell
Set oShell = CreateObject( "WScript.Shell" )

For i = 0 to Application.Profiles.Count-1
  ' ignore local profile
  If Application.Profiles.Item(i).Name <> "Null Provider" Then
    WScript.Echo ""
    WScript.Echo "Metadata profiles available for " _
      & oShell.ExpandEnvironmentStrings("%UserName%")
    WScript.Echo "----------------------------------------"
    WScript.Echo "Profile: " _
      & Application.Profiles.Item(i).Name _
      & ", Host: " & Application.Profiles.Item(i).HostName _
	    & ", Port: " & Application.Profiles.Item(i).Port
    WScript.Echo "----------------------------------------"
    WScript.Echo ""
  End If
Next
Application.Quit
