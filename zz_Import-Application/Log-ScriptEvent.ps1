##########################################################################################################
<#
.SYNOPSIS
   Log to a file in a format that can be read by Trace32.exe / CMTrace.exe 

.DESCRIPTION
   Write a line of data to a script log file in a format that can be parsed by Trace32.exe / CMTrace.exe

   The severity of the logged line can be set as:

        1 - Information
        2 - Warning
        3 - Error

   Warnings will be highlighted in yellow. Errors are highlighted in red.

   The tools to view the log:

   SMS Trace - http://www.microsoft.com/en-us/download/details.aspx?id=18153
   CM Trace - Installation directory on Configuration Manager 2012 Site Server - <Install Directory>\tools\

.EXAMPLE
   Log-ScriptEvent c:\output\update.log "Application of MS15-031 failed" Apply_Patch 3

   This will write a line to the update.log file in c:\output stating that "Application of MS15-031 failed".
   The source component will be Apply_Patch and the line will be highlighted in red as it is an error 
   (severity - 3).

#>
##########################################################################################################

Function Log-ScriptEvent {

#Define and validate parameters
[CmdletBinding()]
Param(
      #Path to the log file
      [parameter(Mandatory=$True)]
      [String]$NewLog,

      #The information to log
      [parameter(Mandatory=$True)]
      [String]$Value,

      #The source of the error
      [parameter(Mandatory=$True)]
      [String]$Component,

      #The severity (1 - Information, 2- Warning, 3 - Error)
      [parameter(Mandatory=$True)]
      [ValidateRange(1,3)]
      [Single]$Severity
      )


#Obtain UTC offset
$DateTime = New-Object -ComObject WbemScripting.SWbemDateTime 
$DateTime.SetVarDate($(Get-Date))
$UtcValue = $DateTime.Value
$UtcOffset = $UtcValue.Substring(21, $UtcValue.Length - 21)


#Create the line to be logged
$LogLine =  "<![LOG[$Value]LOG]!>" +`
            "<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
            "date=`"$(Get-Date -Format M-d-yyyy)`" " +`
            "component=`"$Component`" " +`
            "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
            "type=`"$Severity`" " +`
            "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
            "file=`"`">"

#Write the line to the passed log file
Add-Content -Path $NewLog -Value $LogLine

}

##########################################################################################################
