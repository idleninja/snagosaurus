' Script used to automate volatility dump
' Brett Gross
' *****************************************************
Set args = wscript.arguments
Dim profile, drive_letter, debug, logFile
debug = False
logIt = False
strComputer = "." 
running_Location = left(Wscript.ScriptFullName, len(Wscript.ScriptFullName)-len(Wscript.ScriptName))

if args.count > 0 then
	for i=0 to args.count-1
		if InStr(args.item(i), "drive") > 0 then
			temp_arg = Split(args.item(i), "=")
			key = temp_arg(0)
			val = temp_arg(1)
			running_Location = val & ":\"
		end if
		if InStr(args.item(i), "log") > 0 then
			logIt = True
			if InStr(args.item(i), "=") > 0 then
				temp_arg = Split(args.item(i), "=")
				key = temp_arg(0)
				val = temp_arg(1)
				logFile = val
			else
				logFile = running_Location & "esa_vol.log"
			end if		
		end if
		if Instr(args.item(i), "debug") > 0 then
			debug = True
		end if
	next
else
	'wscript.echo "noargs"
end if



Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=Pkt}!\\" _ 
    & strComputer & "\root\cimv2") 

Set colOperatingSystems = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
	short_version = split(objOperatingSystem.Caption, " ")(2)
	long_version = objOperatingSystem.version
    service_pack = objOperatingSystem.ServicePackMajorVersion
	OSArch = Split(objOperatingSystem.OSArchitecture, "-")(0)
	profile = "Win" & short_version & "SP" & service_pack & "x" & OSArch
	ver = long_version
Next


if logIt then logdata logFile, "**** start ****"
dumpFilename = getTimestampFile()
if logIt then logdata logFile, dumpFilename
if debug then 
	wscript.echo running_Location & "DumpIt_modified.exe"
	'wscript.quit
end if

if logIt then logdata logFile, "Current Directory: " & runCMD("echo %cd%")
if logIt then logdata logFile, "Before execution: " & running_Location & "DumpIt_modified.exe"
Set oExec = WshShell.Exec(running_Location & "DumpIt_modified.exe")

if logIt then logdata logFile, "After execution: status " & oExec.status
if debug then wscript.sleep(10000)
if logIt then logdata logFile, "Before Terminate. Status: " & oExec.status
oExec.Terminate()
if logIt then logdata logFile, "After Terminate. Status: " & oExec.status

' Wait until memory acquisition is complete
Do While oExec.Status = 0
	'if logIt then logdata logFile, "During wait loop. Status: " & oExec.status
	'if logIt then logdata logFile, "During wait loop. StdOut (5 char): " & oExec.StdOut.Read(5)
	'wscript.echo "Status = " & oExec.status
	'wscript.echo "Status = " & oExec.ProcessID
	wscript.sleep(1000)
Loop

dumpFilePathName = running_Location & dumpFilename
if logIt then logdata logFile, "dumpFilePathName = " & dumpFilePathName
If (fso.FileExists(dumpFilePathName)) Then
	sDumpFileName = Split(dumpFilename, ".raw")(0)	
	fso.MoveFile dumpFilePathName, running_Location & sDumpFileName & "_profile=" & profile & "_version=" & ver & ".raw"
else
	if debug then wscript.echo "Error: " & dumpFilePathName & " doesn't exist"
End if


REM **************** Conjunction Junction, what's your function ************************ 
function getTimestampFile()
	ComputerName = WshShell.ExpandEnvironmentStrings("%ComputerName%")
	d = Date
	yr = Year(d)
	mon = padZeros(Month(d))
	dy = padZeros(Day(d))
	t = Time
	h = Hour(t) + GetTimeZoneOffset()
		if h >= 24 then
			h = padZeros(h - 24)
			tmp = Split(DateAdd("d", 1, cdate(mon & "/" & dy & "/" & yr)), "/")
			mon = padZeros(tmp(0))
			dy = padZeros(tmp(1))
			yr = padZeros(tmp(2))
		else
			h = padZeros(h)			
		end if

	m = padZeros(Minute(t))
	s = padZeros(Second(t))

	fdate = yr & mon & dy
	ftime = h & m & s
	getTimestampFile = ComputerName & "-" & fdate & "-" & ftime  & ".raw"
end function

function padZeros(obj)
  if obj < 10 then
	 padZeros = "0" & obj
  else
	 padZeros = obj
  end if 
end function

Function GetTimeZoneOffset()
    Const sComputer = "."

    Dim oWmiService : Set oWmiService = _
        GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
                  & sComputer & "\root\cimv2")

    Dim cTimeZone : Set cTimeZone = _
        oWmiService.ExecQuery("Select * from Win32_TimeZone")

    Dim oTimeZone
    For Each oTimeZone in cTimeZone
        GetTimeZoneOffset = (oTimeZone.Bias / 60) * -1
        Exit For
    Next
End Function

Function logdata(TextFileName, TextToWrite)
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1

  If fso.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      fso.CreateTextFile TextFileName, True
  Else
	  fso.MoveFile TextFileName, TextFileName & ".old"
	  fso.CreateTextFile TextFileName, True
  End If

Set WriteTextFile = fso.OpenTextFile(TextFileName,ForAppending, False)

WriteTextFile.WriteLine TextToWrite
WriteTextFile.Close

End Function
Function runCMD(cmd)
	runCMD = WshShell.exec("cmd /c " & cmd).stdout.readline()
End Function

