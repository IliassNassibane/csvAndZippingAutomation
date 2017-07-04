'----------------------------------------------------------------
' @Script: RCF - inspectieReport.vbs
' @author: Iliass Nassibane, Inspectation
' @desc:   Met deze script wordt er een rapportage uitgevoerd, gezipt en verplaatst op 
'		   een fileserver van de klant.
' @huidige versie: 1.2
'
'v1.0	initiële versie n.a.v. DBmonitoring.vbs
'v1.1	uitgebreide functionaliteit n.a.v. Ultrasoon
'v1.2	uitbreiding door Iliass Nassibane, zipping van bestanden en doorsturen naar een fileshare.
'----------------------------------------------------------------

Option explicit

Dim fso
Dim sScriptPath, iniFile, Logfile, blog
Dim DBserver, Database, sDBuser, sDBpwd
Dim sSQL
Dim sHTTP, sFile, sCSV, sReportFile
Dim strTemp, strTemp2, strTempTXT, iRow, sLine, iProcessed
Dim oCurrentFolder, oFiles, oFile
Dim cFieldSeperator, bAddHeader, bConvertDecimal, bUpdateSQL, bError
'	Dit zijn de variabelen waarmee er een folder kan worden aangemaakt.
Dim fldrLocation, fldrName, newFldr
Dim fldrLocation2, fldrName2, newFldr2
'	Dit zijn de variabelen die gebruikt gaan worden om de zipping sub routine te voeden.
Dim csvInput, folder, zipfilePath, zipfilename
'	Dit zijn de variabelen die gebruikt gaan worden om de bestanden uit de oorspronkelijke locatie te archiveren.
Dim fldrTo, fldrFrom, fldrToMove, subsFolder
Dim folderSubFolder, Collec_Files1, Collec_Files2
Dim datum, dateYear, dateDay, dateMonth, dateFormatted, bestandsNaam 

datum = CDate(date)
dateYear = DatePart("yyyy", datum)
dateMonth = DatePart("m", datum)
dateDay = DatePart("d", datum)

If dateDay < 10 Then
   dateDay = "0" & dateDay
End If
If dateMonth < 10 Then
   dateMonth = "0" & dateMonth
End If
  
dateFormatted = dateYear & dateMonth & dateDay
bestandsNaam = dateFormatted & "_US"

On Error Resume Next

cFieldSeperator = ";"
bAddHeader = vbTrue
bConvertDecimal = vbTrue
bError = vbFalse

'laden instellingen
Set fso = CreateObject("Scripting.FileSystemObject")
sScriptPath = left(WScript.ScriptFullName, LEN(WScript.ScriptFullName) - LEN(WScript.ScriptName) - 1)
If Wscript.Arguments.Count > 0 Then
  inifile = sScriptPath & "\" & Wscript.Arguments(0)
else
  inifile = sScriptPath & "\" & left(WScript.ScriptName, InStrRev(WScript.ScriptName, ".") - 1) & ".ini"
End If

'--------------------------------------------------------------------------------------------------
'	@desc: Variabelen worden gevuld met de benodigde locaties en naamgevingen.
'		   Voor het maken van de locatie waar de pdfs worden geplaatst. 
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'--------------------------------------------------------------------------------------------------

fldrLocation = sScriptPath & "\Source"
fldrName = bestandsNaam
newFldr = fldrLocation & "\" & fldrName

' @desc: checkt of de folderlocatie aanwezig is. Zo niet, dan wordt ie opnieuw aangemaakt.
If not FSO.FolderExists(newFldr) then 
	Call CreateFolder(newFldr)
End if
	
'--------------------------------------------------------------------------------------------------

Logfile = sScriptPath & "\" & GetINIString("Settings", "LogFile", left(WScript.ScriptName, InStrRev(WScript.ScriptName, ".") - 1) & ".log", inifile)
If Wscript.Arguments.Count > 1 Then
  Logfile = sScriptPath & "\" & Wscript.Arguments(1)
End If
strTemp = ucase(GetINIString("Settings", "log", "1", inifile))
bLog = cBool(strTemp = "1") OR cBool(strTemp = "T") OR cBool(strTemp = "W")_
       OR cBool(strTemp = "TRUE") OR cBool(strTemp = "WAAR")

strTemp = ucase(GetINIString("Settings", "UpdateSQL", "1", inifile))
bUpdateSQL = cBool(strTemp = "1") OR cBool(strTemp = "T") OR cBool(strTemp = "W")_
             OR cBool(strTemp = "TRUE") OR cBool(strTemp = "WAAR")

DBserver = GetINIString("Settings", "DBserver", "127.0.0.1", inifile)
Database = GetINIString("Settings", "database", "TempDB", inifile)
sDBUser = GetINIString("Settings", "dbUser", "", inifile)
if (sDBUser <> "") and (sDBUser <> "-") then
  sDBpwd = Decrypt(GetINIString("Settings", "dbpwd", "", inifile), sDBuser)
end if

Call log("program started", 0, logfile)

'meetrun
sSQL = GetIniSection("sql2", "", inifile)
strTemp = SQLexec(DBserver, sDBuser, sDBpwd, Database, sSQL, vbTrue)
WriteFile sScriptPath & "\Source\" & bestandsNaam & ".csv", strTemp
csvInput = sScriptPath & "\Source\" & bestandsNaam & ".csv"

Call csvFormatter(csvInput, ",", ".")

sSQL = GetIniSection("sql1", "", inifile)
strTemp = SQLexec(DBserver, sDBuser, sDBpwd, Database, sSQL, vbTrue)

strTemp = split(strTemp, vbCrLf)
iRow = 0
iProcessed = 0
for each sline in strTemp
  if (sline <> "") and (iRow > 0) then
    bError = vbFalse
    sSQL = GetIniSection("sql8", "", inifile)
    sHTTP = ReplaceValuesHTTP(GetIniString("http2", "Get", "", inifile), strTemp, iRow)
    sFile = sScriptPath & "\" & Filename(ReplaceValues(GetIniString("http2", "output", "", inifile), strTemp, iRow))
    sCSV = filename(ReplaceValues(GetIniString("Settings", "CSVfilename", "", inifile), strTemp, iRow))
    sReportFile = filename(ReplaceValues(GetIniString("Settings", "SaveReport", "", inifile), strTemp, iRow))

    if sHTTP <> "" then
      strTemp2 = HTTPget(sHTTP, sFile, sReportFile)
      Call log("Getting report result : " & strTemp2, 0, logfile)
      if strTemp2 <> "OK" then
        bError = vbTrue
      end if
    end if

    if sCSV <> "" then
      Call log("making CSV", 0, logfile)
      strTemp2 = ReplaceValues(GetIniSection("CSV", "", inifile), strTemp, iRow)
      WriteFile sScriptPath & "\" & sCSV, strTemp2
    end if    

	' SQL3 moet worden aangevuld met de query van Emir.
	' Hiermee wordt namelijk de update uitgevoerd op de database.
    if bError = vbFalse then
      if bUpdateSQL = vbTrue then
        sSQL = GetIniSection("sql3", "", inifile)
        sSQL = ReplaceValues(sSQL, join(strTemp, vbCrLf), iRow)




        wscript.echo sSQL





        strTemp2 = SQLexec(DBserver, sDBuser, sDBpwd, Database, sSQL, vbFalse)
        if strTemp2 = "OK" then
          call log("Updating record OK", 0, logfile)
        end if
      end if
    end if
    iProcessed = iProcessed + 1
  end if
  iRow = iRow + 1
next

subsFolder = sScriptPath & "\Source" & "\" & bestandsNaam
Call EmptyFolderCheck(subsFolder)

'--------------------------------------------------------------------------------------------------
'	@desc: Variabelen worden gevuld met de benodigde locaties en naamgevingen.
'		   Zipping sub routine waarmee het zip bestand wordt aangemaakt. 
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'--------------------------------------------------------------------------------------------------

folder = sScriptPath & "\Source"
zipfilePath = sScriptPath & "\Output"
zipfilename = zipfilePath & "\" & Hour(Now) & fldrName & ".zip"
Call Zipper(zipfilename, folder)

'--------------------------------------------------------------------------------------------------
'	@desc: Sub routine waarmee de gebruikte bestanden na verwerking worden gearchiveerd naar de \Archive locatie
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'-------------------------------------------------------------------------------------------------
fldrLocation2 = sScriptPath & "\Archief"
fldrName2 = Hour(Now) & Minute(Now) & "-" & Day(Now) & "-" & Month(Now) & "-" & DatePart("yyyy", Now()) & "-" & "Source"
newFldr2 = fldrLocation2 & "\" & fldrName2
Call CreateFolder(newFldr2)
Call log("Adres voordat de archivering is begonnen: " & newFldr2 & vbCrLf, 0, logfile)

Set folderSubFolder = folder.SubFolders
Set Collec_Files1 = folder.Files
Set Collec_Files2 = newFldr.Files

If folderSubFolder.Count <> 0 Then			' @desc: checkt of er een subfolder aanwezig is op de locatie.
	If Collec_Files2.Count > 1 Then			' @desc: checkt of de oorspronkelijke locatie met bestanden gevuld is.
		For Each File In Collec_Files1
			fso.MoveFile folder + "\*", newFldr2
			Call log("(Archief) Bestand: " & File & "(" & folder & ")" & " overgezet naar archief." & "(" & newFldr2 & ")", 0, logfile)
		Next

		For Each flder In folderSubFolder
			If FSO.FolderExists(flder) then
				fso.MoveFolder newFldr, newFldr2 + "\"
				Call log("(Archief) Folder: " & flder & "(" & newFldr & ")" & " overgezet naar archief." & "(" & newFldr2 & ")", 0, logfile)
			Else
				Call log("Er is geen subfolder aanwezig.", 0, logfile)
			End if
		Next
	End If
End If
'--------------------------------------------------------------------------------------------------

Call log("Total of " & iProcessed & " records processed", 0, logfile)
Call log("program ended", 0, logfile)

if err.number <> 0 then
  call log ("Unhandled error " & err.number & " : " & err.description)
  err.clear
end if

set FSO = nothing
wscript.quit

'---------------------------------------------------------------------------------------
' Subs
'---------------------------------------------------------------------------------------

Sub WriteINIString(Section, KeyName, Value, FileName)
  Dim INIContents, PosSection, PosEndSection
  
  On Error Resume Next

  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    'Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    'Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Line = KeyName & "=" & Value
        Found = True
      End If
      NewsContents = NewsContents & Line & vbCrLf
    Next

    If isempty(Found) Then
      'key Not found - add it at the end of section
      NewsContents = NewsContents & KeyName & "=" & Value
    Else
      'remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    'Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    'Section Not found. Add section data at the end of file contents.
    If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
      INIContents = INIContents & vbCrLf 
    End If
    INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
      KeyName & "=" & Value
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
End Sub

'---------------------------------------------------------------------------------------

Function GetINIString(Section, KeyName, Default, FileName)
  Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
  
  On Error Resume Next

  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      'Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    End If
  End If
  If isempty(Found) Then Value = Default
  GetINIString = Value
End Function

'---------------------------------------------------------------------------------------

Function GetINISection(Section, Default, FileName)
  Dim INIContents, PosSection, PosEndSection, Value
  
  On error resume next

  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    value = Mid(INIContents, PosSection, PosEndSection - PosSection)
    value = REplace(value, "[" & Section & "]" & vbCrLf, "", vbTextCompare)
  Else
    value = default
  End If
  GetINISection = Value
End Function

'---------------------------------------------------------------------------------------

Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE

  On Error Resume Next

  PosB = InStr(1, sFrom, sStart, 1)

  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function

'---------------------------------------------------------------------------------------

Function GetFile(ByVal FileName)
  Dim FS

  On Error Resume Next

  Set FS = CreateObject("Scripting.FileSystemObject")
  'Go To windows folder If full path Not specified.
  If (InStr(1, FileName, ":\") = 0) And (Left(FileName, 2) <> "\\") Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  if FS.fileexists(filename) then
    GetFile = FS.OpenTextFile(FileName).ReadAll
  else
    GetFile = ""
  end if

  set FS = nothing
End Function

'--------------------------------------------------------------------------------------------------

Function WriteFile(ByVal FileName, ByVal Contents)
  Dim FS, OutStream

  On Error Resume Next

  Set FS = CreateObject("Scripting.FileSystemObject")

  'Go To windows folder If full path Not specified.
  If (InStr(1, FileName, ":\") = 0) And (Left(FileName, 2)<>"\\") Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  Set OutStream = FS.OpenTextFile(FileName, 2, True)
  OutStream.Write Contents
  OutStream.close
  Set FS = nothing
End Function

'---------------------------------------------------------------------------------------

Sub AddFile (sLine, Filename)
  DIM FS, objFSFile

  On Error Resume Next

  Set FS = Wscript.CreateObject("Scripting.FilesystemObject")

  If FS.FileExists(Filename) = True Then
    Set objFSFile = FS.OpenTextFile(Filename, 8, True)
    objFSFile.WriteLine(sLine)
    objFSFile.close
  Else
    Set objFSFile = FS.CreateTextFile(Filename)
    objFSFile.WriteLine(sLine)
    objFSFile.close
  End If
  err.clear

  Set objFSFile = nothing
  set FS = nothing
end sub

'---------------------------------------------------------------------------------------

Function Format(vExpression, sFormat) 
  Dim sResult

  On Error Resume Next

  sResult = vExpression
  With CreateObject("System.Text.StringBuilder")
    .AppendFormat "{0:" & sFormat & "}", sResult
    If Err=0 Then
      sResult = .toString
    else
      sResult = "-"
    end if           
  End With
  Format = sResult
End Function

'--------------------------------------------------------------------------------------------------

function ReplaceVar(ByVal sOrig)

  on error resume next

  ReplaceVar = Replace(sOrig, "%YYYY%", format(now(), "yyyy"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%yyyy%", format(now(), "yyyy"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%YY%", format(now(), "yy"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%yy%", format(now(), "yy"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%MMM%", format(now(), "MMM"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%MM%", format(now(), "MM"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%M%", format(now(), "M"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%DDDD%", format(now(), "dddd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%dddd%", format(now(), "dddd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%DDD%", format(now(), "ddd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%ddd%", format(now(), "ddd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%DD%", format(now(), "dd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%dd%", format(now(), "dd"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%D%", format(now(), "d"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%d%", format(now(), "d"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%hh%", format(now(), "HH"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%HH%", format(now(), "HH"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%NN%", format(now(), "mm"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%nn%", format(now(), "mm"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%mm%", format(now(), "mm"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%SS%", format(now(), "ss"), vbTextCompare)
  ReplaceVar = Replace(ReplaceVar, "%ss%", format(now(), "ss"), vbTextCompare)
end function

'--------------------------------------------------------------------------------------------------

function FileName(ByVal sOrig)

  on error resume next

  FileName = Replace(sOrig, "/", " ", vbTextCompare)
end function

'--------------------------------------------------------------------------------------------------

Sub Log (byval sWhat, iLevel, sfile)
  DIM FS, objFSFile

  on error resume next

  if bLog = vbTrue then
    Set FS = Wscript.CreateObject("Scripting.FilesystemObject")

    If iLevel = 1 then
      sWhat = "WARNING : " & sWhat
    End if
    If iLevel = 2 then
      sWhat = "ERROR : " & sWhat
    End if

    sWhat = format(now(), "yyyy-MM-dd HH:mm:ss") & " " & sWhat

    If FS.FileExists(ReplaceVar(sfile)) = True Then
      Set objFSFile = FS.OpenTextFile(ReplaceVar(sfile), 8, True)
    Else
      Set objFSFile = FS.CreateTextFile(ReplaceVar(sfile))
    End If
    objFSFile.WriteLine(sWhat)
    objFSFile.close

    Set objFSFile = nothing
    set FS = nothing
  end if
end sub

'--------------------------------------------------------------------------------------------------

Sub Mail (slSMTPserver, slMailto, slMailCC, slMailFrom, slOnderwerp, sBody, sHTMLbody, slAttachment, slAttachment2)
  Dim iMsg, iConf, Flds

  Const cdoSendUsingPort = 2

  on error resume next

  set iMsg = CreateObject("CDO.Message")
  set iConf = CreateObject("CDO.Configuration")

  Set Flds = iConf.Fields

' Set the CDOSYS configuration fields to use port 25 on the SMTP server.

  With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = slSMTPServer
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
    .Update
  End With

' Apply the settings to the message.
  With iMsg
    Set .Configuration = iConf
    .To = slMailTo
    if slMailCC <> "" then
      .Cc = slMailCC
    end if
    .From = slMailFrom
    .Subject = slOnderwerp
    .TextBody = sBody
    if sHTMLbody <> "" then
      .HTMLbody = sHTMLBody
    end if
    if slAttachment <> "" then
      .AddAttachment sScriptpath & "\" & slAttachment
    end if
    if slAttachment2 <> "" then
      .AddAttachment sScriptpath & "\" & slAttachment2
    end if

    .Send
  End With

' Clean up variables.
  Set iMsg = Nothing
  Set iConf = Nothing
  Set Flds = Nothing
End Sub

'------------------------------------------------------------------------

Function Decrypt(str, key)
  Dim lenKey, KeyPos, LenStr, x, Newstr
 
  on error resume next

  Newstr = ""
  lenKey = Len(key)
  KeyPos = 1
  LenStr = Len(Str)
 
  str=StrReverse(str)
  For x = LenStr To 1 Step -1
    Newstr = Newstr & chr(asc(Mid(str, x, 1)) - Asc(Mid(key, KeyPos, 1)))
    KeyPos = KeyPos+1
    If KeyPos > lenKey Then KeyPos = 1
  Next
  Newstr=StrReverse(Newstr)
  str = StrReverse(str)
  Decrypt = Newstr
End Function

'---------------------------------------------------------------------------------------

function GetPart(sFrom, iCount, sSeperator)

  Dim sResult2, sTemp2

  on error resume next

  sTemp2 = split(sFrom, sSeperator)
  if iCount <= uBound(sTemp2) then
    sresult2 = sTemp2(iCount)
  else
    sResult2 = "-"
  end if

  GetPart = sResult2
end function

'--------------------------------------------------------------------------------------------------

Function ConvertField(byval sTemp)

  on error resume next
  
  if (bConvertDecimal = vbTrue) AND isNumeric(sTemp) then
    sTemp = Replace(sTemp, ".", ",", vbTextCompare)
  end if
  ConvertField = sTemp
end function

'--------------------------------------------------------------------------------------------------

function SQLExec (byVal server, byVal dbUser, byVal dbPWD, byVal Database, byval sSQL, byval bReturn)
  Dim DB, sConn, rs
  Dim strResult
  Dim iTemp, sTemp

  on error resume next

  Set DB = CreateObject("ADODB.Connection")
  DB.ConnectionTimeout = 30
  DB.CommandTimeout = 60

  strResult = ""

  if (dbUser <> "") and (dbUser <> "-") then
    sConn = "Provider=SQLOLEDB;Data Source=" & server & ";Initial Catalog=" & database & ";User ID=" & dbUser & ";Password=" & dbPWD & ";"
  else
    sConn = "Provider=SQLOLEDB;Data Source=" & server & ";Trusted_Connection=Yes;Initial Catalog=" & database & ";"
  end if
  DB.Open sConn
  if err.number <> 0 then
    Call log("Error connecting to SQL-server " & server, 2, logfile)
    strResult = "Error connecting to SQL-server " & server
    err.clear
  else

    set rs = DB.Execute(sSQL)
    if err.number <> 0 then
      Call log("Error qeurying " & server & " (error " & err.number & ":" & err.description & ")", 2, logfile)
      strResult = "Error qeurying " & server
      err.clear
    else
     if bReturn = vbTrue then
        if not rs.eof then
          rs.movefirst
        end if

        if bAddHeader = vbTrue then
          for iTemp = 0 to (rs.fields.count - 1)
            if iTemp > 0 then
              strResult = strResult & cFieldSeperator & rs.fields(iTemp).name
            else
              strResult = rs.fields(iTemp).name
            end if
          next
          strResult = strResult & vbCrLf
        end if

        do while not rs.eof
          for iTemp = 0 to (rs.fields.count - 1)
            sTemp = rs.fields(iTemp).value
            if IsNull(sTemp) then
              sTemp = ""
            end if
            sTemp = Replace(cStr(sTemp), vbCR, " ", vbTextCompare)
            sTemp = Replace(sTemp, vbLf, " ", vbTextCompare)
            sTemp = ConvertField(sTemp)
            if itemp > 0 then
              strResult = strResult & cFieldSeperator & sTemp
            else
              strResult = strResult & sTemp
            end if
          next
		  rs.movenext
		  if rs.eof then
			strResult = strResult
		  else
		    strResult = strResult & vbCrLf
		  end if
        loop
      else
        if strResult = "" then
          strResult = "OK"
        end if
      end if
    end if
  end if
  DB.close
  Set DB = Nothing
  set rs = nothing

  SQLexec = strResult
end function

'--------------------------------------------------------------------------------------------------

function Table2HTML(byVal strInput)
  Dim sResult, sLine

  on error resume next

  sResult = "<table border=""1"">"
  strInput = split(strInput, vbcrLf)
  for each sLine in strInput
    if trim(sLine) <> "" then
      sResult = sresult & "<TR><TD>" & Replace(sLine, cFieldseperator, "</TD><TD>", vbTextCompare) & "</TD></TR>" & vbCrLf
    end if
  next
  sResult = sResult & "</table>" & vbCrLf & vbCrLf

  Table2HTML = sResult
end function

'--------------------------------------------------------------------------------------------------

Function GetFieldByName(byval strTemp, byval iRow, byval sField)

  Dim sTemp, sTemp2, iTemp, iTemp2
  Dim sResult

  on error resume next
  
  if isArray(strTemp) = vbFalse then
    strTemp = split(strTemp, vbCrLf)
  end if
  sTemp = split(strTemp(0), cFieldSeperator)
  iTemp = 0
  iTemp2 = -1
  for each sTemp2 in sTemp
    if uCase(trim(sTemp2)) = ucase(trim(sField)) then
      iTemp2 = iTemp
    end if
    iTemp = iTemp + 1
  next

  if iTemp2 > -1 then
    sTemp = strTemp(iRow)
    sResult = Getpart(sTemp, iTemp2, cFieldSeperator)
  else
    sResult = "-"
  end if
  GetFieldByName = sResult
end function

'--------------------------------------------------------------------------------------------------

Function IsFieldName(byval strTemp, byval sField)

  Dim sTemp, sTemp2
  Dim bResult

  on error resume next
  
  if isArray(strTemp) = vbFalse then
    strTemp = split(strTemp, vbCrLf)
  end if
  sTemp = split(strTemp(0), cFieldSeperator)

  bResult = vbFalse
  for each sTemp2 in sTemp
    if uCase(trim(sTemp2)) = ucase(trim(sField)) then
      bResult = vbTrue
    end if
  next

  IsFieldName = bResult
end function

'--------------------------------------------------------------------------------------------------

'function ReplaceValues(byval strInput, byval strValues, byval iRow)

'  Dim strResult
'  Dim sTemp, iTemp

'  on error resume next

'  strResult = ReplaceVar(strInput)
'  while instr(1, strResult, "%", vbBinaryCompare) > 0
'    iTemp = instr(1, strResult, "%", vbBinaryCompare)
'    sTemp = mid(strResult, iTemp + 1)
'    sTemp = left(sTemp, instr(1, sTemp, "%", vbBinaryCompare) - 1)
'    if IsFieldName(strValues, sTemp) then

'      strResult = Replace(strResult, "%" & sTemp & "%", GetFieldByName(strValues, iRow, sTemp), vbTextCompare)
'    end if
'  wend
'  ReplaceValues = strResult
'end function

'--------------------------------------------------------------------------------------------------

function ReplaceValues(byval strInput, byval strValues, byval iRow)

  Dim strResult
  Dim sTemp, sTemp2

  on error resume next

  strResult = ReplaceVar(strInput)

  if isArray(strValues) = vbFalse then
    strValues = split(strValues, vbCrLf)
  end if
  sTemp = split(strValues(0), cFieldSeperator)

  for each sTemp2 in sTemp
    strResult = Replace(strResult, "%" & sTemp2 & "%", GetFieldByName(strValues, iRow, sTemp2), vbTextCompare)
  next

  ReplaceValues = strResult
end function

'--------------------------------------------------------------------------------------------------

function ReplaceValuesHTTP(byval strInput, byval strValues, byval iRow)

  Dim strResult
  Dim sTemp, iTemp

  on error resume next

  strResult = ReplaceVar(strInput)
  while instr(1, strResult, "#", vbBinaryCompare) > 0
    iTemp = instr(1, strResult, "#", vbBinaryCompare)
    sTemp = mid(strResult, iTemp + 1)
    sTemp = left(sTemp, instr(1, sTemp, "#", vbBinaryCompare) - 1)
    strResult = Replace(strResult, "#" & sTemp & "#", GetFieldByName(strValues, iRow, sTemp), vbTextCompare)
  wend
  ReplaceValuesHTTP = strResult
end function

'--------------------------------------------------------------------------------------------------

function HTTPget(byval URL, byval sOutput, byval sOutput2)
  Dim bResult
  Dim oXMLHTTP
  Dim oStream

  on error resume next

  Call log("Getting file (" & sOutput & ") through HTTP " & URL , 0, logfile)

  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

  oXMLHTTP.Open "GET", URL, False
  oXMLHTTP.Send

  If oXMLHTTP.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write oXMLHTTP.responseBody
    oStream.SaveToFile sOutput, 2
    if sOutput2 <> "" then
      oStream.SaveToFile sOutput2, 2
      Call log("File (" & sOutput2 & ") created", 0, logfile)
    end if
    oStream.Close
    bResult = "OK"
  else
    bResult = "unsuccessfull"
  End If

  if err.number <> 0 then
    bResult = "error"
    Call log("Error getting file (" & sOutput & ") through HTTP " & URL & " (" & err.number & ":" & err.description & ")", 2, logfile)
    err.clear    
  end if

  HTTPget = bResult
end function 

'--------------------------------------------------------------------------------------------------
'	@desc: Subs en Functions.
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'--------------------------------------------------------------------------------------------------

Sub EmptyFolderCheck(byval subfolder)

	Dim objFSO, objFSOsubfolder, objFSOfiles
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFSOsubfolder = objFSO.GetFolder(subfolder)
	objFSOfiles = objFSOsubfolder.Files.Count
	
	if objFSOfiles = 0 then
		objFSO.DeleteFolder(objFSOsubfolder)
		Call log(objFSOsubfolder & " is leeg en heeft geen bestanden", 2, logfile)
	else 
		Call log(objFSOsubfolder & " is gevuld en heeft bestanden", 0, logfile)
	end if

End sub

'--------------------------------------------------------------------------------------------------
'	@desc: Zipping sub routine waarmee het zip bestand wordt aangemaakt.
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'--------------------------------------------------------------------------------------------------
Sub Zipper(byval zipfile, byval sFolder)
    With CreateObject("Scripting.FileSystemObject")
        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With
	
    With CreateObject("Shell.Application")
		.NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items
			
		Do Until .NameSpace(zipFile).Items.Count = _
				 .NameSpace(sFolder).Items.Count
			WScript.Sleep 1000
		Loop
    End With
	
	objFSO = nothing
	objFSOsubfolder = nothing
	objFSOfiles = nothing
End Sub

'--------------------------------------------------------------------------------------------------
'	@desc: Sub Routine waarmee er een directory wordt aangemaakt op een locatie. Deze directory wordt dan
'		   Gebruikt om de pdfs (reports) in  te kunnen plaatsen.
'	@author: Iliass Nassibane.
'	@date: 18-04-2017.
'-------------------------------------------------------------------------------------------------

Sub CreateFolder(byval newFolder)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CreateFolder(newFolder)
	fso = ""
End Sub

'--------------------------------------------------------------------------------------------------
'	@desc: Function voor het vervangen van een karakter in een string
'	@author: Iliass Nassibane.
'	@date: 04-05-2017.
'-------------------------------------------------------------------------------------------------

Sub csvFormatter(ByVal filepath, ByVal valueToReplace, ByVal replacementVal)
	Dim File1, inFile, strg
	Const ForReading = 1
	Const ForWriting = 2
	
	Set File1 = CreateObject("scripting.FileSystemObject")
	Set inFile = File1.OpenTextFile(filepath, ForReading)
	
	strg = inFile.ReadAll
	inFile.Close
	
	strg = Replace(strg, valueToReplace, replacementVal)
	
	Set inFile = File1.OpenTextFile(filepath, ForWriting)

	inFile.WriteLine strg

	inFile.Close
End Sub