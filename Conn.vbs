Option Explicit
Wscript.Timeout = 600000
Dim ie, WshShell, btn, a, TESTE

'Funções
Sub WaitForLoad
 Do While ie.Busy
  WScript.Sleep (1000)
 Loop
End Sub

'Status da internet
Function internet
Set ie=CreateObject("InternetExplorer.Application")
ie.Visible=0
ie.Navigate "www.google.com.br"
Call WaitForLoad
For Each btn In IE.Document.getElementsByTagName("button")
    If btn.id = "diagnose" Then 
	internet = False
    else
	internet = True
    End IF
Next
ie.quit
End Function

'release e renew
Function Renovar
Set WshShell=CreateObject("wscript.shell")
	WScript.Sleep (3000)
	a = WshShell.Run("ipconfig /flushdns", 0, True)
	a = WshShell.Run("ipconfig /release", 0, True)
	a = WshShell.Run("ipconfig /renew",0, True)	
Set WshShell = Nothing
End Function

Function Rede
Set WshShell=CreateObject("wscript.shell")
	WScript.Sleep (3000)
	a = WshShell.Run("C:\Users\eXplosive\Downloads\DevManView.exe /disable ""ASIX AX88179 USB 3.0 to Gigabit Ethernet Adapter""", 0, True)
	a = WshShell.Run("C:\Users\eXplosive\Downloads\DevManView.exe /enable ""ASIX AX88179 USB 3.0 to Gigabit Ethernet Adapter""", 0, True)
Set WshShell = Nothing
End Function

Function Restart
Set WshShell=CreateObject("wscript.shell")
	WScript.Sleep (3000)
	a = WshShell.Run("shutdown -s -f -t 60 -c ""Sem rede""", 0, True)	
Set WshShell = Nothing
End Function

TESTE = "rede"

do 	
	if internet = False then
		if TESTE = "rede" then
			Rede
			WScript.Sleep (10000)
			if internet = False then
				TESTE = "renovar"
			else
				TESTE = "rede"
			end if
		end if
		if TESTE = "renovar" then
			Renovar
			WScript.Sleep (10000)
			if internet = False then
				TESTE = "restart"
			else
				TESTE = "rede"
			end if
 		end if
		if TESTE = "restart" then
			Restart
		end if					
	else
		TESTE = "rede"
	end if

	WScript.Sleep (600000)
loop