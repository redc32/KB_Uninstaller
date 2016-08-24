'--------------------------------------------------------------------------------------------------------------------
'This script work with:													                             				|
'Client -> min. Windows XP																							|
'Server -> min. Windows Server 2003																					|
'																													|
'The script searches all uninstalled updates and uninstall the update indicted by hiding it in a way 				|
'that is no longer installed. At the end requires a computer restart.												|
'--------------------------------------------------------------------------------------------------------------------

Dim FirstPrompt, FileSystem, hideupdates, NetworkNum
Dim SecondPrompt, ThirdPrompt, Dummy, Check

'Define initial message
FirstPrompt = "Inserire numero della KB da disinstallare e nascondere dai successivi aggiornamenti" & vbCrLf & vbCrLf &_
"Usare la forma 3102429" & vbCrLf & vbCrLf &_
"Assicurarsi di non scrivere KB."

'Check variable for match found or not
Check = 0

Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSystem = CreateObject("Scripting.FileSystemObject")

'First message display with inputbox, setting and popups second message
NetworkNum = InputBox(FirstPrompt,"Ping Automation Tool")
SecondPrompt = "L'operazione di disinstallazione della KB" &NetworkNum & " potrebbe richiedere alcuni minuti. Attendere popup di terminazione procedura."
Dummy = WshShell.Popup (SecondPrompt,1,"Ping Automation Tool",64)

'session creation for search installed software updates
set updateSession = createObject("Microsoft.Update.Session")
set updateSearcher = updateSession.CreateupdateSearcher()

'Search windows updates not hidden, installed and related software

WshShell.Run "Cmd.exe /c start wusa /uninstall /kb:"&NetworkNum &" /quiet /norestart", 0,True

WScript.Echo "Disinstallazione KB completata. Inizio fase di disabilitazione KB"

'Setting KB fullname, as opposed to NetworkNum that only has the number
hideupdates = "KB" &NetworkNum

'Checks for updates to be installed
Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'")
For i = 0 To searchResult.Updates.Count-1
	set update = searchResult.Updates.Item(i)
	'MsgBox hideupdates
	if instr(1, update.Title, hideupdates, vbTextCompare) <> 0 then
		Wscript.echo "Hiding " & hideupdates
		update.IsHidden = True
		Check = 1
		Exit For
	end if
Next

If Check = 0 then 
	Wscript.echo "No match found for " & hideupdates
	WScript.Echo "Procedura completata, buon lavoro."
else
	'Set up the messages
	'
	strText = "E' necessario riavviare il computer per rendere effettive le modifiche. Riavviare ora?"
	strTitle = "Riavvio Computer"
	intType = vbYesNo + vbQuestion + vbDefaultButton2
	'
	' Now display it
	'
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	intResult = objWshShell.Popup(strText, ,strTitle, intType)
	'
	' Process the result
	'
	Select Case intResult
		Case vbYes
			strComputer = "." ' Local Computer
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" &strComputer & "\root\cimv2")
			Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
			For Each objOS in colOS
				objOS.Reboot()
			Next
		Case vbNo
			WScript.Echo "Ricordati di riavviare in seguito il computer"
			WScript.Echo "Procedura completata, buon lavoro."
	End Select
end if

Wscript.Quit