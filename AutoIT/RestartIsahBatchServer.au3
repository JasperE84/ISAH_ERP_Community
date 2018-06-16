#cs ----------------------------------------------------------------------------

 Author:
 	JasperE - https://github.com/JasperE84/ISAH_ERP_Community/

 Tested with:
        ISAH v4.7.1

 License:		
        GNU General Public License version 3
	https://www.gnu.org/licenses/gpl-3.0.en.html

	You may copy, distribute and modify the software as long as you track
	changes/dates in source files. Any modifications to or software including
	(via compiler) GPL-licensed code must also be made available under the GPL
	along with build & install instructions.
        
  Purpose:
        In case a SQL stored proc called by a batch server jobs generates a significant 
        amount of warnings, the ISAH ERP batch server functionality sometimes has issues  
        marking the current or subsequent batch jobs in the queue as active/finished.
        
        Negative effects on subsequent batch jobs can be remediated by periodically 
        restarting the ISAH process in which the batch server runs. This script does 
        this in a relatively safe manner, avoiding terminating the ISAH process in 
        the midst of an active job.
        
        If a batch job is still active, the script will wait for the job to finish.
        
  Instructions:
        - Replace the text ISAHUSERNAME, ISAHPASSWORD & ISAHDBNAME to match your 
          configuration
        - In ISAH's main screen, add a favourite to automatically start batch screen (0590) 
          upon login of batch user 
        - In ISAH's batch dialog (0590), add another favorite with Ctrl+B shortcut to start 
          batch processing
        - Install AutoIT, compile this script, and schedule it with Windows task scheduler

#ce ----------------------------------------------------------------------------

#include <Date.au3>
LogToFile("Script " & @ScriptName & " started")

If WinExists("[CLASS:TFRM0618]") Then
   LogToFile("Batch processing dialog actief, controleer of er nu een batch actief is.")

   While ControlCommand("[CLASS:TFRM0618]", "", "[CLASS:TMemo; INSTANCE:1]", "IsVisible") = 1
	  LogToFile("Actieve batchopdracht, wacht 1 minuut...")
	  Sleep(60000)
   WEnd

   ; Sluit het huidige batchserverdialoog af
   LogToFile("Geen actieve batchopdracht, probeer batchverwerkingsdialoog te sluiten.")
   If WinClose("[CLASS:TFRM0618]") = 1 Then
	  LogToFile("Batchverwerkingsdialoog met success gesloten.")
   Else
	  LogToFile("Kon batchverwerkingsdialoog niet sluiten.")
   EndIf
   Sleep(1000)
Else
   LogToFile("Batchverwerkingsdialoog was niet actief.")
EndIf

; Sluit het Isah hoofdscherm af
If WinExists("[CLASS:TfrmMainMenu]") Then
   LogToFile("Probeer Isah hoofdscherm te sluiten.")

   If WinClose("[CLASS:TfrmMainMenu]") = 1 Then
	  LogToFile("Isah hoofdscherm met success gesloten.")
   Else
	  LogToFile("Kon Isah hoofdscherm niet sluiten, timeout.")
   EndIf

   Sleep(5000)
EndIf

; Forceer isah sluiten als nog draait
If ProcessExists("Isah7.exe") Then
   LogToFile("Isah7.exe proces nog actief, probeer geforceerd te sluiten.")

   Local $secondsWaited = 0
   While ProcessExists("Isah7.exe") And $secondsWaited < 120
	  $secondsWaited = $secondsWaited + 1
	  Local $closeResult = ProcessClose("Isah7.exe")

	  ; ProcessClose() failure? Probeer dan niet opnieuw, zou kunnen voorkomen wanneer er meerdere gebruikers op systeem zijn ingelogd die isah7 hebben draaien
	  If $closeResult <> 1 Then
		 Switch @error
			Case 1
			   LogToFile("Kon Isah7.exe niet sluiten, OpenProcess failed")
			Case 2
			   LogToFile("Kon Isah7.exe niet sluiten, AdjustTokenPrivileges Failed")
			Case 3
			   LogToFile("Kon Isah7.exe niet sluiten, TerminateProcess Failed")
			Case 4
			   LogToFile("Kon Isah7.exe niet sluiten, Cannot verify if process exists")
			Case Else
			   LogToFile("Kon Isah7.exe niet sluiten, errorcode: " & @error)
		 EndSwitch
	  EndIf

	  Sleep(1000)
   WEnd

   If ProcessExists("Isah7.exe") Then
	  LogToFile("FOUT: Kon Isah7.exe definitief niet sluiten")
	  Exit 1
   EndIf
EndIf

; Start Isah met autologin
LogToFile("Start Isah opnieuw op met autologin.")
ShellExecute("C:\Program Files (x86)\Isah\Isah7\Progs\Isah7.exe","/usr:ISAHUSERNAME /pwd:ISAHPASSWORD /dba:""ISAHDBNAME""")

; Wacht op beschikbaar komen hoofdscherm
Local $secondsWaited = 0
While Not WinExists("[CLASS:TfrmMainMenu]") And $secondsWaited < 120
   Sleep(1000)
   $secondsWaited = $secondsWaited + 1
WEnd
If Not WinExists("[CLASS:TfrmMainMenu]") Then
   LogToFile("FOUT: Na twee minuten nog geen handle gevonden naar het Isah hoofdscherm.")
   LogToFile("- Werken de logingegevens in de ShellExecute() parameters hierboven?.")
   Exit 1
Else
   LogToFile("Handle gevonden naar het Isah hoofdscherm");
EndIf

; Wacht op beschikbaar komen batchoverzichtscherm
$secondsWaited = 0
While Not WinExists("[CLASS:TFRM0590]") And $secondsWaited < 120
   Sleep(1000)
   $secondsWaited = $secondsWaited + 1
WEnd
If Not WinExists("[CLASS:TFRM0590]") Then
   LogToFile("FOUT: Na twee minuten nog geen handle gevonden naar het Isah batchopdrachtenscherm 0590.")
   LogToFile("- Is via favorieten op het isah hoofdmenu ingesteld dat batchopdrachten automatisch geopend worden?.")
   Exit 1
Else
   LogToFile("Handle gevonden naar het Isah batchopdrachtenscherm.")
EndIf

; In het batchopdrachten scherm is de sneltoets Ctrl+B via isah favorieten gekoppeld aan bewerken->batchverwerking
; Stuur deze toetscombinatie
; Controleer of batch overzicht form helemaal geinitialiseerd is
$secondsWaited = 0
While Not ControlCommand("[CLASS:TFRM0590]", "", "[CLASS:TcxGridSite; INSTANCE:1]", "IsEnabled", "") And $secondsWaited < 120
   Sleep(1000)
   $secondsWaited = $secondsWaited + 1
WEnd

; Wacht nog 5 seconden om zeker te weten dat het scherm klaar is voor snelkoppelingen (dit bleek helaas proefondervindelijk nodig te zijn)
Sleep(5000)

LogToFile("Sneltoets Ctrl+B wordt verstuurd om de batchverwerking te starten.")
Local $sendResult = ControlSend("[CLASS:TFRM0590]", "", "", "^b")
If $sendResult = 1 Then
   LogToFile("Ctrl+B met succes verstuurd.")
Else
   LogToFile("Kon Ctrl+B niet versturen naar scherm FRM0590.")
   Exit 1
EndIf

LogToFile("Wacht op batchverwerkingsscherm...")
$secondsWaited = 0
While Not WinExists("[CLASS:TFRM0618]") And $secondsWaited < 120
   Sleep(1000)
   $secondsWaited = $secondsWaited + 1

   If $secondsWaited = 60 Then
	  Local $sendResult = ControlSend("[CLASS:TFRM0590]", "", "", "^b")
	  If $sendResult = 1 Then
		 LogToFile("Ctrl+B met succes nogmaals verstuurd, 1 minuut nadat vorige niet reageerde.")
	  Else
		 LogToFile("Kon Ctrl+B niet (nogmaals) versturen naar scherm FRM0590.")
		 Exit 1
	  EndIf
   EndIf
WEnd

If WinExists("[CLASS:TFRM0618]") = 0 Then
   LogToFile("FOUT: Batchverwerkingscherm bestaat niet na 2 minuten.")
Else
   LogToFile("SUCCES: Batchverwerkingscherm bestaat.")
EndIf

; Logfunctie
Func LogToFile($Data, $FileName = -1, $TimeStamp = True)
   If $FileName = -1 Then
	  $FileName = @ScriptDir & '\' & @ScriptName & '.Log.txt'
   EndIf
   $hFile = FileOpen($FileName, 1)
   If $hFile <> -1 Then
	  If $TimeStamp = True Then $Data = _Now() & ' - ' & $Data
	  FileWriteLine($hFile, $Data)
	  FileClose($hFile)
   EndIf
EndFunc

