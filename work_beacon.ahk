; IMPORTANT INFO ABOUT GETTING STARTED: Lines that start with a
; semicolon, such as this one, are comments.  They are not executed.

; This script has a special filename and path because it is automatically
; launched when you run the program directly.  Also, any text file whose
; name ends in .ahk is associated with the program, which means that it
; can be launched simply by double-clicking it.  You can have as many .ahk
; files as you want, located in any folder.  You can also run more than
; one .ahk file simultaneously and each will get its own tray icon.

; SAMPLE HOTKEYS: Below are two sample hotkeys.  The first is Win+Z and it
; launches a web site in the default browser.  The second is Control+Alt+N
; and it launches a new Notepad window (or activates an existing one).  To
; try out these hotkeys, run AutoHotkey again, which will load this file.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#UseHook On
#InstallKeybdHook
#InstallMouseHook
#SingleInstance force
#Persistent
Menu, tray, add
Menu, tray, add, Checkout, do_check_out
SetTitleMatchMode, 2

global LocationRemote := false
global CheckedIn := false
global CheckedOut := true
global Checkin_time
global Checkout_time
global HWIdle
global SWIdle
global MANUAL_checked_inout := false

SetTimer, work_beacon, 10000

#z::Run www.autohotkey.com

!y::
	MANUAL_checked_inout := true
	gosub do_check_out
return

^!n::
IfWinExist Untitled - Notepad
	WinActivate
else
	Run Notepad
return



work_beacon:
        FormatTime, TimeString,, Time
	FormatTime, HWIdle, A_TimeIdlePhysical, ss
	FormatTime, SWIdle, A_TimeIdle, ss
	Checkout_timestamp:=A_Now
	FormatTime, Checkout_time, T12, hh:mm
	FormatTime, Checkout_date,, MM/dd/yyyy
	Checked_durraw := A_Now
	EnvSub, Checked_durraw, Checkin_timestamp, minutes
	;FormatTime, Checkout_dur, Checkout_durraw, m
	Checkout_dur:=Checked_durraw

WinGet, active_id, ID, A
WinGetTitle, active_win, A
	;if (active_id == "")
	;	FileAppend, screen: %wSC_result%  idle: %A_TimeIdle% phys: %A_TimeIdlePhysical%  window: %active_win%  wi_id: %active_id% empty`n, C:\work_beacon.log
	;else
	;	FileAppend, screen: %wSC_result%  idle: %A_TimeIdle% phys: %A_TimeIdlePhysical%  window: %active_win%  wi_id: %active_id% nempty`n, C:\work_beacon.log
	;ToolTip, A_TimeIdle: %A_TimeIdle%`nA_TimeIdlePhysical: %A_TimeIdlePhysical%`nwSC_result:%wSC_result%

	if (CheckedOut){
		FormatTime, Checkin_time, T12,hh:mm
		Checkin_timestamp := A_Now	
		FormatTime, Checkin_date,, MM/dd/yyyy
	}

        ;Tag = DFV64Q1
	If (active_id == "")
	{ 
		if (CheckedIn) {
			gosub do_check_out
		}
	}else If (A_TimeIdle <  2700000)
	{		
		If (A_TimeIdlePhysical > 2700000)
		{
			if (CheckedOut){
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "START [Home Office] %Checkin_date% %Checkin_timeampm%", C:\Users\helgelaurisch\Eigene` Programme\googlecl\, hide
				;Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "Home Office" --when "%Checkin_date% %Checkin_time%" --duration 30 --description "Aufgab" --reminder 0 -v add, C:\Python27\Scripts\, hide
;Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "Home Office" --when "%Checkin_date% %Checkin_time%" --duration %Checkout_dur% --description "Aufgabe" --reminder 0 -v add, C:\Python27\Scripts\, hide
				UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?start_date=%Checkin_date% %Checkin_time%&title=MOL IT START HO&where=Home Office&description=Aufgabe&client=Madsack Online&project=IT Consulting&cmd=checkin, C:\work_beacon.txt
				CheckedOut := false
				CheckedIn := true
				LocationRemote := true
				ToolTip,User is working remote`nRemote: %A_TimeIdle% < 900000 `nLocal:  %A_TimeIdlePhysical% > 900000`n[Home Office] von %Checkin_time% Uhr (%Checkin_timeampm%) bis %Checkout_time% Uhr (%Checkout_timeampm%) 
			}else{
				;ToolTip,[Home Office] von %Checkin_time% Uhr bis %Checkout_time% Uhr (%Checkout_dur%)
			}
			SetTimer, clean_tool, 5000			
		}else{
			if (CheckedOut){
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "START [vor Ort] %Checkin_date% %Checkin_timeampm%", C:\Users\helgelaurisch\Eigene` Programme\googlecl\, hide
				;Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "MOL on site" --when "%Checkin_date% %Checkin_time%" --duration 30 --description "Aufgabe" --reminder 0 -v add, C:\Python27\Scripts\, hide
;Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "MOL on site" --when "%Checkin_date% %Checkin_time%" --duration %Checkout_dur% --description "Aufgabe" --reminder 0 -v add, C:\Python27\Scripts\, hide
				;UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?work_beacon=madsack_laptop&location=remote&cmd=checkin&time=%A_TimeIdle%&physical=%A_TimeIdlePhysical%, C:\work_beacon.txt
				UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?start_date=%Checkin_date% %Checkin_time%&title=MOL IT START OS&where=MOL on site&description=Aufgabe&client=Madsack Online&project=IT Consulting&cmd=checkin, C:\work_beacon.txt
				CheckedOut := false
				CheckedIn := true
				LocationRemote := false
				ToolTip,User is working on site`nRemote: %A_TimeIdle% < 900000 `nLocal:  %A_TimeIdlePhysical% > 900000`n[vor Ort] von %Checkin_time% Uhr (%Checkin_timeampm%) bis %Checkout_time% Uhr (%Checkout_timeampm%) 
			}else{
				;ToolTip,[vor Ort] von %Checkin_time% Uhr bis %Checkout_time% Uhr (%Checkout_dur%)
			}
			SetTimer, clean_tool, 5000
		}
	}
	else
	{
		if (CheckedIn) {
			gosub do_check_out
		}
	}
return

clean_tool:
	ToolTip
	SetTimer, clean_tool, Off
return

do_check_out:

			;Run, %comspec% /c set https_proxy=https://193.30.60.30:3128 && C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "von %checkin_time% Uhr bis %checkout_time% Uhr [vor Ort] Tätigkeit"
			
			if(LocationRemote){
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "[Home Office] [%Checkin_time% - %Checkout_time%] Aufgabenbeschreibung von %Checkin_time% Uhr bis %Checkout_time% Uhr"
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "END [Home Office] Heute %Checkout_time% Uhr", C:\Users\helgelaurisch\Eigene` Programme\googlecl\, hide
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog -t "Aufgabenbeschreibung" "Aufgabenbeschreibung [Home Office] %Checkin_date% %Checkin_timeampm% - %Checkout_timeampm%"
				Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "Home Office" --when "%Checkin_date% %Checkin_time%" --duration %Checkout_dur% --description "Aufgabe" --reminder 0 -v add, C:\Python27\Scripts\, hide
				UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?end_date=%Checkout_date% %Checkout_timea%&title=MOL IT&where=Home Office&description=Aufgabe&client=Madsack Online&project=IT Consulting&cmd=checkout, C:\work_beacon.txt
				ToolTip, [MOL IT: home office] von %Checkin_time% Uhr bis %Checkout_time% Uhr
			}else{
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "[vor Ort] [%Checkin_time% - %Checkout_time%] Aufgabenbeschreibung von %Checkin_time% Uhr bis %Checkout_time% Uhr"
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog "END [vor Ort] at 'Madsack vor Ort' %Checkout_time% today ", C:\Users\helgelaurisch\Eigene` Programme\googlecl\, hide
				;Run, C:\Users\helgelaurisch\Eigene` Programme\googlecl\google.exe calendar add --cal Worklog -t "Aufgabenbeschreibung" "Aufgabenbeschreibung [vor Ort] %Checkin_date% %Checkin_timeampm% - %Checkout_timeampm%"
				Run, C:\Python27\Scripts\gcalcli.py --title "MOL IT" --where "MOL on site" --when "%Checkin_date% %Checkin_time%" --duration %Checkout_dur% --description "Aufgabe" --reminder 0 -v add, C:\Python27\Scripts\, hide
				UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?end_date=%Checkout_date% %Checkout_timea%&title=MOL IT&where=MOL on site&description=Aufgabe&client=Madsack Online&project=IT Consulting&cmd=checkout, C:\work_beacon.txt
				ToolTip, [MOL IT: on Premises] von %Checkin_time% Uhr bis %Checkout_time% Uhr
			}
			;UrlDownloadToFile, *0 http://www.acumenec.de/api/index.php?work_beacon=madsack_laptop&location=remote&cmd=checkout&time=%A_TimeIdle%&physical=%A_TimeIdlePhysical%, C:\work_beacon.txt
			CheckedOut := true
			CheckedIn := false
			FormatTime, Checkin_time,T12,hh:mm
			Checkin_timestamp := A_Now	
			FormatTime, Checkin_date,, MM/dd/yyyy
			SetTimer, clean_tool, 5000
return

; Note: From now on whenever you run AutoHotkey directly, this script
; will be loaded.  So feel free to customize it to suit your needs.

; Please read the QUICK-START TUTORIAL near the top of the help file.
; It explains how to perform common automation tasks such as sending
; keystrokes and mouse clicks.  It also explains more about hotkeys.


;---------------------------------------------------------------------------
; SCRIPT-SPECIFIC HOTKEYS
; --------------------------------------------------------------------------

; All commands in this section are also available from the system tray
; menu. Edit opens the script in Notepad, or in the associated editor.
!^#e::Edit     ; Edit the script by Alt+Ctrl++Win+E.
!^#s::Suspend  ; Toggle hotkeys set by the script by Alt+Ctrl++Win+S.
!^#r::
        Menu,Tray, Icon, %my_ICONDIR%\info.ico
        Reload   ; Reload the script by Alt+Ctrl++Win+R.
return

getScreensaver() { 
 RegRead, op, HKEY_CURRENT_USER, Control Panel\Desktop, ScreenSaveActive 
return, op 
} 

;!^#x::ExitApp  ; Terminate the script by Alt+Ctrl++Win+X.
