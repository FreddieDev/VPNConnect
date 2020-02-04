#SingleInstance force ; Don't allow the script to run multiple times at once and don't warn about replacements
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
YourTokenPin := "" ; Default pin to use (unsafe to save one here)

; Change the tray icon
GLOBE_ICON := 14
Menu, Tray, Icon, shell32.dll, %GLOBE_ICON%

; Ask user for their pin (if no default set)
if (YourTokenPin = "") {
	firstRun := 1
	; Keep asking until user provides a valid token
	while (StrLen(YourTokenPin) != 4) {
		if (!firstRun) {
			MsgBox, Invalid pin - your pin must be 4 digits!
		}
		InputBox, YourTokenPin, MobilePASS, Enter your MobilePASS pin to connect the Cisco VPN,hide,310,150
		
		; Exit if user presses cancel
		if (ErrorLevel) {
			ExitApp
		}
		
		firstRun = 0
	}
}

; Returns true if your WiFi is the office WiFi with intranet
IsAtWork() {
	ClipSaved := ClipboardAll ; Backup clipboard

	; Extract WiFi name from CMD into clipboard into AHK var
	Runwait %comspec% /c netsh wlan show interface | clip,,hide
	WiFiSSID := RegExReplace(clipboard, "s).*?\R\s+SSID\s+:(\V+).*", "$1")

	Clipboard := ClipSaved ; Restore clipboard
	
	whiteListedWiFi := "XS4OFFICE"
	return Trim(WiFiSSID) == whiteListedWiFi
}

; Extracts a one-time-password from MobilePASS using your pin
GetOTP(TokenPin) {
	
	; Close MPass to return to main menu on next launch
	Process, Close, MobilePASS.exe
	sleep 50

	Run, "C:\Program Files (x86)\SafeNet\Authentication\MobilePASS\MobilePASS.exe"
	sleep 1000

	SetControlDelay -1
	; ControlClick, x22 y103, MobilePASS,, LEFT, 1
	WinActivate MobilePASS
	MouseClick, left, 130, 120, 1, 0
	sleep, 50

	ControlSetText, Edit1, %TokenPin%, MobilePASS ; Enter token
	sleep, 50
	ControlClick, Continue, MobilePASS ; Press continue
	sleep 2000

	; Save OTP to var
	ControlGetText, OTP, Edit1, MobilePASS
	sleep 50

	; Kill MobilePASS
	Process, Close, MobilePASS.exe
	
	return OTP
}


; Experimental func to use CMD rather than UI to connect to VPN
; NOTE: 2nd command (OTP) isn't reliably being used
ConnectFromCMD(OTP) {
	; Close active VPN tray UI
	Process, Close, vpnui.exe
	
	; Open new CMD Window and pipe-in username and OTP into VPN app
	VPNPath := "%ProgramFiles(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe"
	CMDToRun := "@echo %~2|@""" . VPNPath . """ connect sslvpnuk.capgemini.com/SSLVPN-Client||" . A_UserName . "||" . OTP
	RunWait, cmd.exe /k %CMDToRun%
}

; Connects Cisco VPN using your one-time-password
ConnectInCisco(OTP) {
	MainMenuTitle := "Cisco AnyConnect Secure Mobility Client"
	ConnectingTitle := "Cisco AnyConnect | "
	TermsTitle := "Cisco AnyConnect"

	RunWait "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
	WinWait, %MainMenuTitle%

	ControlGet, NotConnecting, Enabled,, Connect, %MainMenuTitle%
	
	; Only click connect if the button is enabled
	; Otherwise, forcing this causes "Already connecting" popup error
	if (NotConnecting)
		ControlClick, Connect, %MainMenuTitle% ; Press connect

	WinWaitActive, %ConnectingTitle% ; Wait for password window to show
	Send, %OTP% ; Enter token
	sleep 50
	ControlClick, OK, %ConnectingTitle% ; Press OK


	WinWait, %TermsTitle%, Disconnect ; Wait for password window to show
	ControlClick, Accept, %TermsTitle%, Disconnect ; Press Accept
}

; Accepts Windows Security login windows (Skype/Outlook etc)
AcceptSecurityWindows() {
	WinActivate Windows Security
	sleep 200
	if (WinActive("Windows Security")) {
		MouseClick, left, 120, 275, 1, 0 ; Tick remember me checkbox
		MouseClick, left, 160, 355, 1, 0 ; Press OK
	}
}

; Abort script if connected to office WiFi
if (IsAtWork()) { 
	MsgBox, You can't connect to the VPN from this WiFi network!
	ExitApp
}

; Auto-quit if script doesn't finish fast enough (likely got stuck)
SetTimer, timeOut, 15000


OTP := GetOTP(YourTokenPin)
; ConnectFromCMD(OTP)
ConnectInCisco(OTP)
AcceptSecurityWindows()



return

timeOut:
	ExitApp