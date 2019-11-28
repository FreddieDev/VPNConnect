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
	; Get one time passcode from MobilePASS

	Process, Close, MobilePASS.exe
	sleep 50

	Run "C:\Program Files (x86)\SafeNet\Authentication\MobilePASS\MobilePASS.exe"
	sleep 500

	SetControlDelay -1
	; ControlClick, x22 y103, MobilePASS,, LEFT, 1
	WinActivate MobilePASS
	MouseClick, left, 130, 120, 1, 0
	sleep, 50

	ControlSetText, Edit1, %TokenPin%, MobilePASS ; Enter token
	sleep, 50
	ControlClick, Continue, MobilePASS ; Press continue
	sleep 1000

	; Save OTP to var
	ControlGetText, OTP, Edit1, MobilePASS
	sleep 50

	; Kill MobilePASS
	Process, Close, MobilePASS.exe
	
	return OTP
}

; Connects Cisco VPN using your one-time-password
ConnectInCisco(OTP) {
	MainMenuTitle := "Cisco AnyConnect Secure Mobility Client"
	ConnectingTitle := "Cisco AnyConnect | "
	TermsTitle := "Cisco AnyConnect"

	RunWait "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
	WinWait, %MainMenuTitle%
	sleep 200
	ControlClick, Connect, %MainMenuTitle% ; Press connect

	WinWaitActive, %ConnectingTitle% ; Wait for password window to show
	sleep 500
	Send, %OTP% ; Enter token
	sleep 50
	ControlClick, OK, %ConnectingTitle% ; Press OK


	WinWait, %TermsTitle%, Disconnect ; Wait for password window to show
	sleep 500
	ControlClick, Accept, %TermsTitle%, Disconnect ; Press Accept
}

; Accepts Windows Security login windows (Skype/Outlook etc)
AcceptSecurityWindows() {
	WinActivate Windows Security
	sleep 200
	MouseClick, left, 120, 275, 1, 0 ; Tick checkbox
	MouseClick, left, 160, 355, 1, 0 ; Press OK
}

; Abort script if connected to office WiFi
if (IsAtWork()) { 
	MsgBox, You can't connect to the VPN from this WiFi network!
	ExitApp
}

; Auto-quit if script doesn't finish fast enough (likely got stuck)
SetTimer, timeOut, 15000


OTP := GetOTP(YourTokenPin)
sleep 100
ConnectInCisco(OTP)
sleep 200
AcceptSecurityWindows()



return

timeOut:
	ExitApp