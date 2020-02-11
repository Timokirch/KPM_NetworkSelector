#comments-start
   Quest KPM Network Selector 1.1
   Created by Timo Weberskirch (timo.weberskirch@quest.com)

   **ChangeLog**
   Version 1.1 01/30/2020
	  resolved an issue with spacing in the name of the network interface
   Version 1.0 01/23/2020
	  stable release

#comments-end
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=kace.ico
#AutoIt3Wrapper_Outfile=KPM_NetworkSelector_x86.exe
#AutoIt3Wrapper_Outfile_x64=KPM_NetworkSelector_x64.exe
#AutoIt3Wrapper_Compile_Both=Y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=For switching the network interfaces
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
#AutoIt3Wrapper_Res_ProductVersion=1.1
#AutoIt3Wrapper_Res_LegalCopyright=Quest Software | Timo Weberskirch
#AutoIt3Wrapper_Res_Field=ProductName|Quest KACE SDA Network Selector
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;includes
#include <ButtonConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <file.au3>
#include <array.au3>

;options
Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
Opt("GUIEventOptions", 1)

;ui prep
If Not FileExists(@TempDir & '\KPMNetworkSelector') Then DirCreate(@TempDir & '\KPMNetworkSelector')
FileInstall('logo.jpg', @TempDir & '\KPMNetworkSelector\logo.jpg', 1)

;define global variables
Global $swtitel="KPM Network Selector"
Global $swversion="1.1"
$sHost = @ComputerName
Global $aNetworkAdapters = WMI_ListAllNetworkAdapters($sHost)

Gui()

Func Gui()
	;Create mainwindow wit controls
	Local $intGUIlgth = UBound($aNetworkAdapters) * 20 + 180
	Local $hGUI = GUICreate($swtitel, 360, $intGUIlgth)
	GUISetFont(8, 400, 0, "Arial")
	GUISetBkColor(0xFFFFFF)
	Local $idLOGO = GUICtrlCreatePic(@TempDir & "\KPMNetworkSelector\logo.jpg", 10, 10, 50, 50)
	Local $idTITEL = GUICtrlCreateLabel($swtitel, 70, 10, 300)
	GUICtrlSetFont($idTITEL, 12, $FW_BOLD)
	Local $idINFO = GUICtrlCreateLabel("If no checkbox is used, nothing will change.", 72, 30, 280)
	GUICtrlSetFont($idINFO, 10)
	Local $idVERSION = GUICtrlCreateLabel("v " & $swversion , 10, $intGUIlgth - 25)
	Local $idEXIT = GUICtrlCreateButton("save and exit", 250, $intGUIlgth - 35, 100, 25)
	GUICtrlSetFont($idEXIT, 10)

	; Display the GUI.
	GUISetState(@SW_SHOW, $hGUI)


	;Formatting and building dynamic select list
	_ArrayColInsert($aNetworkAdapters, 0)
	;_ArrayDisplay($aNetworkAdapters)

	Local $idENABLED = GUICtrlCreateLabel("Enabled adapters - disable?", 10, 75, 350, 20)
	GUICtrlSetFont($idENABLED, 10, $FW_BOLD)
	Local $height=95
	Local $i

	For $i = 0 To UBound($aNetworkAdapters) -1
		If $aNetworkAdapters[$i][4] = 1 Then
			$aNetworkAdapters[$i][0] = GUICtrlCreateCheckbox($aNetworkAdapters[$i][1], 10, $height, 350, 20)
			GUICtrlSetFont($aNetworkAdapters[$i][0], 10)
			$height=$height+20
		EndIf
	Next

	$height=$height+20
	Local $iddisabled = GUICtrlCreateLabel("Disabled adapters - enable?", 10, $height, 350, 20)
	GUICtrlSetFont($iddisabled, 10, $FW_BOLD)

	$height=$height+20
	For $i = 0 To UBound($aNetworkAdapters) -1
		If $aNetworkAdapters[$i][4] = 0 Then
			$aNetworkAdapters[$i][0] = GUICtrlCreateCheckbox($aNetworkAdapters[$i][1], 10, $height, 350, 20)
			GUICtrlSetFont($aNetworkAdapters[$i][0], 10)
			$height=$height+20
		EndIf
	Next

	; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $idEXIT

				For $i = 0 To UBound($aNetworkAdapters) -1
					If GuiCtrlRead($aNetworkAdapters[$i][0]) = $GUI_CHECKED Then
						If $aNetworkAdapters[$i][4] = 1 Then
							ShellExecuteWait("netsh.exe", 'int set interface "'& $aNetworkAdapters[$i][1] &'" disable', "", "", @SW_HIDE)
							;MsgBox($MB_SYSTEMMODAL, "Befehl", 'int set interface "'& $aNetworkAdapters[$i][1] &'" disable', 10)
						ElseIf $aNetworkAdapters[$i][4] = 0 Then
							ShellExecuteWait("netsh.exe", 'int set interface "'& $aNetworkAdapters[$i][1] &'" enable', "", "", @SW_HIDE)
						EndIf
					EndIf
				Next

			ExitLoop
        EndSwitch
    WEnd

	; Delete the previous GUI and all controls.
	GUIDelete($hGUI)

EndFunc   ;==>Gui

Func WMI_ListAllNetworkAdapters($sHost)
    Local $objWMIService = ObjGet("winmgmts:\\" & $sHost & "\root\cimv2")
    If @error Then Return SetError(1, 0, 0)
    Local $aStatus[13] = ["Disconnected", "Connecting", "Connected", "Disconnecting", "Hardware not present", "Hardware disabled", "Hardware malfunction", _
                          "Media Disconnected", "Authenticating", "Authentication Succeeded", "Authentication Failed", "Invalid Address", "Credentials Required"]
    ;$colItems = $objWMIService.ExecQuery("SELECT Name, NetConnectionID, NetConnectionStatus FROM Win32_NetworkAdapter WHERE (NetConnectionID = 'Ethernet' OR NetConnectionID = 'Wi-Fi')", "WQL", 0x30)
	$colItems = $objWMIService.ExecQuery("SELECT Name, NetConnectionID, NetConnectionStatus, NetEnabled FROM Win32_NetworkAdapter WHERE PhysicalAdapter = 'True'", "WQL", 0x30)
    Local $aNetworkAdapters[1000][4], $i = 0, $iPointer
    If IsObj($colItems) Then
        For $objItem in $colItems
            With $objItem
                $aNetworkAdapters[$i][0] = .NetConnectionID
                $aNetworkAdapters[$i][1] = .Name
                $aNetworkAdapters[$i][2] = $aStatus[.NetConnectionStatus * 1]
				$aNetworkAdapters[$i][3] = .NetEnabled
            EndWith
            $i += 1
        Next
        ReDim $aNetworkAdapters[$i][4]
        Return $aNetworkAdapters
    Else
        Return SetError(2, 0, 0)
    EndIf
EndFunc ;==>WMI_ListAllNetworkAdapters