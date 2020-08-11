#include-once
#include <_S7_COM.au3>

#cs
Funktionen zum einfachen Erstellen von S7-Hardware-Konfigurationen.
Erweiterung für die UDF: _S7_COM.au3

Apache License Version 2.0

2012 - 2020 by Thorsten Willert

Grundfunktionen:
_S7_HWConfig_Add_Rack: Rack
_S7_HWConfig_Add_CPU: CPU
_S7_HWConfig_Add_CPU_Moduls: CPU-Module
_S7_HWConfig_Add_SubSystem: Sub-System
_S7_HWConfig_Add_SlaveModuls: Slave-Module

Hilfsfunktionen:
_S7_HWConfig_TypeSelect: Auswahl des Bauteil-Typs anhand der Bestellnummer

Spezielle Baugruppen (als Beipiele):
_S7_HWConfig_Add_ET200S: ET200S
_S7_HWConfig_Add_IM153: IM153
_S7_HWConfig_Add_DP_Koppler: DP-Koppler
_S7_HWConfig_AddFestoPP: Festo-Ventilinsel
#ce

Global Enum $_S7_HWConfig_Misc, _
		$_S7_HWConfig_CPU300, _
		$_S7_HWConfig_CPU400, _
		$_S7_HWConfig_DI, _
		$_S7_HWConfig_DO, _
		$_S7_HWConfig_AI, _
		$_S7_HWConfig_AO

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_Rack
; Description ...: Rack
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_Rack($oStation, $sName[, $iNr = 0])
; Parameter(s): .: $oStation    -
;                  $sName       -
;                  $iNr         - Optional: (Default = 0) :
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:05:07 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_Rack($oStation, $sName, $iNr = 0)
	$oStation.Racks.add($sName, "6ES7 390-1???0-0AA0", "", $iNr)
	Return $oStation.Racks($sName)
EndFunc   ;==>_S7_HWConfig_Add_Rack

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_CPU
; Description ...: CPU
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_CPU($oRack, $sName, $sBestellnummer)
; Parameter(s): .: $oRack       -
;                  $sName       -
;                  $sBestellnummer -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:08:06 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_CPU($oRack, $sName, $sBestellnummer)
	Return $oRack.Modules.add($sName, $sBestellnummer, "LATEST", 2)
EndFunc   ;==>_S7_HWConfig_Add_CPU

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_CPU_Moduls
; Description ...: COU-Module
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_CPU_Moduls(ByRef $oRack, ByRef $aData)
; Parameter(s): .: $oRack       -
;                  $aData       -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:09:22 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_CPU_Moduls(ByRef $oRack, ByRef $aData)
	_S7_HWConfig_Add_SlaveModuls($oRack, $aData, 3)
EndFunc   ;==>_S7_HWConfig_Add_CPU_Moduls

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_SubSystem
; Description ...: SubSystem
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_SubSystem($oCPU, $sName[, $iNr = 1[, $iModul = 1]])
; Parameter(s): .: $oCPU        -
;                  $sName       -
;                  $iNr         - Optional: (Default = 1) :
;                  $iModul      - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:10:13 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_SubSystem($oCPU, $sName, $iNr = 1, $iModul = 1)
	Return $oCPU.Modules($iModul).AddSubSystem($sName, 1)
EndFunc   ;==>_S7_HWConfig_Add_SubSystem

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_SlaveModuls
; Description ...: SlaveModul
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_SlaveModuls($oSlave, $aData[, $iOffset = 0])
; Parameter(s): .: $oSlave      -
;                  $aData       - [Name, Bestellnummer, Byte (-1 = Auto), Bit, Kommentar]
;                  $iOffset     - Optional: (Default = 0) :
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:17:18 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_SlaveModuls($oSlave, $aData, $iOffset = 0)
	; $aData  = Name / Bestellnummer / Byte (-1 = Auto) / Bit / Kommentar
	Local $oSlaveModu
	Local $iSt = -1

	For $i = 0 To UBound($aData) - 1
		$iSt = $iOffset + $i
		$oSlaveModu = $oSlave.Modules.add($aData[$i][0], $aData[$i][1], "LATEST", $iSt + 1)

		If $aData[$i][2] > -1 Then
			Switch _S7_HWConfig_TypeSelect($aData[$i][1])
				Case $_S7_HWConfig_DI, $_S7_HWConfig_DO
					If $aData[$i][2] > -1 Then
						$oSlaveModu.LocalInAddresses(0).LogicalAddress = $aData[$i][2]
						$oSlaveModu.LocalInAddresses(0).BitAddress = $aData[$i][3]
					EndIf

				Case $_S7_HWConfig_AI, $_S7_HWConfig_AO
					If $aData[$i][2] > -1 Then $oSlaveModu.LocalInAddresses(0).LogicalAddress = $aData[$i][2]

				Case $_S7_HWConfig_Misc
					If $aData[$i][2] > -1 Then $oSlaveModu.LocalInAddresses(0).LogicalAddress = $aData[$i][2]
			EndSwitch
		EndIf
		$oSlaveModu.RegisterAddresses()
		$oSlaveModu.Comment = $aData[$i][4]
	Next
EndFunc   ;==>_S7_HWConfig_Add_SlaveModuls

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_TypeSelect
; Description ...: Auswahl CPU, DI, DO usw. anhand der Bestellnummer
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_TypeSelect($sType)
; Parameter(s): .: $sType       -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:20:32 CEST 2020
; Related .......: _S7_CONM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_TypeSelect($sType)
	Local $aType[8][2]
	;ET200S
	$aType[0][0] = "6ES7 131"
	$aType[0][1] = $_S7_HWConfig_DI

	$aType[1][0] = "6ES7 132"
	$aType[1][1] = $_S7_HWConfig_DO

	$aType[2][0] = "6ES7 134"
	$aType[2][1] = $_S7_HWConfig_AI

	$aType[3][0] = "6ES7 135"
	$aType[3][1] = $_S7_HWConfig_AO
	; CPU 300
	$aType[4][0] = "6ES7 331"
	$aType[4][1] = $_S7_HWConfig_DI

	$aType[5][0] = "6ES7 332"
	$aType[5][1] = $_S7_HWConfig_DO

	$aType[6][0] = "6ES7 334"
	$aType[6][1] = $_S7_HWConfig_AI

	$aType[7][0] = "6ES7 335"
	$aType[7][1] = $_S7_HWConfig_AO

	For $i = 0 To UBound($aType) - 1
		If StringInStr($sType, $aType[$i][0], 0, 1) Then Return $aType[$i][1]
	Next

	Return $_S7_HWConfig_Misc
EndFunc   ;==>_S7_HWConfig_TypeSelect

; ##############################################################################
; ### Spezielle Bauteile ###

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_ET200S
; Description ...: ET200S
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_ET200S($oSub, $sName, $iPPAdress, $aData)
; Parameter(s): .: $oSub        -
;                  $sName       -
;                  $iPPAdress   -
;                  $aData       -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:11:16 CEST 2020
; Link ..........:
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_ET200S($oSub, $sName, $iPPAdress, $aData)
	Local $oSlave = $oSub.Slaves.add($sName, "6ES7 151-1AA05-0AB0", "LATEST", $iPPAdress)

	_S7_HWConfig_Add_SlaveModuls($oSlave, $aData)
EndFunc   ;==>_S7_HWConfig_Add_ET200S

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_IM153
; Description ...: IM153
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_IM153($oSub, $sName, $iPPAddress, $aData)
; Parameter(s): .: $oSub        -
;                  $sName       -
;                  $iPPAddress  -
;                  $aData       -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:12:51 CEST 2020
; Related .......: _S7_COM
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_IM153($oSub, $sName, $iPPAddress, $aData)
	Local $oSlave = $oSub.Slaves.add($sName, "6ES7 153-1AA03-0XB0", "LATEST", $iPPAddress)

	_S7_HWConfig_Add_SlaveModuls($oSlave, $aData, 3)
EndFunc   ;==>_S7_HWConfig_Add_IM153

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_Add_DP_Koppler
; Description ...: DP Koppler
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_Add_DP_Koppler($oSub, $sName, $iPPAdress)
; Parameter(s): .: $oSub        -
;                  $sName       -
;                  $iPPAdress   -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:02:58 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_Add_DP_Koppler($oSub, $sName, $iPPAdress)
	$oSub.Slaves.add($sName, "SI018070.GSG", "LATEST", $iPPAdress) ; !!! GSD-Datei nicht Bestellnummer
EndFunc   ;==>_S7_HWConfig_Add_DP_Koppler

; #FUNCTION# ===================================================================
; Name ..........: _S7_HWConfig_AddFestoPP
; Description ...: Einbinden einer Festo-Ventilinsel anhand der GSD-Datei
; AutoIt Version : V3.3.14.2
; Syntax ........: _S7_HWConfig_AddFestoPP($oSub, $sName, $iPPAdress, $iOutAddress)
; Parameter(s): .: $oSub        -
;                  $sName       -
;                  $iPPAdress   -
;                  $iOutAddress -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 11 12:03:48 CEST 2020
; Related .......: _S7_COM.au3
; Example .......: Yes
; ==============================================================================
Func _S7_HWConfig_AddFestoPP($oSub, $sName, $iPPAdress, $iOutAddress)
	Local $oSlave = $oSub.Slaves.add($sName, "CPV_0A35.GSD", "", $iPPAdress)
	Local $oSlaveModu = $oSlave.Modules.add("16DA", "Basis:16DA", "LATEST", 1)

	$oSlaveModu.LocalOutAddresses(0).LogicalAddress = Round($iOutAddress, 0)
	$oSlaveModu.RegisterAddresses()
EndFunc   ;==>_S7_HWConfig_AddFestoPP
