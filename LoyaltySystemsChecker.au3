#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon.ico
#pragma compile(ProductVersion, 1.0)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ŒŒŒ  ÎËÌËÍ‡ ÀÃ—')
#pragma compile(FileDescription, œËÎÓÊÂÌËˇ ÔÓ‚ÂÍË ‡·ÓÚÓÒÔÓÒÓ·ÌÓÒÚË ÍÌÓÔÓÍ ÎÓˇÎ¸ÌÓÒÚË)
#pragma compile(LegalCopyright, √‡¯ÍËÌ œ‡‚ÂÎ œ‡‚ÎÓ‚Ë˜ - ÕËÊÌËÈ ÕÓ‚„ÓÓ‰)
#pragma compile(ProductName, LoyaltySystemsChecker)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#include <Array.au3>
#include <FileConstants.au3>
#include "XML.au3"
;~ #include <GuiListView.au3>
;~ #include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
;~ #include <GUIConstantsEx.au3>
;~ #include <ListViewConstants.au3>
;~ #include <StaticConstants.au3>
;~ #include <WindowsConstants.au3>
#include <Date.au3>
#include <File.au3>


$oMyError = ObjEvent("AutoIt.Error","ComErrorHandle")


Local $aTmpDate, $aTmpTime
_DateTimeSplit(_DateAdd('D', -2, _NowCalc()), $aTmpDate, $aTmpTime)
Local $strDate1 = $aTmpDate[3] & "." & $aTmpDate[2] & "." & $aTmpDate[1]

_DateTimeSplit(_DateAdd('D', -1, _NowCalc()), $aTmpDate, $aTmpTime)
Local $strDate2 = $aTmpDate[3] & "." & $aTmpDate[2] & "." & $aTmpDate[1]

CheckLoyaltyReports(GetDataFromProLan())

Func CheckLoyaltyReports($aArray)
	Local $aPosNames[] = [" ‡Á‡Ì¸", _
						" ‡ÒÌÓ‰‡", _
						" ‡ÒÌÓˇÒÍ", _
						"ÕËÊÌËÈ ÕÓ‚„ÓÓ‰", _
						"ÕÓ‚ÓÒË·ËÒÍ", _
						"—.œÂÚÂ·Û„", _
						"—Ó˜Ë", _
						"—ÂÚÂÌÍ‡", _
						"—ÚÛÔËÌÓ", _
						"—Û˘Â‚ÒÍ‡ˇ¬ÁÓÒÎÓÂ", _
						"—Û˘Â‚ÒÍ‡ˇƒÂÚÒÚ‚Ó", _
						"”Ù‡", _
						"‘ÛÌÁÂÌÒÍ‡ˇ"]

	For $sPosName In $aPosNames
		If _ArraySearch($aArray, $sPosName) = -1 Then _
			SendEmail("Õ‡ „ÛÔÔÛ ÔÓ‰‰ÂÊÍË '" & $sPosName & "':" & @CRLF & _
				"«‡ ÔÂ‰˚‰Û˘ËÂ ‰‚‡ ‰Ìˇ ÌÂÚ ËÌÙÓÏ‡ˆËË Ó Ì‡Ê‡ÚËˇı ÍÌÓÔÓÍ ÎÓˇÎ¸ÌÓÒÚË Ò ÓÔÓÒÓÏ:" & _
				"'œÓÂÍÓÏÂÌ‰ÛÂÚÂ ÎË ¬˚ Ì‡¯Û ÍÎËÌËÍÛ ‚‡¯ËÏ ‰ÛÁ¸ˇÏ Ë ÁÌ‡ÍÓÏ˚Ï?'" & @CRLF & @CRLF & _
				"ÕÂÓ·ıÓ‰ËÏÓ ÔÓ‚ÂËÚ¸ ‡·ÓÒÔÓÒÓ·ÌÓÒÚ¸ ÛÍ‡Á‡ÌÌÓ„Ó ÒÂ‚ËÒ‡")
	Next

	SendEmail("ƒ‡ÌÌ˚Â ÔÓ Ì‡Ê‡ÚËˇÏ Ì‡ ÍÌÓÔÍË ÎÓˇÎ¸ÌÓÒÚË Á‡ ÔÂËÓ‰ Ò " & $strDate1 & " ÔÓ " & $strDate2 & ":" & @CRLF & _
		_ArrayToString($aArray), True)
EndFunc


Func GetDataFromProLan()
	Local $strUrl = "http://911.prolan.ru/saas/reports/report.php"
	Local $strReportID = "POS"
	Local $strDtBegin = $strDate1 & " 00:00:00"
	Local $strDtEnd = $strDate2 & " 23:59:59"
	Local $strUseUTC = "0"
	Local $strLoginName = ""
	Local $strPassword = ""
	Local $nQuestionID = 37
	Local $strProxy = "172.16.6.1:8080"

	Local $strBody = "LoginName=" & $strLoginName & _
					 "&Password=" & $strPassword & _
					 "&QuestionID=" & $nQuestionID & _
					 "&ReportID=" & $strReportID & _
					 "&Begin=" & $strDtBegin & _
					 "&End=" & $strDtEnd & _
					 "&UseUTC=" & $strUseUTC

	Local $strXmlResponse = HttpPost($strUrl, $strBody, $strProxy)
	If Not $strXmlResponse Then _
		SendEmail("ÕÂ Û‰‡ÎÓÒ¸ ÔÓÎÛ˜ËÚ¸ ‰‡ÌÌ˚Â Ò Ò‡ÈÚ‡" & @CRLF & @CRLF & _
		"œ‡‡ÏÂÚ˚ Á‡ÔÓÒ‡:" & @CRLF & $strUrl & @CRLF & _
		$strDtBegin & @CRLF & $strDtEnd & @CRLF & $strReportID & $nQuestionID & @CRLF & _
		"œÓÍÒË: " & $strProxy, True)


	Local $strFileName = @ScriptDir & "\response.xml"
	Local $hFile = FileOpen($strFileName, BitOR($FO_OVERWRITE, $FO_ANSI))
	FileWrite($hFile, $strXmlResponse)
	FileClose($hFile)

	Local $resultArray = ParseXmlFileToArray($strFileName)
	If Not IsArray($resultArray) Then _
		SendEmail("ÕÂ Û‰‡ÎÓÒ¸ ‚˚ÔÓÎÌËÚ¸ ‡Á·Ó XML ÓÚ‚ÂÚ‡:" & @CRLF & $strXmlResponse, True)


	If UBound($resultArray, $UBOUND_COLUMNS) <> 7 Then _
		SendEmail("œÓÎÛ˜ÂÌÌ˚È Ï‡ÒÒË‚ ÌÂ ÒÓÓÚ‚ÂÚÒÚ‚ÛÂÚ Á‡‰‡ÌÌÓÈ ¯ËËÌÂ: " & @CRLF & _
			_ArrayToString($resultArray), True)


	Local $result[0][7]

	For $i = 1 To UBound($resultArray, $UBOUND_ROWS) - 1
		Local $tmpArray[1][7]

		For $x = 0 To 6
			$tmpArray[0][$x] = $resultArray[$i + $x][3]
		Next

		_ArrayAdd($result, $tmpArray)
		$i += 6
	Next

	If IsArray($result) Then Return $result

	SendEmail("¬ ÔÓÎÛ˜ÂÌÌÓÏ Ï‡ÒÒË‚Â ÌÂÚ ‰‡ÌÌ˚ı: " & @CRLF & _ArrayToString($resultArray), True)
EndFunc


Func ParseXmlFileToArray($strFileName)
	Local $oXMLDoc = _XML_CreateDOMDocument(Default)
	If @error Then Return

	Local $oXMLDOM_EventsHandler = ObjEvent($oXMLDoc, "XML_DOM_EVENT_")

	_XML_Load($oXMLDoc, $strFileName)
	If @error Then Return

	Local $sXmlAfterTidy = _XML_TIDY($oXMLDoc)
	If @error Then Return

	Local $oNodesColl = _XML_SelectNodes($oXMLDoc, "//Rows/Row/Cell")
	If @error Then Return

	Local $aNodesColl = _XML_Array_GetNodesProperties($oNodesColl)
	If @error Then Return

	Return($aNodesColl)
EndFunc    ;==>Example_1__XML_SelectNodes


Func HttpPost($sURL, $sData = "", $strProxy = "")
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	If $strProxy Then $oHTTP.SetProxy(2, $strProxy)

	$oHTTP.Open("POST", $sURL, False)
	If (@error) Then Return SetError(1, 0, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	$oHTTP.SetRequestHeader("RequestType", "GetXmlLoyaltyReport")
	$oHTTP.SetRequestHeader("Content-Length", StringLen($sData))

	$oHTTP.Send($sData)
	If (@error) Then Return SetError(2, 0, 0)

	If ($oHTTP.Status <> 200) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc


Func ComErrorHandle()
	SendEmail("err.description is: " & @TAB & $oMyError.description  & @CRLF & _
		"err.windescription:"    & @TAB & $oMyError.windescription & @CRLF & _
		"err.number is: "        & @TAB & hex($oMyError.number,8)  & @CRLF & _
		"err.lastdllerror is: "  & @TAB & $oMyError.lastdllerror   & @CRLF & _
		"err.scriptline is: "    & @TAB & $oMyError.scriptline   & @CRLF & _
		"err.source is: "        & @TAB & $oMyError.source       & @CRLF & _
		"err.helpfile is: "      & @TAB & $oMyError.helpfile     & @CRLF & _
		"err.helpcontext is: "   & @TAB & $oMyError.helpcontext, True)
Endfunc


Func SendEmail($sMessage, $bError = False)
	ConsoleWrite("--- Sending email ---" & @CRLF & $sMessage & @CRLF)

	Local $current_pc_name = @ComputerName
	Local $from = "Infoscreen screenshots viewer"
	Local $title = "Infosystems daily report"
	$sMessage &= @CRLF & @CRLF & _
		"---------------------------------------" & @CRLF & _
		"This is automatically generated message" & @CRLF & _
		"Sended from: " & $current_pc_name & @CRLF & _
		"Please do not reply"

	Local $server = ""
	Local $login = ""
	Local $password = ""
	Local $to = ""
	Local $copy = ""
	If $bError Then $to = $copy
	Local $from = "Loyalty systems checker"
	Local $title = "Daily report"

	_INetSmtpMailCom($server, $from, $login, $to, _
		$title, $sMessage, "", $copy, "", $login, $password)

	If $bError Then Exit
EndFunc


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", _
	$as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", _
	$s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	Local $i_Error = 0
	Local $i_Error_desciption = ""

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				$i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
				$as_Body &= $i_Error_desciption & @CRLF
			EndIf
		Next
	EndIf

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
		$objEmail.TextBodyPart.Charset = "utf-8"
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	$objEmail.Configuration.Fields.Update
	$objEmail.Send

	If @error Then Return False
	Return True
EndFunc   ;==>_INetSmtpMailCom