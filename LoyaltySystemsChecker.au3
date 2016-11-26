#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 2.0)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложение для создания отчетов по кнопкам лояльности)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - )
#pragma compile(ProductName, LoyaltySystemsChecker)


#include <Array.au3>
#include <FileConstants.au3>
#include "XML.au3"
#include <DateTimeConstants.au3>
#include <Date.au3>
#include <File.au3>

#Region ====== variables ======
Local $sMailServerBackup = 
Local $sMailLoginBackup = 
Local $sMailPasswordBackup = 
Local $sDeveloperEmail = 

Local $sIniFileName = "settings.ini"
If Not FileExists($sIniFileName) Then _
		SendEmail("Не удается найти файл настроек: " & $sIniFileName, True)

Local $sSectionMail = "mail"
Local $sMailServer = IniRead($sIniFileName, $sSectionMail, "server", $sMailLoginBackup)
Local $sMailLogin = IniRead($sIniFileName, $sSectionMail, "login", $sMailLoginBackup)
Local $sMailPassword = IniRead($sIniFileName, $sSectionMail, "password", $sMailPasswordBackup)

Local $sSectionProlan = "prolan"
Local $sProlanUrl = IniRead($sIniFileName, $sSectionProlan, "url", "")
Local $sProlanReportId = IniRead($sIniFileName, $sSectionProlan, "reportid", "")
Local $sProlanQuestionId = IniRead($sIniFileName, $sSectionProlan, "questionid", "")
Local $sProlanLogin = IniRead($sIniFileName, $sSectionProlan, "login", "")
Local $sProlanPassword = IniRead($sIniFileName, $sSectionProlan, "password", "")
If Not $sProlanUrl Or _
		Not $sProlanReportId Or _
		Not $sProlanQuestionId Or _
		Not $sProlanLogin Or _
		Not $sProlanPassword Then _
		SendEmail("Некорректные значения в секции prolan указаны в файле настроек: " & $sIniFileName, True)

Local $sSectionNotifyReports = "notify_reports"
Local $aNotifyReports = IniReadSection($sIniFileName, $sSectionNotifyReports)

Local $sSectionNotifyReportsAlways = "notify_reports_always"
Local $aNotifyReportsAlways = IniReadSection($sIniFileName, $sSectionNotifyReportsAlways)

Local $sSectionNotifyDailyErrors = "notify_daily_errors"
Local $aNotifyDailyErrors = IniReadSection($sIniFileName, $sSectionNotifyDailyErrors)

Local $sSectionDailyCheck = "daily_check"
Local $aDailyCheck = IniReadSection($sIniFileName, $sSectionDailyCheck)
#EndRegion ====== variables ======

;~ _ArrayDisplay($aNotifyReports)
;~ _ArrayDisplay($aNotifyReportsAlways)
;~ _ArrayDisplay($aNotifyDailyErrors)
;~ _ArrayDisplay($aDailyCheck)

$oMyError = ObjEvent("AutoIt.Error", "ComErrorHandle")

;~ _ArrayDisplay($CmdLine)

If $CmdLine[0] Then
	If $CmdLine[1] = "-week" Then SendReports("week")
	If $CmdLine[1] = "-month" Then SendReports("month")
Else
	CheckLoyaltyReports()
EndIf








Func GetDateFromNow($nValue)
	Local $aTmpDate, $aTmpTime
	_DateTimeSplit(_DateAdd('D', $nValue, _NowCalc()), $aTmpDate, $aTmpTime)
	Return $aTmpDate[3] & "." & $aTmpDate[2] & "." & $aTmpDate[1]
EndFunc


Func SendReports($sPeriod)
	Local $sDate1
	Local $sDate2

	If $sPeriod = "week" Then
		$sDate1 = GetDateFromNow(-7)
		$sDate2 = GetDateFromNow(-1)
	ElseIf $sPeriod = "month" Then
		$sDate1 = "01" & StringRight(GetDateFromNow(-1), 8)
		$sDate2 = GetDateFromNow(-1)
	Else
		SendEmail("Выбран неправильный период формирования отчета (" & $sPeriod & ")", True)
	EndIf

	If Not IsArray($aNotifyReports) Then _
		SendEmail("Массив для отчетов не содержит данных", True)

	Local $aArray = GetDataFromProLan($sDate1, $sDate2)

	Local $sMessageHeader = '<b>Опрос: Порекомендуете ли Вы нашу клинику вашим друзьям и знакомым?</b><br><br>' & _
						    'Отчетный период: начало ' & $sDate1 & ' конец ' & $sDate2 & '<br><br>' & _
						    '<table border="1" cellpadding="5" cellspacing="5">' & _
						    '<tr>' & _
						    '<th>POS</th>' & _
						    '<th>Да</th>' & _
						    '<th>Нет</th>' & _
						    '<th>Затрудняюсь ответить</th>' & _
						    '<th>% Да</th>' & _
						    '<th>% Нет</th>' & _
						    '<th>% Затрудняюсь ответить</th>' & _
							'</tr>'
	Local $sTemplate = '<tr>' & _
						   '<td>@0</td>' & _
						   '<td>@1</td>' & _
						   '<td>@2</td>' & _
						   '<td>@3</td>' & _
						   '<td>@4</td>' & _
						   '<td>@5</td>' & _
						   '<td>@6</td>' & _
						   '</tr>'
	Local $sMessageBody = ""
	Local $sMessageEnding = '</table>'
	Local $sTitle = "Отчет по нажатиям кнопок лояльности"

;~ 	_ArrayDisplay($aArray)

	For $i = 1 To UBound($aNotifyReports, $UBOUND_ROWS) - 1
		Local $sCurrentPos = $aNotifyReports[$i][0]
		Local $sMessage = $sTemplate

		Local $nIndex = _ArraySearch($aArray, $sCurrentPos)
		Local $aArraySlice
		If $nIndex = -1 Then
			Local $aTmp[] = [$sCurrentPos, 0, 0, 0, 0, 0, 0]
			_ArrayTranspose($aTmp)
			$aArraySlice = $aTmp
		Else
			$aArraySlice = _ArrayExtract($aArray, $nIndex, $nIndex)
		EndIf

;~ 		_ArrayDisplay($aArraySlice)

		For $x = 0 To UBound($aArraySlice, $UBOUND_COLUMNS) - 1
			$sMessage = StringReplace($sMessage, "@" & $x, $aArraySlice[0][$x])
		Next

		$sMessageBody &= $sMessage

		SendEmail($sMessageHeader & $sMessage & $sMessageEnding, False, _
			$aNotifyReports[$i][1], "", $sTitle)
	Next

	Local $sTo = ""
	If IsArray($aNotifyReportsAlways) Then
		For $i = 1 To UBound($aNotifyReportsAlways, $UBOUND_ROWS) - 1
			$sTo &= $aNotifyReportsAlways[$i][1] & ";"
		Next
	EndIf

	SendEmail($sMessageHeader & $sMessageBody & $sMessageEnding, False, _
		$sTo, "", $sTitle)

EndFunc


Func CheckLoyaltyReports()
	If Not IsArray($aDailyCheck) Then _
		SendEmail("Не указаны регионы для ежедневной проверки в секции daily_check", True)

	Local $sDate1 = GetDateFromNow(-2)
	Local $sDate2 = GetDateFromNow(-1)
	Local $aArray = GetDataFromProLan($sDate1, $sDate2)

	For $i = 1 To UBound($aDailyCheck, $UBOUND_ROWS) - 1
		Local $sCurrentPos = $aDailyCheck[$i][1]
		If Not $sCurrentPos Then ContinueLoop

		If _ArraySearch($aArray, $sCurrentPos) = -1 Then _
				SendEmail("На группу поддержки '" & $sCurrentPos & "':" & @CRLF & _
				"За предыдущие два дня нет информации о нажатиях кнопок лояльности с опросом: " & _
				"'Порекомендуете ли Вы нашу клинику вашим друзьям и знакомым?'" & @CRLF & @CRLF & _
				"Необходимо проверить рабоспособность указанного сервиса")
	Next

	SendEmail("Данные по нажатиям на кнопки лояльности за период с " & $sDate1 & " по " & $sDate2 & ":" & @CRLF & _
			_ArrayToString($aArray), True)
EndFunc   ;==>CheckLoyaltyReports


Func GetDataFromProLan($sDate1, $sDate2)
	Local $strUrl = "http://911.prolan.ru/saas/reports/report.php"
	Local $strReportID = "POS"
	Local $strDtBegin = $sDate1 & " 00:00:00"
	Local $strDtEnd = $sDate2 & " 23:59:59"
	Local $strUseUTC = "0"
	Local $strLoginName = "ClinicLMC_R"
	Local $strPassword = "XkVd5x54"
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
			SendEmail("Не удалось получить данные с сайта" & @CRLF & @CRLF & _
			"Параметры запроса:" & @CRLF & $strUrl & @CRLF & _
			$strDtBegin & @CRLF & $strDtEnd & @CRLF & $strReportID & " " & $nQuestionID & @CRLF & _
			"Прокси: " & $strProxy, True)

	Local $strFileName = @ScriptDir & "\response.xml"
	Local $hFile = FileOpen($strFileName, BitOR($FO_OVERWRITE, $FO_ANSI))
	FileWrite($hFile, $strXmlResponse)
	FileClose($hFile)

	Local $resultArray = ParseXmlFileToArray($strFileName)
	If Not IsArray($resultArray) Then _
			SendEmail("Не удалось выполнить разбор XML ответа:" & @CRLF & $strXmlResponse, True)

	If UBound($resultArray, $UBOUND_COLUMNS) <> 7 Then _
			SendEmail("Полученный массив не соответствует заданной ширине: " & @CRLF & _
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

	SendEmail("В полученном массиве нет данных: " & @CRLF & _ArrayToString($resultArray), True)
EndFunc   ;==>GetDataFromProLan


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

	Return ($aNodesColl)
EndFunc   ;==>ParseXmlFileToArray


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
EndFunc   ;==>HttpPost


Func ComErrorHandle()
	SendEmail("err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
			"err.number is: " & @TAB & Hex($oMyError.number, 8) & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext, True)
EndFunc   ;==>ComErrorHandle


Func SendEmail($sMessage, $bError = False, $sTo = "", $sCopy = "", $sTitle = "")
	Local $sCurrentPcName = @ComputerName
	Local $sFrom = "Система отчетов по кнопкам лояльности"
	If Not $sTitle Then $sTitle = "Внимание! Имеются ошибки!"

	Local $sEnding = @CRLF & @CRLF & _
			"---------------------------------------" & @CRLF & _
			"Это автоматическое сообщение." & @CRLF & _
			"Пожалуйста, не отвечайте на него." & @CRLF & _
			"Имя системы: " & $sCurrentPcName
	If StringInStr($sMessage, "<") And StringInStr($sMessage, ">") Then _
		$sEnding = StringReplace($sEnding, @CRLF, "<br>")

	$sMessage &= $sEnding

	If Not $sTo Then
		If IsArray($aNotifyDailyErrors) Then
			$sTo = ""
			For $i = 1 To UBound($aNotifyDailyErrors, $UBOUND_ROWS) - 1
				$sTo &= $aNotifyDailyErrors[$i][1] & ";"
			Next
		Else
			$sTo = "stp@7828882.ru"
		EndIf
	EndIf

	If Not $sCopy Then $sCopy = "nn-admin@bzklinika.ru"
	If $bError Then $sTo = $sCopy

	ConsoleWrite($sMessage & @CRLF & $sTo & @CRLF & $sCopy & @CRLF & $sTitle & @CRLF & "------" & @CRLF)

	Return

	If Not _INetSmtpMailCom($sMailServer, $sFrom, $sMailLogin, $sTo, _
			$sTitle, $sMessage, "", $sCopy, "", $sMailLogin, $sMailPassword) Then _
			_INetSmtpMailCom($sMailServerBackup, $sFrom, $sMailLoginBackup, $sTo, _
			$sTitle, $sMessage, "", $sCopy, "", $sMailLoginBackup, $sMailPasswordBackup)

	If $bError Then Exit
EndFunc   ;==>SendEmail


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
