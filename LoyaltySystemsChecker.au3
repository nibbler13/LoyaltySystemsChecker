#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 3.2)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложение для создания отчетов по кнопкам лояльности)
#pragma compile(LegalCopyright,)
#pragma compile(ProductName, LoyaltySystemsChecker)


#include <Array.au3>
#include <FileConstants.au3>
#include "XML.au3"
#include <DateTimeConstants.au3>
#include <Date.au3>
#include <File.au3>
#include <Excel.au3>



$oMyError = ObjEvent("AutoIt.Error", "ComErrorHandle")


#Region ====== variables ======
Local $aErrorNotify = 0

Local $sMailServerBackup = "smtp.budzdorov.ru"
Local $sMailLoginBackup = "infoscreen_screenshots_viewer@nnkk.budzdorov.su"
Local $sMailPasswordBackup = "paqafapy"
Local $sDeveloperEmail = "nn-admin@bzklinika.ru"

Local $sIniFileName = @ScriptDir & "\settings.ini"
If Not FileExists($sIniFileName) Then _
		SendEmail("Не удается найти файл настроек: " & $sIniFileName, True)

Local $sSectionDebug = "debug_mode"
Local $bDebug = (IniRead($sIniFileName, $sSectionDebug, "debug", False) = "1" ? True : False)

Local $sSectionErrorNotify = "error_notify"
$aErrorNotify = IniReadSection($sIniFileName, $sSectionErrorNotify)

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

Local $sSectionClinicsNameConformity = "clinics_name_conformity"
Local $aCilincsNameConformity = IniReadSection($sIniFileName, $sSectionClinicsNameConformity)

Local $aClinicsNamesAndAddressess[0][2]
For $i = 1 To UBound($aCilincsNameConformity, $UBOUND_ROWS) - 1
	Local $aCurrentClinic[1][2]
	$aCurrentClinic[0][0] = $aCilincsNameConformity[$i][0]
	$aCurrentClinic[0][1] = IniReadSection($sIniFileName, $aCilincsNameConformity[$i][1])
	_ArrayAdd($aClinicsNamesAndAddressess, $aCurrentClinic)
Next

Local $sQuestionIDRecommendButton = "37"
Local $sQuestionIDDoctorsQuality = "152"
Local $sReportIDTotal = "TOTAL"
Local $sReportIDPos = "POS"
Local $sReportIDEmployee = "EMPLOYEE"
#EndRegion ====== variables ======


If $CmdLine[0] Then
	If $CmdLine[1] = "-week" Then SendReports("week")
	If $CmdLine[1] = "-month" Then SendReports("month")
Else
	CheckLoyaltyReports()
EndIf



Func GetDateFromNow($nValue)
	Local $aTmpDate, $aTmpTime
	_DateTimeSplit(_DateAdd('D', $nValue, _NowCalc()), $aTmpDate, $aTmpTime)
	If StringLen($aTmpDate[3]) < 2 Then $aTmpDate[3] = "0" & $aTmpDate[3]
	If StringLen($aTmpDate[2]) < 2 Then $aTmpDate[2] = "0" & $aTmpDate[2]
	Return $aTmpDate[3] & "." & $aTmpDate[2] & "." & $aTmpDate[1]
EndFunc


Func SendReports($sPeriod)
	Local $sDateBegin
	Local $sDateEnd

	If $sPeriod = "week" Then
		$sDateBegin = GetDateFromNow(-7)
		$sDateEnd = GetDateFromNow(-1)
	ElseIf $sPeriod = "month" Then
		$sDateBegin = "01" & StringRight(GetDateFromNow(-1), 8)
		$sDateEnd = GetDateFromNow(-1)
	Else
		SendEmail("Выбран неправильный период формирования отчета (" & $sPeriod & ")", True)
	EndIf

	ParseRecommendData($sDateBegin, $sDateEnd)
	ParseDoctorsQualityData($sDateBegin, $sDateEnd)
EndFunc


Func ParseDoctorsQualityData($sDateBegin, $sDateEnd)
	If Not IsArray($aClinicsNamesAndAddressess) Then _
		SendEmail("Массив адресов для рассылки отчетов по мониторам лояльности не содержит данных", True)

	Local $sExcelTemplate = "MerchantReportTemplate.xlsx"
	Local $sExcelTemplateFullPath = @ScriptDir & "\" & $sExcelTemplate
	If Not FileExists($sExcelTemplateFullPath) Then _
		SendEmail("Не удается найти шаблон: " & $sExcelTemplate, True)

	Local $oExcel = _Excel_Open(False, False, False, False, False)
	If Not IsObj($oExcel) Then _
		SendEmail("Не удается создать объект Excel", True)

	Local $aReportsID[] = [$sReportIDTotal, $sReportIDPos, $sReportIDEmployee]
	Local $aDataFromProlan[UBound($aReportsID)]
	For $i = 0 To UBound($aReportsID) - 1
		$aDataFromProlan[$i] = GetDataFromProLan($sDateBegin, $sDateEnd, $sQuestionIDDoctorsQuality, $aReportsID[$i])
	Next

	For $i = 0 To UBound($aClinicsNamesAndAddressess, $UBOUND_ROWS) - 1
		Local $oBook = _Excel_BookOpen($oExcel, $sExcelTemplateFullPath)
		If Not IsObj($oBook) Then _
			SendEmail("Не удается открыть книгу: " & $sExcelTemplateFullPath, True)

		Local $sCurrentName = $aClinicsNamesAndAddressess[$i][0]

		Local $aCurrentMarks = 0
		For $x = UBound($aDataFromProlan) - 1 To 0 Step -1
			Local $nExcelOffset = 0
			If $aReportsID[$x] = $sReportIDTotal Then $nExcelOffset = 1

			Local $nFirstMarkColumn = 1
			If $aReportsID[$x] = $sReportIDEmployee Then
				$nFirstMarkColumn = 2
			EndIf


			Local $aCurrentArray = $aDataFromProlan[$x]
			If $sCurrentName <> "*" Then
				If $aReportsID[$x] = $sReportIDTotal Then
					If IsArray($aCurrentMarks) Then _
						$aCurrentArray = _ArrayExtract($aCurrentMarks, 0, 0, UBound($aCurrentMarks, $UBOUND_COLUMNS) - 10, _
							UBound($aCurrentMarks, $UBOUND_COLUMNS) - 1)
				Else
					$aCurrentArray = GetDataFromArrayByClinicName($aCurrentArray, $sCurrentName, $nFirstMarkColumn, $aCurrentMarks)
				EndIf

			EndIf

			_Excel_RangeWrite($oBook, $aReportsID[$x], $aCurrentArray, _Excel_ColumnToLetter(1 + $nExcelOffset) & 2)
			$oBook.Sheets($aReportsID[$x]).UsedRange.Borders.LineStyle = 1
			$oBook.Sheets($aReportsID[$x]).UsedRange.Borders.Color = 8
			$oBook.Sheets($aReportsID[$x]).UsedRange.Borders.Weight = 2
			$oBook.Sheets($aReportsID[$x]).UsedRange.AutoFilter
		Next


		Local $sResultFileName =  "Отчет по монитору лояльности " & (($sCurrentName == "*") ? "всех клиник" : $sCurrentName) & _
			" за период с " & $sDateBegin & " по " & $sDateEnd & ".xlsx"
		Local $sResultFileFullPath = @ScriptDir & "\" & $sResultFileName
		_Excel_BookSaveAs($oBook, $sResultFileFullPath, 51, True)
		_Excel_BookClose($oBook, False)

		If Not FileExists($sResultFileFullPath) Then _
			SendEmail("Не удалось создать файл отчета: " & _
				$sResultFileFullPath, True)

		Local $sTitle = "Отчет по монитору лояльности"
		Local $sMessage = "Отчет за период с " & $sDateBegin & " по " & $sDateEnd & " во вложении"

		Local $to = GetEmailAddresses($aClinicsNamesAndAddressess[$i][1])
		Local $copy = GetEmailAddresses($aNotifyReportsAlways)
		SendEmail($sMessage, False, $to, $copy, $sTitle, $sResultFileFullPath)
	Next

	_Excel_Close($oExcel, False, True)
EndFunc

Func GetDataFromArrayByClinicName($aArray, $sName, $nFirstMarkColumn, ByRef $aTotalMarks)
	Local $aReturnArray[0][UBound($aArray, $UBOUND_COLUMNS)]
	Local $aMarks = [0, 0, 0, 0, 0]

	For $i = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		If StringInStr($aArray[$i][0], $sName) Then
			For $x = 0 To UBound($aMarks) - 1
				$aMarks[$x] += $aArray[$i][$x + $nFirstMarkColumn]
			Next
			Local $currentRow = _ArrayExtract($aArray, $i, $i)
			_ArrayAdd($aReturnArray, $currentRow)
		EndIf
	Next

	Local $nTotalMarks = 0
	For $mark In $aMarks
		$nTotalMarks += $mark
	Next

	Local $aLastRow[1][UBound($aArray, $UBOUND_COLUMNS)]
	$aLastRow[0][$nFirstMarkColumn - 1] = "Всего"
	If $nTotalMarks Then
		For $i = 0 To UBound($aMarks) - 1
			$aLastRow[0][$nFirstMarkColumn + $i] = $aMarks[$i]
			$aLastRow[0][$nFirstMarkColumn + $i + 5] = StringReplace(Round(($aMarks[$i] / $nTotalMarks) * 100, 2), ".", ",")
		Next
	EndIf

	_ArrayAdd($aReturnArray, $aLastRow)

	$aTotalMarks = $aLastRow
	Return $aReturnArray
EndFunc


Func ParseRecommendData($sDateBegin, $sDateEnd)
	Local $aArray = GetDataFromProLan($sDateBegin, $sDateEnd, $sQuestionIDRecommendButton, $sReportIDPos)

	If Not IsArray($aNotifyReports) Then _
		SendEmail("Массив адресов главных врачей не содержит данных", True)

	Local $sMessageHeader = '<b>Опрос: Порекомендуете ли Вы нашу клинику вашим друзьям и знакомым?</b><br><br>' & _
						    'Отчетный период: начало ' & $sDateBegin & ' конец ' & $sDateEnd & '<br><br>' & _
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

		For $x = 0 To UBound($aArraySlice, $UBOUND_COLUMNS) - 1
			$sMessage = StringReplace($sMessage, "@" & $x, $aArraySlice[0][$x])
		Next

		$sMessageBody &= $sMessage

		SendEmail($sMessageHeader & $sMessage & $sMessageEnding, False, _
			$aNotifyReports[$i][1], "", $sTitle)
	Next

	Local $sTo = GetEmailAddresses($aNotifyReportsAlways)
	SendEmail($sMessageHeader & $sMessageBody & $sMessageEnding, False, _
		$sTo, "", $sTitle)
EndFunc


Func GetEmailAddresses($aArray)
	Local $sAddresses = ""
	If Not IsArray($aArray) Or UBound($aArray, $UBOUND_ROWS) < 2 Then _
		SendEmail("Func GetEmailAddresses($aArray)" & @CRLF & _
			"Массив адресов указан неверно", True)

	For $i = 1 To UBound($aArray, $UBOUND_ROWS) - 1
		$sAddresses &= $aArray[$i][1] & ";"
	Next

	Return $sAddresses
EndFunc


Func CheckLoyaltyReports()
	If Not IsArray($aDailyCheck) Then _
		SendEmail("Не указаны регионы для ежедневной проверки в секции daily_check", True)

	Local $sDate1 = GetDateFromNow(-2)
	Local $sDate2 = GetDateFromNow(-1)
	Local $aArray = GetDataFromProLan($sDate1, $sDate2, $sQuestionIDRecommendButton, $sReportIDPos)

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


Func GetDataFromProLan($sDateBegin, $sDateEnd, $sQuestionID, $sReportID)
	Local $strUrl = "http://911.prolan.ru/saas/reports/report.php"
	Local $strDtBegin = $sDateBegin & " 00:00:00"
	Local $strDtEnd = $sDateEnd & " 23:59:59"

	Local $strUseUTC = "0"
	Local $strLoginName = ""
	Local $strPassword = ""
	Local $strProxy = ""

	Local $strBody = "LoginName=" & $strLoginName & _
			"&Password=" & $strPassword & _
			"&QuestionID=" & $sQuestionID & _
			"&ReportID=" & $sReportID & _
			"&Begin=" & $strDtBegin & _
			"&End=" & $strDtEnd & _
			"&UseUTC=" & $strUseUTC

	Local $sKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
	Local $sValue = RegRead($sKey, "ProxyServer")
	If $sValue <> "" Then $strProxy = $sValue

	Local $strXmlResponse = HttpPost($strUrl, $strBody, $strProxy)

	If Not $strXmlResponse Then _
			SendEmail("Не удалось получить данные с сайта" & @CRLF & @CRLF & _
				"Параметры запроса:" & @CRLF & $strUrl & @CRLF & _
				$strDtBegin & @CRLF & $strDtEnd & @CRLF & $sReportID & " " & $sQuestionID & @CRLF & _
				"Прокси: " & $strProxy, True)

	Local $strFileName = @ScriptDir & "\response.xml"
	If FileExists($strFileName) Then FileDelete($strFileName)

	Local $hFile = FileOpen($strFileName, BitOR($FO_OVERWRITE, $FO_ANSI))
	FileWrite($hFile, $strXmlResponse)
	FileClose($hFile)

	Local $resultArray = ParseXmlFileToArray($strFileName)
	If Not IsArray($resultArray) Then _
			SendEmail("Не удалось выполнить разбор XML ответа:" & @CRLF & $strXmlResponse, True)

	If UBound($resultArray, $UBOUND_COLUMNS) <> 7 Then _
			SendEmail("Полученный массив не соответствует заданной ширине: " & @CRLF & _
			_ArrayToString($resultArray), True)

	Local $nResultsLenght = 6
	If $sQuestionID = $sQuestionIDDoctorsQuality Then
		If $sReportID = $sReportIDEmployee Then $nResultsLenght = 11
		If $sReportID = $sReportIDPos Then $nResultsLenght = 10
		If $sReportID = $sReportIDTotal Then $nResultsLenght = 9
	EndIf

	Local $result[0][$nResultsLenght + 1]

	For $i = 1 To UBound($resultArray, $UBOUND_ROWS) - 1
		Local $tmpArray[1][$nResultsLenght + 1]

		For $x = 0 To $nResultsLenght
			Local $value = $resultArray[$i + $x][3]
			If $x > $nResultsLenght - 5 Then _
				$value = StringReplace($value, ".", ",")
			$tmpArray[0][$x] = $value
		Next

		_ArrayAdd($result, $tmpArray)
		$i += $nResultsLenght
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
	ConsoleWrite("err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
			"err.number is: " & @TAB & Hex($oMyError.number, 8) & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext)
EndFunc   ;==>ComErrorHandle


Func SendEmail($sMessage, $bError = False, $sTo = "", $sCopy = "", $sTitle = "", $sAttachments = "")
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
			$sTo = GetEmailAddresses($aNotifyDailyErrors)
		Else
			$sTo = ""
		EndIf
	EndIf

	If Not $sCopy Then $sCopy = $sDeveloperEmail

	If $bError Then
		$sTo = ""

		If IsArray($aErrorNotify) Then
			For $i = 1 To UBound($aErrorNotify, $UBOUND_ROWS) - 1
				$sTo &= $aErrorNotify[$i][1] & ";"
			Next
		Else
			$sTo = $sDeveloperEmail
		EndIf

		$sCopy = ""
	EndIf

	If $bDebug Then
		$sTo = $sDeveloperEmail
		$sCopy = ""
	EndIf

	If Not _INetSmtpMailCom($sMailServer, $sFrom, $sMailLogin, $sTo, _
			$sTitle, $sMessage, $sAttachments, $sCopy, "", $sMailLogin, $sMailPassword) Then _
			_INetSmtpMailCom($sMailServerBackup, $sFrom, $sMailLoginBackup, $sTo, _
			$sTitle, $sMessage, $sAttachments, $sCopy, "", $sMailLoginBackup, $sMailPasswordBackup)

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
