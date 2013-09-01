' ******************************************
'	日期转换
' ******************************************
Function DateToStr(DateTime, ShowType)
    Dim DateMonth, DateDay, DateHour, DateMinute
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    If Len(DateMonth) < 2 Then DateMonth = "0" & DateMonth
    If Len(DateDay) < 2 Then DateDay = "0" & DateDay
    Select Case ShowType
		Case "Y-m"
			DateToStr = Year(DateTime) & "-" & Month(DateTime)
        Case "Y-m-d"
            DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay
        Case "Y-m-d H:I A"
            Dim DateAMPM
            If DateHour > 12 Then
                DateHour = DateHour - 12
                DateAMPM = "PM"
            Else
                DateHour = DateHour
                DateAMPM = "AM"
            End If
            If Len(DateHour) < 2 Then DateHour = "0" & DateHour
            If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
            DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & " " & DateAMPM
        Case "Y-m-d H:I:S"
            Dim DateSecond
            DateSecond = Second(DateTime)
            If Len(DateHour) < 2 Then DateHour = "0" & DateHour
            If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
            If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
            DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & ":" & DateSecond
        Case "YmdHIS"
            DateSecond = Second(DateTime)
            If Len(DateHour) < 2 Then DateHour = "0" & DateHour
            If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
            If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
            DateToStr = Year(DateTime) & DateMonth & DateDay & DateHour & DateMinute & DateSecond
		Case "Ymd"			
            DateToStr = Year(DateTime) & DateMonth & DateDay 
        Case "ym"
            DateToStr = Right(Year(DateTime), 2) & DateMonth
        Case "d"
            DateToStr = DateDay
        Case Else
            If Len(DateHour) < 2 Then DateHour = "0" & DateHour
            If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
            DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute
    End Select
End Function


Function DateToStr2(DateTime, ShowType)  
	Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond
	Dim FullWeekday, shortWeekday, Fullmonth, Shortmonth, TimeZone1, TimeZone2
	TimeZone1 = "+0800"
	TimeZone2 = "+08:00"
	FullWeekday = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
	shortWeekday = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    Fullmonth = Array("January", "February","March","April", "May", "June", "July", "August", "September", "October", "November", "December")
    Shortmonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

	DateMonth = Month(DateTime)
	DateDay = Day(DateTime)
	DateHour = Hour(DateTime)
	DateMinute = Minute(DateTime)
	DateWeek = Weekday(DateTime)
	DateSecond = Second(DateTime)
	If Len(DateMonth) < 2 Then DateMonth = "0" & DateMonth
	If Len(DateDay) < 2 Then DateDay = "0" & DateDay
	If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
	Select Case ShowType
	Case "Y-m-d"  
		DateToStr2 = Year(DateTime) & "-" & DateMonth & "-" & DateDay
	Case "Y-m-d H:I A"
		Dim DateAMPM
		If DateHour > 12 Then 
			DateHour = DateHour - 12
			DateAMPM = "PM"
		Else
			DateHour = DateHour
			DateAMPM = "AM"
		End If
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour	
		DateToStr2 = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & " " & DateAMPM
	Case "Y-m-d H:I:S"
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour	
		If Len(DateSecond) < 2 Then DateSecond="0" & DateSecond
		DateToStr2 = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & ":" & DateSecond
	Case "YmdHIS"
		DateSecond=Second(DateTime)
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour	
		If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
		DateToStr2 = Year(DateTime) & DateMonth & DateDay & DateHour & DateMinute & DateSecond	
	Case "ym"
		DateToStr2 = Right(Year(DateTime), 2) & DateMonth
	Case "d"
		DateToStr2 = DateDay
    Case "ymd"
        DateToStr2 = Right(Year(DateTime), 4) & DateMonth & DateDay
    Case "mdy" 
        Dim DayEnd
        select Case DateDay
         Case 1 
          DayEnd = "st"
         Case 2
          DayEnd = "nd"
         Case 3
          DayEnd = "rd"
         Case Else
          DayEnd = "th"
        End Select 
        DateToStr2 = Fullmonth(DateMonth - 1) & " " & DateDay & DayEnd & " " & Right(Year(DateTime), 4)
    Case "w,d m y H:I:S" 
		DateSecond = Second(DateTime)
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour	
		If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
        DateToStr2 = shortWeekday(DateWeek - 1) & "," & DateDay & " " & Left(Fullmonth(DateMonth - 1), 3) & " " & Right(Year(DateTime), 4) & " " & DateHour & ":" & DateMinute & ":" & DateSecond & " " & TimeZone1
    Case "y-m-dTH:I:S"
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour	
		If Len(DateSecond) < 2 Then DateSecond="0" & DateSecond
		DateToStr2 = Year(DateTime) & "-" & DateMonth & "-" & DateDay & "T" & DateHour & ":" & DateMinute & ":" & DateSecond & TimeZone2
	Case Else
		If Len(DateHour) < 2 Then DateHour = "0" & DateHour
		DateToStr2 = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute
	End Select
End Function


' 要调用的函数声明 
'根据年份及月份得到每月的总天数 
Function GetDaysInMonth(iMonth, iYear) 
Select Case iMonth 
Case 1, 3, 5, 7, 8, 10, 12 
GetDaysInMonth = 31 
Case 4, 6, 9, 11 
GetDaysInMonth = 30 
Case 2 
If IsDate("February 29, " & iYear) Then 
GetDaysInMonth = 29 
Else 
GetDaysInMonth = 28 
End If 
End Select 
End Function 
'得到一个月开始的日期. 
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) 
Dim dTemp 
dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth) 
GetWeekdayMonthStartsOn = WeekDay(dTemp) 
End Function 
'得到当前一个月的上一个月. 
Function SubtractOneMonth(dDate) 
SubtractOneMonth = DateAdd("m", -1, dDate) 
End Function 
'得到当前一个月的下一个月. 
Function AddOneMonth(dDate) 
AddOneMonth = DateAdd("m", 1, dDate) 
End Function 


Function Date2Chinese(iDate)
	Dim num(10)
	Dim iYear
	Dim iMonth
	Dim iDay

	num(0) = "〇"
	num(1) = "一"
	num(2) = "二"
	num(3) = "三"
	num(4) = "四"
	num(5) = "五"
	num(6) = "六"
	num(7) = "七"
	num(8) = "八"
	num(9) = "九"

	iYear = Year(iDate)
	iMonth = Month(iDate)
	iDay = Day(iDate)
	Date2Chinese = (num(iYear \ 1000) + num((iYear \ 100) Mod 10) + num((iYear\ 10) Mod 10) + num(iYear Mod 10)) & "年"

	If iMonth >= 10 Then
		If iMonth = 10 Then
			Date2Chinese = Date2Chinese & "十" & "月"
		Else
			Date2Chinese = Date2Chinese & "十" & num(iMonth Mod 10) & "月"
		End If
	Else
		Date2Chinese = Date2Chinese & num(iMonth Mod 10) & "月"
	End If

	If iDay >= 10 Then
		If iDay = 10 Then
			Date2Chinese = Date2Chinese & "十" & "日"
		ElseIf iDay = 20 or iDay = 30 Then
			Date2Chinese = Date2Chinese & num(iDay \ 10) & "十" & "日"
		ElseIf iDay > 20 Then
			Date2Chinese = Date2Chinese & num(iDay \ 10) & "十" & num(iDay Mod 10) & "日"
		Else
			Date2Chinese = Date2Chinese & "十" & num(iDay Mod 10) & "日"
		End If
	Else
		Date2Chinese = Date2Chinese & num(iDay Mod 10) & "日"
	End If

End Function

Function Date2ChineseRSS(iDate)
	Dim num(10)
	Dim iYear
	Dim iMonth
	Dim iDay

	num(0) = "〇"
	num(1) = "一"
	num(2) = "二"
	num(3) = "三"
	num(4) = "四"
	num(5) = "五"
	num(6) = "六"
	num(7) = "七"
	num(8) = "八"
	num(9) = "九"

	iYear = Year(iDate)
	iMonth = Month(iDate)
	iDay = Day(iDate)
	Date2ChineseRSS = iYear & "年"

	If iMonth >= 10 Then
		If iMonth = 10 Then
			Date2ChineseRSS = Date2ChineseRSS & "十" & "月"
		Else
			Date2ChineseRSS = Date2ChineseRSS & "十" & num(iMonth Mod 10) & "月"
		End If
	Else
		Date2ChineseRSS = Date2ChineseRSS & num(iMonth Mod 10) & "月"
	End If	

End Function

'时间格式处理
Public Function Format_Time(Byval Tvar,Byval sType)
	dim Tt,sYear,sMonth,sDay,sHour,sMinute,sSecond
	If Not IsDate(Tvar) or sType=0 Then Format_Time = "" : Exit Function
	Tt			= Tvar
	sYear		= Year(Tt)
	sMonth		= Right("0" & Month(Tt),2)
	sDay		= Right("0" & Day(Tt),2)
	sHour		= Right("0" & Hour(Tt),2)
	sMinute		= Right("0" & Minute(Tt),2)
	sSecond		= Right("0" & Second(Tt),2)
	Select Case sType
	Case 1	'2005-10-01 23:45:45
		Format_Time = sYear & "-" & sMonth & "-" & sDay & " " & sHour & ":" & sMinute & ":" & sSecond
	Case 2	'年-月-日 时:分:秒
		Format_Time = sYear & "年" & sMonth & "月" & sDay & "日 " & sHour & "时" & sMinute & "分" & sSecond & "秒"
	Case 3	'2005-10-01
		Format_Time = sYear & "-" & sMonth & "-" & sDay
	Case 4	'2005\10\01
		Format_Time = sYear & "\" & sMonth & "\" & sDay
	Case 5	'10-01 23:45
		Format_Time = sMonth & "-" & sDay & " " & sHour & ":" & sMinute
	Case 6	'2005年10月01日
		Format_Time = sYear & "年" & sMonth & "月" & sDay & "日"
	Case 7	'10-01
		Format_Time = sMonth & "-" & sDay
	Case 8	'20051001234545
		Format_Time = sYear & sMonth & sDay & sHour & sMinute & sSecond
	Case Else
		Format_Time = Tt
	End Select
End Function


' Convert a Date to a string
	' IN  : dDate (date) : source
	' OUT : (string) : destination (format YYYYMMDD)
	Function DateToString(dDate)
		DateToString = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2)
	End Function

	' Convert a DateTime to a string
	' IN  : dDateTime (datetime) : source
	' OUT : (string) : destination (format YYYYMMDD HH:MM:SS)
	Function DateTimeToString(dDateTime)
		DateTimeToString = Year(dDateTime) & Right("0" & Month(dDateTime), 2) & Right("0" & Day(dDateTime), 2) & " " & Right("0" & Hour(dDateTime), 2) & ":" & Right("0" & Minute(dDateTime), 2) & ":" & Right("0" & Second(dDateTime), 2)
	End Function

	' Convert a Time to a string
	' IN  : dTime (time) : source
	' OUT : (string) : destination (format HH:MM:SS)
	Function TimeToString(dTime)
		TimeToString = Right("0" & Hour(dTime), 2) & ":" & Right("0" & Minute(dTime), 2) & ":" & Right("0" & Second(dTime), 2)
	End Function


	' Convert a string to a date or datetime
	' IN  : sDate (string) : source (format YYYYMMDD HH:MM:SS or YYYYMMDD)
	' OUT : (datetime) : destination
	Function StringToDate(strDate)
		Dim dDate, sDate

		sDate = trim(strDate)
		select case Len(sDate)
			case 17
				dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2)) + TimeSerial(Mid(sDate, 10, 2), Mid(sDate, 13, 2), Mid(sDate, 16, 2))
			case 8
				dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2))
			case else
				if isDate(sDate) Then
					dDate = CDate(sDate)
				end if
		End select
		StringToDate = dDate
	End Function
