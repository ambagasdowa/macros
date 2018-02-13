Attribute VB_Name = "Módulo1"
Sub command_1()
    'Config section
    Dim sheetName As String  'Set the name of the sheet where source is fetch
    Dim valueBox As String 'then the trick is in the definition
    sheetName = "source"
    ' Area
    areaBox = Sheets(sheetName).Range("i2").Value + 3
    valueAreaBox = Sheets(sheetName).Range("H" & areaBox).Value
    ' chargeTyppe
    chargeBox = Sheets(sheetName).Range("c2").Value + 3
    valueChargeBox = Sheets(sheetName).Range("B" & chargeBox).Value
      ' Month
    monthBox = Sheets(sheetName).Range("e2").Value + 3
    valueMonthBox = Sheets(sheetName).Range("D" & monthBox).Value
      ' Year
    yearBox = Sheets(sheetName).Range("g2").Value + 3
    valueYearBox = Sheets(sheetName).Range("F" & yearBox).Value


    Dim hoja As Worksheet
    Dim td As PivotTable
    For Each hoja In ThisWorkbook.Worksheets
        For Each td In hoja.PivotTables
            ' Area
            With td.PivotFields("area")
                .ClearAllFilters
                On Error Resume Next
                .CurrentPage = valueAreaBox
            End With
            ' Charge
            With td.PivotFields("Tipo Carga")
                .ClearAllFilters
                On Error Resume Next
                .CurrentPage = valueChargeBox
            End With

            With td.PivotFields("Mes")
                .ClearAllFilters
                On Error Resume Next
                .CurrentPage = valueMonthBox
            End With

            With td.PivotFields("Año")
                .ClearAllFilters
                On Error Resume Next
                .CurrentPage = valueYearBox
            End With

         Next td
     Next
'End If
End Sub

'getLabDay()
Function getlabday(sheetName)
    Dim fecha As Date, last_date As Date, interDate As Date
    Dim fecha_aux As String

    Dim lastDay As String
    Dim UltimoDia As Date
    Dim holidays_range As Range



    'fecha = "01/" & Month(Sheets(sheetName).Range("I4").Value) & "/" & Year(Sheets(sheetName).Range("I4").Value)

    'fecha = DateSerial(Year(Sheets(sheetName).Range("I4").Value), Month(Sheets(sheetName).Range("I4").Value), 1)
    'DateSerial(2004, 6, 30)

    interDate = Date


    fecha_aux = Sheets(sheetName).Range("i4")

    'MsgBox Mid(fecha_aux, 1, 2)

    Dim dia As String, mes As String, annio As String
    Dim WrdArray() As String
    Dim text_string As String
        text_string = fecha_aux
        WrdArray() = Split(text_string, "/", -1)
            'For i = LBound(WrdArray) To UBound(WrdArray)
                'strg = strg & vbNewLine & "Part No. " & i & " - " & WrdArray(i)
            'Next i

    'MsgBox strg



    'MsgBox "dia" & WrdArray(0) & "mes" & WrdArray(1) & "annio" & WrdArray(2)

    'fecha = WrdArray(1) & "/" & WrdArray(0) & "/" & WrdArray(2)
    'fecha = WrdArray(0) & "/" & WrdArray(1) & "/" & WrdArray(2)
    fecha = DateSerial(WrdArray(2), WrdArray(1), WrdArray(0))


    UltimoDia = DateSerial(WrdArray(2), WrdArray(1) + 1, 0)

    lastDay = (Day(UltimoDia))
    'MsgBox lastDay
    'Dim Res As Variant
    'MsgBox WorksheetFunction.NETWORKDAYS_INTL("2016/01/01", "2016/01/31", 11)


    'last_date = lastDay & "/" & Month(Sheets(sheetName).Range("I4").Value) & "/" & Year(Sheets(sheetName).Range("I4").Value)
    'last_date = WrdArray(1) & "/" & lastDay & "/" & WrdArray(2)
    'last_date = lastDay & "/" & WrdArray(1) & "/" & WrdArray(2)
    last_date = DateSerial(WrdArray(2), WrdArray(1), lastDay)

    'MsgBox "fecha => " & fecha
    'MsgBox "last_Date => " & last_date
    'MsgBox "interdate => " & interDate

    'MsgBox NETWORKDAYS_INTL(fecha, last_date, 11, Sheets(sheetName).Range("n3:n25"))
    'MsgBox NETWORKDAYS_INTL(fecha, interDate, 11, Sheets(sheetName).Range("n3:n25"))

    'just in case
    Sheets(sheetName).Range("k4").Value = last_date

    Sheets(sheetName).Range("l4").Value = NETWORKDAYS_INTL(fecha, last_date, 11, Sheets(sheetName).Range("n3:n25"))

    Sheets(sheetName).Range("l5").Value = NETWORKDAYS_INTL(fecha, interDate, 11, Sheets(sheetName).Range("n3:n25"))

    'MsgBox Application.WorksheetFunction.NetworkDays_Intl
    '=NETWORKDAYS(A1+1,B1+1,INDEX(holidays+1,0))
    'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])

End Function




'NETWORKDAYS.ITNL Function
Function NETWORKDAYS_INTL( _
                            start_date As Date, _
                            end_date As Date, _
                            Optional weekend As Variant, _
                            Optional holidays As Variant _
                        ) As Variant

    Dim totalDiff As Integer
    Dim fullWeeks As Integer
    Dim workDays As Integer
    Dim offDays As Integer
    Dim Non_WorkDays(1 To 7) As Boolean
    Dim arHolidays() As String
    Dim noHolidays As Integer
    Dim i As Integer, j As Integer
    Dim cell As Range
    Dim cVal As Variant
    Dim temp As Variant
    Dim DateOrderRev As Integer

    '// ———————————————————————————————————————————————————————————————————————
    '// Check if start date is before end date swap if not
    '// ———————————————————————————————————————————————————————————————————————
    If start_date > end_date Then
        temp = start_date
        start_date = end_date
        end_date = temp
        DateOrderRev = -1
    Else
        DateOrderRev = 1
    End If


    '// 1 = Sunday to 7 = Saturday
    '// ———————————————————————————————————————————————————————————————————————
    '// OPTIONAL ARGUMENT CHECKING 'weekend'
    '// ———————————————————————————————————————————————————————————————————————

    If IsMissing(weekend) Then
        Non_WorkDays(1) = True  '// Sunday
        Non_WorkDays(7) = True  '// Saturday
    '// ———————————————————————————————————————————————————————————————————————
    '// Overloaded Type Checking
    '// ———————————————————————————————————————————————————————————————————————
    '// Argument is astring
    ElseIf TypeName(weekend) = "String" Then
        '// String must contain a total of 7 character either 1's or 0's
        If Len(weekend) = 7 Then
            For i = 1 To 7
                '// Check if a non-(1 or 0) is encountered
                If Mid(weekend, i, 1) <> "1" And Mid(weekend, i, 1) <> "0" Then
                    NETWORKDAYS_INTL = CVErr(xlErrValue)     '// Return #Value!
                    GoTo earlyExit
                End If
                '// NETWORKDAYS.INTL "0000001" = Sunday
                If Mid(weekend, i, 1) = "1" Then
                    If i < 7 Then
                        Non_WorkDays(i + 1) = True
                    Else
                        Non_WorkDays(1) = True
                    End If
                End If
            Next i
        Else
            NETWORKDAYS_INTL = CVErr(xlErrValue)
            GoTo earlyExit
        End If
'Amended By Chuck Hamdan on the date of April 03, 2013
'This amendment is specific to those who want to change the Days off to reflect certain requirements
'Such as, let us say the days off are any of the following:
'======= Days Off =============== Code ==================
'   Saturday & Sunday               1
'   Sunday & Monday                 2
'   Monday & Tuesday                3
'   Tuesday & Wednesday             4
'   Wednesday & Thursday            5
'   Thursday & Friday               6
'   Saturday Only                  17
'   Sunday Only                    11
'   Monday Only                    12
'   Tuesday Only                   13
'   Wednesday Only                 14
'   Thursday Only                  15
'   Friday Only                    16
'   7 Days Work                     0

' For each of those days off would correspond a corresponding code such as:
'       1 for Saturday & Sunday
'       2 for Sunday & Monday
' and so on.
'
' So, we can use a dropdown listbox that list all the days off
' and we would have a cell that would have a formula as listed below:
' =VLOOKUP([The Lookup Value],[The Table Array],2,FALSE)
'
' Check whether TypeName(weekend) = is "Interger" or "Double" or "Range"
'

    ElseIf TypeName(weekend) = "Integer" Or TypeName(weekend) = "Double" Or TypeName(weekend) = "Range" Then
        weekend = Int(weekend)
        If weekend >= 2 And weekend <= 7 Then
            Non_WorkDays(weekend) = True
            Non_WorkDays(weekend - 1) = True
        ElseIf weekend = 1 Then
            Non_WorkDays(1) = True
            Non_WorkDays(7) = True
        ElseIf weekend >= 11 And weekend <= 17 Then
            Non_WorkDays(weekend - 10) = True
        Else
            NETWORKDAYS_INTL = CVErr(xlErrNum)     '//Return #NUM! Error
            GoTo earlyExit
        End If
    Else
        NETWORKDAYS_INTL = CVErr(xlErrValue)
        GoTo earlyExit
    End If
    '// Optional "holidays" argument Handling:
    '// Can be any value or reference to a date value
    '// (Range; Array or single value of a String, Integer, Double, or Date)
    '// ———————————————————————————————————————————————————————————————————————
    '// OPTIONAL ARGUMENT CHECKING 'holidays'
    '// ———————————————————————————————————————————————————————————————————————
    noHolidays = 0

    If Not IsMissing(holidays) Then
        '// ———————————————————————————————————————————————————————————————————
        '// Overloaded Type Checking
        '// ———————————————————————————————————————————————————————————————————
        '// Argument is a Range
        If TypeName(holidays) = "Range" Then
            i = 0
            ReDim arHolidays(1 To holidays.Count)
            For Each cell In holidays
                cVal = cell.Value
                If cVal >= start_date And cVal <= end_date Then
                    arHolidays(i + 1) = cVal
                    i = i + 1
                ElseIf (cVal <> "" And Not IsNumeric(cVal)) And Not IsDate(cVal) Then
                    NETWORKDAYS_INTL = CVErr(xlErrValue)
                    GoTo earlyExit
                End If
            Next cell
            noHolidays = i
        '// Single value multiple types
        '// Argument is a numeric value
        ElseIf IsNumeric(holidays) Then
            holidays = Int(holidays)
            If holidays >= start_date And holidays <= end_date Then
                ReDim arHolidays(1 To 1)
                arHolidays(1) = holidays
                noHolidays = 1
            End If
        '// Argument is a String
        ElseIf TypeName(holidays) = "String" Then
            If DateValue(holidays) >= start_date And DateValue(holidays) <= end_date Then
                ReDim arHolidays(1 To 1)
                arHolidays(1) = DateValue(holidays)
                noHolidays = 1
            End If
        '// Argument is a DATE
        ElseIf TypeName(holidays) = "Date" Then
            If holidays >= start_date And holidays <= end_date Then
                ReDim arHolidays(1 To 1)
                arHolidays(1) = DateValue(holidays)
                noHolidays = 1
            End If
        '// Argument is ARRAY
        ElseIf TypeName(holidays) = "Variant()" Then
            '// Check whats in the Variant Array
            ReDim arHolidays(1 To UBound(holidays))
            j = 0

            For i = 1 To UBound(holidays)
                If TypeName(holidays(i)) = "String" Then
                    cVal = DateValue(holidays(i))
                Else
                    cVal = holidays(i)
                End If

                If cVal >= start_date And cVal <= end_date Then
                    arHolidays(i) = cVal
                    j = j + 1
                ElseIf (cVal <> "" And Not IsNumeric(cVal)) And Not IsDate(cVal) Then
                    NETWORKDAYS_INTL = CVErr(xlErrValue)
                    GoTo earlyExit
                End If
            Next i
            noHolidays = j
        Else
            NETWORKDAYS_INTL = CVErr(xlErrValue)
            GoTo earlyExit
        End If '// Overloaded type checking
    End If  '// IsMissing(holidays)

    If start_date = end_date Then
        If Non_WorkDays(Weekday(start_date)) Then
            NETWORKDAYS_INTL = 0
            GoTo earlyExit
        Else
            NETWORKDAYS_INTL = 1 - noHolidays
            GoTo earlyExit
        End If
    End If

    '// ———————————————————————————————————————————————————————————————————————
    '// Subtract the holidays that fall on a weekend from the total of holidays
    '// ———————————————————————————————————————————————————————————————————————
    If noHolidays > 0 Then
        For i = 1 To noHolidays
            For j = 1 To 7
                If Weekday(arHolidays(i)) = j And Non_WorkDays(j) Then
                    noHolidays = noHolidays - 1
                    Exit For
                End If
            Next j
        Next i
    End If

    offDays = 0
    For i = 1 To 7
        If Non_WorkDays(i) Then offDays = offDays + 1
    Next i

    totalDiff = end_date - start_date + 1

    fullWeeks = Int(totalDiff / 7)
    workDays = ((7 - offDays) * fullWeeks)

    If totalDiff Mod 7 <> 0 Then
        For temp = end_date - (totalDiff Mod 7) + 1 To end_date
            If Non_WorkDays(Weekday(temp)) = False Then
                    workDays = workDays + 1
            End If
        Next
    End If

    NETWORKDAYS_INTL = (workDays - noHolidays) * DateOrderRev

earlyExit:

End Function
