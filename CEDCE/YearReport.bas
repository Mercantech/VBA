Option Explicit

Sub ExportCEDCEYearReportToExcel()

    ' ===========================
    ' SETTINGS
    ' ===========================
    Dim reportYear As Integer: reportYear = 2026
    Dim subjectNeedle As String: subjectNeedle = "CEDCE"
    Dim categoryNeedle As String: categoryNeedle = "CEDCE"
    ' ===========================

    Dim olApp As Outlook.Application
    Dim olNS As Outlook.Namespace
    Dim calFolder As Outlook.MAPIFolder
    Dim calItems As Outlook.Items
    Dim appt As Outlook.AppointmentItem

    Dim xlApp As Object, wb As Object
    Dim wsReport As Object, wsMeet As Object
    Dim wsMonth(1 To 12) As Object
    Dim nextRow(1 To 12) As Long

    Dim filePath As String
    filePath = Environ$("USERPROFILE") & "\Desktop\CEDCE_year_report_" & reportYear & ".xlsx"

    ' --- Outlook ---
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set calFolder = olNS.GetDefaultFolder(olFolderCalendar)

    Set calItems = calFolder.Items
    calItems.IncludeRecurrences = True
    calItems.Sort "[Start]"

    ' --- Excel ---
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    ' Force decimal point in this Excel session
    Dim oldUseSys As Boolean, oldDec As String, oldThou As String
    oldUseSys = xlApp.UseSystemSeparators
    oldDec = xlApp.DecimalSeparator
    oldThou = xlApp.ThousandsSeparator

    xlApp.UseSystemSeparators = False
    xlApp.DecimalSeparator = "."
    xlApp.ThousandsSeparator = ","

    Set wb = xlApp.Workbooks.Add

    ' ===========================
    ' Report sheet (front page)
    ' ===========================
    Set wsReport = wb.Worksheets(1)
    wsReport.Name = "Report"

    ' Meetings sheet
    Set wsMeet = wb.Worksheets.Add(After:=wsReport)
    wsMeet.Name = "Meetings"

    wsMeet.Cells(1, 1).Value = "Subject"
    wsMeet.Cells(1, 2).Value = "Start"
    wsMeet.Cells(1, 3).Value = "End"
    wsMeet.Cells(1, 4).Value = "Hours"
    wsMeet.Cells(1, 5).Value = "ISO Week"
    wsMeet.Cells(1, 6).Value = "ISO Year"
    wsMeet.Cells(1, 7).Value = "Categories"
    wsMeet.Rows(1).Font.Bold = True

    Dim rowMeet As Long: rowMeet = 2

    ' Month sheets
    Dim m As Integer
    For m = 1 To 12
        Set wsMonth(m) = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsMonth(m).Name = MonthName(m, True)
        SetupMonthSheet wsMonth(m), reportYear, m
        nextRow(m) = 4
    Next m

    ' ===========================
    ' Aggregation containers
    ' ===========================
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim monthHours(1 To 12) As Double
    Dim monthCount(1 To 12) As Long

    Dim weekHours As Object: Set weekHours = CreateObject("Scripting.Dictionary")
    Dim weekCount As Object: Set weekCount = CreateObject("Scripting.Dictionary")

    Dim totalHours As Double: totalHours = 0
    Dim totalMeetings As Long: totalMeetings = 0
    Dim maxMeetingHours As Double: maxMeetingHours = 0

    ' ===========================
    ' Loop calendar
    ' ===========================
    Dim apptDate As Date, apptDay As Integer, apptMonth As Integer
    Dim durationHours As Double, dayCol As Long
    Dim uniqKey As String
    Dim isoW As Integer, isoY As Integer
    Dim wKey As String

    For Each appt In calItems
        If appt.Class = olAppointment Then

            Dim subjectMatch As Boolean
            Dim categoryMatch As Boolean
            Dim includeIt As Boolean

            subjectMatch = (InStr(1, appt.Subject, subjectNeedle, vbTextCompare) > 0)
            categoryMatch = CategoriesContain(appt.Categories, categoryNeedle)
            includeIt = (subjectMatch Or categoryMatch)

            If includeIt Then
                apptDate = DateValue(appt.Start)

                If Year(apptDate) = reportYear Then

                    uniqKey = appt.EntryID & "|" & Format(appt.Start, "yyyy-mm-dd hh:nn")

                    If Not seen.Exists(uniqKey) Then
                        seen.Add uniqKey, True

                        apptMonth = Month(apptDate)
                        apptDay = Day(apptDate)
                        durationHours = appt.Duration / 60#
                        dayCol = 2 + apptDay ' C=1 ... AG=31

                        isoW = ISOWeekNumber(appt.Start)
                        isoY = ISOWeekYear(appt.Start)
                        wKey = CStr(isoY) & "-W" & Format(isoW, "00")

                        ' Totals
                        totalHours = totalHours + durationHours
                        totalMeetings = totalMeetings + 1
                        If durationHours > maxMeetingHours Then maxMeetingHours = durationHours

                        ' Month aggregation
                        monthHours(apptMonth) = monthHours(apptMonth) + durationHours
                        monthCount(apptMonth) = monthCount(apptMonth) + 1

                        ' Week aggregation (only weeks touched)
                        If Not weekHours.Exists(wKey) Then
                            weekHours.Add wKey, 0#
                            weekCount.Add wKey, 0&
                        End If
                        weekHours(wKey) = weekHours(wKey) + durationHours
                        weekCount(wKey) = weekCount(wKey) + 1

                        ' Raw meetings log
                        wsMeet.Cells(rowMeet, 1).Value = appt.Subject
                        wsMeet.Cells(rowMeet, 2).Value = appt.Start
                        wsMeet.Cells(rowMeet, 3).Value = appt.End
                        wsMeet.Cells(rowMeet, 4).Value = durationHours
                        wsMeet.Cells(rowMeet, 5).Value = isoW
                        wsMeet.Cells(rowMeet, 6).Value = isoY
                        wsMeet.Cells(rowMeet, 7).Value = appt.Categories
                        rowMeet = rowMeet + 1

                        ' Month sheet row (matrix)
                        With wsMonth(apptMonth)
                            .Range("A" & nextRow(apptMonth) & ":B" & nextRow(apptMonth)).Merge
                            .Cells(nextRow(apptMonth), 1).Value = appt.Subject
                            .Cells(nextRow(apptMonth), dayCol).Value = durationHours
                        End With
                        nextRow(apptMonth) = nextRow(apptMonth) + 1

                    End If
                End If
            End If

        End If
    Next appt

    ' ===========================
    ' Finalize month sheets
    ' ===========================
    For m = 1 To 12
        FinalizeMonthSheet xlApp, wb, wsMonth(m), nextRow(m)
    Next m

    ' Meetings formatting
    wsMeet.Columns("B:C").NumberFormat = "yyyy-mm-dd hh:mm"
    wsMeet.Columns("D:D").NumberFormat = "0.00"
    wsMeet.Columns.AutoFit

    ' ===========================
    ' Build report sheet
    ' ===========================
    BuildReportSheet wsReport, reportYear, subjectNeedle, categoryNeedle, _
                     totalHours, totalMeetings, maxMeetingHours, _
                     monthHours, monthCount, weekHours, weekCount

    ' ===========================
    ' Save & cleanup
    ' ===========================
    wb.SaveAs filePath, 51
    wb.Close False

    xlApp.UseSystemSeparators = oldUseSys
    xlApp.DecimalSeparator = oldDec
    xlApp.ThousandsSeparator = oldThou

    xlApp.Quit

    Set weekCount = Nothing
    Set weekHours = Nothing
    Set seen = Nothing

    Set wsReport = Nothing
    Set wsMeet = Nothing
    For m = 1 To 12
        Set wsMonth(m) = Nothing
    Next m
    Set wb = Nothing
    Set xlApp = Nothing

    MsgBox "Færdig! Årsrapport gemt som CEDCE_year_report_" & reportYear & ".xlsx", vbInformation

End Sub

' ---------- Report helpers ----------

Private Sub BuildReportSheet(ByVal ws As Object, ByVal reportYear As Integer, _
                             ByVal subjectNeedle As String, ByVal categoryNeedle As String, _
                             ByVal totalHours As Double, ByVal totalMeetings As Long, ByVal maxMeetingHours As Double, _
                             ByRef monthHours() As Double, ByRef monthCount() As Long, _
                             ByVal weekHours As Object, ByVal weekCount As Object)

    ws.Cells.Clear

    ' Title
    ws.Range("A1:F1").Merge
    ws.Cells(1, 1).Value = "CEDCE Time Report " & reportYear
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ' Criteria
    ws.Cells(3, 1).Value = "Included if:"
    ws.Cells(3, 2).Value = "Subject contains """ & subjectNeedle & """ OR Category = """ & categoryNeedle & """"

    ' KPI block
    Dim activeMonths As Long, activeWeeks As Long
    activeMonths = CountActiveMonths(monthHours)
    activeWeeks = weekHours.Count

    Dim avgHoursAllMonths As Double
    avgHoursAllMonths = totalHours / 12#

    Dim avgHoursActiveMonths As Double
    If activeMonths > 0 Then avgHoursActiveMonths = totalHours / activeMonths Else avgHoursActiveMonths = 0

    Dim avgHoursActiveWeeks As Double
    If activeWeeks > 0 Then avgHoursActiveWeeks = totalHours / activeWeeks Else avgHoursActiveWeeks = 0

    Dim avgMeetingHours As Double
    If totalMeetings > 0 Then avgMeetingHours = totalHours / totalMeetings Else avgMeetingHours = 0

    ws.Cells(5, 1).Value = "Key figures"
    ws.Cells(5, 1).Font.Bold = True

    ws.Cells(6, 1).Value = "Total hours"
    ws.Cells(6, 2).Value = totalHours

    ws.Cells(7, 1).Value = "Total meetings"
    ws.Cells(7, 2).Value = totalMeetings

    ws.Cells(8, 1).Value = "Avg meeting duration (hours)"
    ws.Cells(8, 2).Value = avgMeetingHours

    ws.Cells(9, 1).Value = "Max meeting duration (hours)"
    ws.Cells(9, 2).Value = maxMeetingHours

    ws.Cells(10, 1).Value = "Avg hours / month (all 12)"
    ws.Cells(10, 2).Value = avgHoursAllMonths

    ws.Cells(11, 1).Value = "Avg hours / month (months with activity)"
    ws.Cells(11, 2).Value = avgHoursActiveMonths

    ws.Cells(12, 1).Value = "Avg hours / week (weeks with activity)"
    ws.Cells(12, 2).Value = avgHoursActiveWeeks

    ws.Range("B6:B12").NumberFormat = "0.00"
    ws.Columns("A:B").AutoFit

    ' Monthly table
    ws.Cells(14, 1).Value = "Monthly totals"
    ws.Cells(14, 1).Font.Bold = True

    ws.Cells(15, 1).Value = "Month"
    ws.Cells(15, 2).Value = "Hours"
    ws.Cells(15, 3).Value = "Meetings"
    ws.Range("A15:C15").Font.Bold = True

    Dim r As Long: r = 16
    Dim m As Integer
    For m = 1 To 12
        ws.Cells(r, 1).Value = MonthName(m, True)
        ws.Cells(r, 2).Value = monthHours(m)
        ws.Cells(r, 3).Value = monthCount(m)
        r = r + 1
    Next m
    ws.Range("B16:B27").NumberFormat = "0.00"
    ws.Columns("A:C").AutoFit

    ' Weekly table (top 25 by hours)
    ws.Cells(14, 5).Value = "Top weeks (by hours)"
    ws.Cells(14, 5).Font.Bold = True

    ws.Cells(15, 5).Value = "Week"
    ws.Cells(15, 6).Value = "Hours"
    ws.Cells(15, 7).Value = "Meetings"
    ws.Range("E15:G15").Font.Bold = True

    Dim keys As Variant, i As Long
    keys = weekHours.Keys
    If weekHours.Count > 0 Then
        SortKeysByWeekHoursDesc keys, weekHours

        Dim limit As Long: limit = 25
        If UBound(keys) - LBound(keys) + 1 < limit Then limit = UBound(keys) - LBound(keys) + 1

        For i = 0 To limit - 1
            ws.Cells(16 + i, 5).Value = keys(i)
            ws.Cells(16 + i, 6).Value = weekHours(keys(i))
            ws.Cells(16 + i, 7).Value = weekCount(keys(i))
        Next i

        ws.Range("F16:F" & (15 + limit)).NumberFormat = "0.00"
        ws.Columns("E:G").AutoFit
    End If

End Sub

Private Function CountActiveMonths(ByRef monthHours() As Double) As Long
    Dim m As Integer, c As Long
    c = 0
    For m = LBound(monthHours) To UBound(monthHours)
        If monthHours(m) > 0.00001 Then c = c + 1
    Next m
    CountActiveMonths = c
End Function

Private Sub SortKeysByWeekHoursDesc(ByRef keys As Variant, ByVal weekHours As Object)
    ' simple in-place sort (desc by hours)
    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CDbl(weekHours(keys(j))) > CDbl(weekHours(keys(i))) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
End Sub

' ---------- Month sheet helpers ----------

Private Sub SetupMonthSheet(ByVal ws As Object, ByVal yearNum As Integer, ByVal monthNum As Integer)

    Dim d As Integer, dayCol As Long

    ws.Range("A1:B1").Merge
    ws.Cells(1, 1).Value = "Event title"
    ws.Rows(1).Font.Bold = True

    For d = 1 To 31
        dayCol = 2 + d
        ws.Cells(1, dayCol).Value = d
    Next d

    ws.Range("A2:B2").Merge
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "Month: " & Format(DateSerial(yearNum, monthNum, 1), "yyyy-mm")

    ws.Columns("A:B").ColumnWidth = 55
    ws.Columns("A:B").WrapText = True
    ws.Range("C:AG").NumberFormat = "0.00"
    ws.Columns("C:AG").ColumnWidth = 4

End Sub

Private Sub FinalizeMonthSheet(ByVal xlApp As Object, ByVal wb As Object, ByVal ws As Object, ByVal nextRow As Long)

    Dim d As Integer, dayCol As Long
    Dim sumRow As Long

    sumRow = nextRow + 1
    ws.Range("A" & sumRow & ":B" & sumRow).Merge
    ws.Cells(sumRow, 1).Value = "TOTAL pr. dag"
    ws.Cells(sumRow, 1).Font.Bold = True

    If nextRow > 4 Then
        For d = 1 To 31
            dayCol = 2 + d
            ws.Cells(sumRow, dayCol).Formula = _
                "=SUM(" & ws.Cells(4, dayCol).Address(False, False) & ":" & ws.Cells(nextRow - 1, dayCol).Address(False, False) & ")"
            ws.Cells(sumRow, dayCol).Font.Bold = True
            ws.Cells(sumRow, dayCol).NumberFormat = "0.00"
        Next d
    End If

    xlApp.Windows(wb.Name).Activate
    ws.Activate
    ws.Range("C4").Select
    xlApp.ActiveWindow.FreezePanes = True

End Sub

' ---------- ISO week helpers ----------

Private Function ISOWeekNumber(d As Date) As Integer
    ISOWeekNumber = DatePart("ww", d, vbMonday, vbFirstFourDays)
End Function

Private Function ISOWeekYear(d As Date) As Integer
    ISOWeekYear = Year(DateAdd("d", 4 - Weekday(d, vbMonday), d))
End Function

' ---------- Categories helper ----------

Private Function CategoriesContain(ByVal categoriesText As String, ByVal categoryName As String) As Boolean
    CategoriesContain = False
    If Len(Trim$(categoriesText)) = 0 Then Exit Function

    Dim parts() As String
    Dim i As Long, p As String

    parts = Split(categoriesText, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If StrComp(p, categoryName, vbTextCompare) = 0 Then
            CategoriesContain = True
            Exit Function
        End If
    Next i
End Function
