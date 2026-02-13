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
    wsMeet.Cells(1, 8).Value = "Description"
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

    Dim weekdayHours(1 To 7) As Double
    Dim weekdayCount(1 To 7) As Long

    Dim totalHours As Double: totalHours = 0
    Dim totalMeetings As Long: totalMeetings = 0
    Dim maxMeetingHours As Double: maxMeetingHours = 0

    ' ===========================
    ' Loop calendar
    ' ===========================
    Dim apptDate As Date, apptDay As Integer, apptMonth As Integer, apptWeekday As Integer
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

                        apptWeekday = Weekday(appt.Start, vbMonday)
                        weekdayHours(apptWeekday) = weekdayHours(apptWeekday) + durationHours
                        weekdayCount(apptWeekday) = weekdayCount(apptWeekday) + 1

                        ' Raw meetings log
                        wsMeet.Cells(rowMeet, 1).Value = appt.Subject
                        wsMeet.Cells(rowMeet, 2).Value = appt.Start
                        wsMeet.Cells(rowMeet, 3).Value = appt.End
                        wsMeet.Cells(rowMeet, 4).Value = durationHours
                        wsMeet.Cells(rowMeet, 5).Value = isoW
                        wsMeet.Cells(rowMeet, 6).Value = isoY
                        wsMeet.Cells(rowMeet, 7).Value = appt.Categories
                        wsMeet.Cells(rowMeet, 8).Value = GetMeetingBodyPreview(appt.Body, 1000)
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
    wsMeet.Columns("H:H").ColumnWidth = 50
    wsMeet.Columns("H:H").WrapText = True
    wsMeet.Columns.AutoFit

    ' ===========================
    ' Build report sheet
    ' ===========================
    BuildReportSheet wsReport, reportYear, subjectNeedle, categoryNeedle, _
                     totalHours, totalMeetings, maxMeetingHours, _
                     monthHours, monthCount, weekHours, weekCount, _
                     weekdayHours, weekdayCount

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
                             ByVal weekHours As Object, ByVal weekCount As Object, _
                             ByRef weekdayHours() As Double, ByRef weekdayCount() As Long)

    ws.Cells.Clear

    ' Title
    ws.Range("A1:F1").Merge
    ws.Cells(1, 1).Value = "CEDCE Time Report " & reportYear
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ' Criteria
    ws.Cells(3, 1).Value = "Included if:"
    ws.Cells(3, 2).Value = "Subject contains """ & subjectNeedle & """ OR Category = """ & categoryNeedle & """"
    ws.Range("A3:B3").Borders.LineStyle = 1

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
    ws.Range("A5:B12").Borders.LineStyle = 1
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
    ws.Range("A14:C27").Borders.LineStyle = 1
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
        ws.Range("E14:G" & (15 + limit)).Borders.LineStyle = 1
        ws.Columns("E:G").AutoFit
    End If

    ' Weekday table (Mon-Sun)
    ws.Cells(14, 9).Value = "Hours per weekday"
    ws.Cells(14, 9).Font.Bold = True
    ws.Cells(15, 9).Value = "Weekday"
    ws.Cells(15, 10).Value = "Hours"
    ws.Cells(15, 11).Value = "Meetings"
    ws.Range("I15:K15").Font.Bold = True
    For i = 1 To 7
        ws.Cells(15 + i, 9).Value = WeekdayName(i, True, vbMonday)
        ws.Cells(15 + i, 10).Value = weekdayHours(i)
        ws.Cells(15 + i, 11).Value = weekdayCount(i)
    Next i
    ws.Range("J16:J22").NumberFormat = "0.00"
    ws.Range("I14:K22").Borders.LineStyle = 1
    ws.Columns("I:K").AutoFit

    ' All weeks (chronological) - overview
    Dim keysChron As Variant, nWeeks As Long
    If weekHours.Count > 0 Then
        keysChron = weekHours.Keys
        SortKeysByWeekChronological keysChron
        nWeeks = UBound(keysChron) - LBound(keysChron) + 1

        ws.Cells(14, 13).Value = "All weeks (chronological)"
        ws.Cells(14, 13).Font.Bold = True
        ws.Cells(15, 13).Value = "Week"
        ws.Cells(15, 14).Value = "Hours"
        ws.Cells(15, 15).Value = "Meetings"
        ws.Range("M15:O15").Font.Bold = True
        For i = 0 To nWeeks - 1
            ws.Cells(16 + i, 13).Value = keysChron(i)
            ws.Cells(16 + i, 14).Value = weekHours(keysChron(i))
            ws.Cells(16 + i, 15).Value = weekCount(keysChron(i))
        Next i
        ws.Range("N16:N" & (15 + nWeeks)).NumberFormat = "0.00"
        ws.Range("M14:O" & (15 + nWeeks)).Borders.LineStyle = 1
        ws.Columns("M:O").AutoFit
    Else
        nWeeks = 0
    End If

    ' ----- Charts -----
    Dim ch As Object
    Const xlColumnClustered As Long = 51
    Const xlLineMarkers As Long = 65

    ' Chart 1: Hours per month (below monthly table)
    Set ch = ws.ChartObjects.Add(10, 420, 320, 200)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData Source:=ws.Range("A16:B27")
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Hours per month"
    ch.Chart.Axes(1).TickLabels.Font.Size = 9
    ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"

    ' Chart 2: Hours per weekday (next to weekday table)
    Set ch = ws.ChartObjects.Add(340, 420, 260, 200)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData Source:=ws.Range("I16:J22")
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Hours per weekday"
    ch.Chart.Axes(1).TickLabels.Font.Size = 9
    ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"

    ' Chart 3: Top 10 weeks (below left chart)
    If weekHours.Count > 0 Then
        Dim topWeeks As Long
        topWeeks = 10
        If limit < topWeeks Then topWeeks = limit
        Set ch = ws.ChartObjects.Add(10, 640, 400, 200)
        ch.Chart.ChartType = xlColumnClustered
        ch.Chart.SetSourceData Source:=ws.Range("E16:F" & (15 + topWeeks))
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.Text = "Top " & topWeeks & " weeks (hours)"
        ch.Chart.Axes(1).TickLabels.Font.Size = 8
        ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"
    End If

    ' Chart 4: Weekly development (line chart - all weeks in chronological order)
    If nWeeks > 0 Then
        Dim tickSpacing As Long
        tickSpacing = 1
        If nWeeks > 15 Then tickSpacing = Int(nWeeks / 12)
        If tickSpacing < 1 Then tickSpacing = 1

        Set ch = ws.ChartObjects.Add(420, 640, 480, 240)
        ch.Chart.ChartType = xlLineMarkers
        ch.Chart.SetSourceData Source:=ws.Range("M16:N" & (15 + nWeeks))
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.Text = "Hours per week (development)"
        ch.Chart.Axes(1).TickLabels.Font.Size = 8
        ch.Chart.Axes(1).TickLabelSpacing = tickSpacing
        ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"
        ch.Chart.Axes(2).MinimumScale = 0
        ch.Chart.SeriesCollection(1).MarkerStyle = 8
        ch.Chart.SeriesCollection(1).MarkerSize = 4
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

Private Sub SortKeysByWeekChronological(ByRef keys As Variant)
    ' in-place sort: 2025-W52 < 2026-W01 < 2026-W02 ...
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim y1 As Long, w1 As Integer, y2 As Long, w2 As Integer
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            ParseWeekKey CStr(keys(j)), y2, w2
            ParseWeekKey CStr(keys(i)), y1, w1
            If (y2 < y1) Or (y2 = y1 And w2 < w1) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
                y1 = y2
                w1 = w2
            End If
        Next j
    Next i
End Sub

Private Sub ParseWeekKey(ByVal wKey As String, ByRef isoYear As Long, ByRef isoWeek As Integer)
    ' "2026-W03" -> isoYear=2026, isoWeek=3
    Dim p As Long
    isoYear = 0
    isoWeek = 0
    p = InStr(1, wKey, "-W", vbTextCompare)
    If p > 1 Then
        isoYear = CLng(Left$(wKey, p - 1))
        isoWeek = CInt(Mid$(wKey, p + 2, 2))
    End If
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

' ---------- Body preview for Meetings sheet ----------

Private Function GetMeetingBodyPreview(ByVal body As String, ByVal maxLen As Long) As String
    Dim s As String
    If Len(body) = 0 Then
        GetMeetingBodyPreview = ""
        Exit Function
    End If
    s = Replace(body, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) > maxLen Then s = Left$(s, maxLen) & "..."
    GetMeetingBodyPreview = s
End Function
