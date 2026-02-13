Option Explicit

' Årsrapport med valgfrit projekt og periode (fra–til dato). Beder om projekt + start/slut før kørsel.
Sub ExportProjectYearReportToExcel()

    ' ===========================
    ' PROJEKTNAVN (bruger angiver)
    ' ===========================
    Dim projectName As String
    projectName = Trim$(InputBox("Indtast projektnavn (bruges til emne/kategori og filnavn):" & vbCrLf & vbCrLf & _
        "Fx: CEDCE, MYRE", "Årsrapport – vælg projekt", "CEDCE"))
    If Len(projectName) = 0 Then
        MsgBox "Ingen projektnavn indtastet. Afbrudt.", vbInformation
        Exit Sub
    End If

    ' ===========================
    ' PERIODE (fra / til dato)
    ' ===========================
    Dim dateFrom As Date, dateTo As Date
    Dim dateFromStr As String, dateToStr As String
    dateFromStr = Trim$(InputBox("Fra dato (yyyy-mm-dd eller dd-mm-yyyy):", "Rapportperiode", Format(DateSerial(Year(Date), 1, 1), "yyyy-mm-dd")))
    If Len(dateFromStr) = 0 Then MsgBox "Ingen startdato. Afbrudt.", vbInformation: Exit Sub
    dateToStr = Trim$(InputBox("Til dato (yyyy-mm-dd eller dd-mm-yyyy):", "Rapportperiode", Format(DateSerial(Year(Date), 12, 31), "yyyy-mm-dd")))
    If Len(dateToStr) = 0 Then MsgBox "Ingen slutdato. Afbrudt.", vbInformation: Exit Sub
    If Not IsDate(dateFromStr) Then MsgBox "Ugyldig fra-dato.", vbExclamation: Exit Sub
    If Not IsDate(dateToStr) Then MsgBox "Ugyldig til-dato.", vbExclamation: Exit Sub
    dateFrom = CDate(dateFromStr)
    dateTo = CDate(dateToStr)
    If dateFrom > dateTo Then
        MsgBox "Fra-dato skal være før eller lig til-dato.", vbExclamation
        Exit Sub
    End If

    ' ===========================
    ' SETTINGS
    ' ===========================
    Dim subjectNeedle As String: subjectNeedle = projectName
    Dim categoryNeedle As String: categoryNeedle = projectName
    ' ===========================

    Dim olApp As Outlook.Application
    Dim olNS As Outlook.Namespace
    Dim calFolder As Outlook.MAPIFolder
    Dim calItems As Outlook.Items
    Dim appt As Outlook.AppointmentItem

    Dim xlApp As Object, wb As Object
    Dim wsReport As Object, wsMeet As Object

    ' Måneder i perioden (dynamisk liste "yyyy-mm")
    Dim monthKeys() As String
    Dim numMonths As Long
    monthKeys = GetMonthKeysInRange(dateFrom, dateTo)
    numMonths = UBound(monthKeys) - LBound(monthKeys) + 1

    Dim wsMonth() As Object
    Dim nextRow() As Long
    ReDim wsMonth(1 To numMonths)
    ReDim nextRow(1 To numMonths)

    Dim monthHours As Object, monthCount As Object
    Set monthHours = CreateObject("Scripting.Dictionary")
    Set monthCount = CreateObject("Scripting.Dictionary")
    Dim mi As Long
    For mi = 1 To numMonths
        monthHours(monthKeys(mi)) = 0#
        monthCount(monthKeys(mi)) = 0&
        nextRow(mi) = 4
    Next mi

    Dim monthKeyToIndex As Object
    Set monthKeyToIndex = CreateObject("Scripting.Dictionary")
    For mi = 1 To numMonths
        monthKeyToIndex(monthKeys(mi)) = mi
    Next mi

    Dim filePath As String
    filePath = Environ$("USERPROFILE") & "\Desktop\" & projectName & "_report_" & Format(dateFrom, "yyyy-mm-dd") & "_til_" & Format(dateTo, "yyyy-mm-dd") & ".xlsx"

    ' --- Hvilken kalender? (egen eller kollegas – åbne kalendre i virksomheden) ---
    Const COMPANY_DOMAIN As String = "mercantec.dk"
    Dim useOwnCalendar As Boolean
    Dim otherCalendarOwner As String
    Dim resp As String
    resp = Trim$(InputBox("Brug din egen kalender?" & vbCrLf & vbCrLf & "J = Ja (egen kalender)" & vbCrLf & "N = Nej – brug en kollegas kalender (alle har åbne kalendre i virksomheden).", "Vælg kalender", "J"))
    If StrComp(UCase$(Left$(resp, 1)), "J", vbBinaryCompare) = 0 Or Len(resp) = 0 Then
        useOwnCalendar = True
    Else
        useOwnCalendar = False
        otherCalendarOwner = Trim$(InputBox("Indtast kollegaens navn eller email (fx ""Anders Jensen"" eller ""anders.jensen""). Ved kun navn/brugernavn bruges @" & COMPANY_DOMAIN, "Anden kalender", ""))
        If Len(otherCalendarOwner) = 0 Then
            MsgBox "Ingen kalender-ejer angivet. Afbrudt.", vbInformation
            Exit Sub
        End If
        ' Kun brugernavn (uden mellemrum)? Så tilføj @domæne. "Anders Jensen" lades stå – Outlook løser visningsnavn.
        If InStr(1, otherCalendarOwner, "@", vbBinaryCompare) = 0 And InStr(1, otherCalendarOwner, " ", vbBinaryCompare) = 0 Then
            otherCalendarOwner = otherCalendarOwner & "@" & COMPANY_DOMAIN
        End If
    End If

    ' --- Outlook ---
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    If useOwnCalendar Then
        Set calFolder = olNS.GetDefaultFolder(olFolderCalendar)
    Else
        Dim recip As Outlook.Recipient
        Set recip = olNS.CreateRecipient(otherCalendarOwner)
        recip.Resolve
        If Not recip.Resolved Then
            MsgBox "Kunne ikke finde kalender-ejeren: """ & otherCalendarOwner & """. Tjek stavning eller brug fuld email-adresse.", vbExclamation
            Exit Sub
        End If
        Set calFolder = olNS.GetSharedDefaultFolder(recip, olFolderCalendar)
    End If

    Set calItems = calFolder.Items
    calItems.IncludeRecurrences = True
    calItems.Sort "[Start]"
    ' Kun hent aftaler i rapport-perioden (undgår at loade hele kalenderen – især vigtigt ved andres kalender)
    Dim restrictFilter As String
    restrictFilter = "[Start] >= '" & Format(dateFrom, "yyyy-mm-dd 00:00") & "' AND [Start] <= '" & Format(dateTo, "yyyy-mm-dd 23:59") & "'"
    Set calItems = calItems.Restrict(restrictFilter)

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

    ' Month sheets (kun for måneder i perioden)
    Dim m As Long, y As Integer, mo As Integer
    For m = 1 To numMonths
        Set wsMonth(m) = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ' Arknavn: "2026-03" eller "Mar 2026" – brug kort format så det er læsbart
        y = CInt(Left$(monthKeys(m), 4))
        mo = CInt(Mid$(monthKeys(m), 6, 2))
        wsMonth(m).Name = MonthName(mo, True) & " " & y
        SetupMonthSheetProject wsMonth(m), y, mo
    Next m

    ' ===========================
    ' Aggregation containers
    ' ===========================
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

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
            categoryMatch = CategoriesContainProject(appt.Categories, categoryNeedle)
            includeIt = (subjectMatch Or categoryMatch)

            If includeIt Then
                apptDate = DateValue(appt.Start)

                If apptDate >= dateFrom And apptDate <= dateTo Then

                    uniqKey = appt.EntryID & "|" & Format(appt.Start, "yyyy-mm-dd hh:nn")

                    If Not seen.Exists(uniqKey) Then
                        seen.Add uniqKey, True

                        apptMonth = Month(apptDate)
                        apptDay = Day(apptDate)
                        durationHours = appt.Duration / 60#
                        dayCol = 2 + apptDay ' C=1 ... AG=31

                        Dim monthKey As String
                        monthKey = Format(DateSerial(Year(apptDate), apptMonth, 1), "yyyy-mm")
                        If monthHours.Exists(monthKey) Then
                            monthHours(monthKey) = monthHours(monthKey) + durationHours
                            monthCount(monthKey) = monthCount(monthKey) + 1
                        End If

                        isoW = ISOWeekNumberProject(appt.Start)
                        isoY = ISOWeekYearProject(appt.Start)
                        wKey = CStr(isoY) & "-W" & Format(isoW, "00")

                        ' Totals
                        totalHours = totalHours + durationHours
                        totalMeetings = totalMeetings + 1
                        If durationHours > maxMeetingHours Then maxMeetingHours = durationHours

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
                        wsMeet.Cells(rowMeet, 8).Value = GetMeetingBodyPreviewProject(appt.Body, 1000)
                        rowMeet = rowMeet + 1

                        ' Month sheet row (matrix) – kun hvis måneden er i perioden
                        If monthKeyToIndex.Exists(monthKey) Then
                            mi = monthKeyToIndex(monthKey)
                            With wsMonth(mi)
                                .Range("A" & nextRow(mi) & ":B" & nextRow(mi)).Merge
                                .Cells(nextRow(mi), 1).Value = appt.Subject
                                .Cells(nextRow(mi), dayCol).Value = durationHours
                            End With
                            nextRow(mi) = nextRow(mi) + 1
                        End If

                    End If
                End If
            End If

        End If
    Next appt

    ' ===========================
    ' Finalize month sheets
    ' ===========================
    For m = 1 To numMonths
        FinalizeMonthSheetProject xlApp, wb, wsMonth(m), nextRow(m)
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
    BuildReportSheetProject wsReport, dateFrom, dateTo, subjectNeedle, categoryNeedle, _
                     totalHours, totalMeetings, maxMeetingHours, _
                     monthKeys, numMonths, monthHours, monthCount, weekHours, weekCount, _
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
    For m = 1 To numMonths
        Set wsMonth(m) = Nothing
    Next m
    Set monthKeyToIndex = Nothing
    Set monthHours = Nothing
    Set monthCount = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    MsgBox "Færdig! Rapport gemt som " & projectName & "_report_" & Format(dateFrom, "yyyy-mm-dd") & "_til_" & Format(dateTo, "yyyy-mm-dd") & ".xlsx", vbInformation

End Sub

' ---------- Report helpers ----------

Private Sub BuildReportSheetProject(ByVal ws As Object, ByVal dateFrom As Date, ByVal dateTo As Date, _
                             ByVal subjectNeedle As String, ByVal categoryNeedle As String, _
                             ByVal totalHours As Double, ByVal totalMeetings As Long, ByVal maxMeetingHours As Double, _
                             ByRef monthKeys() As String, ByVal numMonths As Long, _
                             ByVal monthHours As Object, ByVal monthCount As Object, _
                             ByVal weekHours As Object, ByVal weekCount As Object, _
                             ByRef weekdayHours() As Double, ByRef weekdayCount() As Long)

    ws.Cells.Clear

    ' Title (projekt + periode)
    ws.Range("A1:F1").Merge
    ws.Cells(1, 1).Value = subjectNeedle & " Time Report"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ws.Cells(2, 1).Value = "Rapportperiode: " & Format(dateFrom, "dd. mmm yyyy") & " – " & Format(dateTo, "dd. mmm yyyy")
    ws.Range("A2:F2").Merge

    ' Criteria
    ws.Cells(4, 1).Value = "Included if:"
    ws.Cells(4, 2).Value = "Subject contains """ & subjectNeedle & """ OR Category = """ & categoryNeedle & """"
    ws.Range("A4:B4").Borders.LineStyle = 1

    ' KPI block
    Dim activeMonths As Long, activeWeeks As Long
    activeMonths = CountActiveMonthsProject(monthHours, monthKeys, numMonths)
    activeWeeks = weekHours.Count

    Dim avgHoursAllMonths As Double
    If numMonths > 0 Then avgHoursAllMonths = totalHours / numMonths Else avgHoursAllMonths = 0

    Dim avgHoursActiveMonths As Double
    If activeMonths > 0 Then avgHoursActiveMonths = totalHours / activeMonths Else avgHoursActiveMonths = 0

    Dim avgHoursActiveWeeks As Double
    If activeWeeks > 0 Then avgHoursActiveWeeks = totalHours / activeWeeks Else avgHoursActiveWeeks = 0

    Dim avgMeetingHours As Double
    If totalMeetings > 0 Then avgMeetingHours = totalHours / totalMeetings Else avgMeetingHours = 0

    ws.Cells(6, 1).Value = "Key figures"
    ws.Cells(6, 1).Font.Bold = True

    ws.Cells(7, 1).Value = "Total hours"
    ws.Cells(7, 2).Value = totalHours

    ws.Cells(8, 1).Value = "Total meetings"
    ws.Cells(8, 2).Value = totalMeetings

    ws.Cells(9, 1).Value = "Avg meeting duration (hours)"
    ws.Cells(9, 2).Value = avgMeetingHours

    ws.Cells(10, 1).Value = "Max meeting duration (hours)"
    ws.Cells(10, 2).Value = maxMeetingHours

    ws.Cells(11, 1).Value = "Avg hours / month (hele perioden)"
    ws.Cells(11, 2).Value = avgHoursAllMonths

    ws.Cells(12, 1).Value = "Avg hours / month (måneder med aktivitet)"
    ws.Cells(12, 2).Value = avgHoursActiveMonths

    ws.Cells(13, 1).Value = "Avg hours / week (uger med aktivitet)"
    ws.Cells(13, 2).Value = avgHoursActiveWeeks

    ws.Range("B7:B13").NumberFormat = "0.00"
    ws.Range("A6:B13").Borders.LineStyle = 1
    ws.Columns("A:B").AutoFit

    ' Monthly table (kun måneder i perioden)
    ws.Cells(15, 1).Value = "Monthly totals"
    ws.Cells(15, 1).Font.Bold = True

    ws.Cells(16, 1).Value = "Month"
    ws.Cells(16, 2).Value = "Hours"
    ws.Cells(16, 3).Value = "Meetings"
    ws.Range("A16:C16").Font.Bold = True

    Dim r As Long: r = 17
    Dim m As Long, key As String, y As Integer, mo As Integer
    For m = 1 To numMonths
        key = monthKeys(m)
        y = CInt(Left$(key, 4))
        mo = CInt(Mid$(key, 6, 2))
        ws.Cells(r, 1).Value = MonthName(mo, True) & " " & y
        ws.Cells(r, 2).Value = monthHours(key)
        ws.Cells(r, 3).Value = monthCount(key)
        r = r + 1
    Next m
    ws.Range("B17:B" & (16 + numMonths)).NumberFormat = "0.00"
    ws.Range("A15:C" & (16 + numMonths)).Borders.LineStyle = 1
    ws.Columns("A:C").AutoFit

    ' Weekly table (top 25 by hours)
    Dim lastDataRow As Long
    lastDataRow = 16 + numMonths
    ws.Cells(15, 5).Value = "Top weeks (by hours)"
    ws.Cells(15, 5).Font.Bold = True

    ws.Cells(16, 5).Value = "Week"
    ws.Cells(16, 6).Value = "Hours"
    ws.Cells(16, 7).Value = "Meetings"
    ws.Range("E16:G16").Font.Bold = True

    Dim keys As Variant, i As Long
    keys = weekHours.Keys
    If weekHours.Count > 0 Then
        SortKeysByWeekHoursDescProject keys, weekHours

        Dim limit As Long: limit = 25
        If UBound(keys) - LBound(keys) + 1 < limit Then limit = UBound(keys) - LBound(keys) + 1

        For i = 0 To limit - 1
            ws.Cells(17 + i, 5).Value = keys(i)
            ws.Cells(17 + i, 6).Value = weekHours(keys(i))
            ws.Cells(17 + i, 7).Value = weekCount(keys(i))
        Next i

        ws.Range("F17:F" & (16 + limit)).NumberFormat = "0.00"
        ws.Range("E15:G" & (16 + limit)).Borders.LineStyle = 1
        ws.Columns("E:G").AutoFit
    End If

    ' Weekday table (Mon-Sun)
    ws.Cells(15, 9).Value = "Hours per weekday"
    ws.Cells(15, 9).Font.Bold = True
    ws.Cells(16, 9).Value = "Weekday"
    ws.Cells(16, 10).Value = "Hours"
    ws.Cells(16, 11).Value = "Meetings"
    ws.Range("I16:K16").Font.Bold = True
    For i = 1 To 7
        ws.Cells(16 + i, 9).Value = WeekdayName(i, True, vbMonday)
        ws.Cells(16 + i, 10).Value = weekdayHours(i)
        ws.Cells(16 + i, 11).Value = weekdayCount(i)
    Next i
    ws.Range("J17:J23").NumberFormat = "0.00"
    ws.Range("I15:K23").Borders.LineStyle = 1
    ws.Columns("I:K").AutoFit

    ' All weeks (chronological) - overview
    Dim keysChron As Variant, nWeeks As Long
    If weekHours.Count > 0 Then
        keysChron = weekHours.Keys
        SortKeysByWeekChronologicalProject keysChron
        nWeeks = UBound(keysChron) - LBound(keysChron) + 1

        ws.Cells(15, 13).Value = "All weeks (chronological)"
        ws.Cells(15, 13).Font.Bold = True
        ws.Cells(16, 13).Value = "Week"
        ws.Cells(16, 14).Value = "Hours"
        ws.Cells(16, 15).Value = "Meetings"
        ws.Range("M16:O16").Font.Bold = True
        For i = 0 To nWeeks - 1
            ws.Cells(17 + i, 13).Value = keysChron(i)
            ws.Cells(17 + i, 14).Value = weekHours(keysChron(i))
            ws.Cells(17 + i, 15).Value = weekCount(keysChron(i))
        Next i
        ws.Range("N17:N" & (16 + nWeeks)).NumberFormat = "0.00"
        ws.Range("M15:O" & (16 + nWeeks)).Borders.LineStyle = 1
        ws.Columns("M:O").AutoFit
    Else
        nWeeks = 0
    End If

    ' ----- Charts -----
    Dim ch As Object
    Const xlColumnClustered As Long = 51
    Const xlLineMarkers As Long = 65

    ' Chart 1: Hours per month (kun måneder i perioden)
    Set ch = ws.ChartObjects.Add(10, 420, 320, 200)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData Source:=ws.Range("A17:B" & lastDataRow)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Hours per month"
    ch.Chart.Axes(1).TickLabels.Font.Size = 9
    ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"

    ' Chart 2: Hours per weekday
    Set ch = ws.ChartObjects.Add(340, 420, 260, 200)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData Source:=ws.Range("I17:J23")
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Hours per weekday"
    ch.Chart.Axes(1).TickLabels.Font.Size = 9
    ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"

    ' Chart 3: Top 10 weeks
    If weekHours.Count > 0 Then
        Dim topWeeks As Long
        topWeeks = 10
        If limit < topWeeks Then topWeeks = limit
        Set ch = ws.ChartObjects.Add(10, 640, 400, 200)
        ch.Chart.ChartType = xlColumnClustered
        ch.Chart.SetSourceData Source:=ws.Range("E17:F" & (16 + topWeeks))
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.Text = "Top " & topWeeks & " weeks (hours)"
        ch.Chart.Axes(1).TickLabels.Font.Size = 8
        ch.Chart.Axes(2).TickLabels.NumberFormat = "0.0"
    End If

    ' Chart 4: Weekly development
    If nWeeks > 0 Then
        Dim tickSpacing As Long
        tickSpacing = 1
        If nWeeks > 15 Then tickSpacing = Int(nWeeks / 12)
        If tickSpacing < 1 Then tickSpacing = 1

        Set ch = ws.ChartObjects.Add(420, 640, 480, 240)
        ch.Chart.ChartType = xlLineMarkers
        ch.Chart.SetSourceData Source:=ws.Range("M17:N" & (16 + nWeeks))
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

' Returnerer array af "yyyy-mm" for alle måneder fra dateFrom til dateTo (1-baseret, kronologisk).
Private Function GetMonthKeysInRange(ByVal dateFrom As Date, ByVal dateTo As Date) As String()
    Dim d As Date
    Dim keys() As String
    Dim n As Long, i As Long
    n = (Year(dateTo) - Year(dateFrom)) * 12 + (Month(dateTo) - Month(dateFrom)) + 1
    If n < 1 Then n = 1
    ReDim keys(1 To n)
    d = DateSerial(Year(dateFrom), Month(dateFrom), 1)
    For i = 1 To n
        keys(i) = Format(d, "yyyy-mm")
        d = DateAdd("m", 1, d)
    Next i
    GetMonthKeysInRange = keys
End Function

Private Function CountActiveMonthsProject(ByVal monthHours As Object, ByRef monthKeys() As String, ByVal numMonths As Long) As Long
    Dim m As Long, c As Long, key As String
    c = 0
    For m = 1 To numMonths
        key = monthKeys(m)
        If monthHours.Exists(key) Then
            If CDbl(monthHours(key)) > 0.00001 Then c = c + 1
        End If
    Next m
    CountActiveMonthsProject = c
End Function

Private Sub SortKeysByWeekHoursDescProject(ByRef keys As Variant, ByVal weekHours As Object)
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

Private Sub SortKeysByWeekChronologicalProject(ByRef keys As Variant)
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim y1 As Long, w1 As Integer, y2 As Long, w2 As Integer
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            ParseWeekKeyProject CStr(keys(j)), y2, w2
            ParseWeekKeyProject CStr(keys(i)), y1, w1
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

Private Sub ParseWeekKeyProject(ByVal wKey As String, ByRef isoYear As Long, ByRef isoWeek As Integer)
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

Private Sub SetupMonthSheetProject(ByVal ws As Object, ByVal yearNum As Integer, ByVal monthNum As Integer)
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

Private Sub FinalizeMonthSheetProject(ByVal xlApp As Object, ByVal wb As Object, ByVal ws As Object, ByVal nextRow As Long)
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

Private Function ISOWeekNumberProject(d As Date) As Integer
    ISOWeekNumberProject = DatePart("ww", d, vbMonday, vbFirstFourDays)
End Function

Private Function ISOWeekYearProject(d As Date) As Integer
    ISOWeekYearProject = Year(DateAdd("d", 4 - Weekday(d, vbMonday), d))
End Function

' ---------- Categories helper ----------

Private Function CategoriesContainProject(ByVal categoriesText As String, ByVal categoryName As String) As Boolean
    CategoriesContainProject = False
    If Len(Trim$(categoriesText)) = 0 Then Exit Function
    Dim parts() As String
    Dim i As Long, p As String
    parts = Split(categoriesText, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If StrComp(p, categoryName, vbTextCompare) = 0 Then
            CategoriesContainProject = True
            Exit Function
        End If
    Next i
End Function

' ---------- Body preview ----------

Private Function GetMeetingBodyPreviewProject(ByVal body As String, ByVal maxLen As Long) As String
    Dim s As String
    If Len(body) = 0 Then
        GetMeetingBodyPreviewProject = ""
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
    GetMeetingBodyPreviewProject = s
End Function
