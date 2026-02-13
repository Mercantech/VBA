Option Explicit

' Sletter kalenderaftaler med præcis den angivne emnelinje (hele titlen skal matche).
' Søger kun 5 måneder frem. Nyttigt når mange identiske emner kommer ind uden grund.

Sub DeleteCalendarItemsByTitle()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.Namespace
    Dim calFolder As Outlook.MAPIFolder
    Dim calItems As Outlook.Items
    Dim appt As Outlook.AppointmentItem

    Dim titleToMatch As String
    Dim seenIds As Object
    Dim entryId As Variant
    Dim itemToDelete As Object
    Dim countMatches As Long
    Dim countUnique As Long
    Dim msg As String
    Dim itm As Object

    titleToMatch = Trim$(InputBox("Indtast den præcise kalendertitel (hele emnet skal matche, ikke bare indeholde teksten):" & vbCrLf & vbCrLf & _
        "Søger kun i de næste 5 måneder.", "Slet kalenderemner efter titel", ""))

    If Len(titleToMatch) = 0 Then
        MsgBox "Ingen titel indtastet. Afbrudt.", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrHandler

    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set calFolder = olNS.GetDefaultFolder(olFolderCalendar)

    ' Søg kun 5 måneder frem fra i dag (filtreres i løkken – Restrict med dato er upålidelig i Outlook)
    Dim dateFrom As Date, dateToEnd As Date
    dateFrom = Date
    dateToEnd = DateAdd("d", 1, DateAdd("m", 5, Date))   ' første dag efter 5 måneder

    Set calItems = calFolder.Items
    calItems.IncludeRecurrences = True
    calItems.Sort "[Start]"

    Set seenIds = CreateObject("Scripting.Dictionary")
    countMatches = 0

    For Each itm In calItems
        If itm.Class = olAppointment Then
            Set appt = itm
            If Not appt Is Nothing Then
                ' Stop når vi er forbi søgeperioden (listen er sorteret efter Start)
                If appt.Start >= dateToEnd Then Exit For
                ' Kun med i de næste 5 måneder
                If appt.Start >= dateFrom Then
                    ' Præcis titel-match (hele emnet, ikke "indeholder")
                    If StrComp(Trim$(appt.Subject), titleToMatch, vbTextCompare) = 0 Then
                        countMatches = countMatches + 1
                        If Not seenIds.Exists(appt.EntryID) Then
                            seenIds.Add appt.EntryID, True
                        End If
                    End If
                End If
            End If
        End If
    Next itm

    countUnique = seenIds.Count
    If countUnique = 0 Then
        MsgBox "Ingen aftaler fundet med emnet: """ & titleToMatch & """", vbInformation
        GoTo Cleanup
    End If

    msg = "Fundet " & countMatches & " forekomst(er) (" & countUnique & " unikke serie(r)) med emnet:" & vbCrLf & vbCrLf & _
          """" & titleToMatch & """" & vbCrLf & vbCrLf & _
          "Vil du slette dem alle?"
    If MsgBox(msg, vbYesNo + vbQuestion, "Bekræft sletning") <> vbYes Then
        MsgBox "Sletning annulleret.", vbInformation
        GoTo Cleanup
    End If

    For Each entryId In seenIds.Keys
        On Error Resume Next
        Set itemToDelete = olNS.GetItemFromID(CStr(entryId))
        If Not itemToDelete Is Nothing Then
            itemToDelete.Delete
            Set itemToDelete = Nothing
        End If
        On Error GoTo ErrHandler
    Next entryId

    MsgBox "Slettet " & countUnique & " kalenderemne(r) med emnet """ & titleToMatch & """.", vbInformation

Cleanup:
    Set seenIds = Nothing
    Set calFolder = Nothing
    Set calItems = Nothing
    Set olNS = Nothing
    Set olApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Fejl: " & Err.Description & " (Fejlkode: " & Err.Number & ")", vbCritical
    Resume Cleanup
End Sub
