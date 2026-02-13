VERSION 5.00
Begin {0D452530-1598-11D1-B2F4-00C04FB925C3} UserForm ScriptMenu
   Caption         =   "Script-menu"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   StartUpPosition =   1  'CenterOwner
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} CommandButton btnYearReport
      Caption         =   "CEDCE Årsrapport"
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3480
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} CommandButton btnDeleteByTitle
      Caption         =   "Slet kalender efter titel"
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3480
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label lblTitle
      Caption         =   "Vælg script:"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3480
   End
End
Attribute VB_Name = "ScriptMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Script-menu"
    With Me.lblTitle
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

Private Sub btnYearReport_Click()
    Me.Hide
    On Error Resume Next
    Call ExportCEDCEYearReportToExcel
    If Err.Number <> 0 Then MsgBox "Fejl ved kørsel: " & Err.Description, vbExclamation
    Unload Me
End Sub

Private Sub btnDeleteByTitle_Click()
    Me.Hide
    On Error Resume Next
    Call DeleteCalendarItemsByTitle
    If Err.Number <> 0 Then MsgBox "Fejl ved kørsel: " & Err.Description, vbExclamation
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 0
End Sub
