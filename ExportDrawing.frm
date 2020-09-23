VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExportDrawing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export to Drawing"
   ClientHeight    =   2700
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3612
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3612
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Elements"
      Height          =   1692
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3132
      Begin VB.CheckBox Check3 
         Caption         =   "Add relation labels"
         Height          =   252
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2412
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Add relations"
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   2532
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add columns"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   2532
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2160
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   1212
   End
End
Attribute VB_Name = "ExportDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()

    If Check2.Value = 1 Then
        Check3.Enabled = True
    Else
        Check3.Enabled = False
        Check3.Value = 0
    End If

End Sub

Private Sub Combo1_Click()

    On Error Resume Next
    Label2.Caption = Round(Main.Editor.Width * Combo1.Text) / 10 & " x " & Round(Main.Editor.Height * Combo1.Text) / 10 & " cm"

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "Bitmap file (*.bmp)|*.bmp|"
        .ShowSave
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        Me.Caption = "Saving picture..."
        Me.Refresh

        CreateDrawing 2, Check1.Value, Check2.Value, Check3.Value
        SavePicture Main.Drawing.Picture, .FileName
        
        Me.Caption = "Export to Drawing"
        Screen.MousePointer = vbNormal
        Unload Me
    End With
Err:
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub
