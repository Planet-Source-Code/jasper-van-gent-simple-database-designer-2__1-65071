VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainTableProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table properties"
   ClientHeight    =   2340
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3852
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3852
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   1200
      TabIndex        =   9
      Top             =   1800
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   1212
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   972
   End
   Begin VB.PictureBox Picture1 
      Height          =   252
      Left            =   1320
      ScaleHeight     =   204
      ScaleWidth      =   324
      TabIndex        =   3
      Top             =   600
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2292
   End
   Begin VB.Label Label4 
      Caption         =   "Height:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "Width:"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "Background:"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "MainTableProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Dim tempNewText As String

    Main.Table(ObjectIndex).FillColor = Picture1.BackColor
    ResizeTable (Text2.Text * 10), (Text3.Text * 10)
    
    tempNewText = Replace(Text1.Text, " ", "_")
    
    Main.Combo1.Text = Main.TableLabel(ObjectIndex).Caption
    Main.Combo1.RemoveItem Main.Combo1.ListIndex
    Main.Combo1.AddItem tempNewText
    Main.TableLabel(ObjectIndex).Caption = tempNewText
    Main.Combo1.Text = tempNewText
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub Picture1_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Color = Picture1.BackColor
        .Flags = 1
        .ShowColor
        Picture1.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub
