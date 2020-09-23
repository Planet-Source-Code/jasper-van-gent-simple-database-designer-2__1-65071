VERSION 5.00
Begin VB.Form HelpAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   4944
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4812
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4944
   ScaleWidth      =   4812
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1812
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   4332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   1212
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   828
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   432
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   552
   End
   Begin VB.Image Image1 
      Height          =   1656
      Left            =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "HelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileContent As String

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.webrazor.nl/databasedesigner/", vbNormalFocus

End Sub

Private Sub Form_Load()

    Image1.Picture = Splash.Image1.Picture
    
    Label1.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "Serial: " & ReadIniValue(App.Path & "\Setup.ini", "Application", "Serial")
    Label3.Caption = "Visit homepage"
    Label3.ToolTipText = "http://www.webrazor.nl/"
    
    Me.Width = Image1.Width
    
    Command1.Left = (Me.ScaleWidth - Command1.Width) / 2
    
    Open App.Path & "\Info.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, FileContent
            If Text1.Text = "" Then
                Text1.Text = FileContent
            Else
                Text1.Text = Text1.Text & vbCrLf & FileContent
            End If
        Loop
    Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub Label3_Click()

        Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.webrazor.nl/", vbNormalFocus

End Sub
