VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5676
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5676
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading, please wait..."
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   4812
   End
   Begin VB.Image Image1 
      Height          =   1656
      Left            =   0
      Picture         =   "Splash.frx":0000
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.MousePointer = vbHourglass
    Me.Width = Image1.Width
    Me.Height = Image1.Height
    Shape1.Left = 0
    Shape1.Top = 0
    Shape1.Width = Me.Width
    Shape1.Height = Me.Height
    
    Label1.Left = (Me.Width - Label1.Width) / 2
    Label1.Top = Me.ScaleHeight - (Label1.Height * 2)

End Sub

Private Sub Timer1_Timer()

    i = i + 1
    
    If i = 1 Then
        
        If ReadIniValue(App.Path & "\Setup.ini", "Application", "Serial") = "" Then
            WriteIniValue App.Path & "\Setup.ini", "Application", "Serial", App.Major & App.Minor & "-" & App.Revision & "-" & Format(Now, "yymmddhhnnss")
        End If
        WriteIniValue App.Path & "\Setup.ini", "Application", "Version", App.Major & "." & App.Minor & "." & App.Revision
        WriteIniValue App.Path & "\Setup.ini", "Application", "LastDate", Format(Now, "mm-dd-yyyy")
        WriteIniValue App.Path & "\Setup.ini", "Application", "Root", App.Major & "." & App.Path
        
        Main.StatusBar1.Panels(3).Text = "Paper size: " & Round(Main.Editor.Width) / 10 & " x " & Round(Main.Editor.Height) / 10
        
        'application caption
        Main.Caption = App.Title & " " & App.Major
        Main.mnuHelpAbout.Caption = "&About " & App.Title & "..."
        
        'default settings
        GridSize = 10
        SizeFormat = 10
        ObjectIndex = 0
        ColumnMarge = 0.4
        SmartMoveMode = 0
        
        'load ini settings
        If ReadIniValue(App.Path & "\Setup.ini", "View", "Grid") <> "" Then
            Main.mnuViewGrid.Checked = ReadIniValue(App.Path & "\Setup.ini", "View", "Grid")
        End If
        If ReadIniValue(App.Path & "\Setup.ini", "Edit", "SnapToGrid") <> "" Then
            Main.mnuEditSnap.Checked = ReadIniValue(App.Path & "\Setup.ini", "Edit", "SnapToGrid")
        End If
        If ReadIniValue(App.Path & "\Setup.ini", "Edit", "GridSize") <> "" Then
            GridSize = ReadIniValue(App.Path & "\Setup.ini", "Edit", "GridSize")
        End If
        
        Main.Combo1.AddItem "-- None --"
        Main.Combo1.Text = "-- None --"
        
        Main.Combo2.AddItem "-- None --"
        Main.Combo2.Text = "-- None --"
    
        Main.Combo3.AddItem "Default"
        Main.Combo3.Text = "Default"

        Me.MousePointer = vbNormal
        Unload Me
        Main.Show

    End If

End Sub
