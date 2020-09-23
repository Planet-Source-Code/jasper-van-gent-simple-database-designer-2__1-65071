VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   5676
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   7392
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   7392
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5040
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6120
      TabIndex        =   1
      Top             =   5160
      Width           =   1092
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4812
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   8488
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "MainProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Database"
      TabPicture(1)   =   "MainProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Layout"
      TabPicture(2)   =   "MainProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Data types"
      TabPicture(3)   =   "MainProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Command3"
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(4)=   "Text3"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   -71040
         TabIndex        =   25
         Top             =   720
         Width           =   1572
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   372
         Left            =   -74520
         TabIndex        =   23
         Top             =   3960
         Width           =   972
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   372
         Left            =   -69360
         TabIndex        =   22
         Top             =   720
         Width           =   852
      End
      Begin VB.Frame Frame5 
         Caption         =   "Added types"
         Height          =   4092
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   2772
         Begin VB.ListBox List1 
            Height          =   2928
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   2292
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data handling"
         Height          =   1452
         Left            =   -74760
         TabIndex        =   13
         Top             =   1560
         Width           =   6492
         Begin VB.CheckBox Check1 
            Caption         =   "Use accolades around column names"
            Height          =   252
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   3132
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Engine"
         Height          =   852
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   6492
         Begin VB.ComboBox Combo2 
            Height          =   288
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   2772
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Canvas size (centimeter)"
         Height          =   1212
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   4212
         Begin VB.CheckBox Check2 
            Caption         =   "Lock different sizes"
            Height          =   252
            Left            =   1920
            TabIndex        =   19
            Top             =   720
            Width           =   1932
         End
         Begin VB.ComboBox Combo1 
            Height          =   288
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   1332
         End
         Begin VB.TextBox Text2 
            Height          =   288
            Left            =   960
            TabIndex        =   8
            Text            =   "6"
            Top             =   720
            Width           =   612
         End
         Begin VB.TextBox Text1 
            Height          =   288
            Left            =   960
            TabIndex        =   6
            Text            =   "6"
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label3 
            Caption         =   "Preset:"
            Height          =   252
            Left            =   1920
            TabIndex        =   9
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label2 
            Caption         =   "Height:"
            Height          =   252
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Width:"
            Height          =   252
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   612
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Canvas color"
         Height          =   2652
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4212
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            Height          =   1692
            Left            =   600
            ScaleHeight     =   1644
            ScaleWidth      =   3324
            TabIndex        =   18
            Top             =   360
            Width           =   3372
         End
         Begin VB.PictureBox Picture2 
            Height          =   1692
            Left            =   240
            ScaleHeight     =   1644
            ScaleWidth      =   204
            TabIndex        =   17
            Top             =   360
            Width           =   252
         End
         Begin VB.PictureBox Picture1 
            Height          =   252
            Left            =   240
            ScaleHeight     =   204
            ScaleWidth      =   204
            TabIndex        =   16
            Top             =   2160
            Width           =   252
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1716
            Left            =   600
            Picture         =   "MainProperties.frx":0070
            Top             =   360
            Visible         =   0   'False
            Width           =   3408
         End
         Begin VB.Label Label4 
            Caption         =   "Current color:"
            Height          =   252
            Left            =   600
            TabIndex        =   15
            Top             =   2160
            Width           =   1092
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4092
         Left            =   -74760
         TabIndex        =   26
         Top             =   480
         Width           =   6492
         _ExtentX        =   11451
         _ExtentY        =   7218
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   420
         TabCaption(0)   =   "Tables"
         TabPicture(0)   =   "MainProperties.frx":1C8CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Picture5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Picture4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Combo3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Relations"
         TabPicture(1)   =   "MainProperties.frx":1C8E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label10"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label9"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label12"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Picture6"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Combo4"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Picture7"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Picture8"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   -73200
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   40
            Top             =   1680
            Width           =   372
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00C0FFFF&
            Height          =   252
            Left            =   -73200
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   38
            Top             =   1320
            Width           =   372
         End
         Begin VB.ComboBox Combo3 
            Height          =   288
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   600
            Width           =   852
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   1800
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   30
            Top             =   960
            Width           =   372
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            Height          =   252
            Left            =   1800
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   29
            Top             =   1320
            Width           =   372
         End
         Begin VB.ComboBox Combo4 
            Height          =   288
            Left            =   -73200
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1692
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   -73200
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   27
            Top             =   960
            Width           =   372
         End
         Begin VB.Label Label12 
            Caption         =   "Label font color:"
            Height          =   252
            Left            =   -74520
            TabIndex        =   39
            Top             =   1680
            Width           =   1332
         End
         Begin VB.Label Label5 
            Caption         =   "Header size:"
            Height          =   252
            Left            =   360
            TabIndex        =   37
            Top             =   600
            Width           =   1092
         End
         Begin VB.Label Label6 
            Caption         =   "Header color:"
            Height          =   252
            Left            =   360
            TabIndex        =   36
            Top             =   960
            Width           =   1092
         End
         Begin VB.Label Label7 
            Caption         =   "Head font color:"
            Height          =   252
            Left            =   360
            TabIndex        =   35
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label Label9 
            Caption         =   "Line style:"
            Height          =   252
            Left            =   -74520
            TabIndex        =   34
            Top             =   600
            Width           =   852
         End
         Begin VB.Label Label10 
            Caption         =   "Line color:"
            Height          =   252
            Left            =   -74520
            TabIndex        =   33
            Top             =   960
            Width           =   852
         End
         Begin VB.Label Label11 
            Caption         =   "Label color:"
            Height          =   252
            Left            =   -74520
            TabIndex        =   32
            Top             =   1320
            Width           =   972
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Type:"
         Height          =   252
         Left            =   -71640
         TabIndex        =   24
         Top             =   720
         Width           =   492
      End
   End
End
Attribute VB_Name = "MainProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

    If Combo1.Text = "A4" Then
        Text1.Text = 21
        Text2.Text = 29.7
    ElseIf Combo1.Text = "A5" Then
        Text1.Text = 14.8
        Text2.Text = 21
    ElseIf Combo1.Text = "A3" Then
        Text1.Text = 29.7
        Text2.Text = 42
    ElseIf Combo1.Text = "A2" Then
        Text1.Text = 42
        Text2.Text = 59.4
    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Main.Editor.Width = (Text1.Text * 10)
    Main.Editor.Height = (Text2.Text * 10)
    Main.Editor.BackColor = Picture1.BackColor
    
    'save list with data type
    Open App.Path & "\data.type" For Output As #1
        For i = 0 To List1.ListCount - 1
            Print #1, List1.List(i)
        Next i
    Close #1
    
    'save ini settings
    WriteIniValue App.Path & "\Setup.ini", "Properties", "DatabaseEngine", Combo2.Text
    WriteIniValue App.Path & "\Setup.ini", "Properties", "UseAccolades", Check1.Value
    WriteIniValue App.Path & "\Setup.ini", "Properties", "HeaderSize", Combo3.Text
    WriteIniValue App.Path & "\Setup.ini", "Properties", "HeaderColor", Picture4.BackColor
    WriteIniValue App.Path & "\Setup.ini", "Properties", "HeaderFontColor", Picture5.BackColor
    WriteIniValue App.Path & "\Setup.ini", "Properties", "RelationLineStyle", Combo4.ListIndex
    WriteIniValue App.Path & "\Setup.ini", "Properties", "RelationLineColor", Picture6.BackColor
    WriteIniValue App.Path & "\Setup.ini", "Properties", "RelationLabelColor", Picture7.BackColor
    WriteIniValue App.Path & "\Setup.ini", "Properties", "RelationLabelFontColor", Picture8.BackColor
    
    ResizeElements
    
    Unload Me

End Sub

Private Sub Command3_Click()

    List1.Text = Text3.Text

    If List1.Text = UCase(Text3.Text) Then
        MsgBox "Type already exist!", vbCritical + vbOKOnly, "Add type"
    Else
        List1.AddItem UCase(Text3.Text)
    End If
    
    Text3.Text = ""
    Text3.SetFocus

End Sub

Private Sub Command4_Click()

    List1.RemoveItem List1.ListIndex

End Sub

Private Sub Form_Load()

    'load data.type file
    Open App.Path & "\data.type" For Input As #1
        Do While Not EOF(1)
            Line Input #1, FileContent
            List1.AddItem FileContent
        Loop
    Close #1
    
    Combo4.AddItem "Solid"
    Combo4.AddItem "Dashed"
    Combo4.AddItem "Dotted"
    Combo4.AddItem "Dash-dot"
    Combo4.AddItem "Dash-Dot-Dot"
    
    Text1.Text = Round(Main.Editor.ScaleWidth) / 10
    Text2.Text = Round(Main.Editor.ScaleHeight) / 10
    
    Picture1.BackColor = Main.Editor.BackColor
    
    Combo1.AddItem "A5"
    Combo1.AddItem "A4"
    Combo1.AddItem "A3"
    Combo1.AddItem "A2"

    Combo2.AddItem "MySQL 4.1"
    Combo2.AddItem "MySQL 5"
    Combo2.AddItem "PostgreSQL"
    
    For i = 6 To 10
        Combo3.AddItem i
    Next i

    'load ini
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "DatabaseEngine") <> "" Then
        Combo2.Text = ReadIniValue(App.Path & "\Setup.ini", "Properties", "DatabaseEngine")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "UseAccolades") <> "" Then
        Check1.Value = ReadIniValue(App.Path & "\Setup.ini", "Properties", "UseAccolades")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderSize") <> "" Then
        Combo3.Text = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderSize")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderColor") <> "" Then
        Picture4.BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderColor")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderFontColor") <> "" Then
        Picture5.BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderFontColor")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineStyle") <> "" Then
        Combo4.ListIndex = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineStyle")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineColor") <> "" Then
        Picture6.BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineColor")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelColor") <> "" Then
        Picture7.BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelColor")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelFontColor") <> "" Then
        Picture8.BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelFontColor")
    End If

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

End Sub

Private Sub Picture3_Click()

    Picture1.BackColor = Picture2.BackColor

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Picture2.BackColor = Picture3.Point(X, Y)

End Sub

Private Sub Picture4_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture4.BackColor
        .ShowColor
        Picture4.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture5_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture5.BackColor
        .ShowColor
        Picture5.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture6_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture6.BackColor
        .ShowColor
        Picture6.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture7_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture7.BackColor
        .ShowColor
        Picture7.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture8_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture8.BackColor
        .ShowColor
        Picture8.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Text1_Change()

    If Check2.Value = 1 Then
        Text2.Text = Text1.Text
    End If

End Sub

Private Sub Text2_Change()

    If Check2.Value = 1 Then
        Text1.Text = Text2.Text
    End If

End Sub

Private Sub Text3_GotFocus()

    Command3.Default = True

End Sub
