VERSION 5.00
Begin VB.Form MainPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   3588
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4872
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3588
   ScaleWidth      =   4872
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll1 
      Height          =   288
      Left            =   1404
      Max             =   10000
      Min             =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3120
      Value           =   1
      Width           =   340
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "1"
      Top             =   3120
      Width           =   372
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quality"
      Height          =   852
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   2292
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1812
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   372
      Left            =   2400
      TabIndex        =   6
      Top             =   3000
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientation"
      Height          =   1212
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1932
      Begin VB.OptionButton Option2 
         Caption         =   "Portrait"
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Landscape"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print device"
      Height          =   1092
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4452
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3852
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Copies:"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   612
   End
End
Attribute VB_Name = "MainPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Print project") = vbYes Then
    
        WriteIniValue App.Path & "\Setup.ini", "Print", "Device", Combo1.Text
        WriteIniValue App.Path & "\Setup.ini", "Print", "Quality", Combo2.Text
        WriteIniValue App.Path & "\Setup.ini", "Print", "Landscape", Option1.Value
        WriteIniValue App.Path & "\Setup.ini", "Print", "Portrait", Option2.Value

        CreateDrawing 1, True, True, True
    
        Me.Enabled = False
        Me.MousePointer = vbHourglass
        Me.Caption = "Please wait..."
        
        'set printer device
        For Each Pr In Printers
            If Pr.DeviceName = Combo1.List(Combo1.ListIndex) Then
                Set Printer = Pr
                Exit For
            End If
        Next
        
        'set printer settings
        Printer.ScaleMode = vbCentimeters
        If Combo2.Text = "Draft" Then
            Printer.PrintQuality = vbPRPQDraft
        ElseIf Combo2.Text = "Low" Then
            Printer.PrintQuality = vbPRPQLow
        ElseIf Combo2.Text = "Medium" Then
            Printer.PrintQuality = vbPRPQMedium
        ElseIf Combo2.Text = "High" Then
            Printer.PrintQuality = vbPRPQHigh
        End If
        If Option2 = True Then
            Printer.Orientation = cdlPortrait
        ElseIf Option1 = True Then
            Printer.Orientation = cdlLandscape
        End If

        Printer.Copies = Text1.Text
        Printer.FontName = Main.TableLabel(0).FontName
    
        'create document
        Printer.PaintPicture Main.Drawing.Picture, 0, 0, (Main.Drawing.Width * 0.1), (Main.Drawing.Height * 0.1)
    
        'close document and print
        Printer.EndDoc
        
        Me.Enabled = True
        Me.Caption = "Print"
        Me.MousePointer = vbDefault
        Unload Me
    
    End If

End Sub

Private Sub Form_Load()

    'get printer devices
    For Each Pr In Printers
        Combo1.AddItem Pr.DeviceName
    Next
    Combo1.Text = Printer.DeviceName

    Combo2.AddItem "Draft"
    Combo2.AddItem "Low"
    Combo2.AddItem "Medium"
    Combo2.AddItem "High"
    Combo2.Text = "Medium"

    If ReadIniValue(App.Path & "\Setup.ini", "Print", "Device") <> "" Then
        Combo1.Text = ReadIniValue(App.Path & "\Setup.ini", "Print", "Device")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Print", "Quality") <> "" Then
        Combo2.Text = ReadIniValue(App.Path & "\Setup.ini", "Print", "Quality")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Print", "Landscape") <> "" Then
        Option1.Value = ReadIniValue(App.Path & "\Setup.ini", "Print", "Landscape")
    End If
    If ReadIniValue(App.Path & "\Setup.ini", "Print", "Portrait") <> "" Then
        Option2.Value = ReadIniValue(App.Path & "\Setup.ini", "Print", "Portrait")
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub HScroll1_Change()

    Text1.Text = HScroll1.Value

End Sub
