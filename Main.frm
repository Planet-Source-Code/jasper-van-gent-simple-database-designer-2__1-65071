VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   6744
   ClientLeft      =   48
   ClientTop       =   732
   ClientWidth     =   11256
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6744
   ScaleWidth      =   11256
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   720
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.PictureBox Properties 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   11256
      TabIndex        =   5
      Top             =   6012
      Width           =   11256
      Begin VB.ComboBox Combo3 
         Height          =   288
         Left            =   8160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   2292
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   4560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   2772
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   2772
      End
      Begin VB.Label Label3 
         Caption         =   "Layer:"
         Height          =   252
         Left            =   7560
         TabIndex        =   15
         Top             =   144
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "Relations:"
         Height          =   252
         Left            =   3720
         TabIndex        =   8
         Top             =   144
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Table:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   140
         Width           =   612
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   6504
      Width           =   11256
      _ExtentX        =   19854
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   1764
            Text            =   "Untitled.sdd"
            TextSave        =   "Untitled.sdd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1757
            MinWidth        =   1764
            Text            =   "0 x 0"
            TextSave        =   "0 x 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2815
            MinWidth        =   2822
            Text            =   "Paper size: 0 x 0"
            TextSave        =   "Paper size: 0 x 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Table size: 0 x 0"
            TextSave        =   "Table size: 0 x 0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Centimeters"
            TextSave        =   "Centimeters"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Scale: 1:10"
            TextSave        =   "Scale: 1:10"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox WorkSpace 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      Height          =   5532
      Left            =   0
      ScaleHeight     =   96.732
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   197.697
      TabIndex        =   0
      Top             =   288
      Width           =   11256
      Begin VB.VScrollBar VScroll 
         Height          =   4452
         LargeChange     =   10
         Left            =   10800
         Min             =   1
         SmallChange     =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Value           =   1
         Width           =   252
      End
      Begin VB.HScrollBar HScroll 
         Height          =   252
         LargeChange     =   10
         Left            =   240
         Min             =   1
         SmallChange     =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5160
         Value           =   1
         Width           =   10452
      End
      Begin VB.PictureBox BarEmptyBottom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   10800
         ScaleHeight     =   252
         ScaleWidth      =   252
         TabIndex        =   17
         Top             =   5040
         Width           =   252
      End
      Begin VB.PictureBox BarLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4812
         Left            =   0
         ScaleHeight     =   84.878
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   4.445
         TabIndex        =   4
         Top             =   240
         Width           =   252
         Begin VB.PictureBox LineLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   4.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   3372
            Left            =   -28
            ScaleHeight     =   59.478
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   4.445
            TabIndex        =   14
            Top             =   120
            Width           =   252
         End
      End
      Begin VB.PictureBox BarTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   240
         ScaleHeight     =   4.445
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   190.712
         TabIndex        =   3
         Top             =   0
         Width           =   10812
         Begin VB.PictureBox LineTop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   4.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   252
            Left            =   600
            ScaleHeight     =   4.445
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   70.062
            TabIndex        =   13
            Top             =   -28
            Width           =   3972
         End
      End
      Begin VB.PictureBox Editor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawStyle       =   2  'Dot
         ForeColor       =   &H00FFC0C0&
         Height          =   16838
         Left            =   360
         ScaleHeight     =   296.545
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   209.55
         TabIndex        =   11
         Top             =   360
         Width           =   11906
         Begin VB.PictureBox Drawing 
            AutoRedraw      =   -1  'True
            Height          =   2052
            Left            =   7680
            ScaleHeight     =   35.348
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   28.998
            TabIndex        =   27
            Top             =   1080
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label TableSelector 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   2292
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Label RelationRightType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   4.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Index           =   0
            Left            =   5760
            TabIndex        =   26
            Top             =   2040
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label RelationLeftType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   4.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Index           =   0
            Left            =   5640
            TabIndex        =   25
            Top             =   2040
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label RelationLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Untitled"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   228
            Index           =   0
            Left            =   5640
            TabIndex        =   24
            Top             =   1800
            Visible         =   0   'False
            Width           =   696
         End
         Begin VB.Line RelationVertical 
            Index           =   0
            Visible         =   0   'False
            X1              =   116.417
            X2              =   116.417
            Y1              =   25.4
            Y2              =   42.333
         End
         Begin VB.Line RelationRight 
            Index           =   0
            Visible         =   0   'False
            X1              =   99.483
            X2              =   101.6
            Y1              =   29.633
            Y2              =   29.633
         End
         Begin VB.Line RelationLeft 
            Index           =   0
            Visible         =   0   'False
            X1              =   99.483
            X2              =   101.6
            Y1              =   25.4
            Y2              =   25.4
         End
         Begin VB.Label TableColumnType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "INTEGER(4)"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   192
            Index           =   0
            Left            =   960
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   1308
         End
         Begin VB.Label TableColumnName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Untitled"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   192
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   552
         End
         Begin VB.Shape Selector 
            BorderStyle     =   3  'Dot
            DrawMode        =   2  'Blackness
            Height          =   732
            Left            =   360
            Top             =   2760
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Label TableLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Untitled"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Shape Table 
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   2292
            Index           =   0
            Left            =   240
            Top             =   240
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Shape TableShadow 
            FillStyle       =   0  'Solid
            Height          =   2292
            Index           =   0
            Left            =   360
            Top             =   360
            Visible         =   0   'False
            Width           =   2052
         End
      End
      Begin VB.PictureBox BarEmptyTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   0
         ScaleHeight     =   252
         ScaleWidth      =   252
         TabIndex        =   10
         Top             =   0
         Width           =   252
      End
      Begin VB.PictureBox EditorShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4656
         Left            =   360
         ScaleHeight     =   82.127
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   152.188
         TabIndex        =   12
         Top             =   360
         Width           =   8628
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   480
      Top             =   5520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1708A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1719C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":172AE
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":173C0
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":174D2
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":175E4
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":176F6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17808
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1791A
            Key             =   "Front"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17C6C
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17FBE
            Key             =   "Show"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":18310
            Key             =   "Hide"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":18662
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":18774
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":18CC6
            Key             =   "Relation"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19218
            Key             =   "Columns"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11256
      _ExtentX        =   19854
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Back"
            Object.ToolTipText     =   "Send to Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Front"
            Object.ToolTipText     =   "Bring to Front"
            ImageKey        =   "Front"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Hide"
            Object.ToolTipText     =   "Hide layer"
            ImageKey        =   "Hide"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Show"
            Object.ToolTipText     =   "Show layer"
            ImageKey        =   "Show"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Object.ToolTipText     =   "New table"
            ImageKey        =   "Table"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Columns"
            Object.ToolTipText     =   "Table columns"
            ImageKey        =   "Columns"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Relation"
            Object.ToolTipText     =   "Build new relation"
            ImageKey        =   "Relation"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &Untitled As..."
      End
      Begin VB.Menu mnuFileEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Project export"
         Begin VB.Menu mnuFileExportDrawing 
            Caption         =   "&Create image"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuFileExportEmpty1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileExportSql 
            Caption         =   "&SQL Script"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "I&mport..."
      End
      Begin VB.Menu mnuFileEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileEmpty3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFront 
         Caption         =   "&Bring to Front"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "&Send to Back"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDeselect 
         Caption         =   "&Deselect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "D&elete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSnap 
         Caption         =   "&Snap to Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditGrid 
         Caption         =   "&Change grid"
      End
      Begin VB.Menu mnuEditEmpty3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditProperties 
         Caption         =   "&Table properties"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuLayerS 
      Caption         =   "&Layers"
      Begin VB.Menu mnuLayerNew 
         Caption         =   "&New layer"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuLayerEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLayerShow 
         Caption         =   "&Show"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuLayerHide 
         Caption         =   "&Hide"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuLayerDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectTable 
         Caption         =   "&New table"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuProjectColumns 
         Caption         =   "&Table columns"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuProjectEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectRelation 
         Caption         =   "&Build new relation"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuProjectEmpty3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectProperties 
         Caption         =   "&Properties"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewGrid 
         Caption         =   "&Grid"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBugreport 
         Caption         =   "Bug &report"
      End
      Begin VB.Menu mnuHelpEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuSizeformat 
      Caption         =   "Size"
      Visible         =   0   'False
      Begin VB.Menu mnuSizeformatCentimeters 
         Caption         =   "&Centimeters"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  SIMPLE DATABASE DESIGNER 2 - ALPHA RELEASE
'
'  Developed by Jasper van Gent
'  http://www.webrazor.nl/databasedesigner/
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'define variables
Dim SmartMoveMode As Integer
Dim j As Integer
Dim tempLayerName As String
Dim FileName As String
Dim FileTitle As String
Dim PointX As Single
Dim PointY As Single
Dim NewX As Single
Dim NewY As Single
Dim aFileContent As Variant
Dim aTableColumn1 As Variant
Dim aTableColumn2 As Variant

'hide layer
Function HideLayer()

    ObjectIndex = 0
    DeselectTables
    DeselectRelation

    tempLayerName = Combo3.Text
    Combo3.RemoveItem Combo3.ListIndex
    Combo3.AddItem "--" & tempLayerName
    Combo3.Text = "--" & tempLayerName
    
    'hide table
    For j = 1 To Table.ubound
        'if table is connected with this layer
        If Table(j).Tag = tempLayerName Then
            Table(j).Tag = "--" & tempLayerName
            HideTable (j)
        End If
    Next j
    
    'hide relation
    For j = 1 To RelationLabel.ubound
        'if table is connected with this layer
        If RelationLabel(j).Tag = tempLayerName Then
            RelationLabel(j).Tag = "--" & tempLayerName
            HideRelation (j)
        End If
    Next j

End Function

'show layer
Function ShowLayer()

    ObjectIndex = 0
    DeselectTables
    DeselectRelation

    tempLayerName = Combo3.Text
    Combo3.RemoveItem Combo3.ListIndex
    Combo3.AddItem Replace(tempLayerName, "--", "")
    Combo3.Text = Replace(tempLayerName, "--", "")
    
    'show tables
    For j = 1 To Table.ubound
        If Table(j).Tag = tempLayerName Then
            Table(j).Tag = Replace(tempLayerName, "--", "")
            ShowTable (j)
        End If
    Next j
    
    'show relations
    For j = 1 To RelationLabel.ubound
        If RelationLabel(j).Tag = tempLayerName Then
            RelationLabel(j).Tag = Replace(tempLayerName, "--", "")
            ShowRelation (j)
        End If
    Next j

End Function

'save as
Function SaveAs()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "Simple Database Designer File (*.sdd)|*.sdd|"
        .ShowSave
        FileName = .FileName
        FileTitle = .FileTitle
        SaveFile (FileName)
        mnuFileSaveAs.Caption = "Save " & FileTitle & " &As..."
        mnuFileSave.Enabled = True
    End With
Err:

End Function

'open file
Function OpenFile()

    Dim tempTTableName As String
    Dim tempTLayer As String
    Dim tempTWidth As Integer
    Dim tempTHeight As Integer
    Dim tempTLeft As Integer
    Dim tempTTop As Integer
    Dim tempTColor As String
    
    Dim tempCName As String
    Dim tempCType As String
    
    Dim tempRName As String
    Dim tempRLeftTag As String
    Dim tempRRightTag As String
    Dim tempRLeftTable As String
    Dim tempRLeftColumn As String
    Dim tempRRight As String
    Dim tempRLayer As String
    Dim tempRRightTable As String
    Dim tempRRightColumn As String
    Dim tempRLeftType As Integer
    Dim tempRRightType As Integer

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "Simple Database Designer File (*.sdd)|*.sdd|"
        .ShowOpen
        FileName = .FileName
        FileTitle = .FileTitle
        
        'clear current project
        CreateEmptyProject
        
        Main.Combo3.Clear
        
        'open the selected file
        Open FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, FileContent
                
                'create table
                If Mid(FileContent, 1, 5) = "Table" Then
                    i = 0
                    aFileContent = Split(FileContent, "<|BREAK|>")
                    tempTTableName = aFileContent(1)
                    tempTWidth = aFileContent(5)
                    tempTHeight = aFileContent(6)
                    tempTLeft = aFileContent(3)
                    tempTTop = aFileContent(4)
                    tempTLayer = aFileContent(2)
                    tempTColor = aFileContent(7)
                    CreateTable tempTTableName, tempTWidth, tempTHeight, tempTLeft, tempTTop, tempTLayer, tempTColor

                'create column
                ElseIf Mid(FileContent, 1, 6) = "Column" Then
                    aFileContent = Split(FileContent, "<|BREAK|>")
                    tempCName = aFileContent(1)
                    tempCType = aFileContent(2)
                    AddColumn tempCName, tempCType, Table(Table.ubound).Index, i
                    i = i + 1
                    
                'create relation
                ElseIf Mid(FileContent, 1, 8) = "Relation" Then
                    aFileContent = Split(FileContent, "<|BREAK|>")
                    tempRName = aFileContent(1)
                    tempRLayer = aFileContent(2)
                    tempRLeftTag = aFileContent(3)
                    tempRRightTag = aFileContent(4)
                    tempRLeftType = aFileContent(5)
                    tempRRightType = aFileContent(6)
                    
                    aTableColumn1 = Split(tempRLeftTag, "|")
                    aTableColumn2 = Split(tempRRightTag, "|")
                    
                    tempRLeftTable = aTableColumn1(0)
                    tempRLeftColumn = aTableColumn1(1)
                    tempRRightTable = aTableColumn2(0)
                    tempRRightColumn = aTableColumn2(1)
                    
                    CreateRelation tempRLeftTable, tempRRightTable, tempRLeftColumn, tempRRightColumn, tempRLeftType, tempRRightType, tempRLayer, tempRName
                
                'create layers
                ElseIf Mid(FileContent, 1, 5) = "Layer" Then
                    aFileContent = Split(FileContent, "<|BREAK|>")
                    Combo3.AddItem aFileContent(1)
                End If

            Loop
        Close #1
        
        'hide tables and relations inside hidden layers
        For j = 0 To Combo3.ListCount - 1
            If Mid(Combo3.List(j), 1, 2) = "--" Then
                
                'hide tables
                For k = 1 To Table.ubound
                    If Table(k).Tag = Combo3.List(j) Then
                        HideTable (k)
                    End If
                Next k
                
                'hide relations
                For k = 1 To RelationLabel.ubound
                    If RelationLabel(k).Tag = Combo3.List(j) Then
                        HideRelation (k)
                    End If
                Next k
                
            End If
        Next j
        
        Main.Combo3.Text = "Default"
        
    End With
Err:

End Function

'save file
Function SaveFile(FileName As String)

    Me.Enabled = False
    Me.MousePointer = vbHourglass
    StatusBar1.Panels(1).Text = "Saving " & FileTitle & "..."
    
    Open FileName For Output As #1
        Print #1, "[SIMPLE DATABASE DESIGNER 2]"
        Print #1, "Created: " & Format(Now, "dd-mm-YYYY")
        Print #1, "Version: " & App.Major & "." & App.Minor & "." & App.Revision
        Print #1, ""
        Print #1, "[PROJECT]"
        Print #1, "Project<|BREAK|>" & Editor.BackColor & "<|BREAK|>" & Editor.Width & "<|BREAK|>" & Editor.Height & "<|BREAK|>" & GridSize
        Print #1, ""

        'save layers
        Print #1, "[LAYERS]"
        For i = 0 To Combo3.ListCount - 1
            Print #1, "Layer<|BREAK|>" & Combo3.List(i)
        Next i
        Print #1, ""

        'save tables
        For i = 0 To Combo1.ListCount - 1
            For j = 1 To TableLabel.ubound
                If TableLabel(j).Caption = Combo1.List(i) Then
                    If Table(j).Tag <> "" Then
                        Print #1, "[TABLE]"
                        Print #1, "Table<|BREAK|>" & TableLabel(j).Caption & "<|BREAK|>" & Table(j).Tag & "<|BREAK|>" & Table(j).Left & "<|BREAK|>" & Table(j).Top & "<|BREAK|>" & Table(j).Width & "<|BREAK|>" & Table(j).Height & "<|BREAK|>" & Table(j).FillColor
                        For k = 1 To TableColumnName.ubound
                            If TableColumnName(k).Tag = j Then
                                Print #1, "Column<|BREAK|>" & TableColumnName(k).Caption & "<|BREAK|>" & TableColumnType(k).Caption
                            End If
                        Next k
                        Print #1, ""
                    End If
                End If
            Next j
        Next i
        
        'save relations
        Print #1, "[RELATIONS]"
        For i = 1 To RelationLabel.ubound
            If RelationLabel(i).Tag <> "" Then
                Print #1, "Relation<|BREAK|>" & RelationLabel(i).Caption & "<|BREAK|>" & RelationLabel(i).Tag & "<|BREAK|>" & RelationLeft(i).Tag & "<|BREAK|>" & RelationRight(i).Tag & "<|BREAK|>" & RelationLeftType(i).Caption & "<|BREAK|>" & RelationRightType(i).Caption
            End If
        Next i
    
    Close #1

    StatusBar1.Panels(1).Text = FileTitle
    Me.MousePointer = vbDefault
    Me.Enabled = True

End Function

Private Sub BarEmptyTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuSizeformat
    End If

End Sub

Private Sub Combo1_Click()

    If ObjectIndex = 0 Then
        For i = 1 To TableLabel.ubound
            If TableLabel(i).Caption = Combo1.Text Then
    
                ObjectIndex = i
                SelectTable

                mnuEditCopy.Enabled = True
                mnuEditDelete.Enabled = True
                mnuEditFront.Enabled = True
                mnuEditBack.Enabled = True
                mnuEditDeselect.Enabled = True
                mnuProjectColumns.Enabled = True
    
            End If
        Next i
    End If

End Sub

Private Sub Combo2_Click()

    If RelationIndex = 0 Then
        For i = 1 To RelationLabel.ubound
            If RelationLabel(i).Caption = Combo2.Text Then
    
                RelationIndex = i
                SelectRelation
    
            End If
        Next i
    End If

End Sub

Private Sub Combo3_Click()

    'enable or disable delete menu button
    If Combo3.Text = "Default" Then
        mnuLayerDelete.Enabled = False
        mnuLayerShow.Enabled = False
        mnuLayerHide.Enabled = False
        Toolbar1.Buttons(14).Enabled = False
        Toolbar1.Buttons(15).Enabled = False
        mnuProjectTable.Enabled = True
    Else
        mnuLayerDelete.Enabled = True
        mnuLayerShow.Enabled = True
        mnuLayerHide.Enabled = True
        Toolbar1.Buttons(14).Enabled = True
        Toolbar1.Buttons(15).Enabled = True
        If Left(Combo3.Text, 2) = "--" Then
            mnuProjectTable.Enabled = False
        Else
            mnuProjectTable.Enabled = True
        End If
    End If
    
    'change layer of selected table
    If ObjectIndex <> 0 And Table.ubound <> 0 Then
        Table(ObjectIndex).Tag = Combo3.Text
    End If
    
    'change layer of selected relation
    If RelationIndex <> 0 Then
        RelationLabel(RelationIndex).Tag = Combo3.Text
    End If

End Sub

Private Sub Editor_Click()

    LockWindow Me.hWnd, True
    ObjectIndex = 0
    RelationIndex = 0

    Editor.Visible = False
    
    DeselectTables
    DeselectRelation

    mnuEditCopy.Enabled = False
    mnuEditDelete.Enabled = False
    mnuEditFront.Enabled = False
    mnuEditBack.Enabled = False
    mnuEditDeselect.Enabled = False
    mnuProjectColumns.Enabled = False
    Editor.Visible = True

    LockWindow Me.hWnd, False

End Sub

Private Sub Editor_DragDrop(Source As Control, X As Single, Y As Single)

    LockWindow Me.hWnd, True
    If mnuEditSnap.Checked = True Then
        TableSelector(ObjectIndex).Left = NewX
        TableSelector(ObjectIndex).Top = NewY
    Else
        TableSelector(ObjectIndex).Left = X - PointX
        TableSelector(ObjectIndex).Top = Y - PointY
    End If
    
    Me.Refresh
    LockWindow Me.hWnd, False

End Sub

Private Sub Editor_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    Dim j As Integer
    j = 0

    TableSelector(ObjectIndex).Visible = True

    If mnuEditSnap.Checked = True Then
        
        Dim TempX As Integer
        Dim TempY As Integer
        
        TempX = (X - PointX) / GridSize
        NewX = TempX * GridSize
        TempY = (Y - PointY) / GridSize
        NewY = TempY * GridSize

        Table(ObjectIndex).Left = NewX
        Table(ObjectIndex).Top = NewY
    Else
        
        Table(ObjectIndex).Left = X - PointX
        Table(ObjectIndex).Top = Y - PointY
    
    End If
    
    TableLabel(ObjectIndex).Left = Table(ObjectIndex).Left
    TableLabel(ObjectIndex).Top = Table(ObjectIndex).Top
    TableShadow(ObjectIndex).Left = Table(ObjectIndex).Left + 0.3
    TableShadow(ObjectIndex).Top = Table(ObjectIndex).Top + 0.3
    Selector.Left = Table(ObjectIndex).Left - 0.5
    Selector.Top = Table(ObjectIndex).Top - 0.5
    Selector.Width = Table(ObjectIndex).Width + 1.3
    Selector.Height = Table(ObjectIndex).Height + 1.3
    TableSelector(ObjectIndex).Left = Table(ObjectIndex).Left
    TableSelector(ObjectIndex).Top = Table(ObjectIndex).Top
    
    'move columns to new position
    For i = 1 To Main.TableColumnName.ubound
        'only if column if visible
        If Main.TableColumnName(i).Visible = True Then
            'select column related to current table
            If Main.TableColumnName(i).Tag = ObjectIndex Then
                If j = 0 Then
                    Main.TableColumnName(i).Top = Table(ObjectIndex).Top + TableLabel(ObjectIndex).Height + ColumnMarge
                Else
                    Main.TableColumnName(i).Top = Main.TableColumnName(i - 1).Top + Main.TableColumnName(i - 1).Height + ColumnMarge
                End If
                Main.TableColumnName(i).Left = Table(ObjectIndex).Left + ColumnMarge
                j = 1
            End If
            If Main.TableColumnType(i).Tag = ObjectIndex Then
                Main.TableColumnType(i).Left = Table(ObjectIndex).Left + Table(ObjectIndex).Width - Main.TableColumnType(i).Width - (ColumnMarge * 2)
                Main.TableColumnType(i).Top = Main.TableColumnName(i).Top
            End If
        End If
    Next i

    

    'move relations to new position
    For i = 1 To RelationLabel.ubound
        
        'split left and right columnname tag to get table and column
        aTableColumn1 = Split(RelationLeft(i).Tag, "|")
        aTableColumn2 = Split(RelationRight(i).Tag, "|")

        'get tempColumn1 value
        For j = 1 To Main.TableLabel.ubound
            If Main.TableLabel(j).Caption = aTableColumn1(0) Then
                For k = 1 To Main.TableColumnName.ubound
                    If Main.TableColumnName(k).Caption = aTableColumn1(1) And Main.TableColumnName(k).Tag = Main.TableLabel(j).Index Then
                        tempColumn1 = k
                    End If
                Next k
            End If
        Next j
        
        'get tempColumn2 value
        For j = 1 To Main.TableLabel.ubound
            If Main.TableLabel(j).Caption = aTableColumn2(0) Then
                For k = 1 To Main.TableColumnName.ubound
                    If Main.TableColumnName(k).Caption = aTableColumn2(1) And Main.TableColumnName(k).Tag = Main.TableLabel(j).Index Then
                        tempColumn2 = k
                    End If
                Next k
            End If
        Next j
        
        
        With RelationLeft(i)
            .X1 = TableColumnType(tempColumn1).Left + TableColumnType(tempColumn1).Width
            .X2 = TableColumnType(tempColumn1).Left + TableColumnType(tempColumn1).Width + 3
            .Y1 = TableColumnType(tempColumn1).Top + (TableColumnType(tempColumn1).Height / 2)
            .Y2 = TableColumnType(tempColumn1).Top + (TableColumnType(tempColumn1).Height / 2)
        End With
        With RelationRight(i)
            .X1 = TableColumnType(tempColumn2).Left + TableColumnType(tempColumn2).Width
            .X2 = TableColumnType(tempColumn2).Left + TableColumnType(tempColumn2).Width + 3
            .Y1 = TableColumnType(tempColumn2).Top + (TableColumnType(tempColumn2).Height / 2)
            .Y2 = TableColumnType(tempColumn2).Top + (TableColumnType(tempColumn2).Height / 2)
        End With
        With RelationVertical(i)
            .X1 = RelationLeft(i).X2
            .Y1 = RelationLeft(i).Y1
            .X2 = RelationRight(i).X2
            .Y2 = RelationRight(i).Y1
        End With
        With RelationLabel(i)
            .Left = RelationVertical(i).X1 + ((RelationVertical(i).X2 - RelationVertical(i).X1) / 2) - (.Width / 2)
            .Top = RelationVertical(i).Y1 + ((RelationVertical(i).Y2 - RelationVertical(i).Y1) / 2) - (.Height / 2)
        End With
        With RelationLeftType(i)
            .Top = RelationLeft(i).Y1 - .Height
            .Left = RelationLeft(i).X1 + 1.5
        End With
        With RelationRightType(i)
            .Top = RelationRight(i).Y1 - .Height
            .Left = RelationRight(i).X1 + 1.5
        End With
    Next i

End Sub

Private Sub Editor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    StatusBar1.Panels(2).Text = Round(Round(X, 0) / 10, 2) & " x " & Round(Round(Y, 0) / 10, 2)
    StatusBar1.Panels(4).Text = "Table size: 0 x 0"
    
    If ObjectIndex <> 0 Then
        TableSelector(ObjectIndex).Visible = True
    End If

End Sub

Private Sub Editor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        If ObjectIndex = 0 Then
            PopupMenu mnuProject
        Else
            PopupMenu mnuEdit
        End If
    End If

End Sub

Private Sub Form_Load()
    
    'set window sizes and position
    Main.Width = Screen.Width / 1.4
    Main.Height = Screen.Height / 1.4

    Main.Left = (Screen.Width - Me.ScaleWidth) / 2
    Main.Top = (Screen.Height - Me.ScaleHeight) / 2

    ResizeElements

End Sub

Private Sub Form_Resize()

    LockWindow Me.hWnd, True
    ResizeElements
    Me.Refresh
    LockWindow Me.hWnd, False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub HScroll_Change()

    Editor.Left = 7 - HScroll.Value
    ResizeElements

End Sub

Private Sub mnuEditBack_Click()

    SendTableBack

End Sub

Private Sub mnuEditCopy_Click()

    CopyIndex = ObjectIndex
    mnuEditPaste.Enabled = True
    Toolbar1.Buttons(6).Enabled = True

End Sub

Private Sub mnuEditDelete_Click()

    If ObjectIndex <> 0 Then
        DeleteTable
    ElseIf RelationIndex <> 0 Then
        DeleteRelation
    End If

End Sub

Private Sub mnuEditDeselect_Click()

    If ObjectIndex <> 0 Then
        DeselectTables
    ElseIf RelationIndex <> 0 Then
        DeselectRelation
    End If

End Sub

Private Sub mnuEditFront_Click()

    BringTableFront

End Sub

Private Sub mnuEditGrid_Click()

    Dim NewGridSize As String
    NewGridSize = InputBox("Enter new grid size (milimeters):", "Change grid", GridSize)
    If IsNumeric(NewGridSize) Then
        GridSize = NewGridSize
        WriteIniValue App.Path & "\Setup.ini", "Edit", "GridSize", NewGridSize
        ResizeElements
    End If

End Sub

Private Sub mnuEditPaste_Click()

    PasteTable

End Sub

Private Sub mnuEditProperties_Click()

    Me.Enabled = False
    MainTableProperties.Picture1.BackColor = Table(ObjectIndex).FillColor
    MainTableProperties.Text1.Text = TableLabel(ObjectIndex).Caption
    MainTableProperties.Text2.Text = Round(Table(ObjectIndex).Width) / 10
    MainTableProperties.Text3.Text = Round(Table(ObjectIndex).Height) / 10
    MainTableProperties.Show

End Sub

Private Sub mnuEditSnap_Click()

    If mnuEditSnap.Checked = True Then
        mnuEditSnap.Checked = False
    Else
        mnuEditSnap.Checked = True
    End If

    WriteIniValue App.Path & "\Setup.ini", "Edit", "SnapToGrid", mnuEditSnap.Checked

End Sub

Private Sub mnuFileExit_Click()

    End

End Sub

Private Sub mnuFileExportDrawing_Click()

    ExportDrawing.Show
    Me.Enabled = False

End Sub

Private Sub mnuFileExportSql_Click()

    Me.Enabled = False
    
    'create list
    For i = 1 To Table.ubound
        ExportSQL.List1.AddItem TableLabel(i).Caption
        ExportSQL.List1.Selected(ExportSQL.List1.ListCount - 1) = True
    Next i

    ExportSQL.Show

End Sub

Private Sub mnuFileImport_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "SQL File (*.sql)|*.sql|"
        .DialogTitle = "Import SQL file"
        .ShowOpen
        
        i = 0
        
        'open the selected file
        Open .FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, FileContent
                If Mid(LCase(FileContent), 1, 12) = "create table" Then
                    'create table
                    i = i + 1
                    FileContent = Mid(LCase(FileContent), 14, Len(FileContent) - 15)
                    CreateTable FileContent, Len(FileContent) * 2, 50, GridSize * i, GridSize * i, Main.Combo3.Text, &H80000002
                End If
            Loop
        Close #1
        
    End With
Err:
    Exit Sub

End Sub

Private Sub mnuFileNew_Click()

    Me.Enabled = False
    NewWizard.Show

End Sub

Private Sub mnuFileOpen_Click()

    OpenFile

End Sub

Private Sub mnuFilePrint_Click()

    Me.Enabled = False
    MainPrint.Show

End Sub

Private Sub mnuFilePrintPreview_Click()

    Me.Enabled = False
    CreateDrawing 1, True, True, True
    MainPrintPreview.Picture1.Picture = Drawing.Picture
    MainPrintPreview.Show

End Sub

Private Sub mnuFileSaveAs_Click()

    SaveAs

End Sub

Private Sub mnuHelpAbout_Click()

    Me.Enabled = False
    HelpAbout.Show

End Sub

Function NewLayer()

    Dim sNewLayer As String
    sNewLayer = InputBox("Enter new layer name:", "New layer", "")
    If sNewLayer <> "" Then
        Combo3.AddItem Replace(Replace(sNewLayer, " ", "_"), "--", "")
        Combo3.Text = Replace(Replace(sNewLayer, " ", "_"), "--", "")
    End If

End Function

Private Sub mnuHelpBugreport_Click()

    Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.webrazor.nl/contact/", vbNormalFocus

End Sub

Private Sub mnuLayerDelete_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete layer") = vbYes Then

        'update all related tables to Default layer
        For i = 1 To Table.ubound
            If Table(i).Tag = Combo3.Text Then
                Table(i).Tag = "Default"
            End If
        Next i

        'remove the layer
        Combo3.RemoveItem Combo3.ListIndex
        Combo3.Text = "Default"
        If Combo3.ListCount = 1 Then
            mnuLayerShow.Enabled = False
            mnuLayerHide.Enabled = False
            mnuLayerDelete.Enabled = False
        End If

    End If

End Sub

Private Sub mnuLayerHide_Click()

    HideLayer

End Sub

Private Sub mnuLayerNew_Click()

    NewLayer

End Sub

Private Sub mnuLayerShow_Click()

    ShowLayer

End Sub

Private Sub mnuProjectColumns_Click()

    Me.Enabled = False
    
    'load all added columns
    For i = 1 To TableColumnName.ubound
        If TableColumnName(i).Tag = ObjectIndex Then
            TableColumns.List1.AddItem TableColumnName(i).Caption & " -> " & TableColumnType(i).Caption
        End If
    Next i
    
    TableColumns.Show

End Sub

Private Sub mnuProjectProperties_Click()

    Me.Enabled = False
    MainProperties.Picture3.Width = MainProperties.Image1.Width
    MainProperties.Picture3.Height = MainProperties.Image1.Height
    MainProperties.Picture3.Picture = MainProperties.Image1.Picture
    MainProperties.Show

End Sub

Private Sub mnuProjectRelation_Click()

    Me.Enabled = False
    
    For i = 1 To Table.ubound
        Relations.Combo2.AddItem TableLabel(i).Caption
        Relations.Combo3.AddItem TableLabel(i).Caption
    Next i
    
    Relations.Show

End Sub

Private Sub mnuProjectTable_Click()

    Me.Enabled = False

    'create list
    For i = 1 To Table.ubound
        AddTable.List1.AddItem TableLabel(i).Caption
        AddTable.Combo2.AddItem TableLabel(i).Caption
    Next i

    If ObjectIndex <> 0 Then
        AddTable.List1.Text = Combo1.Text
    End If

    AddTable.Text1.Text = "Untitled_" & Table.Count
    AddTable.Show

End Sub

Private Sub mnuSizeformatCentimeters_Click()

    mnuSizeformatCentimeters.Checked = True
    SizeFormat = 10
    StatusBar1.Panels(5).Text = "Centimeters"
    ResizeElements

End Sub

Private Sub mnuViewGrid_Click()

    If mnuViewGrid.Checked = True Then
        mnuViewGrid.Checked = False
    Else
        mnuViewGrid.Checked = True
    End If
    
    WriteIniValue App.Path & "\Setup.ini", "View", "Grid", mnuViewGrid.Checked

    ResizeElements

End Sub

Private Sub SmartSizeBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Editor.Visible = False
    EditorShadow.Visible = False

End Sub

Private Sub SmartSizeBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SmartMoveMode = 1

End Sub

Private Sub SmartSizeLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SmartMoveMode = 2

End Sub

Private Sub RelationLabel_Click(Index As Integer)

    RelationIndex = Index
    SelectRelation

End Sub

Private Sub RelationLabel_DblClick(Index As Integer)

    Dim newRelationTitle As String
    newRelationTitle = InputBox("Enter new name for relation:", "Relation name", RelationLabel(RelationIndex).Caption)
    If newRelationTitle <> "" Then
        If newRelationTitle <> RelationLabel(RelationIndex).Caption Then
            Combo2.RemoveItem Combo2.ListIndex
            Combo2.AddItem newRelationTitle
            RelationLabel(RelationIndex).Caption = newRelationTitle
            Combo2.Text = newRelationTitle
        End If
    End If

End Sub

Private Sub TableSelector_Click(Index As Integer)
    
    LockWindow Me.hWnd, True
    Editor.Visible = False
    ObjectIndex = Index
    RelationIndex = 0
    
    DeselectRelation
    
    SelectTable

    mnuEditCopy.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditFront.Enabled = True
    mnuEditBack.Enabled = True
    mnuEditDeselect.Enabled = True
    mnuProjectColumns.Enabled = True
    Editor.Visible = True
    
    Me.Refresh
    LockWindow Me.hWnd, False

End Sub

Private Sub TableSelector_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    TableSelector(ObjectIndex).Visible = True

End Sub

Private Sub TableSelector_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    TableSelector(ObjectIndex).Visible = False

End Sub

Private Sub TableSelector_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    PointX = X / 56.7
    PointY = Y / 56.7

    StatusBar1.Panels(4).Text = "Table size: " & Round(Round(TableSelector(Index).Width) / 10, 2) & " x " & Round(Round(TableSelector(Index).Height) / 10, 2)

End Sub

Private Sub TableSelector_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If TableSelector(Index).DragMode = 1 Then
        If Button = 2 Then
            PopupMenu mnuEdit
        End If
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "Delete"
            If ObjectIndex <> 0 Then
                DeleteTable
            ElseIf RelationIndex <> 0 Then
                DeleteRelation
            End If
        
        Case "New"
            NewEmptyProject
        
        Case "Properties"
            If ObjectIndex <> 0 Then
                Me.Enabled = False
                MainTableProperties.Picture1.BackColor = Table(ObjectIndex).FillColor
                MainTableProperties.Text1.Text = TableLabel(ObjectIndex).Caption
                MainTableProperties.Text2.Text = Round(Table(ObjectIndex).Width) / 10
                MainTableProperties.Text3.Text = Round(Table(ObjectIndex).Height) / 10
                MainTableProperties.Show
            Else
                Me.Enabled = False
                MainProperties.Picture3.Width = MainProperties.Image1.Width
                MainProperties.Picture3.Height = MainProperties.Image1.Height
                MainProperties.Picture3.Picture = MainProperties.Image1.Picture
                MainProperties.Show
            End If
        
        Case "Paste"
            PasteTable

        Case "Copy"
            CopyIndex = ObjectIndex
            mnuEditPaste.Enabled = True
            Toolbar1.Buttons(6).Enabled = True

        Case "Open"
            OpenFile
        Case "Front"
            BringTableFront
        Case "Back"
            SendTableBack
        Case "Hide"
            HideLayer
        Case "Show"
            ShowLayer
        Case "Save"
            If FileName = "" Then
                SaveAs
            Else
                SaveFile (FileName)
            End If

        Case "Table"
            Me.Enabled = False
            'create list
            For i = 1 To Table.ubound
                AddTable.List1.AddItem TableLabel(i).Caption
                AddTable.Combo2.AddItem TableLabel(i).Caption
            Next i
            If ObjectIndex <> 0 Then
                AddTable.List1.Text = Combo1.Text
            End If
            AddTable.Text1.Text = "Untitled_" & Table.Count
            AddTable.Show

        Case "Columns"
            Me.Enabled = False
            'load all added columns
            For i = 1 To TableColumnName.ubound
                If TableColumnName(i).Tag = ObjectIndex Then
                    TableColumns.List1.AddItem TableColumnName(i).Caption & " -> " & TableColumnType(i).Caption
                End If
            Next i
            TableColumns.Show
        
        Case "Relation"
            Me.Enabled = False
            For i = 1 To Table.ubound
                Relations.Combo2.AddItem TableLabel(i).Caption
                Relations.Combo3.AddItem TableLabel(i).Caption
            Next i
            Relations.Show
    End Select
End Sub

Private Sub VScroll_Change()

    Editor.Top = 7 - VScroll.Value
    ResizeElements

End Sub

Private Sub WorkSpace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SmartMoveMode = 0

End Sub

Private Sub WorkSpace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuFile
    End If

End Sub
