VERSION 5.00
Begin VB.Form MainPrintPreview 
   Caption         =   "Print preview (A4)"
   ClientHeight    =   6660
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8808
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   117.475
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   155.363
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   5772
      Left            =   120
      ScaleHeight     =   100.965
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   141.182
      TabIndex        =   1
      Top             =   120
      Width           =   8052
      Begin VB.VScrollBar VScroll1 
         Height          =   5172
         LargeChange     =   10
         Left            =   7680
         SmallChange     =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   252
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   252
         LargeChange     =   10
         Left            =   120
         SmallChange     =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   5400
         Width           =   7572
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5616
         Left            =   0
         ScaleHeight     =   99.06
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   70.062
         TabIndex        =   2
         Top             =   0
         Width           =   3972
         Begin VB.Line Vertical 
            BorderStyle     =   3  'Dot
            X1              =   52.917
            X2              =   52.917
            Y1              =   0
            Y2              =   57.15
         End
         Begin VB.Line Horizontal 
            BorderStyle     =   3  'Dot
            X1              =   0
            X2              =   52.917
            Y1              =   57.15
            Y2              =   57.15
         End
         Begin VB.Shape HiddenHorizontal 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   2172
            Left            =   0
            Top             =   3360
            Width           =   3012
         End
         Begin VB.Shape HiddenVertical 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   5532
            Left            =   3120
            Top             =   0
            Width           =   732
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   6120
      Width           =   972
   End
End
Attribute VB_Name = "MainPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Picture1.Left = 0
    Picture1.Top = 0
    Picture1.Width = Main.Editor.Width
    Picture1.Height = Main.Editor.Height
    
    HiddenHorizontal.Top = 297
    HiddenVertical.Left = 210
    
    HScroll1.Max = Main.Editor.Width
    VScroll1.Max = Main.Editor.Height

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Vertical.X1 = 210
    Vertical.X2 = 210
    Vertical.Y1 = 0
    Vertical.Y2 = 297

    Horizontal.X1 = 0
    Horizontal.X2 = 210
    Horizontal.Y1 = 297
    Horizontal.Y2 = 297

    Picture2.Left = 2
    Picture2.Top = 2
    Picture2.Width = Me.ScaleWidth - 4
    Picture2.Height = Me.ScaleHeight - 15

    HScroll1.Left = 0
    HScroll1.Width = Picture2.ScaleWidth - VScroll1.Width
    HScroll1.Top = Picture2.ScaleHeight - HScroll1.Height
    VScroll1.Top = 0
    VScroll1.Left = Picture2.ScaleWidth - VScroll1.Width
    VScroll1.Height = Picture2.ScaleHeight - HScroll1.Height
    
    HiddenVertical.Width = Picture1.ScaleWidth - Vertical.X1
    HiddenVertical.Height = Picture1.ScaleHeight
    HiddenHorizontal.Width = Picture1.ScaleWidth
    HiddenHorizontal.Height = Picture1.ScaleHeight - Horizontal.Y1

    Command1.Top = Me.ScaleHeight - Command1.Height - 3.5

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub HScroll1_Change()

    Picture1.Left = -HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()

    Picture1.Left = -HScroll1.Value

End Sub

Private Sub VScroll1_Change()

    Picture1.Top = -VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()
    
    Picture1.Top = -VScroll1.Value

End Sub
