VERSION 5.00
Begin VB.Form AddTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create new table"
   ClientHeight    =   5436
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4308
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5436
   ScaleWidth      =   4308
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   4920
   End
   Begin VB.Frame Frame2 
      Caption         =   "Custom size (centimeters)"
      Height          =   972
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   3852
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   612
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   612
      End
      Begin VB.Label Label5 
         Caption         =   "Height:"
         Height          =   252
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   612
      End
      Begin VB.Label Label4 
         Caption         =   "Width:"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   492
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4440
      Width           =   2172
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create near"
      Height          =   2412
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   3852
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   1812
      End
      Begin VB.ListBox List1 
         Height          =   1392
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3372
      End
      Begin VB.Label Label2 
         Caption         =   "Position:"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   732
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   1680
      TabIndex        =   6
      Top             =   4920
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2880
      TabIndex        =   7
      Top             =   4920
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Text            =   "Untitled"
      Top             =   240
      Width           =   2532
   End
   Begin VB.Label Label3 
      Caption         =   "Use properties from:"
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Table name:"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "AddTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aColor As String

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Dim tempWidth As Integer
    Dim tempHeight As Integer
    Dim tempLeft As Integer
    Dim tempTop As Integer
    Dim tempColor As String

    tempColor = aColor

    'create with sizes of existing table
    If Combo2.Text <> "" Then
        For i = 1 To Main.TableLabel.UBound
            If Main.TableLabel(i).Caption = Combo2.Text Then
                tempWidth = Main.Table(i).Width
                tempHeight = Main.Table(i).Height
                tempColor = Main.Table(i).FillColor
            End If
        Next i
    
    'create with custom sizes
    ElseIf Text2.Text <> 0 And Text3 <> 0 Then
        tempWidth = (Text2.Text * 10)
        tempHeight = (Text3.Text * 10)
    
    'create with default sizes
    Else
        tempWidth = Len(Text1.Text) * 2
        tempHeight = 40
    End If
    
    'set left and top position from selected table
    If Combo1.Text <> "" Then
        'get selected table properties
        For i = 1 To Main.TableLabel.UBound
            If Main.TableLabel(i).Caption = List1.Text Then
                'exit for loop if selected table was found
                Exit For
            End If
        Next i

        'left
        If Combo1.Text = "Left" Then
            tempLeft = Main.Table(i).Left - tempWidth
            tempTop = Main.Table(i).Top
        'right
        ElseIf Combo1.Text = "Right" Then
            tempLeft = Main.Table(i).Left + tempWidth
            tempTop = Main.Table(i).Top
        'top
        ElseIf Combo1.Text = "Top" Then
            tempLeft = Main.Table(i).Left
            tempTop = Main.Table(i).Top - tempHeight
        'bottom
        ElseIf Combo1.Text = "Bottom" Then
            tempLeft = Main.Table(i).Left
            tempTop = Main.Table(i).Top + tempHeight
        End If

    'default left and top position
    Else
        tempLeft = GridSize
        tempTop = GridSize
    End If

    If tempWidth < 40 Then
        tempWidth = 40
    End If
    
    CreateTable Replace(Text1.Text, " ", "_"), tempWidth, tempHeight, tempLeft, tempTop, Main.Combo3.Text, tempColor

    'enable relation menu item
    If Main.Table.UBound > 1 Then
        Main.mnuProjectRelation.Enabled = True
        Main.Toolbar1.Buttons(19).Enabled = True
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()

    Combo1.AddItem "Left"
    Combo1.AddItem "Right"
    Combo1.AddItem "Bottom"
    Combo1.AddItem "Top"
    SendKeys "{END}+{HOME}"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub Timer1_Timer()
    
    aColor = Rnd() * (16000000 - 100) + 100

End Sub
