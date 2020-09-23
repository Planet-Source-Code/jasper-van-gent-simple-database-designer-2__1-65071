VERSION 5.00
Begin VB.Form Relations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build relation"
   ClientHeight    =   4488
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   8244
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4488
   ScaleWidth      =   8244
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   5520
      TabIndex        =   11
      Top             =   3960
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6720
      TabIndex        =   10
      Top             =   3960
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      Caption         =   "Column 2"
      Height          =   3492
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   3732
      Begin VB.ComboBox Combo3 
         Height          =   288
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2652
      End
      Begin VB.ListBox List2 
         Height          =   2352
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3252
      End
      Begin VB.Label Label3 
         Caption         =   "Table:"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3960
      Width           =   2412
   End
   Begin VB.Frame Frame1 
      Caption         =   "Column 1"
      Height          =   3492
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3732
      Begin VB.ListBox List1 
         Height          =   2352
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   3252
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label Label2 
         Caption         =   "Table:"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Type of relation:"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   1212
   End
End
Attribute VB_Name = "Relations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()

    Dim tempTableId As Integer

    List1.Clear
    
    'get tableid
    For i = 1 To Main.TableLabel.UBound
        If Main.TableLabel(i).Caption = Combo2.Text Then
            tempTableId = i
        End If
    Next i
    
    'search for columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = tempTableId Then
            List1.AddItem Main.TableColumnName(i).Caption
        End If
    Next i

End Sub

Private Sub Combo3_Click()

    Dim tempTableId As Integer

    List2.Clear
    
    'get tableid
    For i = 1 To Main.TableLabel.UBound
        If Main.TableLabel(i).Caption = Combo3.Text Then
            tempTableId = i
        End If
    Next i
    
    'search for columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = tempTableId Then
            List2.AddItem Main.TableColumnName(i).Caption
        End If
    Next i

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()
    
    Dim tempTypeLeft As Integer
    Dim tempTypeRight As Integer

    If Combo1.Text = "On on one" Then
        tempTypeLeft = 0
        tempTypeRight = 0
    ElseIf Combo1.Text = "Column 1 to column 2" Then
        tempTypeLeft = 0
        tempTypeRight = 1
    ElseIf Combo1.Text = "Column 2 to column 1" Then
        tempTypeLeft = 1
        tempTypeRight = 0
    End If


    If Combo2.Text = Combo3.Text Then
        MsgBox "Please select two different tables!", vbCritical + vbOKOnly, "Create relation"
    Else
        If List1.Text = "" Or List2.Text = "" Then
            MsgBox "Please select two columns!", vbCritical + vbOKOnly, "Create relation"
        Else
            CreateRelation Combo2.Text, Combo3.Text, List1.Text, List2.Text, tempTypeLeft, tempTypeRight, Main.Combo3.Text, "Untitled_" & Main.RelationLabel.Count
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()

    Combo1.AddItem "One on one"
    Combo1.AddItem "Column 1 to column 2"
    Combo1.AddItem "Column 2 to column 1"
    
    Combo1.Text = "One on one"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub
