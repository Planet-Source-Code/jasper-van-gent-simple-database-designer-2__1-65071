VERSION 5.00
Begin VB.Form TableColumns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table columns"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   8472
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8472
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   3312
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   4212
   End
   Begin VB.Frame Frame1 
      Caption         =   "Column properties"
      Height          =   3492
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   3732
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   1080
         TabIndex        =   2
         Text            =   "0"
         Top             =   1200
         Width           =   972
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   372
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1092
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1440
         TabIndex        =   6
         Top             =   2880
         Width           =   1092
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   372
         Left            =   2640
         TabIndex        =   5
         Top             =   2880
         Width           =   732
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto increment"
         Height          =   252
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   3012
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Primary Key"
         Height          =   252
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   3012
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2292
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   2292
      End
      Begin VB.Label Label3 
         Caption         =   "Length:"
         Height          =   252
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   252
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   252
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   612
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Height          =   372
      Left            =   5760
      TabIndex        =   9
      Top             =   3840
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6960
      TabIndex        =   10
      Top             =   3840
      Width           =   1212
   End
End
Attribute VB_Name = "TableColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempColumn As String
Dim aListItem As Variant

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    'remove old columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = ObjectIndex Then
            Main.TableColumnName(i).Visible = False
            Main.TableColumnType(i).Visible = False
            Main.TableColumnName(i).Tag = 0
            Main.TableColumnType(i).Tag = 0
        End If
    Next i

    'create new columns
    For i = 0 To List1.ListCount - 1
        aListItem = Split(List1.List(i), " -> ")
        AddColumn aListItem(0), aListItem(1), ObjectIndex, i
    Next i

    Unload Me

End Sub

Private Sub Command3_Click()

    If Text1.Text <> "" Then
        List1.Text = Text1.Text
        If List1.Text <> "" Then
            MsgBox "Column already exists!", vbCritical + vbOKOnly, "Add column"
        Else
            tempColumn = Replace(Text1.Text, " ", "_") & " -> " & Combo1.Text & "(" & Text2.Text & ")"
            tempColumn = tempColumn & " ["
            If Check1.Value = 1 And Check2.Value = 1 Then
                tempColumn = tempColumn & "PA"
            End If
            If Check2.Value = 1 And Check2.Value = 0 Then
                tempColumn = tempColumn & "P"
            End If
            If Check2.Value = 0 And Check2.Value = 1 Then
                tempColumn = tempColumn & "A"
            End If
            tempColumn = tempColumn & "]"
            List1.AddItem tempColumn
        End If
        Text1.SetFocus
        SendKeys "{HOME}+{END}"
    End If

End Sub

Private Sub Command5_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete column") = vbYes Then
        List1.RemoveItem List1.ListIndex
    End If

End Sub

Private Sub Form_Load()

    'load data.type file
    Open App.Path & "\data.type" For Input As #1
        Do While Not EOF(1)
            Line Input #1, FileContent
            Combo1.AddItem FileContent
        Loop
    Close #1
    
    Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub

Private Sub List1_Click()

    If List1.Text <> "" Then
        Command5.Enabled = True
    End If

End Sub

Private Sub Text1_gotfocus()
    Command3.Default = True
End Sub
