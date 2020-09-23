VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExportSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export to SQL Script"
   ClientHeight    =   3804
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3972
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   3972
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   2640
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export tables"
      Height          =   2172
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3492
      Begin VB.ListBox List1 
         Height          =   1560
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   360
         Width           =   3012
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Database name:"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1332
   End
End
Attribute VB_Name = "ExportSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'define strings
Dim Auto_Inc As String
Dim Primary_Key As String
Dim SetEnd As String

'define integers
Dim iTables As Integer
Dim iFields As Integer
Dim k As Integer
Dim iNewFields As Integer

'define arrays
Dim aTable As Variant
Dim aFieldTypeExtra As Variant

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "SQL file (*.sql)|*.sql|"
        .ShowSave
        Open .FileName For Output As #1
            Print #1, "-- SQL export created from Simple Database Designer"
            Print #1, ""
            
            For i = 0 To List1.ListCount - 1
                If List1.Selected(i) = True Then
                    'write SQL script
                    For j = 0 To Main.Table.UBound
                        If Main.TableLabel(j).Caption = List1.List(i) Then
                            Print #1, "-- Table: " & Main.TableLabel(j).Caption
                            Print #1, "CREATE TABLE " & LCase(Text1.Text) & "." & LCase(List1.List(i)) & " ("

                            'reset primary_key value
                            Primary_Key = ""
                            iFields = 0
                            iNewFields = 0

                            'get all fields from table
                            For k = 1 To Main.TableColumnName.UBound
                                If Main.TableColumnName(k).Tag = Main.Table(j).Index Then
                                    iFields = iFields + 1
                                End If
                            Next k
                            
                            For k = 1 To Main.TableColumnName.UBound
                                If Main.TableColumnName(k).Tag = Main.Table(j).Index Then
                                    
                                    iNewFields = iNewFields + 1
                                    Auto_Inc = ""
                                    
                                    'save fields
                                    aFieldTypeExtra = Split(Main.TableColumnType(k).Caption, " ")
                                    'only auto_increment
                                    If aFieldTypeExtra(1) = "[A]" Then
                                        Auto_Inc = " AUTO_INCREMENT"
                                    'primary key and auto_increment
                                    ElseIf aFieldTypeExtra(1) = "[PA]" Then
                                        Auto_Inc = " AUTO_INCREMENT"
                                        If Primary_Key = "" Then
                                            Primary_Key = Main.TableColumnName(k).Caption
                                        Else
                                            Primary_Key = Primary_Key & ", " & Main.TableColumnName(k).Caption
                                        End If
                                    'only primary key
                                    ElseIf aFieldTypeExtra(1) = "[P]" Then
                                        If Primary_Key = "" Then
                                            Primary_Key = Main.TableColumnName(k).Caption
                                        Else
                                            Primary_Key = Primary_Key & ", " & Main.TableColumnName(k).Caption
                                        End If
                                    End If
                                    
                                    'check if there is a comma required at the end or not
                                    If iNewFields < iFields Then
                                        SetEnd = ","
                                    Else
                                        SetEnd = ""
                                    End If
                                    
                                    Print #1, "  " & Main.TableColumnName(k).Caption & " " & UCase(aFieldTypeExtra(0)) & " NOT NULL" & Auto_Inc & SetEnd
                                
                                End If
                            Next k
                            
                            'save primary key
                            If Primary_Key <> "" Then
                                Print #1, "  PRIMARY KEY(" & Primary_Key & ")"
                            End If
                            
                            Print #1, ")"
                            Print #1, ""
                            Print #1, ""
                        End If
                    Next j
                End If
            Next i
            
        Close #1

        Unload Me

    End With
Err:
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me

End Sub
