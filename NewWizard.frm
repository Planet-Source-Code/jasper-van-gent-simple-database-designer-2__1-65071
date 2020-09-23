VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form NewWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New project wizard"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6372
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6372
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&New empty project"
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   3840
      TabIndex        =   2
      Top             =   3840
      Width           =   1092
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3372
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   5948
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   420
      TabCaption(0)   =   "Create from template"
      TabPicture(0)   =   "NewWizard.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "No templates avaliable..."
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   5652
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1092
   End
End
Attribute VB_Name = "NewWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Command3_Click()

    CreateEmptyProject
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Main.Enabled = True
    Main.Show
    Unload Me
    
End Sub
