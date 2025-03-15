VERSION 5.00
Begin VB.Form frm_menu_principal 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   3360
      TabIndex        =   2
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frm_menu_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   IsLoading True

   Dim I As Long
   For I = 1 To 9999999
    DoEvents
   Next I

   IsLoading False

End Sub
