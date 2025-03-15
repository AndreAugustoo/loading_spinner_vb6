VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm_loading_spinner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1000
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   600
      Width           =   975
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1575
         Left            =   -120
         TabIndex        =   2
         Top             =   -120
         Width           =   1455
         ExtentX         =   2566
         ExtentY         =   2778
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Line borda 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   2760
      X2              =   2760
      Y1              =   120
      Y2              =   600
   End
End
Attribute VB_Name = "frm_loading_spinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   Call AjustarBorda(Me)
   
   WebBrowser1.Visible = True
   
   Picture1.BorderStyle = 0
   Picture1.Visible = True
   
   WebBrowser1.Navigate "file:///C:/Projects/VB6/LoadingSpinner/loading.gif"

End Sub

Public Sub AjustarBorda(P_Formulario As Form)
    On Error Resume Next
    
    Const Margem As Integer = 8
    
    Load P_Formulario.borda(0)
    Load P_Formulario.borda(1)
    Load P_Formulario.borda(2)
    Load P_Formulario.borda(3)
    
    With P_Formulario.borda(0)
        .X1 = 0
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = 0
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(1)
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(2)
        .X1 = P_Formulario.ScaleWidth - Margem
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(3)
        .X1 = 0
        .Y1 = P_Formulario.ScaleHeight - Margem
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
End Sub
