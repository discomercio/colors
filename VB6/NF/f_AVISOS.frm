VERSION 5.00
Begin VB.Form f_AVISOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AVISOS"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pnAvisos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4305
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9075
      Begin VB.TextBox c_Avisos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4200
      Picture         =   "f_AVISOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1425
   End
End
Attribute VB_Name = "f_AVISOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConfirma_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()

    c_Avisos = ""
    If sAvisosAExibir <> "" Then
        c_Avisos = sAvisosAExibir
        End If
        
End Sub
