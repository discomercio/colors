VERSION 5.00
Begin VB.Form f_CD 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELEÇÃO DE CENTRO DE DISTRIBUIÇÃO"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cb_emit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   6540
   End
   Begin VB.PictureBox b_OK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5640
      Picture         =   "f_CD.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1800
      Width           =   1035
   End
   Begin VB.PictureBox b_NAO 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   720
      Picture         =   "f_CD.frx":03DE
      ScaleHeight     =   480
      ScaleWidth      =   1035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label l_emit 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o Emitente a ser utilizado"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2505
   End
End
Attribute VB_Name = "f_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEmitPadrao As String

Private Sub b_NAO_Click()

'   se os valores do vetor "usuario" estiverem vazios, significa um cancelamento
'   no login, portanto, finalizar o programa
'   senão, apenas fechar a tela
    If Trim(usuario.emit) = "" Then
        End
    Else
        Unload Me
        End If
'   ~~~

End Sub

Private Sub b_ok_Click()

Dim s As String
Dim i As Integer

    On Error GoTo BOK_TRATA_ERRO
    
    If cb_emit = "" Then
        cb_emit.SetFocus
        Exit Sub
        End If

    Screen.MousePointer = vbHourglass
    cb_emit.BackColor = COR_CINZA
    b_OK.Visible = False
    b_NAO.Visible = False


    ' ATRIBUIÇÃO DO EMITENTE E DA UF
    usuario.emit = Mid$(cb_emit, 1, Len(cb_emit) - 5)
    usuario.emit_uf = Mid$(cb_emit, Len(cb_emit) - 2, 2)
    i = cb_emit.ListIndex
    usuario.emit_id = vEmitsUsuario(i).c2
    txtFixoEspecifico = vEmitsUsuario(i).c3

    ' SE O EMITENTE SELECIONADO FOR DIFERENTE DO PADRÃO, GRAVAR
    If Not configura_registry_usuario_emit_padrao(strEmitPadrao, usuario.emit_id) Then
        aviso "Não foi possível gravar as configurações do Emitente selecionado para futuros acessos!"
        End If

    cb_emit.BackColor = COR_VERDE_ESCURO

    Unload Me

    Screen.MousePointer = vbDefault

Exit Sub
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BOK_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
   'ENCERRA O PROGRAMA
   '~~~
    End
   '~~~
    Exit Sub
    
    
End Sub

Private Sub cb_emit_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        b_ok_Click
        End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        b_NAO_Click
        Exit Sub
        End If
        
End Sub

Private Sub Form_Activate()

Dim s As String

    On Error GoTo FA_TRATA_ERRO
    
    If painel_ativo Is Me Then Exit Sub
    
    Set painel_ativo = Me

    Screen.MousePointer = Default

    If cor_fundo_padrao <> "" Then Me.BackColor = cor_fundo_padrao
    
    cb_emit.SetFocus

Exit Sub



FA_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    MsgBox s, vbOKOnly + vbCritical, App.Title
    Exit Sub

End Sub


Private Sub Form_Load()

Dim i As Integer
Dim s As String
Dim achou As Boolean

    On Error GoTo FL_TRATA_ERRO
    
    cb_emit.Clear
    For i = LBound(vEmitsUsuario) To UBound(vEmitsUsuario)
        cb_emit.AddItem Trim$(vEmitsUsuario(i).c1)
        Next
    
 '> CENTRALIZA O PAINEL NA TELA
    ScaleMode = vbPixels
    Line (1, 1)-(ScaleWidth - 3, b_OK.top - 8), COR_BRANCO, B
    Line (2, 2)-(ScaleWidth - 2, b_OK.top - 7), COR_CINZA_ESCURO, B

    left = painel_principal.left + (painel_principal.Width - Width) \ 2
    top = painel_principal.top + (painel_principal.Height - Height) \ 2
    
'> OBTÉM O EMITENTE PADRÃO, SE HOUVER
    strEmitPadrao = ""
    If le_registry_usuario_emit_padrao(strEmitPadrao) Then
        i = LBound(vEmitsUsuario)
        achou = False
        Do While (Not achou) And (i <= UBound(vEmitsUsuario))
            If vEmitsUsuario(i).c2 = strEmitPadrao Then
                cb_emit.ListIndex = i
                achou = True
                End If
            i = i + 1
            Loop
        End If
        
    Exit Sub
    
FL_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    MsgBox s, vbOKOnly + vbCritical, App.Title
    Exit Sub
    

    
End Sub
