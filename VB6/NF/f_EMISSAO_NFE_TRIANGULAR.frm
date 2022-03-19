VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form f_EMISSAO_NFE_TRIANGULAR 
   Caption         =   "Nota Fiscal"
   ClientHeight    =   13920
   ClientLeft      =   165
   ClientTop       =   -6450
   ClientWidth     =   20625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13920
   ScaleWidth      =   20625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_operacoes_pendentes 
      Caption         =   "Operações Pendentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   18960
      TabIndex        =   340
      Top             =   10800
      Width           =   1395
   End
   Begin VB.Frame pnInfoFilaPedido 
      Height          =   570
      Left            =   16320
      TabIndex        =   335
      Top             =   10800
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton b_fila_remove 
         Height          =   390
         Left            =   1920
         Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   341
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton b_fila_play 
         Height          =   390
         Left            =   120
         Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":04BD
         Style           =   1  'Graphical
         TabIndex        =   336
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblInfoFilaPedido 
         AutoSize        =   -1  'True
         Caption         =   "Retornar à Fila"
         Height          =   195
         Left            =   720
         TabIndex        =   337
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame pn_aviso_operacao_concluida 
      Height          =   1695
      Left            =   18000
      TabIndex        =   331
      Top             =   2280
      Width           =   2295
      Begin VB.Label lbl_aviso_operacao_concluida 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Pedido com Operação Triangular Concluída"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1335
         Left            =   120
         TabIndex        =   332
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton b_cancela_triangular 
      Caption         =   "Cancela Emissão Triangular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   240
      TabIndex        =   303
      Top             =   11520
      Width           =   2355
   End
   Begin VB.CommandButton b_fechar 
      Caption         =   "&Fechar"
      Height          =   450
      Left            =   16320
      TabIndex        =   302
      Top             =   13320
      Width           =   2475
   End
   Begin VB.Frame pnItens 
      Caption         =   "Itens"
      Height          =   4335
      Left            =   120
      TabIndex        =   99
      ToolTipText     =   "1605"
      Top             =   6240
      Width           =   20340
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   283
         TabStop         =   0   'False
         Top             =   3600
         Width           =   525
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   282
         TabStop         =   0   'False
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   281
         TabStop         =   0   'False
         Top             =   3600
         Width           =   4260
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   280
         TabStop         =   0   'False
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   279
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   278
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   11
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   277
         Top             =   3600
         Width           =   2235
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   3315
         Width           =   525
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   3315
         Width           =   4260
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   272
         TabStop         =   0   'False
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   271
         TabStop         =   0   'False
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   10
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   270
         Top             =   3315
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   9
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   269
         Top             =   3030
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   8
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   268
         Top             =   2745
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   7
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   267
         Top             =   2460
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   6
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   266
         Top             =   2175
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   5
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   265
         Top             =   1890
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   4
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   264
         Top             =   1605
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   3
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   263
         Top             =   1320
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   2
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   262
         Top             =   1035
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   1
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   261
         Top             =   750
         Width           =   2235
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   0
         Left            =   5850
         MaxLength       =   500
         TabIndex        =   260
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox c_vl_total_geral 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10005
         Locked          =   -1  'True
         TabIndex        =   259
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   258
         TabStop         =   0   'False
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   257
         TabStop         =   0   'False
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   256
         TabStop         =   0   'False
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   255
         TabStop         =   0   'False
         Top             =   3030
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   254
         TabStop         =   0   'False
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   3030
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   252
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   251
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   250
         TabStop         =   0   'False
         Top             =   2745
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   2745
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   247
         TabStop         =   0   'False
         Top             =   2745
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   246
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   243
         TabStop         =   0   'False
         Top             =   2460
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   242
         TabStop         =   0   'False
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   241
         TabStop         =   0   'False
         Top             =   2460
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   239
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   238
         TabStop         =   0   'False
         Top             =   2175
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   2175
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   2175
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   233
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   1890
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   1890
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   1605
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   1605
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   1035
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   1035
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   208
         TabStop         =   0   'False
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   750
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   750
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   10005
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   465
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   465
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   8085
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   465
         Width           =   4260
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   705
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   465
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   480
         Width           =   525
      End
      Begin VB.TextBox c_total_volumes 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5850
         MaxLength       =   15
         TabIndex        =   198
         Top             =   3885
         Width           =   735
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   0
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":0794
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":0796
         Style           =   2  'Dropdown List
         TabIndex        =   197
         Top             =   465
         Width           =   3195
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   196
         Top             =   465
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   195
         Top             =   750
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   194
         Top             =   1035
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   193
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   192
         Top             =   1605
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   191
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   190
         Top             =   2175
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   189
         Top             =   2460
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   188
         Top             =   2745
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   187
         Top             =   3030
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   186
         Top             =   3315
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   12615
         MaxLength       =   4
         TabIndex        =   185
         Top             =   3600
         Width           =   525
      End
      Begin VB.TextBox c_vl_total_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11310
         Locked          =   -1  'True
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   183
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   182
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   181
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   180
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   179
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   178
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   177
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   176
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   175
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   174
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   173
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   11310
         MaxLength       =   18
         TabIndex        =   172
         Top             =   3600
         Width           =   1305
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   1
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":0798
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":079A
         Style           =   2  'Dropdown List
         TabIndex        =   171
         Top             =   750
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   2
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":079C
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":079E
         Style           =   2  'Dropdown List
         TabIndex        =   170
         Top             =   1035
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   3
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07A0
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07A2
         Style           =   2  'Dropdown List
         TabIndex        =   169
         Top             =   1320
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   4
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07A4
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07A6
         Style           =   2  'Dropdown List
         TabIndex        =   168
         Top             =   1605
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   5
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07A8
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07AA
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   1890
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   6
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07AC
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07AE
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   2175
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   7
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07B0
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07B2
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   2460
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   8
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07B4
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07B6
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   2745
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   9
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07B8
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07BA
         Style           =   2  'Dropdown List
         TabIndex        =   163
         Top             =   3030
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   10
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07BC
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07BE
         Style           =   2  'Dropdown List
         TabIndex        =   162
         Top             =   3315
         Width           =   3195
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   11
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":07C0
         Left            =   13140
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":07C2
         Style           =   2  'Dropdown List
         TabIndex        =   161
         Top             =   3600
         Width           =   3195
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   160
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   159
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   158
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   157
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   156
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   155
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   154
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   153
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   152
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   151
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   150
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   16335
         MaxLength       =   8
         TabIndex        =   149
         Top             =   465
         Width           =   885
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   17220
         TabIndex        =   148
         Top             =   465
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   17220
         TabIndex        =   147
         Top             =   750
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   17220
         TabIndex        =   146
         Top             =   1035
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   17220
         TabIndex        =   145
         Top             =   1320
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   17220
         TabIndex        =   144
         Top             =   1605
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   17220
         TabIndex        =   143
         Top             =   1890
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   17220
         TabIndex        =   142
         Top             =   2175
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   17220
         TabIndex        =   141
         Top             =   2460
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   17220
         TabIndex        =   140
         Top             =   2745
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   17220
         TabIndex        =   139
         Top             =   3030
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   17220
         TabIndex        =   138
         Top             =   3315
         Width           =   585
      End
      Begin VB.ComboBox cb_ICMS_item 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   17220
         TabIndex        =   137
         Top             =   3600
         Width           =   585
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   136
         Top             =   465
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   135
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   134
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   133
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   132
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   131
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   130
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   129
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   128
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   127
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   126
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   17805
         MaxLength       =   15
         TabIndex        =   125
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   0
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   124
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   1
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   123
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   2
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   122
         Top             =   1035
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   3
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   121
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   4
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   120
         Top             =   1605
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   5
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   119
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   6
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   118
         Top             =   2175
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   7
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   117
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   8
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   116
         Top             =   2745
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   9
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   115
         Top             =   3030
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   10
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   114
         Top             =   3315
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   11
         Left            =   18750
         MaxLength       =   6
         TabIndex        =   113
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   112
         Top             =   465
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   111
         Top             =   750
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   110
         Top             =   1035
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   109
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   108
         Top             =   1605
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   107
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   106
         Top             =   2175
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   105
         Top             =   2460
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   104
         Top             =   2745
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   103
         Top             =   3030
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   102
         Top             =   3315
         Width           =   525
      End
      Begin VB.TextBox c_fcp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   101
         Top             =   3600
         Width           =   525
      End
      Begin VB.TextBox c_vl_total_icms 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   16335
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1425
      End
      Begin VB.Label l_tit_produto_obs 
         AutoSize        =   -1  'True
         Caption         =   "Informações Adicionais"
         Height          =   195
         Left            =   5865
         TabIndex        =   301
         Top             =   255
         Width           =   1635
      End
      Begin VB.Label l_tit_vl_total_geral 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9450
         TabIndex        =   300
         Top             =   3930
         Width           =   450
      End
      Begin VB.Label l_tit_vl_total 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total"
         Height          =   195
         Left            =   10530
         TabIndex        =   299
         Top             =   255
         Width           =   765
      End
      Begin VB.Label l_tit_vl_unitario 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário"
         Height          =   195
         Left            =   9045
         TabIndex        =   298
         Top             =   255
         Width           =   945
      End
      Begin VB.Label l_tit_qtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   8220
         TabIndex        =   297
         Top             =   255
         Width           =   345
      End
      Begin VB.Label l_tit_descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1605
         TabIndex        =   296
         Top             =   255
         Width           =   720
      End
      Begin VB.Label l_tit_produto 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   195
         Left            =   720
         TabIndex        =   295
         Top             =   255
         Width           =   555
      End
      Begin VB.Label l_tit_fabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabric"
         Height          =   195
         Left            =   195
         TabIndex        =   294
         Top             =   255
         Width           =   435
      End
      Begin VB.Label l_tit_total_volumes 
         AutoSize        =   -1  'True
         Caption         =   "Volumes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5025
         TabIndex        =   293
         Top             =   3930
         Width           =   720
      End
      Begin VB.Label l_tit_CFOP 
         AutoSize        =   -1  'True
         Caption         =   "CFOP"
         Height          =   195
         Left            =   13155
         TabIndex        =   292
         Top             =   255
         Width           =   420
      End
      Begin VB.Label l_tit_CST 
         AutoSize        =   -1  'True
         Caption         =   "CST"
         Height          =   195
         Left            =   12720
         TabIndex        =   291
         Top             =   255
         Width           =   315
      End
      Begin VB.Label l_tit_vl_outras_despesas_acessorias 
         AutoSize        =   -1  'True
         Caption         =   "Desp Acessórias"
         Height          =   195
         Left            =   11415
         TabIndex        =   290
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label l_tit_NCM 
         AutoSize        =   -1  'True
         Caption         =   "NCM"
         Height          =   195
         Left            =   16350
         TabIndex        =   289
         Top             =   255
         Width           =   360
      End
      Begin VB.Label l_tit_ICMS_item 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
         Height          =   195
         Left            =   17235
         TabIndex        =   288
         Top             =   255
         Width           =   390
      End
      Begin VB.Label l_tit_xPed 
         AutoSize        =   -1  'True
         Caption         =   "xPed"
         Height          =   195
         Left            =   17820
         TabIndex        =   287
         Top             =   255
         Width           =   360
      End
      Begin VB.Label l_tit_nItemPed 
         AutoSize        =   -1  'True
         Caption         =   "nItemPed"
         Height          =   195
         Left            =   18750
         TabIndex        =   286
         Top             =   240
         Width           =   675
      End
      Begin VB.Label l_tit_FCP 
         AutoSize        =   -1  'True
         Caption         =   "%FCP"
         Height          =   195
         Left            =   19560
         TabIndex        =   285
         Top             =   240
         Width           =   420
      End
      Begin VB.Label l_tit_vl_total_icms 
         AutoSize        =   -1  'True
         Caption         =   "Total ICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15240
         TabIndex        =   284
         Top             =   3930
         Width           =   960
      End
   End
   Begin VB.Frame pnParcelasEmBoletos 
      Caption         =   "Parcelas em Boletos"
      Height          =   3135
      Left            =   10560
      TabIndex        =   88
      Top             =   10680
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox c_numparc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   94
         Top             =   2040
         Width           =   945
      End
      Begin VB.CommandButton b_parc_edicao_ok 
         Height          =   390
         Left            =   240
         Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":07C4
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2565
         Width           =   690
      End
      Begin VB.CommandButton b_parc_edicao_cancela 
         Height          =   390
         Left            =   1560
         Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":0A16
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2565
         Width           =   690
      End
      Begin VB.CommandButton b_recalculaparc 
         Caption         =   "&Reagendar Parcelas Seguintes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         TabIndex        =   91
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox c_valorparc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   90
         Top             =   2040
         Width           =   1545
      End
      Begin VB.TextBox c_dataparc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   89
         Top             =   2040
         Width           =   1260
      End
      Begin MSComctlLib.ListView lvParcBoletos 
         Height          =   1335
         Left            =   120
         TabIndex        =   95
         Top             =   360
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label l_tit_valorparc 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   3720
         TabIndex        =   98
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label l_tit_dataparc 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   1560
         TabIndex        =   97
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label l_tit_numparc 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   240
         TabIndex        =   96
         Top             =   1800
         Width           =   540
      End
   End
   Begin VB.CommandButton b_imprime_remessa 
      Caption         =   "&Emitir NFe de Remessa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   16320
      TabIndex        =   87
      Top             =   12120
      Width           =   2475
   End
   Begin VB.Timer contagem 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   10920
   End
   Begin VB.Frame pnDanfe 
      Caption         =   "DANFE"
      Height          =   2265
      Left            =   18960
      TabIndex        =   82
      Top             =   11520
      Width           =   1455
      Begin VB.TextBox c_pedido_danfe 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Top             =   720
         Width           =   1170
      End
      Begin VB.CommandButton b_danfe 
         Caption         =   "D&ANFE"
         Height          =   390
         Left            =   240
         TabIndex        =   83
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label l_tit_pedido_Danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
         Height          =   195
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.TextBox c_dados_adicionais_remessa 
      Height          =   1050
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   80
      Top             =   12720
      Width           =   3420
   End
   Begin VB.TextBox c_dados_adicionais_venda 
      Height          =   1050
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   78
      Top             =   10920
      Width           =   3420
   End
   Begin VB.CommandButton b_emissao_automatica 
      Caption         =   "&Painel Emissão Automática"
      Height          =   450
      Left            =   16320
      TabIndex        =   62
      Top             =   12720
      Width           =   2475
   End
   Begin VB.Frame pnNumeroNFe 
      Caption         =   "Última NFe emitida"
      Height          =   1725
      Left            =   2760
      TabIndex        =   55
      Top             =   10680
      Width           =   4020
      Begin VB.Label l_num_NF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2130
         TabIndex        =   61
         Top             =   450
         Width           =   1710
      End
      Begin VB.Label l_tit_num_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   2145
         TabIndex        =   60
         Top             =   240
         Width           =   525
      End
      Begin VB.Label l_serie_NF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   180
         TabIndex        =   59
         Top             =   450
         Width           =   1710
      End
      Begin VB.Label l_tit_serie_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   195
         TabIndex        =   58
         Top             =   240
         Width           =   585
      End
      Begin VB.Label l_emitente_NF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   180
         TabIndex        =   57
         Top             =   1275
         Width           =   3660
      End
      Begin VB.Label l_tit_emitente_NF 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   195
         TabIndex        =   56
         Top             =   1065
         Width           =   615
      End
   End
   Begin VB.Frame pn_recebedor 
      Caption         =   "Recebedor"
      Height          =   2655
      Left            =   120
      TabIndex        =   31
      Top             =   3480
      Width           =   20340
      Begin VB.ComboBox cb_icms_remessa 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10680
         TabIndex        =   338
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cb_loc_dest_remessa 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   333
         Top             =   600
         Width           =   2460
      End
      Begin VB.CheckBox chk_InfoComprador 
         Caption         =   "Utilizar Identificação do Comprador"
         Height          =   600
         Left            =   120
         TabIndex        =   328
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame pn_endereco_edicao 
         Caption         =   "Edição do Endereço"
         Height          =   1335
         Left            =   120
         TabIndex        =   309
         Top             =   1200
         Width           =   20100
         Begin VB.CommandButton b_end_edicao_limpa 
            Height          =   390
            Left            =   15600
            Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":0E89
            Style           =   1  'Graphical
            TabIndex        =   320
            Top             =   855
            Width           =   810
         End
         Begin VB.CommandButton b_end_edicao_ok 
            Height          =   390
            Left            =   16470
            Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":12DC
            Style           =   1  'Graphical
            TabIndex        =   319
            Top             =   855
            Width           =   810
         End
         Begin VB.CommandButton b_end_edicao_cancela 
            Height          =   390
            Left            =   17340
            Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":152E
            Style           =   1  'Graphical
            TabIndex        =   318
            Top             =   855
            Width           =   810
         End
         Begin VB.TextBox c_end_edicao_uf 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   14760
            MaxLength       =   2
            TabIndex        =   317
            Text            =   "SP"
            Top             =   900
            Width           =   585
         End
         Begin VB.TextBox c_end_edicao_cidade 
            Height          =   285
            Left            =   8115
            MaxLength       =   60
            TabIndex        =   316
            Text            =   "São Paulo"
            Top             =   900
            Width           =   5730
         End
         Begin VB.TextBox c_end_edicao_bairro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   315
            Text            =   "Rua do João da Silva"
            Top             =   900
            Width           =   5730
         End
         Begin VB.TextBox c_end_edicao_complemento 
            Height          =   285
            Left            =   14760
            MaxLength       =   60
            TabIndex        =   314
            Text            =   "Apartamento 99"
            Top             =   300
            Width           =   3390
         End
         Begin VB.TextBox c_end_edicao_numero 
            Height          =   285
            Left            =   11520
            MaxLength       =   60
            TabIndex        =   313
            Text            =   "999"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox c_end_edicao_logradouro 
            Height          =   285
            Left            =   4800
            MaxLength       =   60
            TabIndex        =   312
            Text            =   "Rua do João da Silva"
            Top             =   300
            Width           =   5730
         End
         Begin VB.CommandButton b_cep_pesquisar 
            Height          =   390
            Left            =   2520
            Picture         =   "f_EMISSAO_NFE_TRIANGULAR.frx":19A1
            Style           =   1  'Graphical
            TabIndex        =   311
            Top             =   255
            Width           =   810
         End
         Begin VB.TextBox c_end_edicao_cep 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            TabIndex        =   310
            Text            =   "00000-000"
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label l_tit_end_edicao_uf 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   14415
            TabIndex        =   327
            Top             =   945
            Width           =   255
         End
         Begin VB.Label l_tit_end_edicao_cidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7425
            TabIndex        =   326
            Top             =   945
            Width           =   600
         End
         Begin VB.Label l_tit_end_edicao_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   540
            TabIndex        =   325
            Top             =   945
            Width           =   510
         End
         Begin VB.Label l_tit_end_edicao_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   13530
            TabIndex        =   324
            Top             =   345
            Width           =   1140
         End
         Begin VB.Label l_tit_end_edicao_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11205
            TabIndex        =   323
            Top             =   345
            Width           =   225
         End
         Begin VB.Label l_tit_end_edicao_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3885
            TabIndex        =   322
            Top             =   345
            Width           =   825
         End
         Begin VB.Label l_tit_end_edicao_cep 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   675
            TabIndex        =   321
            Top             =   345
            Width           =   375
         End
      End
      Begin VB.CommandButton b_editar_endereco 
         Caption         =   "Editar E&ndereço"
         Height          =   450
         Left            =   18120
         TabIndex        =   54
         Top             =   1800
         Width           =   1755
      End
      Begin VB.TextBox c_rg_dest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   8955
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   600
         Width           =   1530
      End
      Begin VB.Frame pn_endereco_recebedor 
         Caption         =   "Endereço do Recebedor"
         Height          =   1020
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   19740
         Begin VB.Label l_tit_end_recebedor_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   285
            Width           =   885
         End
         Begin VB.Label l_end_recebedor_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Rua do João da Silva"
            Height          =   195
            Left            =   1155
            TabIndex        =   52
            Top             =   285
            Width           =   1530
         End
         Begin VB.Label l_tit_end_recebedor_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9795
            TabIndex        =   51
            Top             =   285
            Width           =   285
         End
         Begin VB.Label l_end_recebedor_numero 
            AutoSize        =   -1  'True
            Caption         =   "999"
            Height          =   195
            Left            =   10170
            TabIndex        =   50
            Top             =   285
            Width           =   270
         End
         Begin VB.Label l_tit_end_recebedor_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12930
            TabIndex        =   49
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label l_end_recebedor_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Apartamento 99"
            Height          =   195
            Left            =   14220
            TabIndex        =   48
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label l_tit_end_recebedor_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   495
            TabIndex        =   47
            Top             =   585
            Width           =   570
         End
         Begin VB.Label l_end_recebedor_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Vila dos Testadores"
            Height          =   195
            Left            =   1155
            TabIndex        =   46
            Top             =   585
            Width           =   1395
         End
         Begin VB.Label l_tit_end_recebedor_cidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7875
            TabIndex        =   45
            Top             =   585
            Width           =   660
         End
         Begin VB.Label l_end_recebedor_cidade 
            AutoSize        =   -1  'True
            Caption         =   "São Paulo"
            Height          =   195
            Left            =   8625
            TabIndex        =   44
            Top             =   585
            Width           =   735
         End
         Begin VB.Label l_tit_end_recebedor_uf 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   13815
            TabIndex        =   43
            Top             =   585
            Width           =   315
         End
         Begin VB.Label l_end_recebedor_uf 
            AutoSize        =   -1  'True
            Caption         =   "SP"
            Height          =   195
            Left            =   14220
            TabIndex        =   42
            Top             =   585
            Width           =   210
         End
         Begin VB.Label l_tit_end_recebedor_cep 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15525
            TabIndex        =   41
            Top             =   585
            Width           =   435
         End
         Begin VB.Label l_end_recebedor_cep 
            AutoSize        =   -1  'True
            Caption         =   "00000-000"
            Height          =   195
            Left            =   16050
            TabIndex        =   40
            Top             =   585
            Width           =   765
         End
      End
      Begin VB.TextBox c_cnpj_cpf_dest 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1470
         MaxLength       =   18
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   600
         Width           =   2010
      End
      Begin VB.ComboBox cb_natureza_recebedor 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":1BF3
         Left            =   14280
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":1BF5
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   600
         Width           =   5970
      End
      Begin VB.TextBox c_nome_dest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   3555
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   600
         Width           =   5250
      End
      Begin VB.Label l_tit_aliquota_icms_remessa 
         AutoSize        =   -1  'True
         Caption         =   "ICMS Remessa"
         Height          =   195
         Left            =   10575
         TabIndex        =   339
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label l_tit_loc_dest_remessa 
         AutoSize        =   -1  'True
         Caption         =   "Local de Destino (Remessa)"
         Height          =   195
         Left            =   11880
         TabIndex        =   334
         Top             =   390
         Width           =   1995
      End
      Begin VB.Label l_tit_rg_dest 
         AutoSize        =   -1  'True
         Caption         =   "RG /IE"
         Height          =   195
         Left            =   9360
         TabIndex        =   306
         Top             =   390
         Width           =   510
      End
      Begin VB.Label l_tit_cnpj_cpf_dest 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF Recebedor"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         Top             =   390
         Width           =   1620
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Operação (Remessa)"
         Height          =   195
         Left            =   15240
         TabIndex        =   37
         Top             =   390
         Width           =   2415
      End
      Begin VB.Label l_tit_nome_dest 
         AutoSize        =   -1  'True
         Caption         =   "Nome Recebedor"
         Height          =   195
         Left            =   5280
         TabIndex        =   36
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame pn_comprador 
      Caption         =   "Comprador"
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   20340
      Begin VB.Frame pn_dados_comprador 
         Caption         =   "Dados do Comprador"
         Height          =   900
         Left            =   120
         TabIndex        =   63
         Top             =   2280
         Width           =   19740
         Begin VB.Label l_tit_comprador_nome 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   305
            Top             =   285
            Width           =   555
         End
         Begin VB.Label l_comprador_nome 
            AutoSize        =   -1  'True
            Caption         =   "João da Silva"
            Height          =   195
            Left            =   840
            TabIndex        =   304
            Top             =   285
            Width           =   960
         End
         Begin VB.Label l_end_comprador_cep 
            AutoSize        =   -1  'True
            Caption         =   "00000-000"
            Height          =   195
            Left            =   16050
            TabIndex        =   77
            Top             =   585
            Width           =   765
         End
         Begin VB.Label l_tit_end_comprador_cep 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15525
            TabIndex        =   76
            Top             =   585
            Width           =   435
         End
         Begin VB.Label l_end_comprador_uf 
            AutoSize        =   -1  'True
            Caption         =   "SP"
            Height          =   195
            Left            =   14220
            TabIndex        =   75
            Top             =   585
            Width           =   210
         End
         Begin VB.Label l_tit_end_comprador_uf 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   13815
            TabIndex        =   74
            Top             =   585
            Width           =   315
         End
         Begin VB.Label l_end_comprador_cidade 
            AutoSize        =   -1  'True
            Caption         =   "São Paulo"
            Height          =   195
            Left            =   8640
            TabIndex        =   73
            Top             =   585
            Width           =   735
         End
         Begin VB.Label l_tit_end_comprador_cidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7875
            TabIndex        =   72
            Top             =   600
            Width           =   660
         End
         Begin VB.Label l_end_comprador_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Vila dos Testadores"
            Height          =   195
            Left            =   900
            TabIndex        =   71
            Top             =   585
            Width           =   1395
         End
         Begin VB.Label l_tit_end_comprador_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   585
            Width           =   570
         End
         Begin VB.Label l_end_comprador_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Apartamento 99"
            Height          =   195
            Left            =   15105
            TabIndex        =   69
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label l_tit_end_comprador_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   13815
            TabIndex        =   68
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label l_end_comprador_numero 
            AutoSize        =   -1  'True
            Caption         =   "999"
            Height          =   195
            Left            =   12170
            TabIndex        =   67
            Top             =   285
            Width           =   270
         End
         Begin VB.Label l_tit_end_comprador_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11795
            TabIndex        =   66
            Top             =   285
            Width           =   285
         End
         Begin VB.Label l_end_comprador_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Rua do João da Silva"
            Height          =   195
            Left            =   6155
            TabIndex        =   65
            Top             =   285
            Width           =   1530
         End
         Begin VB.Label l_tit_end_comprador_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5180
            TabIndex        =   64
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.TextBox c_chave_nfe_ref 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10095
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "f_EMISSAO_NFE_TRIANGULAR.frx":1BF7
         Top             =   1890
         Width           =   9720
      End
      Begin VB.ComboBox cb_finalidade 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":1C24
         Left            =   120
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":1C26
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1890
         Width           =   8301
      End
      Begin VB.Frame pnZerarAliquotas 
         Height          =   1365
         Left            =   13560
         TabIndex        =   15
         Top             =   360
         Width           =   6375
         Begin VB.ComboBox cb_zerar_COFINS 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1170
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   900
            Width           =   5055
         End
         Begin VB.ComboBox cb_zerar_PIS 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1170
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   270
            Width           =   5055
         End
         Begin VB.Label l_tit_zerar_COFINS 
            AutoSize        =   -1  'True
            Caption         =   "Zerar COFINS"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   975
            Width           =   1005
         End
         Begin VB.Label l_tit_zerar_PIS 
            AutoSize        =   -1  'True
            Caption         =   "Zerar PIS"
            Height          =   195
            Left            =   450
            TabIndex        =   18
            Top             =   345
            Width           =   675
         End
      End
      Begin VB.ComboBox cb_tipo_NF 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   630
         Width           =   2340
      End
      Begin VB.ComboBox cb_frete 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10095
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   2460
      End
      Begin VB.TextBox c_ipi 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   12
         Top             =   630
         Width           =   1020
      End
      Begin VB.ComboBox cb_icms 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   11
         Top             =   630
         Width           =   975
      End
      Begin VB.ComboBox cb_natureza 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "f_EMISSAO_NFE_TRIANGULAR.frx":1C28
         Left            =   120
         List            =   "f_EMISSAO_NFE_TRIANGULAR.frx":1C2A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1260
         Width           =   8301
      End
      Begin VB.TextBox c_pedido 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   9
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   630
         Width           =   1650
      End
      Begin VB.ComboBox cb_loc_dest 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7215
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   630
         Width           =   2460
      End
      Begin VB.Label l_tit_IE 
         AutoSize        =   -1  'True
         Caption         =   "IE"
         Height          =   195
         Left            =   12960
         TabIndex        =   330
         Top             =   420
         Width           =   150
      End
      Begin VB.Label l_IE 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NC"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   12720
         TabIndex        =   329
         Top             =   630
         Width           =   585
      End
      Begin VB.Label l_tit_emitente_uf 
         AutoSize        =   -1  'True
         Caption         =   "UF do Emitente"
         Height          =   195
         Left            =   8760
         TabIndex        =   308
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label l_emitente_uf 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   8865
         TabIndex        =   307
         Top             =   1635
         Width           =   825
      End
      Begin VB.Label l_tit_chave_nfe_ref 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso NFe Referenciada"
         Height          =   195
         Left            =   10110
         TabIndex        =   30
         Top             =   1680
         Width           =   2610
      End
      Begin VB.Label l_tit_finalidade 
         AutoSize        =   -1  'True
         Caption         =   "Finalidade"
         Height          =   195
         Left            =   135
         TabIndex        =   29
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label l_tit_loc_dest 
         AutoSize        =   -1  'True
         Caption         =   "Local de Destino (Venda)"
         Height          =   195
         Left            =   7230
         TabIndex        =   28
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label l_tit_tipo_NF 
         AutoSize        =   -1  'True
         Caption         =   "Tipo do Documento Fiscal"
         Height          =   195
         Left            =   2070
         TabIndex        =   27
         Top             =   420
         Width           =   1860
      End
      Begin VB.Label l_tit_frete 
         AutoSize        =   -1  'True
         Caption         =   "Frete por Conta"
         Height          =   195
         Left            =   10110
         TabIndex        =   26
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label l_tit_aliquota_IPI 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota IPI"
         Height          =   195
         Left            =   5910
         TabIndex        =   25
         Top             =   420
         Width           =   840
      End
      Begin VB.Label l_tit_aliquota_icms 
         AutoSize        =   -1  'True
         Caption         =   "ICMS Venda"
         Height          =   195
         Left            =   4575
         TabIndex        =   24
         Top             =   420
         Width           =   900
      End
      Begin VB.Label l_tit_natureza 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Operação (Venda)"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   1050
         Width           =   2220
      End
      Begin VB.Label l_tit_pedido 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.CommandButton b_dummy 
      Appearance      =   0  'Flat
      Caption         =   "b_dummy"
      Height          =   345
      Left            =   5445
      TabIndex        =   3
      Top             =   0
      Width           =   1350
   End
   Begin VB.CommandButton b_imprime_venda 
      Caption         =   "&Emitir NFe de Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   16320
      TabIndex        =   2
      Top             =   11520
      Width           =   2475
   End
   Begin VB.Timer relogio 
      Interval        =   1000
      Left            =   1200
      Top             =   12360
   End
   Begin VB.Frame pnPedidoInfo 
      Caption         =   "Informações do Pedido"
      Height          =   1185
      Left            =   2760
      TabIndex        =   0
      Top             =   12600
      Width           =   3930
      Begin VB.TextBox c_info_pedido 
         Height          =   780
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   3660
      End
   End
   Begin VB.Label infocontagem 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tempo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   240
      TabIndex        =   86
      Top             =   10680
      Visible         =   0   'False
      Width           =   2340
      WordWrap        =   -1  'True
   End
   Begin VB.Label l_tit_dados_adicionais_remessa 
      AutoSize        =   -1  'True
      Caption         =   "Dados Adicionais (Nota de Remessa)"
      Height          =   195
      Left            =   6960
      TabIndex        =   81
      Top             =   12480
      Width           =   2640
   End
   Begin VB.Label l_tit_dados_adicionais_venda 
      AutoSize        =   -1  'True
      Caption         =   "Dados Adicionais (Nota de Venda)"
      Height          =   195
      Left            =   6960
      TabIndex        =   79
      Top             =   10680
      Width           =   2445
   End
   Begin VB.Label info 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   6
      Top             =   12360
      Width           =   2340
      WordWrap        =   -1  'True
   End
   Begin VB.Label hoje 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.00.0000"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   270
      TabIndex        =   5
      Top             =   13065
      Width           =   2340
   End
   Begin VB.Label agora 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   270
      TabIndex        =   4
      Top             =   13440
      Width           =   2340
   End
   Begin VB.Menu mnu_ARQUIVO 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_emissao_automatica 
         Caption         =   "&Modo de Emissão Automática"
      End
      Begin VB.Menu mnu_FECHAR 
         Caption         =   "&Fechar"
      End
   End
End
Attribute VB_Name = "f_EMISSAO_NFE_TRIANGULAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modulo_inicializacao_ok As Boolean

Private Const FONTNAME_IMPRESSAO = "Tahoma"
Private Const FONTSIZE_IMPRESSAO = 8
Private Const FONTBOLD_IMPRESSAO = True
Private Const FONTITALIC_IMPRESSAO = False
Private Const FORMATO_PERCENTUAL = "##0.00"

Dim usar_endereco_editado As Boolean
Dim endereco_editado__cep As String
Dim endereco_editado__logradouro As String
Dim endereco_editado__numero As String
Dim endereco_editado__complemento As String
Dim endereco_editado__bairro As String
Dim endereco_editado__cidade As String
Dim endereco_editado__uf As String
Dim endereco_comprador__nome As String
Dim endereco_comprador__cnpj_cpf As String
Dim endereco_comprador__rg As String
Dim endereco_comprador__cep As String
Dim endereco_comprador__logradouro As String
Dim endereco_comprador__numero As String
Dim endereco_comprador__complemento As String
Dim endereco_comprador__bairro As String
Dim endereco_comprador__cidade As String
Dim endereco_comprador__uf As String
Dim endereco_recebedor__nome As String
Dim endereco_recebedor__cnpj_cpf As String
Dim endereco_recebedor__rg As String
Dim endereco_recebedor__cep As String
Dim endereco_recebedor__logradouro As String
Dim endereco_recebedor__numero As String
Dim endereco_recebedor__complemento As String
Dim endereco_recebedor__bairro As String
Dim endereco_recebedor__cidade As String
Dim endereco_recebedor__uf As String

Dim pedido_anterior As String
Dim MaxSegundosNotaTriangular As Long
Dim dt_hr_inicio_emissao As Date
Dim blnEmissaoOK As Boolean
Dim pedido_ultima_emissao As String
Dim blnEsperaNFTriangular As Boolean

Dim lngIdNFeTriangular As Long
Dim lngSerieNFeTriangular As Long
Dim lngNumVendaNFeTriangular As Long
Dim lngNumRemessaNFeTriangular As Long

Dim inumparcela As Integer
Dim v_pedido_manual_boleto() As String
Dim v_parcela_manual_boleto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim blnExisteParcelamentoBoleto As Boolean

Private Sub tab_stop_configura()
Dim i As Integer

    b_dummy.TabIndex = 0
    b_danfe.TabIndex = 0
    c_pedido_danfe.TabIndex = 0
    b_operacoes_pendentes.TabIndex = 0
    b_fechar.TabIndex = 0
    b_emissao_automatica.TabIndex = 0
    b_imprime_remessa.TabIndex = 0
    b_imprime_venda.TabIndex = 0
    b_parc_edicao_cancela.TabIndex = 0
    b_parc_edicao_ok.TabIndex = 0
    c_dataparc.TabIndex = 0
    lvParcBoletos.TabIndex = 0
    c_dados_adicionais_remessa.TabIndex = 0
    c_dados_adicionais_venda.TabIndex = 0
    c_info_pedido.TabIndex = 0
    b_cancela_triangular.TabIndex = 0
    c_vl_total_icms.TabIndex = 0
    c_vl_total_outras_despesas_acessorias.TabIndex = 0
    c_vl_total_geral.TabIndex = 0
    c_total_volumes.TabIndex = 0
    
    For i = c_produto.UBound To c_produto.LBound Step -1
        c_fcp(i).TabIndex = 0
        c_nItemPed(i).TabIndex = 0
        c_xPed(i).TabIndex = 0
        c_nItemPed(i).TabIndex = 0
        cb_ICMS_item(i).TabIndex = 0
        cb_CFOP(i).TabIndex = 0
        c_CST(i).TabIndex = 0
        c_vl_outras_despesas_acessorias(i).TabIndex = 0
        c_vl_total(i).TabIndex = 0
        c_vl_unitario(i).TabIndex = 0
        c_qtde(i).TabIndex = 0
        c_produto_obs(i).TabIndex = 0
        c_descricao(i).TabIndex = 0
        c_produto(i).TabIndex = 0
        c_fabricante(i).TabIndex = 0
        Next
        
    b_end_edicao_cancela.TabIndex = 0
    b_end_edicao_ok.TabIndex = 0
    b_end_edicao_limpa.TabIndex = 0
    c_end_edicao_uf.TabIndex = 0
    c_end_edicao_cidade.TabIndex = 0
    c_end_edicao_bairro.TabIndex = 0
    c_end_edicao_complemento.TabIndex = 0
    c_end_edicao_numero.TabIndex = 0
    c_end_edicao_logradouro.TabIndex = 0
    b_cep_pesquisar.TabIndex = 0
    c_end_edicao_cep.TabIndex = 0
    b_editar_endereco.TabIndex = 0
    cb_natureza_recebedor.TabIndex = 0
    cb_loc_dest_remessa.TabIndex = 0
    cb_icms_remessa.TabIndex = 0
    c_rg_dest.TabIndex = 0
    c_nome_dest.TabIndex = 0
    c_cnpj_cpf_dest.TabIndex = 0
    cb_zerar_COFINS.TabIndex = 0
    cb_zerar_PIS.TabIndex = 0
    c_chave_nfe_ref.TabIndex = 0
    cb_finalidade.TabIndex = 0
    cb_natureza.TabIndex = 0
    cb_frete.TabIndex = 0
    cb_loc_dest.TabIndex = 0
    c_ipi.TabIndex = 0
    cb_icms.TabIndex = 0
    cb_tipo_NF.TabIndex = 0
    c_pedido.TabIndex = 0



End Sub

Private Sub CriaListaParcelasEmBoletos()
   Dim clmX As ColumnHeader

    lvParcBoletos.ListItems.Clear
    
    'criar a coluna oculta e as três colunas visíveis
    Set clmX = lvParcBoletos.ColumnHeaders.Add()
    clmX.Text = "oculto"
    Set clmX = lvParcBoletos.ColumnHeaders.Add()
    clmX.Text = "Parcela"
    clmX.Alignment = lvwColumnRight
    Set clmX = lvParcBoletos.ColumnHeaders.Add()
    clmX.Text = "Forma"
    clmX.Alignment = lvwColumnLeft
    Set clmX = lvParcBoletos.ColumnHeaders.Add()
    clmX.Text = "Dt Vencto"
    clmX.Alignment = lvwColumnCenter
    Set clmX = lvParcBoletos.ColumnHeaders.Add()
    clmX.Text = "Valor"
    clmX.Alignment = lvwColumnRight
    
    'diminuir a largura da primeira coluna
    lvParcBoletos.ColumnHeaders(1).Width = 0
    lvParcBoletos.ColumnHeaders(2).Width = lvParcBoletos.ColumnHeaders(2).Width * 0.5

End Sub

Private Sub AdicionaListaParcelasEmBoletos(lista_parc() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO)
    Dim itmX As ListItem
    Dim i As Integer
    Dim existeBoleto As Boolean
    
    lvParcBoletos.ListItems.Clear
    c_numparc.Text = ""
    c_dataparc.Text = ""
    c_valorparc.Text = ""
    b_parc_edicao_ok.Enabled = False

    'se não houver parcelamento, sair
    If (UBound(lista_parc) = 0) And (lista_parc(0).intNumDestaParcela = 0) Then Exit Sub
    
    'verificar se existe parcela em boleto; se não existir, sair
    existeBoleto = False
    i = LBound(lista_parc)
    Do While Not existeBoleto And (i <= UBound(lista_parc))
        If lista_parc(i).id_forma_pagto = ID_FORMA_PAGTO_BOLETO Then
            existeBoleto = True
            End If
        i = i + 1
        Loop
    If Not existeBoleto Then Exit Sub
    
    pnParcelasEmBoletos.Visible = True
    blnExisteParcelamentoBoleto = True
    
    For i = LBound(lista_parc) To UBound(lista_parc)
        Set itmX = lvParcBoletos.ListItems.Add()
        itmX.SubItems(1) = lista_parc(i).intNumDestaParcela
        itmX.SubItems(2) = descricao_opcao_forma_pagamento(lista_parc(i).id_forma_pagto)
        itmX.SubItems(3) = lista_parc(i).dtVencto
        itmX.SubItems(4) = formata_moeda(lista_parc(i).vlValor)
        Next i
End Sub

Private Sub ObtemParcelaSelecionada(ByRef parcnum As Integer, ByRef parcdata As String, ByRef parcvalor As String)
    
    parcnum = lvParcBoletos.SelectedItem.SubItems(1)
    parcdata = lvParcBoletos.SelectedItem.SubItems(3)
    parcvalor = lvParcBoletos.SelectedItem.SubItems(4)

End Sub

Private Sub AtualizaParcelaSelecionada(ByRef parcnum As Integer, _
                                    ByRef parcdata As String, _
                                    ByRef parcvalor As String, _
                                    ByRef lista_parc() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO)
    Dim i As Integer
    For i = LBound(lista_parc) To UBound(lista_parc)
        If lista_parc(i).intNumDestaParcela = parcnum Then
            lvParcBoletos.ListItems.Item(parcnum).SubItems(3) = parcdata
            lvParcBoletos.ListItems.Item(parcnum).SubItems(4) = parcvalor
            lista_parc(i).dtVencto = CDate(parcdata)
            lista_parc(i).vlValor = converte_para_currency(parcvalor)
            Exit For
            End If
        Next

End Sub


Private Sub b_cancela_triangular_Click()

    trata_botao_cancela_triangular
    
End Sub

Private Sub b_cep_pesquisar_Click()

    trata_botao_pesquisa_cep
    
End Sub

Private Sub b_danfe_Click()

Const NomeDestaRotina = "b_danfe_Click()"
Dim s As String

    On Error GoTo B_DANFE_CLICK_TRATA_ERRO
    
    If Trim$(c_pedido_danfe) = "" Then
        aviso_erro "Informe o nº do pedido do qual deseja consultar a DANFE!!"
        c_pedido_danfe.SetFocus
        Exit Sub
        End If
    
    DANFE_CONSULTA_parametro_emitente Trim$(c_pedido_danfe)
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
B_DANFE_CLICK_TRATA_ERRO:
'========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
 
End Sub

Private Sub b_editar_endereco_Click()

    pn_endereco_edicao.Visible = True
    
End Sub

Private Sub b_emissao_automatica_Click()
    
    fechar_modo_emissao_nfe_triangular

End Sub

Private Sub b_end_edicao_cancela_Click()
    
    trata_botao_endereco_edicao_cancela

End Sub

Private Sub b_end_edicao_limpa_Click()

    trata_botao_endereco_edicao_limpa
    
End Sub

Private Sub b_end_edicao_ok_Click()

    trata_botao_endereco_edicao_ok
    
End Sub

Private Sub b_fechar_Click()
Dim s As String

        If (lngIdNFeTriangular > 0) And contagem.Enabled Then
            If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
                aviso_erro s
                End If
            End If

   '~~~
    End
   '~~~
End Sub

Private Sub b_fila_play_Click()

    sPedidoTriangular = ""
    fechar_modo_emissao_nfe_triangular
    
End Sub

Private Sub b_fila_remove_Click()

    trata_botao_fila_remove
    
End Sub

Private Sub b_imprime_remessa_Click()

    NFe_emite_remessa
    
End Sub

Private Sub b_imprime_venda_Click()
Dim lngId As Long
Dim strUsuario As String
Dim msg_erro As String

    If NFeExisteNotaTriangularEmEmissao(lngId, strUsuario, msg_erro) And _
        (lngId <> lngIdNFeTriangular) Then
        aviso "Nota triangular sendo emitida pelo usuário " & strUsuario
        Exit Sub
    Else
        If msg_erro <> "" Then aviso msg_erro
        End If
        
    NFe_emite_venda
    
End Sub

Private Sub b_operacoes_pendentes_Click()

    sAvisosAExibir = RetornaOperacoesTriangularesPendentes
    If sAvisosAExibir <> "" Then sAvisosAExibir = sAvisosAExibir & vbCrLf & vbCrLf
    sAvisosAExibir = sAvisosAExibir & RetornaNumeracaoRemessaPendente
    If sAvisosAExibir <> "" Then
        f_AVISOS.Show vbModal, Me
    Else
        aviso "Não foram encontradas operações triangulares pendentes!!"
        End If
    
End Sub

Private Sub b_parc_edicao_cancela_Click()

    c_numparc.Text = ""
    c_dataparc.Text = ""
    c_valorparc.Text = ""
    
    b_parc_edicao_ok.Enabled = False

End Sub

Private Sub b_parc_edicao_ok_Click()

    If Trim(c_dataparc) = "" Then
        aviso "Data da parcela não pode estar em branco!!!"
        c_dataparc.SetFocus
        End If
        
    If CDate(c_dataparc) < Date Then
        aviso "Data não pode ser anterior ao dia atual!!!"
        c_dataparc.SetFocus
        End If
        
    If CDate(c_dataparc) < Date + 5 Then
        aviso "Data não pode ser inferior a um período de 05 dias!!!"
        c_dataparc.SetFocus
        End If
        
    If Trim(c_valorparc) = "" Then
        aviso "Valor da parcela não pode estar em branco!!!"
        c_valorparc.SetFocus
        End If
    
    AtualizaParcelaSelecionada CInt(c_numparc), c_dataparc, c_valorparc, v_parcela_manual_boleto()
        
    'se a primeira parcela foi alterada, habilita o botão para recálculo das demais parcelas
    If CInt(c_numparc) = 1 Then b_recalculaparc.Enabled = True

End Sub

Private Sub b_recalculaparc_Click()
    Dim i As Integer
    Dim dtUltimoPagtoCalculado As Date
    Dim posicao_tela As Integer
    
    If Not confirma("Confirma o reagendamento das parcelas seguintes?") Then Exit Sub
    
    dtUltimoPagtoCalculado = v_parcela_manual_boleto(LBound(v_parcela_manual_boleto)).dtVencto
    
    For i = LBound(v_parcela_manual_boleto) + 1 To UBound(v_parcela_manual_boleto)
        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
        v_parcela_manual_boleto(i).dtVencto = dtUltimoPagtoCalculado
        posicao_tela = v_parcela_manual_boleto(i).intNumDestaParcela
        lvParcBoletos.ListItems.Item(posicao_tela).SubItems(3) = dtUltimoPagtoCalculado
        Next

End Sub

Private Sub c_chave_nfe_ref_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If

End Sub

Private Sub c_cnpj_cpf_dest_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_nome_dest.SetFocus
        Exit Sub
        End If

End Sub

Private Sub c_cnpj_cpf_dest_LostFocus()

Const NomeDestaRotina = "c_cnpj_cpf_dest_LostFocus()"
Dim s As String
Dim s_aux As String
Dim t As ADODB.Recordset
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strRamal As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim blnSubstituirEndEntrega As Boolean


    On Error GoTo C_CNPJ_CPF_DEST_LF_TRATA_ERRO
    
    If Trim$(c_cnpj_cpf_dest) = "" Then Exit Sub
    
    c_nome_dest = ""
    c_rg_dest = ""
    
    If Not cnpj_cpf_ok(c_cnpj_cpf_dest) Then
        aviso_erro "CNPJ/CPF inválido!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
        
    c_cnpj_cpf_dest = cnpj_cpf_formata(c_cnpj_cpf_dest)
        
    
'   PESQUISANDO BANCO DE DADOS
    aguarde INFO_EXECUTANDO, "pesquisando CNPJ/CPF no banco de dados"
    
  ' T_CLIENTE
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT" & _
            " *" & _
        " FROM t_CLIENTE" & _
        " WHERE" & _
            " (cnpj_cpf = '" & retorna_so_digitos(c_cnpj_cpf_dest) & "')"
    t.Open s, dbc, , , adCmdText
    If t.EOF Then
        aguarde INFO_NORMAL, m_id
        aviso_erro "CNPJ/CPF não cadastrado!! Cadastre ou preencha manualmente o nome do recebedor"
        'c_cnpj_cpf_dest.SetFocus
        c_dados_adicionais_venda = obtem_dados_adicionais_venda(lngIdNFeTriangular)
        c_dados_adicionais_venda.ToolTipText = c_dados_adicionais_venda
        'parei aqui, rever lhgx
        GoSub C_CNPJ_CPF_DEST_LF_FECHA_TABELAS
        Exit Sub
        End If
    
    c_nome_dest = Trim$("" & t("nome"))
    If Len(retorna_so_digitos(c_cnpj_cpf_dest)) = 14 Then
        c_rg_dest = Trim$("" & t("ie"))
    Else
        c_rg_dest = Trim$("" & t("rg"))
        End If
    
    blnSubstituirEndEntrega = False
    If (chk_InfoComprador.Value = 0) And _
        (l_end_recebedor_logradouro <> "") And _
        ((l_end_recebedor_logradouro <> UCase$(Trim$("" & t("endereco")))) Or _
        (l_end_recebedor_numero <> UCase$(Trim$("" & t("endereco_numero")))) Or _
        (l_end_recebedor_cep <> cep_formata(retorna_so_digitos(Trim$("" & t("cep")))))) Then
        s_aux = UCase$(Trim$("" & t("endereco"))) & ", " & UCase$(Trim$("" & t("endereco_numero"))) & ", " & _
                UCase$(Trim$("" & t("cidade"))) & ", " & UCase$(Trim$("" & t("uf"))) & ", " & _
                cep_formata(retorna_so_digitos(Trim$("" & t("cep"))))
        s = "Deseja utilizar o endereço do cadastro do cliente para a remessa?" & vbCrLf & vbCrLf & "(" & s_aux & ")"
        If confirma(s) Then
            s = "Forneça a senha para utilizar o endereço abaixo:" & vbCrLf & vbCrLf & s_aux
            f_CONFIRMACAO_VIA_SENHA.strMensagemInformativa = s
            f_CONFIRMACAO_VIA_SENHA.strSenhaCorreta = usuario.senha
            f_CONFIRMACAO_VIA_SENHA.Show vbModal, Me
            If f_CONFIRMACAO_VIA_SENHA.blnResultadoFormOk Then blnSubstituirEndEntrega = True
            End If
        End If
    
    If (l_end_recebedor_logradouro = "") Or blnSubstituirEndEntrega Then

        limpa_dados_endereco_cadastro
        'limpa_dados_endereco_editado
        limpa_campos_endereco_edicao
        'c_dados_adicionais_venda = ""
                
        endereco_recebedor__logradouro = UCase$(Trim$("" & t("endereco")))
        endereco_recebedor__numero = UCase$(Trim$("" & t("endereco_numero")))
        endereco_recebedor__complemento = UCase$(Trim$("" & t("endereco_complemento")))
        endereco_recebedor__bairro = UCase$(Trim$("" & t("bairro")))
        endereco_recebedor__cep = cep_formata(retorna_so_digitos(Trim$("" & t("cep"))))
        endereco_recebedor__cidade = UCase$(Trim$("" & t("cidade")))
        endereco_recebedor__uf = UCase$(Trim$("" & t("uf")))
    
        l_end_recebedor_cep = endereco_recebedor__cep
        l_end_recebedor_logradouro = endereco_recebedor__logradouro
        l_end_recebedor_numero = endereco_recebedor__numero
        l_end_recebedor_complemento = endereco_recebedor__complemento
        l_end_recebedor_bairro = endereco_recebedor__bairro
        l_end_recebedor_cidade = endereco_recebedor__cidade
        l_end_recebedor_uf = endereco_recebedor__uf
    
        
        'preencher os campos de telefone
        strTelCel = telefone_formata_2(retorna_so_digitos(Trim$("" & t("tel_cel"))))
        strTelRes = telefone_formata_2(retorna_so_digitos(Trim$("" & t("tel_res"))))
        strTelCom = telefone_formata_2(retorna_so_digitos(Trim$("" & t("tel_com"))))
        strTelCom2 = telefone_formata_2(retorna_so_digitos(Trim$("" & t("tel_com_2"))))
        If strTelCel <> "" Then
            strDDD = retorna_so_digitos(Trim$("" & t("ddd_cel")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            If (Len(strDDD) = 2) Then strTelCel = "(" & strDDD & ")" & strTelCel
            End If
        If strTelRes <> "" Then
            strDDD = retorna_so_digitos(Trim$("" & t("ddd_res")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            If (Len(strDDD) = 2) Then strTelRes = "(" & strDDD & ")" & strTelRes
            End If
        If strTelCom <> "" Then
            strDDD = retorna_so_digitos(Trim$("" & t("ddd_com")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            strRamal = retorna_so_digitos(Trim$("" & t("ramal_com")))
            If (Len(strDDD) = 2) Then strTelCom = "(" & strDDD & ") " & strTelCom
            If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom = strTelCom & " R: " & strRamal
            End If
        If strTelCom2 <> "" Then
            strDDD = retorna_so_digitos(Trim$("" & t("ddd_com_2")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            strRamal = retorna_so_digitos(Trim$("" & t("ramal_com_2")))
            If (Len(strDDD) = 2) Then strTelCom2 = "(" & strDDD & ") " & strTelCom2
            If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom2 = strTelCom2 & " R: " & strRamal
            End If
    
        s = ""
        If UCase$(Trim$("" & t("tipo"))) = ID_PF Then
            strSufixoRes = "Tel Res: "
            strSufixoCom = "Tel Com: "
        Else
            strSufixoRes = "Tel: "
            strSufixoCom = "Tel: "
            End If
        If (strTelCel <> "") And (strTelRes <> "") Then s = strSufixoRes & strTelRes
        If ((strTelCel <> "") Or (strTelRes <> "")) And (strTelCom <> "") Then
            If s <> "" Then s = s & " / "
            s = s & strSufixoCom & strTelCom
            End If
        If ((strTelCel <> "") Or (strTelRes <> "") Or (strTelCom <> "")) And (strTelCom2 <> "") Then
            If s <> "" Then s = s & " / "
            s = s & strSufixoCom & strTelCom2
            End If
    
        End If
        
    c_dados_adicionais_venda = obtem_dados_adicionais_venda(lngIdNFeTriangular)
    c_dados_adicionais_venda.ToolTipText = c_dados_adicionais_venda
        
    GoSub C_CNPJ_CPF_DEST_LF_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_CNPJ_CPF_DEST_LF_FECHA_TABELAS:
'================================
    bd_desaloca_recordset t, True
    Return
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_CNPJ_CPF_DEST_LF_TRATA_ERRO:
'=============================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub C_CNPJ_CPF_DEST_LF_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Private Sub c_dados_adicionais_venda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_dados_adicionais_venda.SetFocus
        Exit Sub
        End If

End Sub

Private Sub c_dataparc_LostFocus()

    c_dataparc = Trim$(c_dataparc)
    If c_dataparc = "" Then Exit Sub
    
    data_ok c_dataparc

End Sub


Private Sub c_ipi_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_loc_dest.SetFocus
        Exit Sub
        End If

End Sub

Private Sub c_nome_dest_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(retorna_so_digitos(c_cnpj_cpf_dest.Text)) = 14 Then
            cb_natureza_recebedor.SetFocus
        Else
            c_rg_dest.SetFocus
            End If
        Exit Sub
        End If

End Sub

Private Sub c_nome_dest_LostFocus()

    c_dados_adicionais_venda = obtem_dados_adicionais_venda(lngIdNFeTriangular)
    c_dados_adicionais_venda.ToolTipText = c_dados_adicionais_venda
    
End Sub

Private Sub c_pedido_KeyPress(KeyAscii As Integer)
Dim executa_tab As Boolean
Dim s As String
Dim c As String

    If KeyAscii = 13 Then
    '   COMO O CAMPO ACEITA MÚLTIPLAS LINHAS, SÓ VAI P/ O PRÓXIMO CAMPO APÓS 2 "ENTER's" CONSECUTIVOS
        executa_tab = True
    '   CURSOR ESTÁ NO FINAL DO TEXTO (IGNORA "ENTER's" SUBSEQUENTES NO TEXTO) ?
        s = Mid$(c_pedido.Text, c_pedido.SelStart + 1)
        s = Replace$(s, vbCr, "")
        s = Replace$(s, vbLf, "")
        s = Trim$(s)
        If s <> "" Then executa_tab = False
    '   CARACTER ANTERIOR É "ENTER" ?
        If c_pedido.SelStart > 0 Then
            c = Mid$(c_pedido.Text, c_pedido.SelStart, 1)
            If (c <> Chr$(13)) And (c <> Chr$(10)) Then executa_tab = False
            End If
        
        If Not c_pedido.MultiLine Then
            c_pedido = normaliza_num_pedido(c_pedido)
            If Len(c_pedido) > 0 Then c_pedido.SelStart = Len(c_pedido)
            executa_tab = True
            End If
        
        If executa_tab Then
            KeyAscii = 0
            c_produto_obs(c_produto_obs.LBound).SetFocus
            End If
        
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_pedido(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
    KeyAscii = maiuscula(KeyAscii)
    
End Sub

Private Sub c_pedido_LostFocus()
Dim s As String

    c_pedido = Trim$(c_pedido)
    
    s = normaliza_num_pedido(c_pedido)
    If s <> "" Then
        c_pedido = s
        If Not pedido_eh_do_emitente_atual(c_pedido) Then Exit Sub
        End If
    
    pedido_preenche_dados_tela c_pedido
      
End Sub

Private Sub c_rg_dest_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_icms_remessa.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_finalidade_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_chave_nfe_ref.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_frete_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_natureza.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_icms_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_ipi.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_icms_remessa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_natureza_recebedor.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_loc_dest_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_frete.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_natureza_Click()
    ' Se o código de natureza da operação inicia com 1 ou 5, trata-se de uma operação interna;
    ' se o código de natureza da operação inicia com 2 ou 6, trata-se de uma operação interestadual
    Dim digito As String
    Dim s_cfop As String
    
    digito = left(Trim(cb_natureza.Text), 1)
    If (digito = "1") Or (digito = "5") Then cb_loc_dest.ListIndex = 0
    If (digito = "2") Or (digito = "6") Then cb_loc_dest.ListIndex = 1
    
    s_cfop = left(Trim(cb_natureza.Text), 5)
    If s_cfop = ("5.915") Or s_cfop = ("6.152") Or s_cfop = ("5.949") Or _
       s_cfop = ("6.117") Or s_cfop = ("6.923") Or s_cfop = ("6.910") Then
       cb_zerar_COFINS.ListIndex = 4
       cb_zerar_PIS.ListIndex = 4
    Else
       cb_zerar_COFINS.ListIndex = 0
       cb_zerar_PIS.ListIndex = 0
       End If

End Sub

Private Sub cb_natureza_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_finalidade.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_natureza_recebedor_Click()
    ' Se o código de natureza da operação inicia com 1 ou 5, trata-se de uma operação interna;
    ' se o código de natureza da operação inicia com 2 ou 6, trata-se de uma operação interestadual
    Dim digito As String
    
    digito = left(Trim(cb_natureza_recebedor.Text), 1)
    If (digito = "1") Or (digito = "5") Then cb_loc_dest_remessa.ListIndex = 0
    If (digito = "2") Or (digito = "6") Then cb_loc_dest_remessa.ListIndex = 1

End Sub

Private Sub cb_natureza_recebedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_dados_adicionais_venda.SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_tipo_NF_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_icms.SetFocus
        Exit Sub
        End If
    
End Sub

Private Sub contagem_Timer()

    gerencia_temmpo_limite_emissao
    
End Sub

Private Sub Form_Activate()

    If Not modulo_inicializacao_ok Then
        
      ' OK !!
        modulo_inicializacao_ok = True
        
        tab_stop_configura
        
        relogio_Timer
        
        aguarde INFO_EXECUTANDO, "iniciando aplicativo"
                   
    '   PREPARA CAMPOS/CARREGA DADOS INICIAIS
        formulario_inicia
        
    '   LIMPA CAMPOS/POSICIONA DEFAULTS
        formulario_limpa
        
    '   COR DE FUNDO
        If cor_fundo_padrao <> "" Then
            Me.BackColor = cor_fundo_padrao
            End If
    
    '   DADOS DA ÚLTIMA NFe EMITIDA
        l_serie_NF = f_MAIN.l_serie_NF
        l_num_NF = f_MAIN.l_num_NF
        l_emitente_NF = f_MAIN.l_emitente_NF
        
        Caption = Caption & " v" & m_id_versao
        
        If DESENVOLVIMENTO Then
            Caption = Caption & "  (Versão Exclusiva de Desenvolvimento/Homologação)"
            End If
        
        aguarde INFO_NORMAL, m_id
        End If
    
'   EXIBIR UF DO EMITENTE SELECIONADO NO LABEL EM DESTAQUE
    l_emitente_uf.Caption = usuario.emit_uf

    If sPedidoTriangular <> "" Then
        pnInfoFilaPedido.Visible = True
        c_pedido = sPedidoTriangular
        pedido_preenche_dados_tela c_pedido
        End If
        
    If sPedidoDANFETelaAnterior <> "" Then
        c_pedido_danfe = sPedidoDANFETelaAnterior
        sPedidoDANFETelaAnterior = ""
        End If
    If sNFAnteriorSerie <> "" Then
        l_serie_NF = sNFAnteriorSerie
        sNFAnteriorSerie = ""
    End If
    If sNFAnteriorNumero <> "" Then
        l_num_NF = sNFAnteriorNumero
        sNFAnteriorNumero = ""
        End If
    If sNFAnteriorEmitente <> "" Then
        l_emitente_NF = sNFAnteriorEmitente
        sNFAnteriorEmitente = ""
        End If

    
End Sub

Private Sub Form_Load()

    Set painel_ativo = Me
    Set painel_principal = Me

    b_dummy.top = -500

    modulo_inicializacao_ok = False
    
    ScaleMode = vbPixels

    CriaListaParcelasEmBoletos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim s As String

    If (lngIdNFeTriangular > 0) And contagem.Enabled Then
        s = "Existe nota fiscal triangular sendo emitida. Esta ação cancelará a emissão desta nota." & _
            vbCrLf & _
            "Continua assim mesmo?"
        If Not confirma(s) Then Cancel = 1 'Exit Sub
        If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
            aviso_erro s
            End If
        End If
        
    If c_pedido_danfe <> "" Then
        sPedidoDANFETelaAnterior = c_pedido_danfe
        End If
    
End Sub

Private Sub lvParcBoletos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim parcnum As Integer
    Dim parcdata As String
    Dim parcvalor As String
    
    ObtemParcelaSelecionada parcnum, parcdata, parcvalor
    c_numparc.Text = Str(parcnum)
    c_dataparc.Text = parcdata
    c_valorparc.Text = parcvalor
    b_parc_edicao_ok.Enabled = True
    b_recalculaparc.Enabled = False

End Sub

Private Sub mnu_emissao_automatica_Click()

    fechar_modo_emissao_nfe_triangular
    
End Sub

Private Sub mnu_FECHAR_Click()
    
    fechar_programa
    
End Sub


Sub fechar_modo_emissao_nfe_triangular()
Dim s As String

    If sPedidoTriangular <> "" Then
        s = "Deseja encaminhar o pedido " & sPedidoTriangular & " para a emissão padrão (sem operação triangular)?"
        If Not confirma(s) Then sPedidoTriangular = ""
    Else
        If ha_dados_preenchidos Then
            s = "Os dados preenchidos serão perdidos se o painel for alternado para o modo de emissão automática!!" & _
                vbCrLf & _
                "Continua assim mesmo?"
            If Not confirma(s) Then Exit Sub
            End If
        If (lngIdNFeTriangular > 0) And contagem.Enabled Then
            s = "Existe nota fiscal triangular sendo emitida. Esta ação cancelará a emissão desta nota." & _
                vbCrLf & _
                "Continua assim mesmo?"
            If Not confirma(s) Then Exit Sub
            If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
                aviso_erro s
                End If
            lngIdNFeTriangular = 0
            End If
        End If
    
    Unload f_EMISSAO_NFE_TRIANGULAR

End Sub

Sub fechar_programa()
Dim s As String

    If (lngIdNFeTriangular > 0) And contagem.Enabled Then
        s = "Existe nota fiscal triangular sendo emitida. Esta ação cancelará a emissão desta nota." & _
            vbCrLf & _
            "Continua assim mesmo?"
        If Not confirma(s) Then Exit Sub
        If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
            aviso_erro s
            End If
        End If

'   FECHA BANCO DE DADOS
    BD_Fecha
    BD_CEP_Fecha
    BD_Assist_Fecha
    
'   ENCERRA PROGRAMA
    End

End Sub

Function ha_dados_preenchidos() As Boolean
Const NomeDestaRotina = "ha_dados_preenchidos()"
Dim i As Integer
Dim s As String

    ha_dados_preenchidos = True
    
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_fabricante(i)) <> "" Then Exit Function
        If Trim$(c_produto(i)) <> "" Then Exit Function
        If Trim$(c_qtde(i)) <> "" Then Exit Function
        If converte_para_currency(c_vl_unitario(i)) <> 0 Then Exit Function
        Next
    
    If Trim$(c_dados_adicionais_venda) <> "" Then Exit Function
    
'   TODO
    
    ha_dados_preenchidos = False
    
End Function


Private Sub relogio_Timer()

Dim s As String

    s = left$(Time$, 5)
    If Val(right$(Time$, 1)) Mod 2 Then Mid$(s, 3, 1) = " "
    agora = s

    hoje = Format$(Date, "dd/mm/yyyy")

End Sub


Sub formulario_inicia()

' CONSTANTES
Const NomeDestaRotina = "formulario_inicia()"

Dim s As String
Dim s_aux As String
Dim msg_erro As String
Dim v_CFOP() As TIPO_LISTA_CFOP
Dim i As Integer
Dim j As Integer
Dim i_qtde As Integer
Dim vAliquotas() As String

    On Error GoTo FI_TRATA_ERRO
    
'   FINALIDADE DE EMISSÃO
'   ~~~~~~~~~~~~~~~~~~~~~
    cb_finalidade.Clear
    cb_finalidade.AddItem "1 - NFe Normal"
    cb_finalidade.AddItem "2 - NFe Complementar"
    cb_finalidade.AddItem "3 - NFe de Ajuste"
    cb_finalidade.AddItem "4 - Devolução de Mercadoria"

'   CHAVE DE ACESSO NFE REFERENCIADA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    c_chave_nfe_ref = ""
                
'   TIPO DO DOCUMENTO FISCAL
'   ~~~~~~~~~~~~~~~~~~~~~~~~
    cb_tipo_NF.Clear
    cb_tipo_NF.AddItem "0 - ENTRADA"
    cb_tipo_NF.AddItem "1 - SAÍDA"
    
'   LOCAL DE DESTINO DA VENDA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cb_loc_dest.Clear
    cb_loc_dest.AddItem "1 - INTERNA"
    cb_loc_dest.AddItem "2 - INTERESTADUAL"
    cb_loc_dest.AddItem "3 - EXTERIOR"
    
'   LOCAL DE DESTINO DA REMESSA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cb_loc_dest_remessa.Clear
    cb_loc_dest_remessa.AddItem "1 - INTERNA"
    cb_loc_dest_remessa.AddItem "2 - INTERESTADUAL"
    cb_loc_dest_remessa.AddItem "3 - EXTERIOR"

'   NATUREZA DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~
    cb_natureza.Clear
    cb_natureza_recebedor.Clear
    For j = cb_CFOP.LBound To cb_CFOP.UBound
        cb_CFOP(j).Clear
        cb_CFOP(j).AddItem ""
        Next
    
    ReDim v_CFOP(0)
    If Not le_arquivo_CFOP(v_CFOP(), msg_erro) Then
        s = "Falha ao ler arquivo com a relação de C.F.O.P. !!" & _
            vbCrLf & "Não é possível continuar !!"
        If msg_erro <> "" Then s = s & vbCrLf & vbCrLf & msg_erro
        aviso_erro s
       '~~~
        End
       '~~~
        End If
       
    i_qtde = 0
    For i = LBound(v_CFOP) To UBound(v_CFOP)
        With v_CFOP(i)
            If .codigo <> "" Then
                i_qtde = i_qtde + 1
                End If
            End With
        Next
    
    If i_qtde = 0 Then
        s = "Não foi fornecida a relação de C.F.O.P. !!" & _
            vbCrLf & "Não é possível continuar !!"
        aviso_erro s
       '~~~
        End
       '~~~
        End If
    
    For i = LBound(v_CFOP) To UBound(v_CFOP)
        With v_CFOP(i)
            If .descricao <> "" Then
                s = .codigo & String$(1, " ") & iniciais_em_maiusculas(.descricao)
                cb_natureza.AddItem s
                cb_natureza_recebedor.AddItem s
                For j = cb_CFOP.LBound To cb_CFOP.UBound
                    cb_CFOP(j).AddItem s
                    Next
                End If
            End With
        Next
       
'   ALÍQUOTAS ICMS
'   ~~~~~~~~~~~~~
    s_aux = retorna_lista_aliquotas_ICMS
    If s_aux <> "" Then
        cb_icms.Clear
        cb_icms_remessa.Clear
        vAliquotas = Split(s_aux, vbCrLf)
        For i = LBound(vAliquotas) To UBound(vAliquotas)
            cb_icms.AddItem vAliquotas(i)
            cb_icms_remessa.AddItem vAliquotas(i)
            Next
    Else
        cb_icms.Clear
        cb_icms.AddItem "0"
        cb_icms.AddItem "4"
        cb_icms.AddItem "7"
        cb_icms.AddItem "12"
        cb_icms.AddItem "17"
        cb_icms.AddItem "18"
        cb_icms.AddItem "20"
        
        cb_icms_remessa.Clear
        cb_icms_remessa.AddItem "0"
        cb_icms_remessa.AddItem "4"
        cb_icms_remessa.AddItem "7"
        cb_icms_remessa.AddItem "12"
        cb_icms_remessa.AddItem "17"
        cb_icms_remessa.AddItem "18"
        cb_icms_remessa.AddItem "20"
        End If
        
    For i = cb_ICMS_item.LBound To cb_ICMS_item.UBound
        cb_ICMS_item(i).Clear
        cb_ICMS_item(i).AddItem ""
        For j = 0 To (cb_icms.ListCount - 1)
            If Trim$(cb_icms.List(j)) <> "" Then cb_ICMS_item(i).AddItem cb_icms.List(j)
            Next
        Next
    
'   FRETE POR CONTA
'   ~~~~~~~~~~~~~~~
    cb_frete.Clear
    'cb_frete.AddItem "0 - EMITENTE"
    'cb_frete.AddItem "1 - DESTINATÁRIO"
    cb_frete.AddItem "0 - Contratação do Remetente (CIF)"
    cb_frete.AddItem "1 - Contratação do Destinatário (FOB)"
    cb_frete.AddItem "2 - Contratação de Terceiros"
    cb_frete.AddItem "3 - Transporte Próprio Remetente"
    cb_frete.AddItem "4 - Transporte Próprio Destinatário"
    cb_frete.AddItem "9 - Sem Ocorrência"
    
'   ZERAR PIS/COFINS
'   ~~~~~~~~~~~~~~~~
    cb_zerar_PIS.Clear
    cb_zerar_PIS.AddItem "  "
    cb_zerar_PIS.AddItem "04 - Op. tributável (tributação monofásica (alíquota zero))"
    cb_zerar_PIS.AddItem "06 - Op. tributável (alíquota zero)"
    cb_zerar_PIS.AddItem "07 - Op. isenta da contribuição"
    cb_zerar_PIS.AddItem "08 - Op. sem incidência da contribuição"
    cb_zerar_PIS.AddItem "09 - Op. com suspensão da contribuição"
    
    cb_zerar_COFINS.Clear
    cb_zerar_COFINS.AddItem "  "
    cb_zerar_COFINS.AddItem "04 - Op. tributável (tributação monofásica (alíquota zero))"
    cb_zerar_COFINS.AddItem "06 - Op. tributável (alíquota zero)"
    cb_zerar_COFINS.AddItem "07 - Op. isenta da contribuição"
    cb_zerar_COFINS.AddItem "08 - Op. sem incidência da contribuição"
    cb_zerar_COFINS.AddItem "09 - Op. com suspensão da contribuição"
    
''   DADOS ADICIONAIS - VENDA
''   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    With c_dados_adicionais_venda
'        .FontName = FONTNAME_IMPRESSAO
'        .FontSize = FONTSIZE_IMPRESSAO
'        .FontBold = FONTBOLD_IMPRESSAO
'        .FontItalic = FONTITALIC_IMPRESSAO
'        End With
'
''   DADOS ADICIONAIS - REMESSA
''   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    With c_dados_adicionais_remessa
'        .FontName = FONTNAME_IMPRESSAO
'        .FontSize = FONTSIZE_IMPRESSAO
'        .FontBold = FONTBOLD_IMPRESSAO
'        .FontItalic = FONTITALIC_IMPRESSAO
'        End With
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FI_TRATA_ERRO:
'=============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Sub formulario_limpa()

Dim s As String
Dim s_aux As String
Dim i As Integer
Dim aliquota_icms As Single

    pedido_anterior = ""
    
'   CNPJ/CPF DESTINATÁRIO
'   ~~~~~~~~~~~~~~~~~~~~~
    c_cnpj_cpf_dest = ""
    c_nome_dest = ""
    c_rg_dest = ""
    
'   ITENS
'   ~~~~~
    c_vl_total_outras_despesas_acessorias = ""
    c_vl_total_geral = ""
    c_total_volumes = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_xPed(i) = ""
        c_nItemPed(i) = ""
        cb_ICMS_item(i).ListIndex = -1
        cb_ICMS_item(i) = ""
        c_NCM(i) = ""
        cb_CFOP(i).ListIndex = -1
        c_CST(i) = ""
        c_fabricante(i) = ""
        c_fabricante(i).ForeColor = vbBlack
        c_produto(i) = ""
        c_produto(i).ForeColor = vbBlack
        c_descricao(i) = ""
        c_descricao(i).ForeColor = vbBlack
        c_qtde(i) = ""
        c_vl_unitario(i) = ""
        c_vl_total(i) = ""
        c_vl_outras_despesas_acessorias(i) = ""
        c_produto_obs(i) = ""
        Next

'   FINALIDADE DE EMISSÃO
'   ~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "1 -"
    For i = 0 To cb_finalidade.ListCount - 1
        If left$(cb_finalidade.List(i), Len(s)) = s Then
            cb_finalidade.ListIndex = i
            Exit For
            End If
        Next
    
'   CHAVE DE ACESSO DA NFE REFERENCIADA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    c_chave_nfe_ref = ""
    
'   TIPO DO DOCUMENTO FISCAL
'   ~~~~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "1 -"
    For i = 0 To cb_tipo_NF.ListCount - 1
        If left$(cb_tipo_NF.List(i), Len(s)) = s Then
            cb_tipo_NF.ListIndex = i
            Exit For
            End If
        Next
    
'   LOCAL DE DESTINO DA OPERAÇÃO - VENDA E REMESSA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "1 -"
    For i = 0 To cb_loc_dest.ListCount - 1
        If left$(cb_loc_dest.List(i), Len(s)) = s Then
            cb_loc_dest.ListIndex = i
            Exit For
            End If
        Next
    For i = 0 To cb_loc_dest_remessa.ListCount - 1
        If left$(cb_loc_dest_remessa.List(i), Len(s)) = s Then
            cb_loc_dest_remessa.ListIndex = i
            Exit For
            End If
        Next
        
'   NATUREZA DA OPERAÇÃO - COMPRADOR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "6.108"
    For i = 0 To cb_natureza.ListCount - 1
        If left$(cb_natureza.List(i), Len(s)) = s Then
            cb_natureza.ListIndex = i
            Exit For
            End If
        Next
'   LOCAL DE DESTINO (Venda) - Interestadual
    cb_loc_dest.ListIndex = 1
        
'   NATUREZA DA OPERAÇÃO - RECEBEDOR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "6.923"
    For i = 0 To cb_natureza_recebedor.ListCount - 1
        If left$(cb_natureza_recebedor.List(i), Len(s)) = s Then
            cb_natureza_recebedor.ListIndex = i
            Exit For
            End If
        Next
'   LOCAL DE DESTINO (Remessa) - Interestadual
    cb_loc_dest_remessa.ListIndex = 1

'   ALÍQUOTAS ICMS
'   ~~~~~~~~~~~~~
'   DEFAULT VENDA
    s = "18"
    'por solicitação do Ricardo, vamos fixar a alíquota do ES em 12% no combobox do ICMS de venda,
    'portanto, não utilizaremos a busca na tabela de alíquotas (LHGX, 18/12/2017)
    'If obtem_aliquota_ICMS_UF_destino(usuario.emit_uf, aliquota_icms, s_aux) Then
    If False Then
        s = CStr(aliquota_icms)
    Else
        Select Case usuario.emit_uf
            Case "ES": s = "12"
            Case "MG": s = "18"
            Case "MS": s = "17"
            Case "RJ": s = "20"
            Case "SP": s = "18"
            Case "TO": s = "18"
            Case Else: s = "18"
            End Select
        End If
    
    For i = 0 To cb_icms.ListCount - 1
        If cb_icms.List(i) = s Then
            cb_icms.ListIndex = i
            Exit For
            End If
        Next
    
'   DEFAULT REMESSA
    s = "0"
    
    For i = 0 To cb_icms_remessa.ListCount - 1
        If cb_icms_remessa.List(i) = s Then
            cb_icms_remessa.ListIndex = i
            Exit For
            End If
        Next
     
'   ALÍQUOTA IPI
'   ~~~~~~~~~~~~
    c_ipi = ""
    
'   ZERAR PIS/COFINS
'   ~~~~~~~~~~~~~~~~
    cb_zerar_PIS.ListIndex = 0
    cb_zerar_COFINS.ListIndex = 0
    

'   FRETE POR CONTA
'   ~~~~~~~~~~~~~~~
'   DEFAULT
    s = "0 -"
    For i = 0 To cb_frete.ListCount - 1
        If left$(cb_frete.List(i), Len(s)) = s Then
            cb_frete.ListIndex = i
            Exit For
            End If
        Next


'   INFORMAÇOES DO PEDIDO
'   ~~~~~~~~~~~~~~~~~~~~~
    c_info_pedido = ""
           
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    c_dados_adicionais_venda = ""
    c_dados_adicionais_remessa = ""
    
'   FRAMES DE ENDEREÇO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~
    limpa_dados_endereco_comprador
    c_cnpj_cpf_dest = ""
    c_nome_dest = ""
    limpa_dados_endereco_cadastro
    limpa_campos_endereco_edicao
    pn_endereco_edicao.Visible = False
    
'   PARCELAS DE BOLETOS
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~
    blnExisteParcelamentoBoleto = False
    pnParcelasEmBoletos.Visible = False

'   REABILITANDO CONTROLES DESABILITADOS EM CASO DE EMISSÃO DE NOTA DE REMESSA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    chk_InfoComprador.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_cnpj_cpf_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_nome_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_rg_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    cb_natureza_recebedor.Enabled = True
    b_editar_endereco.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_dados_adicionais_venda.Locked = False
    
'   INFO CONTRIBUINTE
'   ~~~~~~~~~~~~~~~~~
    l_IE.Caption = ""
    
    
'   INFO OPERAÇÃO TRIANGULAR CONCLUÍDA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pn_aviso_operacao_concluida.Visible = False
    
'   FOCO INICIAL
'   ~~~~~~~~~~~~
    c_pedido = ""
    c_pedido.Enabled = True
    c_pedido.SetFocus
    
End Sub

Sub formulario_limpa_campos_itens_pedido()
Dim i As Integer
    
    c_vl_total_outras_despesas_acessorias = ""
    c_vl_total_geral = ""
    c_total_volumes = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_xPed(i) = ""
        c_nItemPed(i) = ""
        cb_ICMS_item(i).ListIndex = -1
        cb_ICMS_item(i) = ""
        c_NCM(i) = ""
        cb_CFOP(i).ListIndex = -1
        c_CST(i) = ""
        c_fabricante(i) = ""
        c_produto(i) = ""
        c_descricao(i) = ""
        c_qtde(i) = ""
        c_vl_unitario(i) = ""
        c_vl_total(i) = ""
        c_produto_obs(i) = ""
        c_vl_outras_despesas_acessorias(i) = ""
        Next
        
End Sub

Sub limpa_campos_endereco_edicao()
    
    usar_endereco_editado = False
    
    c_end_edicao_cep = ""
    c_end_edicao_logradouro = ""
    c_end_edicao_numero = ""
    c_end_edicao_complemento = ""
    c_end_edicao_bairro = ""
    c_end_edicao_cidade = ""
    c_end_edicao_uf = ""
    
End Sub

Sub limpa_dados_endereco_cadastro()

'    c_cnpj_cpf_dest = ""
'    c_nome_dest = ""
    c_rg_dest = ""
    l_end_recebedor_logradouro = ""
    l_end_recebedor_numero = ""
    l_end_recebedor_complemento = ""
    l_end_recebedor_bairro = ""
    l_end_recebedor_cidade = ""
    l_end_recebedor_uf = ""
    l_end_recebedor_cep = ""

End Sub

Sub limpa_dados_endereco_comprador()

    l_comprador_nome = ""
    l_end_comprador_logradouro = ""
    l_end_comprador_numero = ""
    l_end_comprador_complemento = ""
    l_end_comprador_bairro = ""
    l_end_comprador_cidade = ""
    l_end_comprador_uf = ""
    l_end_comprador_cep = ""

End Sub

Sub trata_botao_cancela_triangular()
Dim s As String

    
    If Trim(c_pedido) = "" Then
        aviso_erro "Pedido não informado!!"
        Exit Sub
        End If
        
    s = "Confirma o cancelamento da operação triangular atual?"
    
    If confirma(s) Then
        infocontagem.Visible = False
        c_pedido.Enabled = True
        contagem.Enabled = False
        blnEsperaNFTriangular = False
        
        If lngIdNFeTriangular > 0 Then
            If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
                aviso_erro s
                End If
            lngIdNFeTriangular = 0
            End If
        formulario_limpa
'        If sPedidoTriangular <> "" Then
'            s = "Deseja encaminhar o pedido " & sPedidoTriangular & " para a emissão padrão (sem operação triangular)?"
'            If Not confirma(s) Then sPedidoTriangular = ""
'            fechar_modo_emissao_nfe_triangular
'            End If
        If sPedidoTriangular <> "" Then fechar_modo_emissao_nfe_triangular
        End If
    
End Sub

Sub trata_botao_endereco_edicao_limpa()

    limpa_campos_endereco_edicao
    c_end_edicao_cep.SetFocus

End Sub

Sub trata_botao_endereco_edicao_cancela()

    pn_endereco_edicao.Visible = False

End Sub

Sub trata_botao_endereco_edicao_ok()
    
    If Trim$(c_end_edicao_cep) = "" Then
        aviso_erro "Preencha o CEP!!"
        c_end_edicao_cep.SetFocus
        Exit Sub
        End If
    
    If Len(retorna_so_digitos(c_end_edicao_cep)) <> 8 Then
        aviso_erro "CEP com tamanho inválido!!"
        c_end_edicao_cep.SetFocus
        Exit Sub
        End If
    
    If Trim$(c_end_edicao_logradouro) = "" Then
        aviso_erro "O campo endereço está vazio!!"
        c_end_edicao_logradouro.SetFocus
        Exit Sub
        End If
    
    If Trim$(c_end_edicao_numero) = "" Then
        aviso_erro "O campo número do endereço está vazio!!"
        c_end_edicao_numero.SetFocus
        Exit Sub
        End If
    
    If Trim$(c_end_edicao_bairro) = "" Then
        aviso_erro "O campo bairro está vazio!!"
        c_end_edicao_bairro.SetFocus
        Exit Sub
        End If
    
    If Trim$(c_end_edicao_cidade) = "" Then
        aviso_erro "O campo cidade está vazio!!"
        c_end_edicao_cidade.SetFocus
        Exit Sub
        End If
    
    If Not UF_ok(c_end_edicao_uf) Then
        aviso_erro "UF inválida!!"
        c_end_edicao_uf.SetFocus
        Exit Sub
        End If
    
    endereco_editado__cep = cep_formata(retorna_so_digitos(c_end_edicao_cep))
    l_end_recebedor_cep = endereco_editado__cep
    
    endereco_editado__logradouro = Trim$(c_end_edicao_logradouro)
    l_end_recebedor_logradouro = endereco_editado__logradouro
    
    endereco_editado__numero = Trim$(c_end_edicao_numero)
    l_end_recebedor_numero = endereco_editado__numero
    
    endereco_editado__complemento = Trim$(c_end_edicao_complemento)
    l_end_recebedor_complemento = endereco_editado__complemento
    
    endereco_editado__bairro = Trim$(c_end_edicao_bairro)
    l_end_recebedor_bairro = endereco_editado__bairro
    
    endereco_editado__cidade = Trim$(c_end_edicao_cidade)
    l_end_recebedor_cidade = endereco_editado__cidade
    
    endereco_editado__uf = Trim$(c_end_edicao_uf)
    l_end_recebedor_uf = endereco_editado__uf
    
    usar_endereco_editado = True
    
    pn_endereco_edicao.Visible = False
    
    c_dados_adicionais_venda = obtem_dados_adicionais_venda(lngIdNFeTriangular)
    c_dados_adicionais_venda.ToolTipText = c_dados_adicionais_venda
    
End Sub

Sub trata_botao_pesquisa_cep()

Dim s As String
Dim s_cep As String
Dim t As ADODB.Recordset

    On Error GoTo TBPESQCEP_TRATA_ERRO
    
    s_cep = retorna_so_digitos(Trim$(c_end_edicao_cep))
    
    limpa_campos_endereco_edicao
        
    c_end_edicao_cep = cep_formata(s_cep)
    
    If Trim$(s_cep) = "" Then Exit Sub
   
    If Len(s_cep) <> 8 Then
        aviso_erro "CEP informado com tamanho inválido!!"
        c_end_edicao_cep.SetFocus
        Exit Sub
        End If
    
    
    On Error GoTo TBPESQCEP_TRATA_ERRO_COM_FECHA_TABELAS
    
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT" & _
            " '1_LOGRADOURO' AS tabela_origem," & _
            " Logr.CEP_DIG AS cep," & _
            " Logr.UFE_SG AS uf," & _
            " Loc.LOC_NOSUB AS localidade," & _
            " Bai.BAI_NO AS bairro_extenso," & _
            " Bai.BAI_NO_ABREV AS bairro_abreviado," & _
            " Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," & _
            " Logr.LOG_NO AS logradouro_nome," & _
            " Logr.LOG_COMPLEMENTO AS logradouro_complemento" & _
        " FROM LOG_LOGRADOURO Logr" & _
            " LEFT JOIN LOG_BAIRRO Bai ON (Logr.BAI_NU_SEQUENCIAL_INI = Bai.BAI_NU_SEQUENCIAL)" & _
            " LEFT JOIN LOG_LOCALIDADE Loc ON (Logr.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" & _
        " WHERE" & _
            " (Logr.CEP_DIG = '" & s_cep & "')"
    
    s = s & _
        " UNION " & _
        "SELECT" & _
            " '2_LOCALIDADE' AS tabela_origem," & _
            " CEP_DIG AS cep," & _
            " UFE_SG AS uf," & _
            " LOC_NOSUB AS localidade," & _
            " '' AS bairro_extenso," & _
            " '' AS bairro_abreviado," & _
            " '' AS logradouro_tipo," & _
            " '' AS logradouro_nome," & _
            " '' AS logradouro_complemento" & _
        " FROM LOG_LOCALIDADE" & _
        " WHERE" & _
            " (CEP_DIG = '" & s_cep & "')"
            
'   CONSULTA DADOS DA TABELA ANTIGA, POIS ELA É MANTIDA P/ MANTER FUNCIONANDO O CADASTRAMENTO MANUAL DE CEP'S
    s = s & _
        " UNION " & _
        "SELECT" & _
            " '3_LOGRADOURO' AS tabela_origem," & _
            " cep8_log" & SQL_COLLATE_CASE_ACCENT & " AS cep," & _
            " uf_log" & SQL_COLLATE_CASE_ACCENT & " AS uf," & _
            " nome_local" & SQL_COLLATE_CASE_ACCENT & " AS localidade," & _
            " extenso_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_extenso," & _
            " abrev_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_abreviado," & _
            " abrev_tipo" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_tipo," & _
            " nome_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_nome," & _
            " comple_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_complemento" & _
        " FROM t_CEP_LOGRADOURO " & _
        " WHERE" & _
            " (cep8_log = '" & s_cep & "')"
            
    s = s & _
        " ORDER BY" & _
            " tabela_origem," & _
            " cep"
    
    t.Open s, dbcCep, , , adCmdText
    If t.EOF Then
        GoSub TBPESQCEP_FECHA_TABELAS
        aviso_erro "CEP " & cep_formata(s_cep) & " NÃO foi encontrado na base de dados!!"
        c_end_edicao_cep.SetFocus
        Exit Sub
        End If
    
    c_end_edicao_logradouro = Trim$(Trim$("" & t("logradouro_tipo")) & " " & Trim$("" & t("logradouro_nome")))
    c_end_edicao_bairro = Trim$("" & t("bairro_extenso"))
    If Trim$(c_end_edicao_bairro) = "" Then
        c_end_edicao_bairro = Trim$("" & t("bairro_abreviado"))
        End If
    c_end_edicao_cidade = Trim$("" & t("localidade"))
    c_end_edicao_uf = Trim$("" & t("uf"))
    
    GoSub TBPESQCEP_FECHA_TABELAS
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBPESQCEP_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBPESQCEP_TRATA_ERRO_COM_FECHA_TABELAS:
'======================================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub TBPESQCEP_FECHA_TABELAS
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBPESQCEP_FECHA_TABELAS:
'=======================
    bd_desaloca_recordset t, True
    Return
    
End Sub

Function ObtemRetornoOperacaoTriangularConcluida(ByVal lNumeroNF As Long, _
                                            ByVal lSerieNF As Long, _
                                            ByRef iRetorno As Integer, _
                                            ByRef sMensagem As String) As Boolean
Dim s As String
Dim s_aux As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String

Dim cmdNFeSituacao As New ADODB.Command
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim dbcNFe As ADODB.Connection

    On Error GoTo NFE_OROTC_TRATA_ERRO
    
    ObtemRetornoOperacaoTriangularConcluida = False

    aguarde INFO_EXECUTANDO, "Obtendo informações sobre a operação triangular realizada anteriormente"
    
  ' t_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
  
  '   CONEXÃO AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   PREPARA COMMAND
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)


    s = "SELECT" & _
        " razao_social," & _
        " NFe_T1_servidor_BD," & _
        " NFe_T1_nome_BD," & _
        " NFe_T1_usuario_BD," & _
        " NFe_T1_senha_BD" & _
    " FROM t_NFE_EMITENTE" & _
    " WHERE" & _
        " (id = " & usuario.emit_id & ")"
    If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    If t_NFE_EMITENTE.EOF Then
        aviso_erro "Falha na localização dos dados do emitente"
        GoTo NFE_OROTC_FECHA_TABELAS
        End If
        
    strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
    strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
    strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
    strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
    strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
    
    decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
    s = "Provider=" & BD_OLEDB_PROVIDER & _
        ";Data Source=" & strNfeT1ServidorBd & _
        ";Initial Catalog=" & strNfeT1NomeBd & _
        ";User Id=" & strNfeT1UsuarioBd & _
        ";Password=" & s_aux
    If dbcNFe.State <> adStateClosed Then dbcNFe.Close
    dbcNFe.Open s

'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    Set cmdNFeSituacao.ActiveConnection = dbcNFe

    strNumeroNfNormalizado = NFeFormataNumeroNF(lNumeroNF)
    strSerieNfNormalizado = NFeFormataSerieNF(lSerieNF)

'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
    cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
    Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
    iRetorno = rsNFeRetornoSPSituacao("Retorno")
    sMensagem = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))

    ObtemRetornoOperacaoTriangularConcluida = True
    
    GoTo NFE_OROTC_FECHA_TABELAS
        
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_OROTC_FECHA_TABELAS:
'=======================
  ' RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    
  ' COMMAND
    bd_desaloca_command cmdNFeSituacao
    
  ' CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
    
    aguarde INFO_NORMAL, m_id
    
    Exit Function

    
NFE_OROTC_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    GoTo NFE_OROTC_FECHA_TABELAS
    

End Function

Sub pedido_preenche_dados_tela(ByVal pedido As String)
Dim s_resp As String
Dim s_end_entrega As String
Dim s_end_entrega_uf As String
Dim s_end_cliente_uf As String
Dim s_erro As String
Dim s As String
Dim i As Integer
Dim i_status_triangular As Integer
Dim lngId As Long
Dim strUsuario As String
Dim blnNotadeVendaEmitida As Boolean
Dim receb_cnpj_cpf As String
Dim receb_nome As String
Dim receb_rg As String
Dim receb_endereco As String
Dim receb_numero As String
Dim receb_complemento As String
Dim receb_bairro As String
Dim receb_cidade As String
Dim receb_uf As String
Dim receb_cep As String
Dim s_venda_natop As String
Dim s_venda_frete As String
Dim s_venda_destop As String
Dim i_retorno_venda As Integer
Dim i_retorno_remessa As Integer
Dim s_retorno_venda As String
Dim s_retorno_remessa As String
Dim strLogPedido As String
Dim strLogComplemento As String
Dim s_NFe_texto_constar As String

    On Error GoTo PPDT_TRATA_ERRO
    
    If Trim$(pedido) = "" Then Exit Sub

    If pedido_anterior = Trim$(pedido) Then Exit Sub
    pedido_anterior = Trim$(pedido)
    
    If pedido <> "" Then
        'verificar se os dados do cliente devem vir da memorização no pedido
        If (param_pedidomemorizacaoenderecos.campo_inteiro) Then
            If Not obtem_info_pedido_triangular_memorizada(pedido, s_resp, s_end_entrega, s_end_entrega_uf, s_NFe_texto_constar, s_end_cliente_uf, s_erro) Then
                If s_erro <> "" Then
                    aviso_erro s_erro
                    c_pedido = ""
                    pedido_anterior = ""
                    c_pedido.SetFocus
                    Exit Sub
                    End If
                End If
        Else
            If Not obtem_info_pedido_triangular(pedido, s_resp, s_end_entrega, s_end_entrega_uf, s_NFe_texto_constar, s_end_cliente_uf, s_erro) Then
                If s_erro <> "" Then
                    aviso_erro s_erro
                    c_pedido = ""
                    pedido_anterior = ""
                    c_pedido.SetFocus
                    Exit Sub
                    End If
                End If
            End If
        End If
    
'   VERIFICAR SE A NOTA DE VENDA JÁ FOI EMITIDA (em caso positivo, carregar apenas para emissão da nota de remessa)
    blnNotadeVendaEmitida = pedido_vinculado_a_nf_triangular(pedido, i_status_triangular, lngIdNFeTriangular, lngSerieNFeTriangular, lngNumVendaNFeTriangular, lngNumRemessaNFeTriangular, receb_cnpj_cpf, receb_nome, receb_rg, receb_endereco, receb_numero, receb_complemento, receb_bairro, receb_cidade, receb_uf, receb_cep)
    
    If Not blnNotadeVendaEmitida Then
    '   SE O PARÂMETRO NF_MaxSegundos_EsperaNotaTriangular FOR ZERO OU INEXISTENTE, NÃO HAVERÁ BLOQUEIO NA EMISSÃO DA NOTA TRIANGULAR
        blnEsperaNFTriangular = False
        MaxSegundosNotaTriangular = retorna_max_segundos_nota_triangular
        If MaxSegundosNotaTriangular > 0 Then blnEsperaNFTriangular = True
        If blnEsperaNFTriangular Then
        '   SE O USUÁRIO NÃO ESTIVER EMITINDO NOTA TRIANGULAR, VERIFICA SE EXISTE OUTRA NOTA TRIANGULAR SENDO EMITIDA
            If NFeExisteNotaTriangularEmEmissao(lngId, strUsuario, s_erro) Then
                'Em caso positivo, informar e sair
                aviso "Nota triangular sendo emitida pelo usuário " & strUsuario
                c_pedido = ""
                pedido_anterior = ""
                c_pedido.SetFocus
                Exit Sub
            Else
                'Em caso negativo, inserir registro para emissão de nota triangular
                If Not insere_registro_nfe_triangular(lngIdNFeTriangular, lngSerieNFeTriangular, lngNumVendaNFeTriangular, lngNumRemessaNFeTriangular, pedido, True, s_erro) Or _
                    Not atualiza_nfe_triangular_venda(lngIdNFeTriangular, ST_NFT_EM_PROCESSAMENTO, s_erro) Then
                    aviso_erro s_erro
                    c_pedido.SetFocus
                    Exit Sub
                    End If
                dt_hr_inicio_emissao = Now
                blnEmissaoOK = False
                infocontagem.Caption = ""
                infocontagem.Visible = True
                c_pedido.Enabled = False
                contagem.Enabled = True
                End If
            End If
    Else
        'se a nota de remessa já foi emitida, não há necessidade de bloquear emissão de outras notas
        blnEmissaoOK = True
        infocontagem.Caption = ""
        infocontagem.Visible = False
        c_pedido.Enabled = True
        contagem.Enabled = False
        End If
    
'   EXIBE OS ITENS DO PEDIDO NA TELA
    formulario_exibe_itens_pedido_triangular Trim$(pedido)
                
    If blnNotadeVendaEmitida Then
        endereco_recebedor__nome = receb_nome
        endereco_recebedor__cnpj_cpf = receb_cnpj_cpf
        endereco_recebedor__rg = receb_rg
        endereco_recebedor__logradouro = receb_endereco
        endereco_recebedor__numero = receb_numero
        endereco_recebedor__complemento = receb_complemento
        endereco_recebedor__bairro = receb_bairro
        endereco_recebedor__cidade = receb_cidade
        endereco_recebedor__uf = receb_uf
        endereco_recebedor__cep = receb_cep
        End If
        
    c_info_pedido = s_resp
    l_comprador_nome.Caption = endereco_comprador__nome
    l_end_comprador_logradouro.Caption = endereco_comprador__logradouro
    l_end_comprador_numero.Caption = endereco_comprador__numero
    l_end_comprador_complemento.Caption = endereco_comprador__complemento
    l_end_comprador_bairro.Caption = endereco_comprador__bairro
    l_end_comprador_cidade.Caption = endereco_comprador__cidade
    l_end_comprador_uf.Caption = endereco_comprador__uf
    l_end_comprador_cep.Caption = endereco_comprador__cep
    l_end_recebedor_logradouro.Caption = endereco_recebedor__logradouro
    l_end_recebedor_numero.Caption = endereco_recebedor__numero
    l_end_recebedor_complemento.Caption = endereco_recebedor__complemento
    l_end_recebedor_bairro.Caption = endereco_recebedor__bairro
    l_end_recebedor_cidade.Caption = endereco_recebedor__cidade
    l_end_recebedor_uf.Caption = endereco_recebedor__uf
    l_end_recebedor_cep.Caption = endereco_recebedor__cep
    If blnNotadeVendaEmitida Then
        'se existe nota de venda emitida, usar os dados preenchidos do recebedor
        c_cnpj_cpf_dest = cnpj_cpf_formata(endereco_recebedor__cnpj_cpf)
        c_nome_dest = endereco_recebedor__nome
        c_rg_dest = endereco_recebedor__rg
    ElseIf param_pedidomemorizacaoenderecos.campo_inteiro = 1 Then
        'se ainda não existe nota de venda emitida e existe memorização de endereço, usar os dados da t_PEDIDO
        c_cnpj_cpf_dest = cnpj_cpf_formata(endereco_recebedor__cnpj_cpf)
        c_nome_dest = endereco_recebedor__nome
        c_rg_dest = endereco_recebedor__rg
    Else
        'no caso de pessoa jurídica, a opção será pesquisar/preencher os dados do recebedor
        chk_InfoComprador.Value = 0
        chk_InfoComprador.Enabled = False
        c_cnpj_cpf_dest = ""
        c_nome_dest = ""
        'no caso de pessoa física, preencher os campos de CPF e Nome com os dados do comprador
        If Len(endereco_comprador__cnpj_cpf) < 14 Then
            chk_InfoComprador.Value = 1
            chk_InfoComprador.Enabled = True
            c_cnpj_cpf_dest = cnpj_cpf_formata(endereco_comprador__cnpj_cpf)
            c_rg_dest = endereco_comprador__rg
            c_nome_dest = endereco_comprador__nome
            End If
        End If
    
    c_dados_adicionais_venda = obtem_dados_adicionais_venda(lngIdNFeTriangular)
    c_dados_adicionais_venda.ToolTipText = c_dados_adicionais_venda
    c_dados_adicionais_remessa = obtem_dados_adicionais_remessa(lngIdNFeTriangular)
    c_dados_adicionais_remessa.ToolTipText = c_dados_adicionais_remessa

    'verificar se existe informação de parcelas em boleto
    If (param_geracaoboletos.campo_texto = "Manual") Then
        If pedido <> "" Then
            ReDim v_pedido_manual_boleto(0)
            v_pedido_manual_boleto(UBound(v_pedido_manual_boleto)) = pedido
            blnExisteParcelamentoBoleto = False
            pnParcelasEmBoletos.Visible = False
            If ExisteDadosParcelasPagto(pedido, s_erro) And _
                consultaDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
                AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
                If blnExisteParcelamentoBoleto Then
                    pnParcelasEmBoletos.Visible = True
                    pnParcelasEmBoletos.Enabled = False
                    c_dataparc.Enabled = False
                    End If
            ElseIf geraDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
                AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
                If blnExisteParcelamentoBoleto Then
                    pnParcelasEmBoletos.Visible = True
                    pnParcelasEmBoletos.Enabled = True
                    c_dataparc.Enabled = True
                    End If
                End If
            End If
        End If

    'AJUSTES NA TELA, DEPENDENDO DO STATUS DA OPERAÇÃO TRIANGULAR
    '(se for em processamento, tem que testar se ambos foram emitidos [venda e remessa] e fazer um tratamento diferente para cada)
    b_imprime_venda.Enabled = True
    b_imprime_remessa.Enabled = False
    b_cancela_triangular.Enabled = True
    chk_InfoComprador.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_cnpj_cpf_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_nome_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_rg_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    cb_natureza_recebedor.Enabled = True
    b_editar_endereco.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
    c_dados_adicionais_venda.Locked = False
    c_dados_adicionais_remessa.Locked = False
    
    Select Case i_status_triangular
        Case ST_NFT_EM_PROCESSAMENTO:
            b_cancela_triangular.Enabled = True
            b_cancela_triangular.Visible = True
            If blnNotadeVendaEmitida Then
                chk_InfoComprador.Value = 0
                chk_InfoComprador.Enabled = False
                c_cnpj_cpf_dest.Enabled = False
                c_nome_dest.Enabled = False
                c_rg_dest.Enabled = False
                b_editar_endereco.Enabled = False
                c_dados_adicionais_venda.Locked = True
                c_dados_adicionais_remessa.Locked = False
                pn_aviso_operacao_concluida.Visible = False
                b_imprime_venda.Enabled = False
                b_imprime_remessa.Enabled = True
                If obtem_dados_nf_venda(lngNumVendaNFeTriangular, lngSerieNFeTriangular, s_venda_destop, s_venda_frete, s_venda_natop) Then
                    'ajustando 'local de destino da venda'
                    For i = 0 To cb_loc_dest.ListCount - 1
                        If left$(cb_loc_dest.List(i), Len(s_venda_destop)) = s_venda_destop Then
                            cb_loc_dest.ListIndex = i
                            Exit For
                            End If
                        Next
                    'ajustando 'frete por conta'
                    For i = 0 To cb_frete.ListCount - 1
                        If left$(cb_frete.List(i), Len(s_venda_frete)) = s_venda_frete Then
                            cb_frete.ListIndex = i
                            Exit For
                            End If
                        Next
                    'ajustando 'natureza da operação'
                    For i = 0 To cb_natureza.ListCount - 1
                        If left$(cb_natureza.List(i), Len(s_venda_natop)) = s_venda_natop Then
                            cb_natureza.ListIndex = i
                            Exit For
                            End If
                        Next
                    End If
            Else
                chk_InfoComprador.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
                c_cnpj_cpf_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
                c_nome_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
                c_rg_dest.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
                cb_natureza_recebedor.Enabled = True
                b_editar_endereco.Enabled = param_pedidomemorizacaoenderecos.campo_inteiro = 0
                c_dados_adicionais_venda.Locked = False
                c_dados_adicionais_remessa.Locked = False
                pn_aviso_operacao_concluida.Visible = False
                b_imprime_venda.Enabled = True
                b_imprime_remessa.Enabled = False
            '   ICMS
            '   DEFAULT
                s = "18"
                Select Case usuario.emit_uf
                    Case "ES": s = "12"
                    Case "MG": s = "18"
                    Case "MS": s = "17"
                    Case "RJ": s = "20"
                    Case "SP": s = "18"
                    Case "TO": s = "18"
                    Case Else: s = "18"
                    End Select
                For i = 0 To cb_icms.ListCount - 1
                    If cb_icms.List(i) = s Then
                        cb_icms.ListIndex = i
                        Exit For
                        End If
                    Next
                End If
        Case ST_NFT_EMITIDA:
            If Not ObtemRetornoOperacaoTriangularConcluida(lngNumVendaNFeTriangular, lngSerieNFeTriangular, i_retorno_venda, s_retorno_venda) Or _
            Not ObtemRetornoOperacaoTriangularConcluida(lngNumRemessaNFeTriangular, lngSerieNFeTriangular, i_retorno_remessa, s_retorno_remessa) Then
                aviso_erro "Não foi possível consultar situação da operação na Target, tente novamente mais tarde"
                b_cancela_triangular.Enabled = False
                chk_InfoComprador.Value = 0
                chk_InfoComprador.Enabled = False
                c_cnpj_cpf_dest.Enabled = False
                c_nome_dest.Enabled = False
                c_rg_dest.Enabled = False
                cb_natureza_recebedor.Enabled = False
                b_editar_endereco.Enabled = False
                c_dados_adicionais_venda.Locked = True
                c_dados_adicionais_remessa.Locked = True
                lbl_aviso_operacao_concluida.Caption = "Situação da operação na Target desconhecida"
                pn_aviso_operacao_concluida.Visible = True
                b_imprime_venda.Enabled = False
                b_imprime_remessa.Enabled = False
            Else
                'primeira situação: ambas as notas foram emitidas anteriormente com sucesso
                '(exibir na tela)
                If (i_retorno_venda = 1) And (i_retorno_remessa = 1) Then
                    b_cancela_triangular.Enabled = False
                    chk_InfoComprador.Value = 0
                    chk_InfoComprador.Enabled = False
                    c_cnpj_cpf_dest.Enabled = False
                    c_nome_dest.Enabled = False
                    c_rg_dest.Enabled = False
                    cb_natureza_recebedor.Enabled = False
                    b_editar_endereco.Enabled = False
                    c_dados_adicionais_venda.Locked = True
                    c_dados_adicionais_remessa.Locked = True
                    lbl_aviso_operacao_concluida.Caption = "Pedido com Operação Triangular Concluída"
                    pn_aviso_operacao_concluida.Visible = True
                    b_imprime_venda.Enabled = False
                    b_imprime_remessa.Enabled = False
                'segunda situação: alguma das notas está em processamento
                '(solicitar ação do usuário)
                 ElseIf ((i_retorno_venda <> 1) And ((InStr(s_retorno_venda, "Em processamento") > 0) Or (InStr(s_retorno_venda, "Aguardando processamento") > 0))) Or _
                        ((i_retorno_remessa <> 1) And (InStr(s_retorno_remessa, "Em processamento") > 0) Or (InStr(s_retorno_remessa, "Aguardando processamento") > 0)) Then
                    s_erro = "Pedido " & pedido & ":" & vbCrLf
                    If InStr(s_retorno_venda, "processamento") > 0 Then
                        s_erro = s_erro & "Informação sobre a nota de venda nº " & CStr(lngNumVendaNFeTriangular) & vbCrLf & _
                            "(Mensagem: " & s_retorno_venda & ")" & vbCrLf
                        End If
                    If InStr(s_retorno_remessa, "processamento") > 0 Then
                        s_erro = s_erro & "Informação sobre a nota de remessa nº " & CStr(lngNumRemessaNFeTriangular) & vbCrLf & _
                            "(Mensagem: " & s_retorno_remessa & ")" & vbCrLf
                        End If
                    s_erro = s_erro & vbCrLf & "Deseja efetuar uma nova emissão de notas fiscais para este pedido?"
                    If Not confirma(s_erro) Then
                        formulario_limpa
                        Exit Sub
                        End If
                    s = "Forneça a senha para CANCELAR a operação triangular"
                    f_CONFIRMACAO_VIA_SENHA.strMensagemInformativa = s
                    f_CONFIRMACAO_VIA_SENHA.strSenhaCorreta = usuario.senha
                    f_CONFIRMACAO_VIA_SENHA.Show vbModal, Me
                    If f_CONFIRMACAO_VIA_SENHA.blnResultadoFormOk Then
                        If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_USUARIO, s) Then
                            aviso_erro s
                        Else
                            aviso "Operação triangular cancelada, para realizar nova operação redigite o número do pedido (" & pedido & ")"
                            'grava log OP_LOG_NFE_EMISSAO_TRIANGULAR
                            aguarde INFO_EXECUTANDO, "gravando log"
                            strLogPedido = pedido
                            strLogComplemento = "Cancelamento de Operação Triangular Mediante Uso de Senha" & _
                                                ";Nota de venda=" & CStr(lngNumVendaNFeTriangular) & _
                                                "; Código retorno de venda=" & CStr(i_retorno_venda) & _
                                                "; Mensagem retorno de venda=" & s_retorno_venda & _
                                                "; Nota de remessa=" & CStr(lngNumRemessaNFeTriangular) & _
                                                "; Código retorno de remessa=" & CStr(i_retorno_remessa) & _
                                                "; Mensagem retorno de remessa=" & s_retorno_remessa
                            Call grava_log(usuario.id, "", strLogPedido, "", OP_LOG_NFE_EMISSAO_TRIANGULAR, strLogComplemento)
                            End If
                        End If
                    aguarde INFO_NORMAL, m_id
                    lngIdNFeTriangular = 0
                    formulario_limpa
               'terceira situação: ambas as notas apresentaram erro na emissão e a mensagem de retorno foi fornecida
                '(avisar e solicitar cancelamento)
                ElseIf ((i_retorno_venda <> 1) And (s_retorno_venda <> "")) Or _
                        ((i_retorno_remessa <> 1) And (s_retorno_remessa <> "")) Then
                    s_erro = "Pedido " & pedido & ":" & vbCrLf
                    If i_retorno_venda <> 1 Then
                        s_erro = s_erro & "Problemas na emissão da nota de venda nº " & CStr(lngNumVendaNFeTriangular) & vbCrLf & _
                            "(Mensagem: " & s_retorno_venda & ")" & vbCrLf
                        End If
                    If i_retorno_remessa <> 1 Then
                        s_erro = s_erro & "Problemas na emissão da nota de remessa nº " & CStr(lngNumRemessaNFeTriangular) & vbCrLf & _
                            "(Mensagem: " & s_retorno_remessa & ")" & vbCrLf
                        End If
                    's_erro = s_erro & vbCrLf & "A operação será AUTOMATICAMENTE CANCELADA pelo sistema."
                    s_erro = s_erro & vbCrLf & "Favor CANCELAR a emissão triangular."
                    aviso s_erro
                    'If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_SISTEMA, s) Then
                    '    aviso_erro s
                    '    End If
                    'lngIdNFeTriangular = 0
                    'formulario_limpa
                    b_imprime_venda.Enabled = False
                    b_imprime_remessa.Enabled = False
                Else
                    'situação de possível falta de retorno do servidor da Target; tentar novamente mais tarde
                    aviso_erro "Não foi possível consultar situação da operação na Target, tente novamente mais tarde"
                    b_cancela_triangular.Enabled = False
                    chk_InfoComprador.Value = 0
                    chk_InfoComprador.Enabled = False
                    c_cnpj_cpf_dest.Enabled = False
                    c_nome_dest.Enabled = False
                    c_rg_dest.Enabled = False
                    cb_natureza_recebedor.Enabled = False
                    b_editar_endereco.Enabled = False
                    c_dados_adicionais_venda.Locked = True
                    c_dados_adicionais_remessa.Locked = True
                    lbl_aviso_operacao_concluida.Caption = "Situação da operação na Target desconhecida"
                    pn_aviso_operacao_concluida.Visible = True
                    b_imprime_venda.Enabled = False
                    b_imprime_remessa.Enabled = False
                    End If
                End If
        End Select

    Exit Sub
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPDT_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    Exit Sub


End Sub

Function obtem_info_pedido_triangular(ByVal pedido As String, ByRef strResposta As String, ByRef strEndEntregaFormatado As String, ByRef strEndEntregaUf As String, ByRef strTextoConstar As String, ByRef strEndClienteUf As String, ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_pedido_triangular()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_nome As String
Dim s_cnpj_cpf As String
Dim s_ie_rg As String
Dim s_obs_1 As String
Dim s_info As String
Dim s_endereco As String
Dim s_end_linha_1 As String
Dim s_end_linha_2 As String
Dim s_end_linha_3 As String
Dim s_end_entrega As String
Dim pedido_a As String
Dim s_id_cliente As String
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strRamal As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim strInfoIE As String
Dim strNFeTextoConstar As String

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo OIPT_TRATA_ERRO
    
    obtem_info_pedido_triangular = False
    strMsgErro = ""
    strResposta = ""
    strEndEntregaFormatado = ""
    strEndEntregaUf = ""
    strEndClienteUf = ""
    l_IE = ""
    
    pedido = Trim$("" & pedido)
    pedido = normaliza_num_pedido(pedido)
    
    If pedido = "" Then
        strMsgErro = "Não foi informado o número do pedido!"
        Exit Function
        End If
        
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_DESTINATARIO (PODE SER T_CLIENTE OU T_LOJA)
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    endereco_comprador__nome = ""
    endereco_comprador__cnpj_cpf = ""
    endereco_comprador__rg = ""
    endereco_comprador__logradouro = ""
    endereco_comprador__numero = ""
    endereco_comprador__complemento = ""
    endereco_comprador__bairro = ""
    endereco_comprador__cep = ""
    endereco_comprador__cidade = ""
    endereco_comprador__uf = ""
    endereco_recebedor__logradouro = ""
    endereco_recebedor__numero = ""
    endereco_recebedor__complemento = ""
    endereco_recebedor__bairro = ""
    endereco_recebedor__cep = ""
    endereco_recebedor__cidade = ""
    endereco_recebedor__uf = ""
    s_endereco = ""
    s_nome = ""
    s_cnpj_cpf = ""
    s_ie_rg = ""
    s_obs_1 = ""
    s_end_entrega = ""
        

'   VERIFICA O PEDIDO
    s_id_cliente = ""
    pedido_a = ""
    s_erro = ""
    s = "SELECT" & _
            " pedido, st_entrega, id_cliente, obs_1, st_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep, NFe_texto_constar" & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & Trim$(pedido) & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Pedido " & Trim$(pedido) & " não está cadastrado !!"
    Else
'   TEXTO A CONSTAR NA NOTA FISCAL
    strNFeTextoConstar = Trim("" & t_PEDIDO("NFe_texto_constar"))
    
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA DE VENDA
    s_id_cliente = Trim$("" & t_PEDIDO("id_cliente"))
    s = "SELECT * FROM t_CLIENTE WHERE (id='" & s_id_cliente & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        strMsgErro = "Cliente com nº registro " & s_id_cliente & " não foi encontrado!!"
        GoSub OIPT_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
    
    '   ENDEREÇO DE ENTREGA
        endereco_recebedor__logradouro = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco")))
        endereco_recebedor__numero = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_numero")))
        endereco_recebedor__complemento = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_complemento")))
        endereco_recebedor__bairro = UCase$(Trim("" & t_PEDIDO("EndEtg_bairro")))
        endereco_recebedor__cep = cep_formata(retorna_so_digitos(Trim("" & t_PEDIDO("EndEtg_cep"))))
        endereco_recebedor__cidade = UCase$(Trim("" & t_PEDIDO("EndEtg_cidade")))
        endereco_recebedor__uf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
        If (CLng(t_PEDIDO("st_end_entrega")) <> 0) And (endereco_recebedor__uf <> strEndClienteUf) Then
            s_end_entrega = formata_endereco(endereco_recebedor__logradouro, endereco_recebedor__numero, endereco_recebedor__complemento, endereco_recebedor__bairro, endereco_recebedor__cidade, endereco_recebedor__uf, endereco_recebedor__cep)
            strEndEntregaFormatado = s_end_entrega
            strEndEntregaUf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
            If s_end_entrega <> "" Then s_end_entrega = vbCrLf & "ENTREGA: " & s_end_entrega
        Else
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Operação triangular interestadual não caracterizada!!"
            End If
    
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        If Trim$("" & t_PEDIDO("obs_1")) <> "" Then
            If s_obs_1 <> "" Then s_obs_1 = s_obs_1 & vbCrLf
            s = Trim$("" & t_PEDIDO("obs_1"))
            s = substitui_caracteres(s, vbCr, " ")
            s = substitui_caracteres(s, vbLf, " ")
            s_obs_1 = s_obs_1 & s
            End If
        End If
    
    s = "SELECT pedido, fabricante, produto FROM t_PEDIDO_ITEM WHERE (pedido='" & Trim$(pedido) & "')"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    If t_PEDIDO_ITEM.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Não foi encontrado nenhum produto relacionado ao pedido " & Trim$(pedido) & "!!"
        End If
        
'   ENCONTROU ERRO ?
    If s_erro <> "" Then
        strMsgErro = s_erro
        GoSub OIPT_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        
'   PREENCHE DADOS DO DESTINATÁRIO DA NOTA DE VENDA
    endereco_comprador__nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))
    endereco_comprador__logradouro = UCase$(Trim$("" & t_DESTINATARIO("endereco")))
    endereco_comprador__numero = UCase$(Trim$("" & t_DESTINATARIO("endereco_numero")))
    endereco_comprador__complemento = UCase$(Trim$("" & t_DESTINATARIO("endereco_complemento")))
    s_endereco = endereco_comprador__logradouro
    If endereco_comprador__numero <> "" Then s_endereco = s_endereco & ", " & endereco_comprador__numero
    If endereco_comprador__complemento <> "" Then s_endereco = s_endereco & " " & endereco_comprador__complemento

'   BAIRRO
    endereco_comprador__bairro = UCase$(Trim$("" & t_DESTINATARIO("bairro")))

'   CEP
    endereco_comprador__cep = cep_formata(retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep"))))

'   CIDADE
    endereco_comprador__cidade = UCase$(Trim$("" & t_DESTINATARIO("cidade")))

'   UF
    endereco_comprador__uf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   CNPJ/CPF
    s_cnpj_cpf = Trim$("" & t_DESTINATARIO("cnpj_cpf"))
    endereco_comprador__cnpj_cpf = s_cnpj_cpf

'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    endereco_comprador__rg = s_ie_rg
        
    'preencher os campos de telefone
    strTelCel = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_cel"))))
    strTelRes = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res"))))
    strTelCom = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com"))))
    strTelCom2 = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2"))))
    If strTelCel <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_cel")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelCel = "(" & strDDD & ")" & strTelCel
        End If
    If strTelRes <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelRes = "(" & strDDD & ")" & strTelRes
        End If
    If strTelCom <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com")))
        If (Len(strDDD) = 2) Then strTelCom = "(" & strDDD & ") " & strTelCom
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom = strTelCom & " R: " & strRamal
        End If
    If strTelCom2 <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com_2")))
        If (Len(strDDD) = 2) Then strTelCom2 = "(" & strDDD & ") " & strTelCom2
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom2 = strTelCom2 & " R: " & strRamal
        End If

    
    s_end_linha_1 = s_endereco
    If (s_end_linha_1 <> "") And (endereco_comprador__bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & endereco_comprador__bairro
    
    s_end_linha_2 = endereco_recebedor__cidade
    If (s_end_linha_2 <> "") And (endereco_comprador__uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & endereco_comprador__uf
    If (s_end_linha_2 <> "") And (endereco_comprador__cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & endereco_comprador__cep
        
    s_end_linha_3 = ""
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PF Then
        strSufixoRes = "Tel Res: "
        strSufixoCom = "Tel Com: "
    Else
        strSufixoRes = "Tel: "
        strSufixoCom = "Tel: "
        End If
    If (strTelCel <> "") And (strTelRes <> "") Then s_end_linha_3 = strSufixoRes & strTelRes
    If ((strTelCel <> "") Or (strTelRes <> "")) And (strTelCom <> "") Then
        If s_end_linha_3 <> "" Then s_end_linha_3 = s_end_linha_3 & " / "
        s_end_linha_3 = s_end_linha_3 & strSufixoCom & strTelCom
        End If
    If ((strTelCel <> "") Or (strTelRes <> "") Or (strTelCom <> "")) And (strTelCom2 <> "") Then
        If s_end_linha_3 <> "" Then s_end_linha_3 = s_end_linha_3 & " / "
        s_end_linha_3 = s_end_linha_3 & strSufixoCom & strTelCom2
        End If

        
    If (s_end_linha_1 <> "") And ((s_end_linha_2 <> "") Or (s_end_linha_3 <> "")) Then s_end_linha_1 = s_end_linha_1 & vbCrLf
    If (s_end_linha_2 <> "") And (s_end_linha_3 <> "") Then s_end_linha_2 = s_end_linha_2 & vbCrLf
    
    s_info = s_nome & vbCrLf
    
    If s_cnpj_cpf <> "" Then s_info = s_info & "CNPJ/CPF: " & cnpj_cpf_formata(s_cnpj_cpf) & vbCrLf
    If s_ie_rg <> "" Then s_info = s_info & "IE/RG: " & s_ie_rg & vbCrLf
            
    s_info = s_info & _
             s_end_linha_1 & s_end_linha_2 & s_end_linha_3 & _
             s_end_entrega & vbCrLf & vbCrLf & _
             "OBSERVAÇÕES I" & vbCrLf & _
             s_obs_1

    
'   INFORMAÇÃO SE É CONTRIBUINTE DE ICMS
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        Select Case t_DESTINATARIO("contribuinte_icms_status")
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO: strInfoIE = "NC"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM: strInfoIE = "C"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO: strInfoIE = "I"
            Case Else: strInfoIE = ""
            End Select
    Else
        Select Case t_DESTINATARIO("produtor_rural_status")
            Case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM: strInfoIE = "PR"
            Case Else: strInfoIE = ""
            End Select
        End If
    l_IE.Caption = strInfoIE
    
    GoSub OIPT_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

    strResposta = s_info
    obtem_info_pedido_triangular = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPT_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIPT_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPT_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_DESTINATARIO, True
    Return
    
End Function

Function obtem_info_pedido_triangular_memorizada(ByVal pedido As String, ByRef strResposta As String, ByRef strEndEntregaFormatado As String, ByRef strEndEntregaUf As String, ByRef strTextoConstar As String, ByRef strEndClienteUf As String, ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_pedido_triangular_memorizada()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_nome As String
Dim s_cnpj_cpf As String
Dim s_ie_rg As String
Dim s_obs_1 As String
Dim s_info As String
Dim s_endereco As String
Dim s_end_linha_1 As String
Dim s_end_linha_2 As String
Dim s_end_linha_3 As String
Dim s_end_entrega As String
Dim pedido_a As String
Dim s_id_cliente As String
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strRamal As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim strInfoIE As String
Dim strNFeTextoConstar As String

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo OIPTM_TRATA_ERRO
    
    obtem_info_pedido_triangular_memorizada = False
    strMsgErro = ""
    strResposta = ""
    strEndEntregaFormatado = ""
    strEndEntregaUf = ""
    strEndClienteUf = ""
    l_IE = ""
    
    pedido = Trim$("" & pedido)
    pedido = normaliza_num_pedido(pedido)
    
    If pedido = "" Then
        strMsgErro = "Não foi informado o número do pedido!"
        Exit Function
        End If
        
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_DESTINATARIO (PODE SER T_CLIENTE OU T_LOJA)
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    endereco_comprador__nome = ""
    endereco_comprador__cnpj_cpf = ""
    endereco_comprador__rg = ""
    endereco_comprador__logradouro = ""
    endereco_comprador__numero = ""
    endereco_comprador__complemento = ""
    endereco_comprador__bairro = ""
    endereco_comprador__cep = ""
    endereco_comprador__cidade = ""
    endereco_comprador__uf = ""
    endereco_recebedor__nome = ""
    endereco_recebedor__cnpj_cpf = ""
    endereco_recebedor__rg = ""
    endereco_recebedor__logradouro = ""
    endereco_recebedor__numero = ""
    endereco_recebedor__complemento = ""
    endereco_recebedor__bairro = ""
    endereco_recebedor__cep = ""
    endereco_recebedor__cidade = ""
    endereco_recebedor__uf = ""
    'c_cnpj_cpf_dest
    s_endereco = ""
    s_nome = ""
    s_cnpj_cpf = ""
    s_ie_rg = ""
    s_obs_1 = ""
    s_end_entrega = ""
        

'   VERIFICA O PEDIDO
    s_id_cliente = ""
    pedido_a = ""
    s_erro = ""
    s = "SELECT" & _
            " pedido, st_entrega, id_cliente, obs_1, st_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep, NFe_texto_constar, " & _
            " EndEtg_nome as nome, EndEtg_cnpj_cpf as cnpj_cpf, EndEtg_tipo_pessoa as tipo_pessoa, EndEtg_ie as ie, EndEtg_rg as rg " & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & Trim$(pedido) & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Pedido " & Trim$(pedido) & " não está cadastrado !!"
    Else
'   TEXTO A CONSTAR NA NOTA FISCAL
    strNFeTextoConstar = Trim("" & t_PEDIDO("NFe_texto_constar"))
    
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA DE VENDA
    s = "SELECT" & _
            " pedido, id_cliente, st_memorizacao_completa_enderecos, endereco_uf as uf, endereco_cnpj_cpf as cnpj_cpf, " & _
            " endereco_logradouro as endereco, " & _
            " endereco_bairro as bairro, " & _
            " endereco_cidade as cidade, " & _
            " endereco_cep as cep, " & _
            " endereco_numero, " & _
            " endereco_complemento, " & _
            " endereco_email as email, endereco_email_xml as email_xml, " & _
            " endereco_nome as nome, " & _
            " endereco_ddd_res as ddd_res, endereco_tel_res as tel_res, " & _
            " endereco_ddd_com as ddd_com, endereco_tel_com as tel_com, endereco_ramal_com as ramal_com, " & _
            " endereco_ddd_cel as ddd_cel, endereco_tel_cel as tel_cel, " & _
            " endereco_ddd_com_2 as ddd_com_2, endereco_tel_com_2 as tel_com_2, endereco_ramal_com_2 as ramal_com_2, " & _
            " endereco_tipo_pessoa as tipo, " & _
            " endereco_contribuinte_icms_status as contribuinte_icms_status, " & _
            " endereco_produtor_rural_status as produtor_rural_status, " & _
            " endereco_ie as ie, " & _
            " endereco_rg as rg, " & _
            " endereco_contato as contato " & _
        " FROM t_PEDIDO" & _
        " WHERE (pedido = '" & Trim$(pedido) & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        strMsgErro = "Problemas na localização do endereço memorizado no pedido " & Trim$(pedido) & "!!"
        GoSub OIPTM_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
    
    '   ENDEREÇO DE ENTREGA
        endereco_recebedor__nome = UCase$(Trim("" & t_PEDIDO("nome")))
        endereco_recebedor__cnpj_cpf = Trim("" & t_PEDIDO("cnpj_cpf"))
        If Trim("" & t_PEDIDO("tipo_pessoa")) = "PJ" Then
            endereco_recebedor__rg = Trim("" & t_PEDIDO("ie"))
        Else
            endereco_recebedor__rg = Trim("" & t_PEDIDO("rg"))
            End If
        endereco_recebedor__logradouro = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco")))
        endereco_recebedor__numero = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_numero")))
        endereco_recebedor__complemento = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_complemento")))
        endereco_recebedor__bairro = UCase$(Trim("" & t_PEDIDO("EndEtg_bairro")))
        endereco_recebedor__cep = cep_formata(retorna_so_digitos(Trim("" & t_PEDIDO("EndEtg_cep"))))
        endereco_recebedor__cidade = UCase$(Trim("" & t_PEDIDO("EndEtg_cidade")))
        endereco_recebedor__uf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
        If (CLng(t_PEDIDO("st_end_entrega")) <> 0) And (endereco_recebedor__uf <> strEndClienteUf) Then
            s_end_entrega = formata_endereco(endereco_recebedor__logradouro, endereco_recebedor__numero, endereco_recebedor__complemento, endereco_recebedor__bairro, endereco_recebedor__cidade, endereco_recebedor__uf, endereco_recebedor__cep)
            strEndEntregaFormatado = s_end_entrega
            strEndEntregaUf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
            If s_end_entrega <> "" Then s_end_entrega = vbCrLf & "ENTREGA: " & s_end_entrega
        Else
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Operação triangular interestadual não caracterizada!!"
            End If
    
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        If Trim$("" & t_PEDIDO("obs_1")) <> "" Then
            If s_obs_1 <> "" Then s_obs_1 = s_obs_1 & vbCrLf
            s = Trim$("" & t_PEDIDO("obs_1"))
            s = substitui_caracteres(s, vbCr, " ")
            s = substitui_caracteres(s, vbLf, " ")
            s_obs_1 = s_obs_1 & s
            End If
        End If
    
    s = "SELECT pedido, fabricante, produto FROM t_PEDIDO_ITEM WHERE (pedido='" & Trim$(pedido) & "')"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    If t_PEDIDO_ITEM.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Não foi encontrado nenhum produto relacionado ao pedido " & Trim$(pedido) & "!!"
        End If
        
'   ENCONTROU ERRO ?
    If s_erro <> "" Then
        strMsgErro = s_erro
        GoSub OIPTM_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        
'   PREENCHE DADOS DO DESTINATÁRIO DA NOTA DE VENDA
    endereco_comprador__nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))
    endereco_comprador__logradouro = UCase$(Trim$("" & t_DESTINATARIO("endereco")))
    endereco_comprador__numero = UCase$(Trim$("" & t_DESTINATARIO("endereco_numero")))
    endereco_comprador__complemento = UCase$(Trim$("" & t_DESTINATARIO("endereco_complemento")))
    s_endereco = endereco_comprador__logradouro
    If endereco_comprador__numero <> "" Then s_endereco = s_endereco & ", " & endereco_comprador__numero
    If endereco_comprador__complemento <> "" Then s_endereco = s_endereco & " " & endereco_comprador__complemento

'   BAIRRO
    endereco_comprador__bairro = UCase$(Trim$("" & t_DESTINATARIO("bairro")))

'   CEP
    endereco_comprador__cep = cep_formata(retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep"))))

'   CIDADE
    endereco_comprador__cidade = UCase$(Trim$("" & t_DESTINATARIO("cidade")))

'   UF
    endereco_comprador__uf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   CNPJ/CPF
    s_cnpj_cpf = Trim$("" & t_DESTINATARIO("cnpj_cpf"))
    endereco_comprador__cnpj_cpf = s_cnpj_cpf

'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    endereco_comprador__rg = s_ie_rg
        
    'preencher os campos de telefone
    strTelCel = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_cel"))))
    strTelRes = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res"))))
    strTelCom = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com"))))
    strTelCom2 = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2"))))
    If strTelCel <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_cel")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelCel = "(" & strDDD & ")" & strTelCel
        End If
    If strTelRes <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelRes = "(" & strDDD & ")" & strTelRes
        End If
    If strTelCom <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com")))
        If (Len(strDDD) = 2) Then strTelCom = "(" & strDDD & ") " & strTelCom
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom = strTelCom & " R: " & strRamal
        End If
    If strTelCom2 <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com_2")))
        If (Len(strDDD) = 2) Then strTelCom2 = "(" & strDDD & ") " & strTelCom2
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom2 = strTelCom2 & " R: " & strRamal
        End If

    
    s_end_linha_1 = s_endereco
    If (s_end_linha_1 <> "") And (endereco_comprador__bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & endereco_comprador__bairro
    
    s_end_linha_2 = endereco_comprador__cidade
    If (s_end_linha_2 <> "") And (endereco_comprador__uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & endereco_comprador__uf
    If (s_end_linha_2 <> "") And (endereco_comprador__cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & endereco_comprador__cep
        
    s_end_linha_3 = ""
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PF Then
        strSufixoRes = "Tel Res: "
        strSufixoCom = "Tel Com: "
    Else
        strSufixoRes = "Tel: "
        strSufixoCom = "Tel: "
        End If
    If (strTelCel <> "") And (strTelRes <> "") Then s_end_linha_3 = strSufixoRes & strTelRes
    If ((strTelCel <> "") Or (strTelRes <> "")) And (strTelCom <> "") Then
        If s_end_linha_3 <> "" Then s_end_linha_3 = s_end_linha_3 & " / "
        s_end_linha_3 = s_end_linha_3 & strSufixoCom & strTelCom
        End If
    If ((strTelCel <> "") Or (strTelRes <> "") Or (strTelCom <> "")) And (strTelCom2 <> "") Then
        If s_end_linha_3 <> "" Then s_end_linha_3 = s_end_linha_3 & " / "
        s_end_linha_3 = s_end_linha_3 & strSufixoCom & strTelCom2
        End If

        
    If (s_end_linha_1 <> "") And ((s_end_linha_2 <> "") Or (s_end_linha_3 <> "")) Then s_end_linha_1 = s_end_linha_1 & vbCrLf
    If (s_end_linha_2 <> "") And (s_end_linha_3 <> "") Then s_end_linha_2 = s_end_linha_2 & vbCrLf
    
    s_info = s_nome & vbCrLf
    
    If s_cnpj_cpf <> "" Then s_info = s_info & "CNPJ/CPF: " & cnpj_cpf_formata(s_cnpj_cpf) & vbCrLf
    If s_ie_rg <> "" Then s_info = s_info & "IE/RG: " & s_ie_rg & vbCrLf
            
    s_info = s_info & _
             s_end_linha_1 & s_end_linha_2 & s_end_linha_3 & _
             s_end_entrega & vbCrLf & vbCrLf & _
             "OBSERVAÇÕES I" & vbCrLf & _
             s_obs_1

    
'   INFORMAÇÃO SE É CONTRIBUINTE DE ICMS
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        Select Case t_DESTINATARIO("contribuinte_icms_status")
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO: strInfoIE = "NC"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM: strInfoIE = "C"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO: strInfoIE = "I"
            Case Else: strInfoIE = ""
            End Select
    Else
        Select Case t_DESTINATARIO("produtor_rural_status")
            Case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM: strInfoIE = "PR"
            Case Else: strInfoIE = ""
            End Select
        End If
    l_IE.Caption = strInfoIE
    
    GoSub OIPTM_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

    strResposta = s_info
    obtem_info_pedido_triangular_memorizada = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPTM_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIPTM_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPTM_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_DESTINATARIO, True
    Return
    
End Function



Sub formulario_exibe_itens_pedido_triangular(ByVal pedido_selecionado As String)
Const NomeDestaRotina = "formulario_exibe_itens_pedido_triangular()"
Dim s As String
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim intIndice As Integer
Dim vl_unitario As Currency
Dim vl_total As Currency
Dim vl_total_geral As Currency
Dim intQtde As Integer
Dim lngTotalVolumes As Long
Dim n As Long


    On Error GoTo FEIPT_TRATA_ERRO
    
'   LIMPA OS CAMPOS
    formulario_limpa_campos_itens_pedido
    
    If Trim$(pedido_selecionado) = "" Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"

'   T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

'   T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   VERIFICA SE O PEDIDO ESTÁ CADASTRADO
    s = "SELECT" & _
            " pedido" & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & pedido_selecionado & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        aviso_erro "O pedido " & pedido_selecionado & " NÃO está cadastrado!!"
        GoSub FEIPT_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        c_pedido.SetFocus
        Exit Sub
        End If
    
'   OBTÉM OS ITENS DO PEDIDO
    s = "SELECT" & _
            " tPI.fabricante," & _
            " tPI.produto," & _
            " tPI.descricao," & _
            " tPI.qtde_volumes," & _
            " tPI.preco_NF," & _
            " tEI.ncm," & _
            " tEI.cst," & _
            " Sum(tEM.qtde) AS qtde"
    s = s & _
        " FROM t_PEDIDO_ITEM tPI" & _
            " INNER JOIN t_ESTOQUE_MOVIMENTO tEM ON (tPI.pedido=tEM.pedido) AND (tPI.fabricante=tEM.fabricante) AND (tPI.produto=tEM.produto)" & _
            " INNER JOIN t_ESTOQUE_ITEM tEI ON (tEM.id_estoque=tEI.id_estoque) AND (tEM.fabricante=tEI.fabricante) AND (tEM.produto=tEI.produto)"
    s = s & _
        " WHERE" & _
            " (tPI.pedido = '" & Trim$(pedido_selecionado) & "')" & _
            " AND (anulado_status=0)" & _
            " AND (estoque <> '" & ID_ESTOQUE_DEVOLUCAO & "')" & _
            " AND (preco_NF > 0)"
    s = s & _
        " GROUP BY" & _
            " tPI.fabricante," & _
            " tPI.produto," & _
            " tPI.descricao," & _
            " tPI.qtde_volumes," & _
            " tPI.preco_NF," & _
            " tEI.ncm," & _
            " tEI.cst"
    s = s & _
        " ORDER BY" & _
            " tPI.produto," & _
            " tEI.ncm," & _
            " tEI.cst"
    
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    intIndice = c_produto.LBound
    Do While Not t_PEDIDO_ITEM.EOF
    '   VERIFICA SE AINDA HÁ LINHAS DISPONÍVEIS
        If intIndice > c_produto.UBound Then
            aviso_erro "O pedido " & pedido_selecionado & " possui mais itens do que o permitido!!"
            GoSub FEIPT_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
            
        c_fabricante(intIndice) = Trim$("" & t_PEDIDO_ITEM("fabricante"))
        c_produto(intIndice) = Trim$("" & t_PEDIDO_ITEM("produto"))
        c_descricao(intIndice) = Trim$("" & t_PEDIDO_ITEM("descricao"))
        
        c_CST(intIndice) = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PEDIDO_ITEM("cst")))
        c_NCM(intIndice) = Trim$("" & t_PEDIDO_ITEM("ncm"))
        
        intQtde = t_PEDIDO_ITEM("qtde")
        c_qtde(intIndice) = CStr(intQtde)
        
        n = 0
        If IsNumeric(t_PEDIDO_ITEM("qtde_volumes")) Then n = CLng(t_PEDIDO_ITEM("qtde_volumes"))
        lngTotalVolumes = lngTotalVolumes + (n * intQtde)
        
        vl_unitario = t_PEDIDO_ITEM("preco_NF")
        c_vl_unitario(intIndice) = formata_moeda(vl_unitario)
        
        vl_total = intQtde * vl_unitario
        c_vl_total(intIndice) = formata_moeda(vl_total)
        
        vl_total_geral = vl_total_geral + vl_total
        
        intIndice = intIndice + 1
        t_PEDIDO_ITEM.MoveNext
        Loop
    
    c_vl_total_geral = formata_moeda(vl_total_geral)
    c_total_volumes = CStr(lngTotalVolumes)
    
    GoSub FEIPT_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FEIPT_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub FEIPT_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FEIPT_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    Return
    
End Sub

Function pedido_eh_do_emitente_atual(ByVal pedido_selecionado As String) As Boolean
    Const NomeDestaRotina = "pedido_eh_do_emitente_atual()"
    Dim s As String
    Dim s_cd As String
    Dim t_PEDIDO As ADODB.Recordset
    
    On Error GoTo PEDCA_TRATA_ERRO
    
    pedido_eh_do_emitente_atual = False
    
'   T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   VERIFICA SE O PEDIDO ESTÁ CADASTRADO
    s = "SELECT" & _
            " *" & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & pedido_selecionado & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        aviso_erro "O pedido " & pedido_selecionado & " NÃO está cadastrado!!"
        GoSub PEDCA_FECHA_TABELAS
        c_pedido.SetFocus
        Exit Function
        End If
    
'   VERIFICA SE PEDIDO PODE SER EMITIDO NO EMITENTE SELECIONADO
    If (usuario.emit_id <> Trim$("" & t_PEDIDO("id_nfe_emitente"))) Then
        aviso_erro "Pedido não pode ser emitido no Emitente atual (" & usuario.emit & ")!!"
            GoSub PEDCA_FECHA_TABELAS
            Exit Function
        End If
    
    pedido_eh_do_emitente_atual = True
    
    GoSub PEDCA_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PEDCA_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub PEDCA_FECHA_TABELAS
    aviso_erro s
    Exit Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PEDCA_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    Return
    
End Function

Function retorna_max_segundos_nota_triangular() As Long
Dim maxseg As Long
Dim s As String
Dim t As ADODB.Recordset

    On Error GoTo OMSNT_TRATA_ERRO
    
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT *" & _
        " FROM t_PARAMETRO" & _
        " WHERE" & _
            " (id = 'NF_MaxSegundos_EsperaNotaTriangular')"
    
    t.Open s, dbc, , , adCmdText
    If t.EOF Then
        GoSub OMSNT_FECHA_TABELAS
        maxseg = 0
    Else
        maxseg = t("campo_inteiro")
        End If

    GoSub OMSNT_FECHA_TABELAS
    
    retorna_max_segundos_nota_triangular = maxseg
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OMSNT_TRATA_ERRO:
'======================================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub OMSNT_FECHA_TABELAS
    aviso_erro s
    retorna_max_segundos_nota_triangular = 0
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OMSNT_FECHA_TABELAS:
'=======================
    bd_desaloca_recordset t, True
    Return
    
End Function


Sub gerencia_temmpo_limite_emissao()
Dim lngTempo As Long
Dim strTempoRestante As String
Dim strHoras As String
Dim strMinutos As String
Dim strSegundos As String
Dim s As String

    lngTempo = MaxSegundosNotaTriangular - DateDiff("s", dt_hr_inicio_emissao, Now)
    
    If Not blnEmissaoOK And (lngTempo > 0) Then
        strHoras = Trim(Str(lngTempo \ 3600))
        If Len(strHoras) <= 1 Then strHoras = "0" & strHoras
        lngTempo = lngTempo - (lngTempo \ 3600) * 3600
        strMinutos = Trim(Str(lngTempo \ 60))
        lngTempo = lngTempo - (lngTempo \ 60) * 60
        If Len(strMinutos) <= 1 Then strMinutos = "0" & strMinutos
        strSegundos = Trim(Str(lngTempo))
        If Len(strSegundos) <= 1 Then strSegundos = "0" & strSegundos
        strTempoRestante = strHoras & ":" & strMinutos & ":" & strSegundos
        infocontagem.Caption = "Tempo para emissão: " & vbCrLf & strTempoRestante
    Else
        infocontagem.Visible = False
        b_cancela_triangular.Visible = False
        b_cancela_triangular.Enabled = False
        c_pedido.Enabled = True
        contagem.Enabled = False
        If Not blnEmissaoOK Then aviso "Encerrado o tempo para a emissão triangular!!!"
        formulario_limpa
        c_pedido = ""
        If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_CANCELADA_TIMEOUT, s) Then
            aviso_erro s
            End If
        lngIdNFeTriangular = 0
        'se está atendendo à fila de impressão, fechar o formulário e retornar ao form principal
        'para dar sequência à fila
        If sPedidoTriangular <> "" Then
            fechar_modo_emissao_nfe_triangular
            End If
        End If

End Sub

Function insere_registro_nfe_triangular(ByRef IdNFeTriangular As Long, _
                                        ByRef NFserie As Long, _
                                        ByRef NFNumvenda As Long, _
                                        ByRef NFNumremessa As Long, _
                                        ByRef StrNumPedido As String, _
                                        ByVal blnPesquisarNumeracao As Boolean, _
                                        ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Gera o registro para emissão de nota fiscal triangular, que vai controlar a impressão
'|
'|- insere o registro na tabela t_NFe_TRIANGULAR sem preencher as numerações das notas (motivo: garantir numeração não repetida)
'|- se blnPesquisarNumeracao é verdadeiro
'|  - obtem o NSU valor x na t_FIN_CONTROLE
'|  - atribui x + 1 para a nota de venda e x + 2 para a nota de remessa
'|  - atualiza numerações das  notas na tabela t_NFe_TRIANGULAR
'|
'|  Parâmetros:
'|      IdNFeTriangular: id do registro de emissão de nota fiscal triangular
'|      NFserie: número de série das notas de venda e remessa
'|      NFNumvenda: número da nota de venda
'|      NFNumremessa: número da nota de remessa
'|      strNumPedido: número do pedido
'|      blnPesquisarNumeracao: indica se a numeração será obtida em t_FIN_CONTROLE
'|      strMsgErro: mensagem de erro ocorrido no processo, caso haja
'|
'|  Retorno da função:
'|      true: sucesso ao inserir o registro
'|      false: falha ao inserir o registro
'|

Dim lngRecordsAffected As Long
Dim t As ADODB.Recordset
Dim strSql As String
    
    On Error GoTo IRNT_TRATA_ERRO
    
    insere_registro_nfe_triangular = False
    strMsgErro = ""
    
'   ~~~~~~~~~~~~~~~~~
    dbc.BeginTrans
'   ~~~~~~~~~~~~~~~~~
    
    'criar o NSU_T_NFe_TRIANGULAR no projeto depois
    If Not geraNsu("t_NFe_TRIANGULAR", IdNFeTriangular, strMsgErro) Then
        If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
        strMsgErro = "Falha ao criar o registro para geração de NSU da nota triangular!!" & strMsgErro
        'GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
        '   ~~~~~~~~~~~~~~~~~
            dbc.RollbackTrans
        '   ~~~~~~~~~~~~~~~~~
        aviso_erro strMsgErro
        Exit Function
        End If
        
'   SE HAVERÁ PESQUISA DA NUMERAÇÃO, ZERAR VARIÁVEIS
    If blnPesquisarNumeracao Then
        NFNumvenda = 0
        NFNumremessa = 0
        End If
        
'   NÃO ESTÁ CADASTRADO, ENTÃO CADASTRA AGORA
    strSql = "INSERT INTO t_NFe_TRIANGULAR (" & _
                "id, id_nfe_emitente, Nfe_serie_venda, Nfe_numero_venda, Nfe_serie_remessa, Nfe_numero_remessa, " & _
                "pedido, Nfe_venda_emissao_status, Nfe_venda_emissao_data_hora, Nfe_venda_emissao_usuario, " & _
                "emissao_status, usuario_emissao_status, usuario_cadastro" & _
            ") VALUES (" & _
                "'" & CStr(IdNFeTriangular) & "'," & _
                usuario.emit_id & ", " & _
                CStr(NFserie) & "," & _
                CStr(NFNumvenda) & "," & _
                CStr(NFserie) & "," & _
                CStr(NFNumremessa) & "," & _
                "'" & StrNumPedido & "'," & _
                CStr(ST_NFT_EM_PROCESSAMENTO) & "," & _
                "GETDATE()," & _
                "'" & usuario.id & "'," & _
                CStr(ST_NFT_EM_PROCESSAMENTO) & "," & _
                "'" & usuario.id & "'," & _
                "'" & usuario.id & "'" & _
            ")"
    Call dbc.Execute(strSql, lngRecordsAffected)
    If lngRecordsAffected <> 1 Then
        strMsgErro = "Falha ao criar o registro para emissão de nota triangular!!"
    '   ~~~~~~~~~~~~~~~~~
        dbc.RollbackTrans
    '   ~~~~~~~~~~~~~~~~~
        Exit Function
        End If
        
'   ~~~~~~~~~~~~~~~~~
    dbc.CommitTrans
'   ~~~~~~~~~~~~~~~~~
        
    Sleep 200
    
    If blnPesquisarNumeracao Then
    '   RECORDSET
        Set t = New ADODB.Recordset
        With t
            .CursorType = BD_CURSOR_SOMENTE_LEITURA
            .LockType = BD_POLITICA_LOCKING
            .CacheSize = BD_CACHE_CONSULTA
            End With
    
        strSql = "SELECT" & _
                    " n.NFe_serie_NF," & _
                    " n.NFe_numero_NF" & _
                " FROM t_NFE_EMITENTE e" & _
                " INNER JOIN t_NFE_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" & _
                " WHERE" & _
                    " (e.id=" & usuario.emit_id & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open strSql, dbc, , , adCmdText
        If Not t.EOF Then
            NFserie = t("NFe_serie_NF")
            NFNumvenda = t("NFe_numero_NF") + 1
            NFNumremessa = t("NFe_numero_NF") + 2
            End If
    
    '   TENTA ATUALIZAR O BANCO DE DADOS
        strSql = "UPDATE t_NFe_TRIANGULAR SET" & _
                    " Nfe_serie_venda = " & CStr(NFserie) & ", " & _
                    " Nfe_numero_venda = " & CStr(NFNumvenda) & ", " & _
                    " Nfe_serie_remessa = " & CStr(NFserie) & ", " & _
                    " Nfe_numero_remessa = " & CStr(NFNumremessa) & _
                " WHERE" & _
                    " (id = " & CStr(IdNFeTriangular) & ")"
        Call dbc.Execute(strSql, lngRecordsAffected)
        If lngRecordsAffected <> 1 Then
            strMsgErro = "Falha ao atualizar o registro de nota fiscal triangular!!"
            GoSub IRNT_FECHA_TABELAS
            Exit Function
            End If
        End If
    
    insere_registro_nfe_triangular = True
    
    GoSub IRNT_FECHA_TABELAS
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IRNT_TRATA_ERRO:
'~~~~~~~~~~~~~
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    GoSub IRNT_FECHA_TABELAS
    Exit Function
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IRNT_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function

Function atualiza_nfe_triangular_geral(ByRef IdNFeTriangular As Long, ByRef NFeStatus As Integer, ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Atualiza o o status geral do registro para emissão de nota fiscal triangular
'|
'|  Parâmetros:
'|      IdNFeTriangular: id do registro de emissão de nota fiscal triangular
'|      NFeStatus: novo status geral do registro
'|      strMsgErro: mensagem de erro ocorrido no processo, caso haja
'|
'|  Retorno da função:
'|      true: sucesso ao atualizar o registro
'|      false: falha ao atualizar o registro
'|

Dim lngRecordsAffected As Long
Dim strSql As String
    
    On Error GoTo ANTG_TRATA_ERRO
    
    atualiza_nfe_triangular_geral = False
    strMsgErro = ""
    
'   TENTA ATUALIZAR O BANCO DE DADOS
    strSql = "UPDATE t_NFe_TRIANGULAR SET" & _
                " emissao_status = " & CStr(NFeStatus) & ", " & _
                " dt_hr_status = getdate(), " & _
                " usuario_emissao_status = '" & usuario.id & "'" & _
            " WHERE" & _
                " (id = " & CStr(IdNFeTriangular) & ")"
    Call dbc.Execute(strSql, lngRecordsAffected)
    If lngRecordsAffected <> 1 Then
        strMsgErro = "Falha ao atualizar o registro de nota fiscal triangular!!"
        Exit Function
        End If
    
    atualiza_nfe_triangular_geral = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ANTG_TRATA_ERRO:
'~~~~~~~~~~~~~~~
    strMsgErro = "Falha na atualização da nota fiscal triangular!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    Exit Function
    
    
End Function

Function atualiza_nfe_triangular_venda(ByRef IdNFeTriangular As Long, _
                                        ByRef NFeStatus As Integer, _
                                        ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Atualiza o o status da nota de venda do registro para emissão de nota fiscal triangular
'|
'|  Parâmetros:
'|      IdNFeTriangular: id do registro de emissão de nota fiscal triangular
'|      NFeStatus: novo status da nota de venda do registro
'|      strMsgErro: mensagem de erro ocorrido no processo, caso haja
'|
'|  Retorno da função:
'|      true: sucesso ao atualizar o registro
'|      false: falha ao atualizar o registro
'|

Dim lngRecordsAffected As Long
Dim strSql As String
    
    On Error GoTo ANTV_TRATA_ERRO
    
    atualiza_nfe_triangular_venda = False
    strMsgErro = ""
    
'   TENTA ATUALIZAR O BANCO DE DADOS
    strSql = "UPDATE t_NFe_TRIANGULAR SET" & _
                " Nfe_venda_emissao_status = " & CStr(NFeStatus) & ", " & _
                " Nfe_venda_emissao_data_hora = getdate(), " & _
                " Nfe_venda_emissao_usuario = '" & usuario.id & "'" & _
            " WHERE" & _
                " (id = " & CStr(IdNFeTriangular) & ")"
    Call dbc.Execute(strSql, lngRecordsAffected)
    If lngRecordsAffected <> 1 Then
        strMsgErro = "Falha ao atualizar o registro de nota fiscal triangular!!"
        Exit Function
        End If
    
    atualiza_nfe_triangular_venda = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ANTV_TRATA_ERRO:
'~~~~~~~~~~~~~~~
    strMsgErro = "Falha na atualização da nota fiscal triangular!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    Exit Function
    
    
End Function

Function atualiza_nfe_triangular_remessa(ByRef IdNFeTriangular As Long, ByRef NFeStatus As Integer, ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Atualiza o o status da nota de remessa do registro para emissão de nota fiscal triangular
'|
'|  Parâmetros:
'|      IdNFeTriangular: id do registro de emissão de nota fiscal triangular
'|      NFeStatus: novo status da nota de remessa do registro
'|      strMsgErro: mensagem de erro ocorrido no processo, caso haja
'|
'|  Retorno da função:
'|      true: sucesso ao atualizar o registro
'|      false: falha ao atualizar o registro
'|

Dim lngRecordsAffected As Long
Dim strSql As String
    
    On Error GoTo ANTR_TRATA_ERRO
    
    atualiza_nfe_triangular_remessa = False
    strMsgErro = ""
    
'   TENTA ATUALIZAR O BANCO DE DADOS
    strSql = "UPDATE t_NFe_TRIANGULAR SET" & _
                " Nfe_remessa_emissao_status = " & CStr(NFeStatus) & ", " & _
                " Nfe_remessa_emissao_data_hora = getdate(), " & _
                " Nfe_remessa_emissao_usuario = '" & usuario.id & "'" & _
            " WHERE" & _
                " (id = " & CStr(IdNFeTriangular) & ")"
    Call dbc.Execute(strSql, lngRecordsAffected)
    If lngRecordsAffected <> 1 Then
        strMsgErro = "Falha ao atualizar o registro de nota fiscal triangular!!"
        Exit Function
        End If
    
    atualiza_nfe_triangular_remessa = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ANTR_TRATA_ERRO:
'~~~~~~~~~~~~~~~
    strMsgErro = "Falha na atualização da nota fiscal triangular!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    Exit Function
    
    
End Function

Function atualiza_nfe_triangular_inf_adicionais(ByRef IdNFeTriangular As Long, _
                                        ByVal recebedor_cnpj_cpf As String, _
                                        ByVal recebedor_nome As String, _
                                        ByVal recebedor_rg As String, _
                                        ByVal recebedor_endereco As String, _
                                        ByVal recebedor_numero As String, _
                                        ByVal recebedor_complemento As String, _
                                        ByVal recebedor_bairro As String, _
                                        ByVal recebedor_cidade As String, _
                                        ByVal recebedor_uf As String, _
                                        ByVal recebedor_cep As String, _
                                        ByVal infAdic_venda As String, _
                                        ByVal infAdic_remessa As String, _
                                        ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Atualiza o o status da nota de venda do registro para emissão de nota fiscal triangular
'|
'|  Parâmetros:
'|      IdNFeTriangular: id do registro de emissão de nota fiscal triangular
'|      recebedor_cnpj_cpf: CNPJ/CPF do recebedor da mercadoria
'|      recebedor_nome: Nome do recebedor da mercadoria
'|      recebedor_rg: RG do recebedor da mercadoria
'|      recebedor_endereco: logradouro do endereço de entrega da mercadoria
'|      recebedor_numero: número do endereço de entrega da mercadoria
'|      recebedor_complemento: complemento do endereço de entrega da mercadoria
'|      recebedor_bairro: bairro do endereço de entrega da mercadoria
'|      recebedor_cidade: cidade do endereço de entrega da mercadoria
'|      recebedor_uf: uf do endereço de entrega da mercadoria
'|      recebedor_cep: cep do endereço de entrega da mercadoria
'|      infAdic_venda: Informações adicionais presentes na nota fiscal de venda
'|      infAdic_remessa: Informações adicionais presentes na nota fiscal de remessa
'|      strMsgErro: mensagem de erro ocorrido no processo, caso haja
'|
'|  Retorno da função:
'|      true: sucesso ao atualizar o registro
'|      false: falha ao atualizar o registro
'|

Dim lngRecordsAffected As Long
Dim strSql As String
Dim strAtualizacoes As String
    
    On Error GoTo ANTIA_TRATA_ERRO
    
    atualiza_nfe_triangular_inf_adicionais = False
    strMsgErro = ""
    
    strAtualizacoes = ""
    If Trim(recebedor_cnpj_cpf) <> "" Then
        strAtualizacoes = strAtualizacoes & " recebedor_cnpj_cpf = '" & Trim(recebedor_cnpj_cpf) & "'"
        End If
    If Trim(recebedor_nome) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_nome = '" & bd_filtra_aspas(Trim(recebedor_nome)) & "'"
        End If
    If Trim(recebedor_rg) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_rg = '" & Trim(recebedor_rg) & "'"
        End If
    If Trim(recebedor_endereco) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_endereco = '" & bd_filtra_aspas(Trim(recebedor_endereco)) & "'"
        End If
    If Trim(recebedor_numero) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_numero = '" & Trim(recebedor_numero) & "'"
        End If
    If Trim(recebedor_complemento) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_complemento = '" & bd_filtra_aspas(Trim(recebedor_complemento)) & "'"
        End If
    If Trim(recebedor_bairro) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_bairro = '" & bd_filtra_aspas(Trim(recebedor_bairro)) & "'"
        End If
    If Trim(recebedor_cidade) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_cidade = '" & bd_filtra_aspas(Trim(recebedor_cidade)) & "'"
        End If
    If Trim(recebedor_uf) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_uf = '" & Trim(recebedor_uf) & "'"
        End If
    If Trim(recebedor_cep) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " recebedor_cep = '" & retorna_so_digitos(recebedor_cep) & "'"
        End If
    If Trim(infAdic_venda) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " infAdic_venda = '" & bd_filtra_aspas(Trim(infAdic_venda)) & "'"
        End If
    If Trim(infAdic_remessa) <> "" Then
        If strAtualizacoes <> "" Then strAtualizacoes = strAtualizacoes & ","
        strAtualizacoes = strAtualizacoes & " infAdic_remessa = '" & bd_filtra_aspas(Trim(infAdic_remessa)) & "'"
        End If
    
'   TENTA ATUALIZAR O BANCO DE DADOS
    If strAtualizacoes <> "" Then
        strSql = "UPDATE t_NFe_TRIANGULAR SET" & _
                    strAtualizacoes & _
                " WHERE" & _
                    " (id = " & CStr(IdNFeTriangular) & ")"
        Call dbc.Execute(strSql, lngRecordsAffected)
        If lngRecordsAffected <> 1 Then
            strMsgErro = "Falha ao atualizar o registro de nota fiscal triangular!!"
            Exit Function
            End If
        End If
        
    atualiza_nfe_triangular_inf_adicionais = True
    
Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ANTIA_TRATA_ERRO:
'~~~~~~~~~~~~~~~
    strMsgErro = "Falha na atualização complementar da nota fiscal triangular!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    Exit Function
    
    
End Function

Function obtem_dados_adicionais_venda(ByRef IdNFeTriangular As Long) As String
Dim t As ADODB.Recordset
Dim strSql As String
Dim s As String
Dim s_num_remessa As String

    On Error GoTo ODAV_TRATA_ERRO
    
'   RECORDSET
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    strSql = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (id=" & CStr(IdNFeTriangular) & ")"
    If t.State <> adStateClosed Then t.Close
    t.Open strSql, dbc, , , adCmdText
    s = ""
    s_num_remessa = "???"
    If (Not t.EOF) Then
        'EFETUAR CARREGAMENTO SE A NOTA DE VENDA FOI EMITIDA
        s_num_remessa = t("Nfe_numero_remessa")
        If (t("Nfe_venda_emissao_status") = ST_NFT_EMITIDA) Then s = Trim("" & t("infAdic_venda"))
        End If
    
    'se nota de venda não foi emitida, preencher o texto com base nos campos preenchidos em tela
    '(manter informações adicionais em branco se o CNPJ/CPF não estiver preenchido)
    If (s = "") And (retorna_so_digitos(c_cnpj_cpf_dest) <> "") Then
        s = "MERCADORIA SERA ENTREGUE POR CONTA E ORDEM A: " & vbCrLf
        s = s & c_nome_dest & vbCrLf
        s = s & IIf(Len(retorna_so_digitos(c_cnpj_cpf_dest)) = 14, "CNPJ: ", "CPF: ") & c_cnpj_cpf_dest
        If Trim$(c_rg_dest) <> "" Then s = s & " / IE/RG: " & c_rg_dest
        s = s & vbCrLf
        s = s & formata_endereco(l_end_recebedor_logradouro, l_end_recebedor_numero, l_end_recebedor_complemento, l_end_recebedor_bairro, l_end_recebedor_cidade, l_end_recebedor_uf, l_end_recebedor_cep) & vbCrLf
        s = s & "ATRAVES DA NOSSA NOTA FISCAL DE REMESSA No " & s_num_remessa & " EMITIDA EM " & Format$(Date, FORMATO_DATA) & vbCrLf
        s = s & "EMISSAO NOS TERMOS DO ARTIGO 129, DO RICMS/SP/2000" & vbCrLf
        End If
    
    obtem_dados_adicionais_venda = s
    
    GoSub ODAV_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODAV_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    GoSub ODAV_FECHA_TABELAS
    Exit Function
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODAV_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function

Function obtem_dados_adicionais_remessa(ByRef IdNFeTriangular As Long) As String
Dim t As ADODB.Recordset
Dim strSql As String
Dim s As String
Dim s_num_venda As String
Dim s_dt_emissao_nota_venda As String

    On Error GoTo ODAR_TRATA_ERRO
    
    s = ""
    
'   RECORDSET
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    strSql = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (id=" & CStr(IdNFeTriangular) & ")"
    If t.State <> adStateClosed Then t.Close
    t.Open strSql, dbc, , , adCmdText
    s_num_venda = "???"
    s_dt_emissao_nota_venda = Format$(Date, FORMATO_DATA)
    If Not t.EOF Then
        If t("Nfe_remessa_emissao_status") = ST_NFT_EMITIDA Then
            'EFETUAR CARREGAMENTO SE A NOTA DE REMESSA FOI EMITIDA
            s = Trim("" & t("infAdic_remessa"))
        Else
            s_num_venda = t("Nfe_numero_venda")
            s_dt_emissao_nota_venda = Format$(t("Nfe_venda_emissao_data_hora"), FORMATO_DATA)
            End If
        End If
        
    
    'se a nota de remessa não foi emitida anteriormente, utilizar os dados do recebedor para
    'preencher as informações adicionais
    If s = "" Then
        s = "MERCADORIA REMETIDA DIRETAMENTE POR CONTA E ORDEM DE: " & vbCrLf
        s = s & l_comprador_nome & vbCrLf
        s = s & formata_endereco(l_end_comprador_logradouro, l_end_comprador_numero, endereco_comprador__complemento, endereco_comprador__bairro, endereco_comprador__cidade, endereco_comprador__uf, endereco_comprador__cep) & vbCrLf
        s = s & "ATRAVES DA NOSSA NOTA FISCAL DE VENDA No " & s_num_venda & " EMITIDA EM " & s_dt_emissao_nota_venda & vbCrLf
        s = s & "OS IMPOSTOS SOB ESTA NOTA, FORAM DESTACADOS EM NOSSA NOTA FISCAL DE VENDA No " & s_num_venda & " EMITIDA EM " & s_dt_emissao_nota_venda & vbCrLf
        s = s & "EMISSAO NOS TERMOS DO ARTIGO 129, DO RICMS/SP/2000" & vbCrLf
        End If
        
        
    obtem_dados_adicionais_remessa = s
    
    GoSub ODAR_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODAR_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    GoSub ODAR_FECHA_TABELAS
    Exit Function
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODAR_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function

Function obtem_dados_nf_venda(ByVal NumNFVenda As Long, _
                                        ByVal SerieNFVenda As Long, _
                                        ByRef DestinoOp As String, _
                                        ByRef FretePor As String, _
                                        ByRef NatOp As String) As Boolean
Dim t As ADODB.Recordset
Dim strSql As String
Dim s As String

    On Error GoTo ODNFV_TRATA_ERRO
    
    obtem_dados_nf_venda = False
    
    DestinoOp = ""
    FretePor = ""
    NatOp = ""
    
'   RECORDSET
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    strSql = "SELECT" & _
                " *" & _
            " FROM t_NFE_IMAGEM" & _
            " WHERE" & _
                " (NFe_numero_NF=" & CStr(NumNFVenda) & ")" & _
                " AND (NFe_serie_NF=" & CStr(SerieNFVenda) & ")" & _
            " AND (st_anulado = 0) " & _
            " ORDER BY id DESC"
    If t.State <> adStateClosed Then t.Close
    t.Open strSql, dbc, , , adCmdText
    If Not t.EOF Then
        DestinoOp = Trim("" & t("ide__idDest"))
        FretePor = Trim("" & t("transp__modFrete"))
        End If
        
    strSql = "SELECT" & _
                " *" & _
            " FROM t_NFE_EMISSAO" & _
            " WHERE" & _
                " (NFe_numero_NF=" & CStr(NumNFVenda) & ")" & _
                " AND (NFe_serie_NF=" & CStr(SerieNFVenda) & ")" & _
            " AND (st_anulado = 0) " & _
            " ORDER BY id DESC"
    If t.State <> adStateClosed Then t.Close
    t.Open strSql, dbc, , , adCmdText
    If Not t.EOF Then
        NatOp = Trim("" & t("natureza_operacao_codigo"))
        End If
        
    
        
    obtem_dados_nf_venda = True
    
    GoSub ODNFV_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODNFV_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    GoSub ODNFV_FECHA_TABELAS
    Exit Function
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODNFV_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function


Sub DANFE_CONSULTA_parametro_emitente(ByVal relacaoPedidos As String)

' CONSTANTES
Const NomeDestaRotina = "DANFE_consulta_parametro_emitente()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_alerta_erro As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim strNFeMsgRetornoSP As String
Dim strPedido As String

Dim i As Integer
Dim j As Integer
Dim ic As Integer
Dim qtde_pedidos As Integer
Dim intIdNfeEmitente As Integer
Dim lngNFeSerieNF As Long
Dim lngNFeNumeroNF As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

Dim blnOperacaoNaoTriangular As Boolean

' VETORES
Dim v() As String
Dim v_pedido() As String
Dim v_danfe() As String

' BANCO DE DADOS
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim t_NFE_TRIANGULAR As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim dbcNFe As ADODB.Connection

    On Error GoTo DANFE_CONSULTA_PARAM_EMITENTE_TRATA_ERRO
    
    relacaoPedidos = normaliza_lista_pedidos(relacaoPedidos)
    
    ReDim v_danfe(0)
    v_danfe(UBound(v_danfe)) = ""
    
    ReDim v_pedido(0)
    v_pedido(UBound(v_pedido)) = ""
        
    qtde_pedidos = 0
    
    v = Split(relacaoPedidos, vbCrLf)
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
        '   REPETIDO ?
            For j = LBound(v_pedido) To UBound(v_pedido)
                If Trim$(v(i)) = v_pedido(j) Then
                    aviso_erro "Pedido " & Trim$(v(i)) & " está repetido na lista!!"
                    c_pedido_danfe.SetFocus
                    Exit Sub
                    End If
                Next
                
            If v_pedido(UBound(v_pedido)) <> "" Then ReDim Preserve v_pedido(UBound(v_pedido) + 1)
            v_pedido(UBound(v_pedido)) = Trim$(v(i))
            qtde_pedidos = qtde_pedidos + 1
            End If
        Next
    
    If qtde_pedidos = 0 Then
        aviso_erro "Informe o número do pedido!!"
        c_pedido_danfe.SetFocus
        Exit Sub
        End If
    
  ' t_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
  
  ' T_NFE_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
  ' T_NFE_TRIANGULAR
    Set t_NFE_TRIANGULAR = New ADODB.Recordset
    With t_NFE_TRIANGULAR
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
'   CONEXÃO AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    blnOperacaoNaoTriangular = True
    
'----------------------------------------------------------------------------------
'INÍCIO DO TRECHO PARA EMISSÃO DE DANFE's RELACIONADAS A OPERAÇÕES TRIANGULARES
'----------------------------------------------------------------------------------
    If blnNotaTriangularAtiva Then
'   PARA CADA PEDIDO DA LISTA, OBTÉM E EXIBE A DANFE
        s_alerta_erro = ""
        For ic = LBound(v_pedido) To UBound(v_pedido)
            strPedido = Trim$(v_pedido(ic))
            If strPedido <> "" Then
                aguarde INFO_EXECUTANDO, "consultando situação da NFe"
                
                s = "SELECT" & _
                        " id_nfe_emitente," & _
                        " NFe_serie_venda," & _
                        " NFe_numero_venda," & _
                        " NFe_serie_remessa," & _
                        " NFe_numero_remessa" & _
                    " FROM t_NFe_TRIANGULAR" & _
                    " WHERE" & _
                        " (pedido = '" & strPedido & "')" & _
                        " AND emissao_status in (" & CStr(ST_NFT_EM_PROCESSAMENTO) & ", " & CStr(ST_NFT_EMITIDA) & ")" & _
                    " ORDER BY" & _
                        " id DESC"
                If t_NFE_TRIANGULAR.State <> adStateClosed Then t_NFE_TRIANGULAR.Close
                t_NFE_TRIANGULAR.Open s, dbc, , , adCmdText
                If t_NFE_TRIANGULAR.EOF Then
                    'If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                    's_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": não foi localizada nenhuma NFe Triangular emitida!!"
                    GoTo PROXIMO_PEDIDO_TRI
                    End If
                    
                blnOperacaoNaoTriangular = False
                
                intIdNfeEmitente = t_NFE_TRIANGULAR("id_nfe_emitente")
                
                s = "SELECT" & _
                        " razao_social," & _
                        " NFe_T1_servidor_BD," & _
                        " NFe_T1_nome_BD," & _
                        " NFe_T1_usuario_BD," & _
                        " NFe_T1_senha_BD" & _
                    " FROM t_NFE_EMITENTE" & _
                    " WHERE" & _
                        " (id = " & CStr(intIdNfeEmitente) & ")"
                If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
                t_NFE_EMITENTE.Open s, dbc, , , adCmdText
                If t_NFE_EMITENTE.EOF Then
                    If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                    s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao localizar o registro em t_NFE_EMITENTE (id=" & CStr(intIdNfeEmitente) & ")!!"
                    GoTo PROXIMO_PEDIDO_TRI
                    End If
                    
                strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
                strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
                strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
                strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
                strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
                
                decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
                s = "Provider=" & BD_OLEDB_PROVIDER & _
                    ";Data Source=" & strNfeT1ServidorBd & _
                    ";Initial Catalog=" & strNfeT1NomeBd & _
                    ";User Id=" & strNfeT1UsuarioBd & _
                    ";Password=" & s_aux
                If dbcNFe.State <> adStateClosed Then dbcNFe.Close
                dbcNFe.Open s
                
            '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
                Set cmdNFeSituacao.ActiveConnection = dbcNFe
                
                Do While Not t_NFE_TRIANGULAR.EOF
                    
                    lngNFeSerieNF = t_NFE_TRIANGULAR("NFe_serie_venda")
                    lngNFeNumeroNF = t_NFE_TRIANGULAR("NFe_numero_venda")
                    
                    strNumeroNfNormalizado = NFeFormataNumeroNF(lngNFeNumeroNF)
                    strSerieNfNormalizado = NFeFormataSerieNF(lngNFeSerieNF)
                    
                    'Emissão da nota de venda
                    If (strNumeroNfNormalizado <> "") And _
                        confirma("Confirma a consulta da nota de VENDA nº " & strNumeroNfNormalizado & "?") Then
                    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
                        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
                        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
                        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
                        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
                        strNFeMsgRetornoSP = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
                        
                        If intNfeRetornoSP <> 1 Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": não é possível consultar a DANFE nº " & strNumeroNfNormalizado & " porque a NFe encontra-se na seguinte situação: " & CStr(intNfeRetornoSP) & " - " & strNFeMsgRetornoSP
                            GoTo PROXIMA_NFE_TRI
                            End If
                                        
                        aguarde INFO_EXECUTANDO, "copiando PDF do DANFE"
                        Set cmdNFeDanfe.ActiveConnection = dbcNFe
                        cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
                        cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
                        Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
                        If rsNFeRetornoSPDanfe.EOF Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": o conteúdo da DANFE nº " & strNumeroNfNormalizado & " não foi encontrado!!"
                            GoTo PROXIMA_NFE_TRI
                            End If
                        
                        strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & "_" & strPedido & ".pdf"
                        strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
                        
                        If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                            If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                                s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                                GoTo PROXIMA_NFE_TRI
                                End If
                            End If
                        
                        strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                        If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                            If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                                s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                                GoTo PROXIMA_NFE_TRI
                                End If
                            End If
                        
                        lFileHandle = FreeFile
                        Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                        lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                        lngOffset = 0
                        Do While lngOffset < lngFileSize
                            bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                            Put #lFileHandle, , bytFile()
                            lngOffset = lngOffset + CHUNK_SIZE
                            Loop
                        
                        If v_danfe(UBound(v_danfe)) <> "" Then ReDim Preserve v_danfe(UBound(v_danfe) + 1)
                        v_danfe(UBound(v_danfe)) = strNomeArqCompletoDanfe
                        
                        Close #lFileHandle
                        End If
                    
                    lngNFeSerieNF = t_NFE_TRIANGULAR("NFe_serie_remessa")
                    lngNFeNumeroNF = t_NFE_TRIANGULAR("NFe_numero_remessa")
                    
                    strNumeroNfNormalizado = NFeFormataNumeroNF(lngNFeNumeroNF)
                    strSerieNfNormalizado = NFeFormataSerieNF(lngNFeSerieNF)
                    
                    'Emissão da nota de remessa
                    If (strNumeroNfNormalizado <> "") And _
                        confirma("Confirma a consulta da nota de REMESSA nº " & strNumeroNfNormalizado & "?") Then
                    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
                        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
                        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
                        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
                        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
                        strNFeMsgRetornoSP = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
                        
                        If intNfeRetornoSP <> 1 Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": não é possível consultar a DANFE nº " & strNumeroNfNormalizado & " porque a NFe encontra-se na seguinte situação: " & CStr(intNfeRetornoSP) & " - " & strNFeMsgRetornoSP
                            GoTo PROXIMA_NFE_TRI
                            End If
                                        
                        aguarde INFO_EXECUTANDO, "copiando PDF do DANFE"
                        Set cmdNFeDanfe.ActiveConnection = dbcNFe
                        cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
                        cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
                        Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
                        If rsNFeRetornoSPDanfe.EOF Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": o conteúdo da DANFE nº " & strNumeroNfNormalizado & " não foi encontrado!!"
                            GoTo PROXIMA_NFE_TRI
                            End If
                        
                        strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & "_" & strPedido & ".pdf"
                        strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
                        
                        If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                            If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                                s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                                GoTo PROXIMA_NFE_TRI
                                End If
                            End If
                        
                        strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                        If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                            If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                                s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                                GoTo PROXIMA_NFE_TRI
                                End If
                            End If
                        
                        lFileHandle = FreeFile
                        Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                        lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                        lngOffset = 0
                        Do While lngOffset < lngFileSize
                            bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                            Put #lFileHandle, , bytFile()
                            lngOffset = lngOffset + CHUNK_SIZE
                            Loop
                        
                        If v_danfe(UBound(v_danfe)) <> "" Then ReDim Preserve v_danfe(UBound(v_danfe) + 1)
                        v_danfe(UBound(v_danfe)) = strNomeArqCompletoDanfe
                        
                        Close #lFileHandle
                        End If
                
PROXIMA_NFE_TRI:
'===============
                    t_NFE_TRIANGULAR.MoveNext
                    Loop
                End If
                
PROXIMO_PEDIDO_TRI:
'==================
            Next
        End If
'----------------------------------------------------------------------------------
'FIM DO TRECHO PARA EMISSÃO DE DANFE's RELACIONADAS A OPERAÇÕES TRIANGULARES
'----------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
'INÍCIO DO TRECHO PARA EMISSÃO DE DANFE's RELACIONADAS A OPERAÇÕES NÃO TRIANGULARES
'----------------------------------------------------------------------------------
    If blnOperacaoNaoTriangular Then
    '   PARA CADA PEDIDO DA LISTA, OBTÉM E EXIBE A DANFE
        s_alerta_erro = ""
        For ic = LBound(v_pedido) To UBound(v_pedido)
            strPedido = Trim$(v_pedido(ic))
            If strPedido <> "" Then
                aguarde INFO_EXECUTANDO, "consultando situação da NFe"
                
                s = "SELECT" & _
                        " id_nfe_emitente," & _
                        " NFe_serie_NF," & _
                        " NFe_numero_NF" & _
                    " FROM t_NFe_EMISSAO" & _
                    " WHERE" & _
                        " (pedido = '" & strPedido & "')" & _
                    " ORDER BY" & _
                        " id DESC"
                If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
                t_NFe_EMISSAO.Open s, dbc, , , adCmdText
                If t_NFe_EMISSAO.EOF Then
                    If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                    s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": não foi localizada nenhuma NFe emitida!!"
                    GoTo PROXIMO_PEDIDO
                    End If
                    
                intIdNfeEmitente = t_NFe_EMISSAO("id_nfe_emitente")
                
                s = "SELECT" & _
                        " razao_social," & _
                        " NFe_T1_servidor_BD," & _
                        " NFe_T1_nome_BD," & _
                        " NFe_T1_usuario_BD," & _
                        " NFe_T1_senha_BD" & _
                    " FROM t_NFE_EMITENTE" & _
                    " WHERE" & _
                        " (id = " & CStr(intIdNfeEmitente) & ")"
                If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
                t_NFE_EMITENTE.Open s, dbc, , , adCmdText
                If t_NFE_EMITENTE.EOF Then
                    If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                    s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao localizar o registro em t_NFE_EMITENTE (id=" & CStr(intIdNfeEmitente) & ")!!"
                    GoTo PROXIMO_PEDIDO
                    End If
                    
                strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
                strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
                strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
                strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
                strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
                
                decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
                s = "Provider=" & BD_OLEDB_PROVIDER & _
                    ";Data Source=" & strNfeT1ServidorBd & _
                    ";Initial Catalog=" & strNfeT1NomeBd & _
                    ";User Id=" & strNfeT1UsuarioBd & _
                    ";Password=" & s_aux
                If dbcNFe.State <> adStateClosed Then dbcNFe.Close
                dbcNFe.Open s
                
            '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
                Set cmdNFeSituacao.ActiveConnection = dbcNFe
                
                Do While Not t_NFe_EMISSAO.EOF
                    lngNFeSerieNF = t_NFe_EMISSAO("NFe_serie_NF")
                    lngNFeNumeroNF = t_NFe_EMISSAO("NFe_numero_NF")
                    
                    strNumeroNfNormalizado = NFeFormataNumeroNF(lngNFeNumeroNF)
                    strSerieNfNormalizado = NFeFormataSerieNF(lngNFeSerieNF)
                    
                '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
                    cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
                    cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
                    Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
                    intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
                    strNFeMsgRetornoSP = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
                    
                    If intNfeRetornoSP <> 1 Then
                        If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                        s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": não é possível consultar a DANFE nº " & strNumeroNfNormalizado & " porque a NFe encontra-se na seguinte situação: " & CStr(intNfeRetornoSP) & " - " & strNFeMsgRetornoSP
                        GoTo PROXIMA_NFE
                        End If
                                    
                    aguarde INFO_EXECUTANDO, "copiando PDF do DANFE"
                    Set cmdNFeDanfe.ActiveConnection = dbcNFe
                    cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
                    cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
                    Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
                    If rsNFeRetornoSPDanfe.EOF Then
                        If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                        s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": o conteúdo da DANFE nº " & strNumeroNfNormalizado & " não foi encontrado!!"
                        GoTo PROXIMA_NFE
                        End If
                    
                    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & "_" & strPedido & ".pdf"
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
                    
                    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                            GoTo PROXIMA_NFE
                            End If
                        End If
                    
                    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                            If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                            s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                            GoTo PROXIMA_NFE
                            End If
                        End If
                    
                    lFileHandle = FreeFile
                    Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                    lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                    lngOffset = 0
                    Do While lngOffset < lngFileSize
                        bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                        Put #lFileHandle, , bytFile()
                        lngOffset = lngOffset + CHUNK_SIZE
                        Loop
                    
                    If v_danfe(UBound(v_danfe)) <> "" Then ReDim Preserve v_danfe(UBound(v_danfe) + 1)
                    v_danfe(UBound(v_danfe)) = strNomeArqCompletoDanfe
                    
                    Close #lFileHandle
                
PROXIMA_NFE:
'===========
                    t_NFe_EMISSAO.MoveNext
                    Loop
                End If
                
PROXIMO_PEDIDO:
'==============
            Next
    
    
        End If
'----------------------------------------------------------------------------------
'FIM DO TRECHO PARA EMISSÃO DE DANFE's RELACIONADAS A OPERAÇÕES NÃO TRIANGULARES
'----------------------------------------------------------------------------------

    GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
    
    aguarde INFO_EXECUTANDO, "exibindo PDF do DANFE"
    
    For ic = LBound(v_danfe) To UBound(v_danfe)
        If Trim$(v_danfe(ic)) <> "" Then
            If Not start_doc(Trim$(v_danfe(ic)), s_erro) Then
                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                s_alerta_erro = s_alerta_erro & "Falha ao exibir o arquivo PDF do DANFE (" & Trim$(v_danfe(ic)) & "): " & s_erro
                End If
            End If
        Next
    
'   HOUVE ERROS?
    If s_alerta_erro <> "" Then aviso_erro s_alerta_erro
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS:
'===========================================
  ' RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset t_NFE_TRIANGULAR, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPDanfe, True
    
  ' COMMAND
    bd_desaloca_command cmdNFeSituacao
    bd_desaloca_command cmdNFeDanfe
    
  ' CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
    
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DANFE_CONSULTA_PARAM_EMITENTE_TRATA_ERRO:
'========================================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub


Sub NFe_emite_venda()
' __________________________________________________________________________________________
'|
'|  EMITE A NOTA FISCAL ELETRÔNICA (NFe) DE VENDA COM BASE NO PEDIDO
'|  ESPECIFICADO E NOS DEMAIS PARÂMETROS PREENCHIDOS MANUALMENTE.
'|
'|  OS PRODUTOS (T_PEDIDO_ITEM) COM PRECO_NF = R$ 0,00 SÃO
'|  RELATIVOS A BRINDES E DEVEM SER TOTALMENTE IGNORADOS.
'|  OS BRINDES ACOMPANHAM OS OUTROS PRODUTOS DENTRO DA MESMA CAIXA.
'|

' CONSTANTES
Const NomeDestaRotina = "NFe_emite_venda()"
Const MAX_LINHAS_NOTA_FISCAL_DEFAULT = 19
Const NFE_AMBIENTE_PRODUCAO = "1" '1-Produção  2-Homologação
Const NFE_AMBIENTE_HOMOLOGACAO = "2" '1-Produção  2-Homologação
'Const NFE_FINALIDADE_NFE = "1" '1-Normal  2-Complementar  3-Ajuste
Const NFE_INDFINAL_CONSUMIDOR_NORMAL = "0"
Const NFE_INDFINAL_CONSUMIDOR_FINAL = "1"


' STRINGS
Dim NFE_AMBIENTE As String
Dim c As String
Dim s As String
Dim s_confirma As String
Dim s_aux As String
Dim s_msg As String
Dim s_serie_NF_aux As String
Dim s_numero_NF_aux As String
Dim s_erro As String
Dim s_erro_aux As String
Dim strCampo As String
Dim strCnpjCpfAux As String
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim strRamal As String
Dim strConfirmacaoEtgImediata As String
Dim strIcms As String
Dim strSerieNf As String
Dim strSerieNfNormalizado As String
Dim strNumeroNf As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfTriangular As String
Dim strNumeroNfTriangular As String
Dim strEmitenteNf As String
Dim strIdCliente As String
Dim strFabricanteAnterior As String
Dim strProdutoAnterior As String
Dim strPedidoAnterior As String
Dim strLoja As String
Dim strOrigemUF As String
Dim strDestinoUF As String
Dim strPresComprador As String
Dim strConfirmacaoObs2 As String
Dim strTransportadoraId As String
Dim strTransportadoraCnpj As String
Dim strTransportadoraRazaoSocial As String
Dim strTransportadoraIE As String
Dim strTransportadoraUF As String
Dim strTransportadoraEmail As String
Dim strListaPedidosSemTransportadora As String
Dim strListaPedidosComTransportadora As String
Dim strTipoParcelamento As String
Dim strLogPedido As String
Dim strLogComplemento As String
Dim strNFeCodFinalidade As String
Dim strNFeCodFinalidadeAux As String
Dim strNFeChaveAcessoNotaReferenciada As String
Dim strNFeArquivo As String
Dim strNFeTagOperacional As String
Dim strNFeTagIdentificacao As String
Dim strNFeTagDestinatario As String
Dim strNFeTagEndEntrega As String
Dim strNFeTagBlocoProduto As String
Dim strNFeTagDet As String
Dim strNFeTagIcms As String
Dim strNFeCst As String
Dim strNFeTagPis As String
Dim strNFeTagCofins As String
Dim strNFeTagIcmsUFDest As String
Dim strNFeTagValoresTotais As String
Dim strNFeTagTransp As String
Dim strNFeTagTransporta As String
Dim strNFeTagVol As String
Dim strNFeTagFat As String
Dim strNFeTagDup As String
Dim strNFeTagInfAdicionais As String
Dim strNFeTagPag As String
Dim strNFeInfAdicQuadroProdutos As String
Dim strNFeInfAdicQuadroInfAdic As String
Dim strCfopCodigo As String
Dim strCfopCodigoFormatado As String
Dim strCfopDescricao As String
Dim strCfopCodigoAux As String
Dim strCfopCodigoFormatadoAux As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim strDestinatarioCnpjCpf As String
Dim strEndEtgEndereco As String
Dim strEndEtgEnderecoNumero As String
Dim strEndEtgEnderecoComplemento As String
Dim strEndEtgBairro As String
Dim strEndEtgCidade As String
Dim strEndEtgUf As String
Dim strEndEtgCep As String
Dim strEndEtgEnderecoCompletoFormatado As String
Dim strEndClienteUf As String
Dim strEmitenteCidade As String
Dim strEmitenteUf As String
Dim strNFeMsgRetornoSPSituacao As String
Dim strNFeMsgRetornoSPEmite As String
Dim strNFeMsgRetornoSPEmiteTamAjustadoBD As String
Dim strCodStatusInutilizacao As String
Dim strListaSugeridaMunicipiosIBGE As String
Dim strTextoCubagem As String
Dim strZerarPisCst As String
Dim strZerarCofinsCst As String
Dim strInfoAdicIbpt As String
Dim strEmailXML As String
Dim strNFeRef As String
Dim strInfoAdicParc As String
Dim strPedidoBSMarketplace As String
Dim strMarketplaceCodOrigem As String
Dim strMarketPlaceCNPJ As String
Dim strMarketPlaceCadIntTran As String
Dim strPagtoAntecipadoStatus As Integer
Dim strPagtoAntecipadoQuitadoStatus As Integer
Dim s_Texto_DIFAL_UF As String

' FLAGS
Dim blnAchou As Boolean
Dim blnTemPedidoComTransportadora As Boolean
Dim blnTemPedidoSemTransportadora As Boolean
Dim blnTemPedidoComStBemUsoConsumo As Boolean
Dim blnTemPedidoSemStBemUsoConsumo As Boolean
Dim blnTemPagtoPorBoleto As Boolean
Dim blnImprimeDadosFatura As Boolean
Dim blnIsDestinatarioPJ As Boolean
Dim blnTemEndEtg As Boolean
Dim blnHaProdutoCstIcms60 As Boolean
Dim blnErro As Boolean
Dim blnExibirTotalTributos As Boolean
Dim blnHaProdutoSemDadosIbpt As Boolean
Dim blnExisteMemorizacaoEndereco As Boolean
Dim blnNotadeCompromisso As Boolean
Dim blnRemessaEntregaFutura As Boolean
Dim blnIgnorarDIFAL As Boolean

' CONTADORES
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim n As Long
Dim ic As Integer
Dim intNumItem As Integer
Dim intIdNfeEmitente As Integer
Dim iQtdConfirmaDuvidaEmit As Integer

' QUANTIDADES
Dim qtde As Long
Dim total_volumes As Long
Dim qtde_pedidos As Integer
Dim qtde_linhas_nf As Integer
Dim idx As Integer
Dim lngMax As Long
Dim lngAffectedRecords As Long
Dim MAX_LINHAS_NOTA_FISCAL As Integer

' CÓDIGOS E NSU
Dim intNfeRetornoSPSituacao As Integer
Dim intNfeRetornoSPEmite As Integer
Dim lngNsuNFeEmissao As Long
Dim lngNsuNFeImagem As Long
Dim lngNFeUltNumeroNfEmitido As Long
Dim lngNFeUltSerieEmitida As Long
Dim lngNFeSerieManual As Long
Dim lngNFeNumeroNfManual As Long
Dim intContribuinteICMS As Integer
Dim intAnoPartilha As Integer
Dim intImprimeIntermediadorAusente As Integer

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_PEDIDO_ITEM_DEVOLVIDO As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset
Dim t_TRANSPORTADORA As ADODB.Recordset
Dim t_IBPT As ADODB.Recordset
Dim t_NFe_EMITENTE_X_LOJA As ADODB.Recordset
'Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim t_T1_NFE_INUTILIZA As ADODB.Recordset
Dim t_CODIGO_DESCRICAO As ADODB.Recordset
Dim t_NFe_UF_PARAMETRO As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPEmite As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeEmite As New ADODB.Command
Dim dbcNFe As ADODB.Connection

' MOEDA
Dim vl_unitario As Currency
Dim vl_total_produtos As Currency
Dim vl_total_BC_ICMS As Currency
Dim vl_total_BC_ICMS_ST As Currency
Dim vl_BC_ICMS As Currency
Dim vl_BC_ICMS_ST As Currency
Dim vl_BC_ICMS_ST_Ret As Currency
Dim vl_ICMS As Currency
Dim vl_ICMSDeson As Currency
Dim vl_ICMS_ST As Currency
Dim vl_ICMS_ST_Ret As Currency
Dim vl_IPI As Currency
Dim vl_total_ICMS As Currency
Dim vl_total_ICMSDeson As Currency
Dim vl_total_ICMS_ST As Currency
Dim vl_total_IPI As Currency
Dim vl_aux As Currency
Dim vl_total_outras_despesas_acessorias As Currency
Dim vl_BC_PIS As Currency
Dim vl_PIS As Currency
Dim vl_total_PIS As Currency
Dim vl_BC_COFINS As Currency
Dim vl_COFINS As Currency
Dim vl_total_COFINS As Currency
Dim vl_estimado_tributos As Currency
Dim vl_total_estimado_tributos As Currency
Dim vl_total_NF As Currency
Dim vl_fcp As Currency
Dim vl_ICMS_UF_dest As Currency
Dim vl_ICMS_UF_remet As Currency
Dim vl_ICMS_diferencial_interestadual As Currency
Dim vl_ICMS_diferencial_aux As Currency
Dim vl_total_FCPUFDest As Currency
Dim vl_total_ICMSUFDest As Currency
Dim vl_total_ICMSUFRemet As Currency
Dim vl_total_vFCP As Currency
Dim vl_total_vFCPST As Currency
Dim vl_total_vFCPSTRet As Currency
Dim vl_total_vIPIDevol As Currency


' PERCENTUAL
Dim perc_ICMS As Single
Dim perc_ICMS_ST As Single
Dim perc_ICMS_ST_aux As Single
Dim perc_IPI As Single
Dim perc_PIS As Single
Dim perc_COFINS As Single
Dim perc_IBPT As Single
Dim perc_aux As Single
Dim perc_ICMS_interna_UF_dest As Single
Dim perc_ICMS_UF_dest As Single
Dim perc_ICMS_UF_remet As Single
Dim perc_fcp As Single
Dim perc_ICMS_diferencial_interestadual As Single

' REAL
Dim peso_aux As Single
Dim total_peso_bruto As Single
Dim total_peso_liquido As Single
Dim cubagem_aux As Single
Dim cubagem_bruto As Single
Dim aliquota_icms_interestadual As Single

' VETORES
Dim v() As String
Dim v_pedido() As String
Dim v_nf() As TIPO_LINHA_NOTA_FISCAL
Dim v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim v_nf_confere() As TIPO_LINHA_NOTA_FISCAL
Dim v_flagDadosTelaJaLido() As Boolean
Dim vListaNFeRef() As String

' DADOS DE IMAGEM DA NFE
Dim rNFeImg As TIPO_NFe_IMG
Dim vNFeImgItem() As TIPO_NFe_IMG_ITEM
Dim vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP
Dim vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA
Dim vNFeImgPag() As TIPO_NFe_IMG_PAG

    On Error GoTo NFE_EMITE_TRATA_ERRO
            
    c_pedido = normaliza_lista_pedidos(c_pedido)
    
    If Not pedido_eh_do_emitente_atual(c_pedido) Then Exit Sub
    
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_produto(i)) <> "" Then
            If converte_para_currency(c_vl_outras_despesas_acessorias(i)) < 0 Then
                aviso_erro "O valor das outras despesas acessórias do produto " & Trim$(c_produto(i)) & " não pode ser negativo!!"
                c_vl_outras_despesas_acessorias(i).SetFocus
                Exit Sub
                End If
            End If
        Next
    
    
    If DESENVOLVIMENTO Then
        NFE_AMBIENTE = NFE_AMBIENTE_HOMOLOGACAO
    Else
        NFE_AMBIENTE = NFE_AMBIENTE_PRODUCAO
        End If
        
    ReDim v_pedido(0)
    v_pedido(UBound(v_pedido)) = ""
    
    ReDim vNFeImgItem(0)
    ReDim vNFeImgTagDup(0)
    ReDim vNFeImgNFeRef(0)
    ReDim vNFeImgPag(0)
    
    qtde_pedidos = 0
    iQtdConfirmaDuvidaEmit = 0
    
    strNFeArquivo = ""
    strNFeTagOperacional = ""
    strNFeTagIdentificacao = ""
    strNFeTagDestinatario = ""
    strNFeTagEndEntrega = ""
    strNFeTagBlocoProduto = ""
    strNFeTagValoresTotais = ""
    strNFeTagTransp = ""
    strNFeTagTransporta = ""
    strNFeTagInfAdicionais = ""
    strNFeInfAdicQuadroProdutos = ""
    strNFeInfAdicQuadroInfAdic = ""
    strNFeTagDup = ""
    strNFeTagFat = ""
    
    blnTemPedidoComStBemUsoConsumo = False
    blnTemPedidoSemStBemUsoConsumo = False
    blnTemPedidoComTransportadora = False
    blnTemPedidoSemTransportadora = False
    blnTemPagtoPorBoleto = False
    blnImprimeDadosFatura = False
    strListaPedidosSemTransportadora = ""
    strListaPedidosComTransportadora = ""
    
    v = Split(c_pedido, vbCrLf)
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
        '   REPETIDO?
            For j = LBound(v_pedido) To UBound(v_pedido)
                If Trim$(v(i)) = v_pedido(j) Then
                    aviso_erro "Pedido " & Trim$(v(i)) & " está repetido na lista!!"
                    c_pedido.SetFocus
                    Exit Sub
                    End If
                Next
                
            If v_pedido(UBound(v_pedido)) <> "" Then ReDim Preserve v_pedido(UBound(v_pedido) + 1)
            v_pedido(UBound(v_pedido)) = Trim$(v(i))
            qtde_pedidos = qtde_pedidos + 1
            End If
        Next
    
    If qtde_pedidos = 0 Then
        aviso_erro "Informe o número do pedido!!"
        c_pedido.SetFocus
        Exit Sub
        End If
        
    If qtde_pedidos > 1 Then
        aviso_erro "É possível emitir a NFe de apenas 1 pedido por vez!!"
        c_pedido.SetFocus
        Exit Sub
        End If
    
    rNFeImg.pedido = c_pedido
    
'   OBTÉM TIPO DO DOCUMENTO FISCAL
'   LHGX - fixando tipo do documento em SAÍDA [no form, cb_tipo_NF estará desabilitado](verificar com Bonshop se está correto)
    rNFeImg.ide__tpNF = left$(Trim$(cb_tipo_NF), 1)
    If rNFeImg.ide__tpNF = "" Then
        aviso_erro "Selecione o tipo de documento fiscal (entrada ou saída)!!"
        Exit Sub
        End If
        
    If rNFeImg.ide__tpNF = "0" Then
        s = "A NFe que será emitida será de ENTRADA!!" & vbCrLf & "Continua com a emissão da NFe?"
        If Not confirma(s) Then
            Exit Sub
            End If
        End If
        
        
'>  NATUREZA DA OPERAÇÃO
    s = UCase$(cb_natureza)
    strCfopCodigoFormatado = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = " " Then Exit For
        strCfopCodigoFormatado = strCfopCodigoFormatado & c
        Next
        
    strCfopCodigo = retorna_so_digitos(strCfopCodigoFormatado)
    strCfopDescricao = Trim$(Mid$(s, Len(strCfopCodigoFormatado) + 1, Len(s) - Len(strCfopCodigoFormatado)))
        
'>  LOCAL DE DESTINO DA OPERAÇÃO
    rNFeImg.ide__idDest = left$(Trim$(cb_loc_dest), 1)
        
'>  FINALIDADE DE EMISSÃO
    strNFeCodFinalidade = left$(Trim$(cb_finalidade), 1)
    If strNFeCodFinalidade = "" Then
        aviso_erro "Selecione a finalidade da NFe!!"
        Exit Sub
        End If
        
'>  CNPJ/CPF DESTINATÁRIO
    If Trim(c_cnpj_cpf_dest) = "" Then
        aviso_erro "Preencha o CNPJ/CPF do recebedor!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
        
'>  NOME DESTINATÁRIO
    If Trim(c_nome_dest) = "" Then
        aviso_erro "Preencha o nome do recebedor!!"
        c_nome_dest.SetFocus
        Exit Sub
        End If
    
    strNFeCodFinalidadeAux = retorna_finalidade_nfe(strCfopCodigo)
    If strNFeCodFinalidade <> strNFeCodFinalidadeAux Then
        s = "Possível divergência encontrada na finalidade da NFe:" & vbCrLf & _
            "Finalidade selecionada: " & strNFeCodFinalidade & " - " & descricao_finalidade_nfe(strNFeCodFinalidade) & vbCrLf & _
            "Finalidade recomendada para o CFOP " & strCfopCodigoFormatado & ": " & strNFeCodFinalidadeAux & " - " & descricao_finalidade_nfe(strNFeCodFinalidadeAux) & _
            vbCrLf & vbCrLf & _
            "Continua com a emissão da NFe?"
        If Not confirma(s) Then
            Exit Sub
            End If
        End If
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM_DEVOLVIDO
    Set t_PEDIDO_ITEM_DEVOLVIDO = New ADODB.Recordset
    With t_PEDIDO_ITEM_DEVOLVIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_DESTINATARIO
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_TRANSPORTADORA
    Set t_TRANSPORTADORA = New ADODB.Recordset
    With t_TRANSPORTADORA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_IBPT
    Set t_IBPT = New ADODB.Recordset
    With t_IBPT
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

  ' T_NFE_EMITENTE_X_LOJA
    Set t_NFe_EMITENTE_X_LOJA = New ADODB.Recordset
    With t_NFe_EMITENTE_X_LOJA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'  ' T_FIN_BOLETO_CEDENTE
'    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
'    With t_FIN_BOLETO_CEDENTE
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
  
  ' T_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
   
   ' T_NFE_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
  
  ' T_NFe_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
  
  ' T_T1_NFE_INUTILIZA
    Set t_T1_NFE_INUTILIZA = New ADODB.Recordset
    With t_T1_NFE_INUTILIZA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_CODIGO_DESCRICAO
    Set t_CODIGO_DESCRICAO = New ADODB.Recordset
    With t_CODIGO_DESCRICAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' t_NFe_UF_PARAMETRO
    Set t_NFe_UF_PARAMETRO = New ADODB.Recordset
    With t_NFe_UF_PARAMETRO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  

'   VERIFICA CADA UM DOS PEDIDOS
    strIdCliente = ""
    strPedidoAnterior = ""
    strLoja = ""
    s_erro = ""
    strConfirmacaoObs2 = ""
    strConfirmacaoEtgImediata = ""
    strTransportadoraId = ""
    strPedidoBSMarketplace = ""
    rNFeImg.ide__indPag = "2" ' Forma de pagamento: outros
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            s = "SELECT" & _
                    " t_PEDIDO.pedido," & _
                    " t_PEDIDO.loja," & _
                    " t_PEDIDO.st_entrega," & _
                    " t_PEDIDO.id_cliente," & _
                    " t_PEDIDO.obs_2," & _
                    " t_PEDIDO.transportadora_id," & _
                    " t_PEDIDO.StBemUsoConsumo," & _
                    " t_PEDIDO.st_etg_imediata," & _
                    " t_PEDIDO.pedido_bs_x_marketplace," & _
                    " t_PEDIDO.marketplace_codigo_origem," & _
                    " t_PEDIDO.PagtoAntecipadoQuitadoStatus," & _
                    " t_PEDIDO__BASE.PagtoAntecipadoStatus," & _
                    " t_PEDIDO__BASE.tipo_parcelamento," & _
                    " t_PEDIDO__BASE.av_forma_pagto," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_entrada," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_prestacao," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_demais_prest," & _
                    " t_PEDIDO__BASE.pu_forma_pagto" & _
                " FROM t_PEDIDO" & _
                    " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
                        " ON (SUBSTRING(t_PEDIDO.pedido,1," & CStr(TAM_MIN_ID_PEDIDO) & ")=t_PEDIDO__BASE.pedido)" & _
                " WHERE" & _
                    " (t_PEDIDO.pedido='" & Trim$(v_pedido(i)) & "')"
            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
            t_PEDIDO.Open s, dbc, , , adCmdText
            If t_PEDIDO.EOF Then
                If s_erro <> "" Then s_erro = s_erro & vbCrLf
                s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " não está cadastrado!!"
            Else
                strLoja = Trim$("" & t_PEDIDO("loja"))
                
                strPedidoBSMarketplace = Trim$("" & t_PEDIDO("pedido_bs_x_marketplace"))
                strMarketplaceCodOrigem = Trim$("" & t_PEDIDO("marketplace_codigo_origem"))
                
                strPagtoAntecipadoStatus = Trim$("" & CStr(t_PEDIDO("PagtoAntecipadoStatus")))
                strPagtoAntecipadoQuitadoStatus = Trim$("" & CStr(t_PEDIDO("PagtoAntecipadoQuitadoStatus")))
                
                If CLng(t_PEDIDO("StBemUsoConsumo")) = 1 Then
                    blnTemPedidoComStBemUsoConsumo = True
                Else
                    blnTemPedidoSemStBemUsoConsumo = True
                    End If
                    
                If (Trim$("" & t_PEDIDO("obs_2")) <> "") And (Not IsLetra(Trim$("" & t_PEDIDO("obs_2")))) Then
                    If strConfirmacaoObs2 <> "" Then strConfirmacaoObs2 = strConfirmacaoObs2 & vbCrLf
                    strConfirmacaoObs2 = strConfirmacaoObs2 & Trim$("" & t_PEDIDO("pedido")) & " preenchido com: " & Trim$("" & t_PEDIDO("obs_2"))
                    End If
                    
                If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " está cancelado!!"
                    End If
                    
                If CLng(t_PEDIDO("st_etg_imediata")) <> 2 Then
                    If strConfirmacaoEtgImediata <> "" Then strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & vbCrLf
                    strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & "Pedido " & Trim$(v_pedido(i)) & " NÃO está definido para 'Entrega Imediata'!!"
                    End If
                
                strTipoParcelamento = Trim$("" & t_PEDIDO("tipo_parcelamento"))
                If strTipoParcelamento = CStr(COD_FORMA_PAGTO_A_VISTA) Then
                    rNFeImg.ide__indPag = "0"  ' A vista
                    If Trim$("" & t_PEDIDO("av_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
                    rNFeImg.ide__indPag = "1"  ' A prazo
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_entrada")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_prestacao")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
                    rNFeImg.ide__indPag = "1"  ' A prazo
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_prim_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_demais_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
                    rNFeImg.ide__indPag = "2"  ' Outros
                    If Trim$("" & t_PEDIDO("pu_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    End If
                
                If Trim$("" & t_PEDIDO("transportadora_id")) = "" Then
                    blnTemPedidoSemTransportadora = True
                    If strListaPedidosSemTransportadora <> "" Then strListaPedidosSemTransportadora = strListaPedidosSemTransportadora & ", "
                    strListaPedidosSemTransportadora = strListaPedidosSemTransportadora & Trim$(v_pedido(i))
                Else
                    blnTemPedidoComTransportadora = True
                    If strListaPedidosComTransportadora <> "" Then strListaPedidosComTransportadora = strListaPedidosComTransportadora & ", "
                    strListaPedidosComTransportadora = strListaPedidosComTransportadora & Trim$(v_pedido(i))
                    
                    If strTransportadoraId = "" Then
                        strTransportadoraId = Trim$("" & t_PEDIDO("transportadora_id"))
                    Else
                        If strTransportadoraId <> Trim$("" & t_PEDIDO("transportadora_id")) Then
                            If s_erro <> "" Then s_erro = s_erro & vbCrLf
                            s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " informa uma transportadora diferente!!"
                            End If
                        End If
                    End If
                    
            '   TODOS OS PEDIDOS DEVEM PERTENCER AO MESMO CLIENTE
                If strIdCliente = "" Then
                    strIdCliente = Trim$("" & t_PEDIDO("id_cliente"))
                    strPedidoAnterior = Trim$("" & t_PEDIDO("pedido"))
                    End If
                If strIdCliente <> Trim$("" & t_PEDIDO("id_cliente")) Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " pertence a um cliente diferente que o pedido " & strPedidoAnterior & "!!"
                    End If
                End If
            
            s = "SELECT " & _
                    "pedido, " & _
                    "fabricante, " & _
                    "produto" & _
                " FROM t_PEDIDO_ITEM" & _
                " WHERE" & _
                    " (pedido='" & Trim$(v_pedido(i)) & "')" & _
                    " AND (preco_NF > 0)"
            If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
            t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
            If t_PEDIDO_ITEM.EOF Then
                If s_erro <> "" Then s_erro = s_erro & vbCrLf
                s_erro = s_erro & "Não foi encontrado nenhum produto relacionado ao pedido " & Trim$(v_pedido(i)) & "!!"
                End If
            End If
            
            'obter as informações de marketplace
            If (param_nfintermediador.campo_inteiro = 1) And (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") Then
                s = "SELECT o.codigo, o.descricao, og.parametro_campo_texto, og.parametro_2_campo_texto, og.parametro_3_campo_flag  " & _
                    "FROM (select * from t_CODIGO_DESCRICAO where grupo = 'PedidoECommerce_Origem') o  " & _
                        "INNER JOIN (select * from t_CODIGO_DESCRICAO where grupo = 'PedidoECommerce_Origem_Grupo') og  " & _
                        "on o.codigo_pai = og.codigo " & _
                    "WHERE o.codigo = '" & strMarketplaceCodOrigem & "'"
                If t_CODIGO_DESCRICAO.State <> adStateClosed Then t_CODIGO_DESCRICAO.Close
                t_CODIGO_DESCRICAO.Open s, dbc, , , adCmdText
                If t_CODIGO_DESCRICAO.EOF Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Problema na identificação do marketplace do pedido " & Trim$(v_pedido(i)) & "!!"
                Else
                    strMarketPlaceCNPJ = Trim$("" & t_CODIGO_DESCRICAO("parametro_campo_texto"))
                    strMarketPlaceCadIntTran = Trim$("" & t_CODIGO_DESCRICAO("parametro_2_campo_texto"))
                    intImprimeIntermediadorAusente = t_CODIGO_DESCRICAO("parametro_3_campo_flag")
                    End If
                    
                End If

            
        Next
        
    If s_erro = "" Then
        If blnTemPedidoComTransportadora And blnTemPedidoSemTransportadora Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Há pedido(s) com transportadora cadastrada (" & strListaPedidosComTransportadora & ") e há pedido(s) sem transportadora cadastrada (" & strListaPedidosSemTransportadora & ")!!"
            End If
        End If
        
'   ENCONTROU ERRO?
    If s_erro <> "" Then
        aviso_erro s_erro
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
        
'   OBTÉM OS DADOS DO EMITENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~
    If strLoja = "" Then
        aviso_erro "Falha ao obter o nº da loja do pedido!!"
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
                
    If usuario.emit_id <> "" Then
        intIdNfeEmitente = CInt(usuario.emit_id)
            
        s = "SELECT" & _
                " id," & _
                " razao_social," & _
                " cidade," & _
                " uf," & _
                " NFe_T1_servidor_BD," & _
                " NFe_T1_nome_BD," & _
                " NFe_T1_usuario_BD," & _
                " NFe_T1_senha_BD" & _
            " FROM t_NFE_EMITENTE" & _
            " WHERE" & _
                " (id = " & CStr(intIdNfeEmitente) & ")"
        
        t_NFE_EMITENTE.Open s, dbc, , , adCmdText
        If t_NFE_EMITENTE.EOF Then
            aviso_erro "Dados do emitente não foram localizados no BD (id=" & CStr(intIdNfeEmitente) & ")!!"
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            strEmitenteNf = Trim$("" & t_NFE_EMITENTE("razao_social"))
            strEmitenteCidade = Trim$("" & t_NFE_EMITENTE("cidade"))
            strEmitenteUf = Trim$("" & t_NFE_EMITENTE("uf"))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            End If
    Else
    '   OBTÉM O EMITENTE PADRÃO
        s = "SELECT" & _
                " id," & _
                " razao_social," & _
                " cidade," & _
                " uf," & _
                " NFe_T1_servidor_BD," & _
                " NFe_T1_nome_BD," & _
                " NFe_T1_usuario_BD," & _
                " NFe_T1_senha_BD" & _
            " FROM t_NFE_EMITENTE" & _
            " WHERE" & _
                " (NFe_st_emitente_padrao = 1)"
        
        t_NFE_EMITENTE.Open s, dbc, , , adCmdText
        If t_NFE_EMITENTE.EOF Then
            aviso_erro "Não há emitente padrão definido no sistema!!"
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        ElseIf t_NFE_EMITENTE.RecordCount > 1 Then
            aviso_erro "Há mais de 1 emitente padrão definido no sistema!!"
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            intIdNfeEmitente = t_NFE_EMITENTE("id")
            strEmitenteNf = Trim$("" & t_NFE_EMITENTE("razao_social"))
            strEmitenteCidade = Trim$("" & t_NFE_EMITENTE("cidade"))
            strEmitenteUf = Trim$("" & t_NFE_EMITENTE("uf"))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            End If
        End If
   
    rNFeImg.id_nfe_emitente = intIdNfeEmitente
   
    
    'OBTÉM O INDICADOR DE PRESENÇA DO COMPRADOR NO ESTABELECIMENTO COMERCIAL NO MOMENTO DA OPERAÇÃO
    'se loja for 201 (E-Commerce), indicador será 2 (Internet); senão, indicador será 3 (Teleatendimento)
    strPresComprador = ""
    If strLoja = "201" Then
        strPresComprador = "2"
    Else
        strPresComprador = "3"
        End If

    ' OBTÉM UF DO EMITENTE (pegar UF do emitente padrão, conforme conversa entre Hamilton e Luiz em 21/10/2014)
    strOrigemUF = strEmitenteUf
        
        
'   CONEXÃO AO BD NFE
'   ~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "conectando ao banco dados de NFe"
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
    decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
    s = "Provider=" & BD_OLEDB_PROVIDER & _
        ";Data Source=" & strNfeT1ServidorBd & _
        ";Initial Catalog=" & strNfeT1NomeBd & _
        ";User Id=" & strNfeT1UsuarioBd & _
        ";Password=" & s_aux
    dbcNFe.Open s
    
        
'   VERIFICA SE O PEDIDO JÁ TEM NFe EMITIDA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set cmdNFeSituacao.ActiveConnection = dbcNFe
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            s = "SELECT DISTINCT" & _
                    " NFe_serie_NF," & _
                    " NFe_numero_NF" & _
                " FROM t_NFe_EMISSAO" & _
                " WHERE" & _
                    " (pedido = '" & Trim$(v_pedido(i)) & "')" & _
                " ORDER BY" & _
                    " NFe_serie_NF," & _
                    " NFe_numero_NF"
            If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
            t_NFe_EMISSAO.Open s, dbc, , , adCmdText
            
            s_msg = ""
            j = 0
            Do While Not t_NFe_EMISSAO.EOF
                j = j + 1
                s_serie_NF_aux = NFeFormataSerieNF(Trim$("" & t_NFe_EMISSAO("NFe_serie_NF")))
                s_numero_NF_aux = NFeFormataNumeroNF(Trim$("" & t_NFe_EMISSAO("NFe_numero_NF")))
                
                cmdNFeSituacao.Parameters("NFe") = s_numero_NF_aux
                cmdNFeSituacao.Parameters("Serie") = s_serie_NF_aux
                Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
                intNfeRetornoSPSituacao = rsNFeRetornoSPSituacao("Retorno")
                strNFeMsgRetornoSPSituacao = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
                        
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & CStr(j) & ") " & _
                    "Série: " & s_serie_NF_aux & _
                    ", Nº: " & s_numero_NF_aux & _
                    ", Situação: " & intNfeRetornoSPSituacao & " - " & strNFeMsgRetornoSPSituacao
                t_NFe_EMISSAO.MoveNext
                Loop
                
            If s_msg <> "" Then
                s_msg = "O pedido " & Trim$(v_pedido(i)) & " já possui NFe que se encontra na seguinte situação:" & vbCrLf & s_msg
                s_msg = s_msg & vbCrLf & vbCrLf & "Continua com a emissão desta NFe?"
                If Not confirma(s_msg) Then
                    GoSub NFE_EMITE_FECHA_TABELAS
                    aguarde INFO_NORMAL, m_id
                    Exit Sub
                    End If
                End If
            End If
        Next
        
           
'   O(S) PEDIDO(S) ESTÁ COM 'ENTREGA IMEDIATA' IGUAL A 'NÃO'?
    If strConfirmacaoEtgImediata <> "" Then
        strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & _
                                    vbCrLf & vbCrLf & "Continua com a emissão da NFe?"
        If Not confirma(strConfirmacaoEtgImediata) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
        
'   SE HÁ PEDIDO COM O CAMPO "OBSERVAÇÕES II" JÁ PREENCHIDO, DEVE AVISAR E PEDIR CONFIRMAÇÃO ANTES DE PROSSEGUIR
'   A CONFIRMAÇÃO É FEITA SOMENTE P/ NOTAS DE SAÍDA, POIS EM NOTAS DE ENTRADA O Nº DA NFe NÃO É ANOTADO NO CAMPO
'   OBS_2 DO PEDIDO, MAS SIM NOS ITENS DEVOLVIDOS, QUANDO APLICÁVEL.
'   0-Entrada  1-Saída
    If rNFeImg.ide__tpNF = "1" Then
        If strConfirmacaoObs2 <> "" Then
            strConfirmacaoObs2 = "O campo " & Chr$(34) & "Observações II" & Chr$(34) & " já está preenchido nos seguintes pedidos:" & _
                                 vbCrLf & strConfirmacaoObs2 & _
                                 vbCrLf & vbCrLf & "Continua com a emissão da NFe?"
            If Not confirma(strConfirmacaoObs2) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
        
        
'   NO CASO DE UM PRODUTO APARECER EM VÁRIOS PEDIDOS E O PREÇO DE VENDA FOR DIFERENTE,
'   DEVE PEDIR UMA CONFIRMAÇÃO AO OPERADOR ANTES DE USAR A MÉDIA DO PREÇO DE VENDA
    If qtde_pedidos > 1 Then
        s = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
        s = "SELECT " & _
                "fabricante, " & _
                "produto, " & _
                "preco_NF, " & _
                "t_PEDIDO_ITEM.pedido, " & _
                "descricao" & _
            " FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
            " WHERE" & _
                " (" & s & ")" & _
                " AND (preco_NF > 0)" & _
            " ORDER BY " & _
                "fabricante, " & _
                "produto, " & _
                "t_PEDIDO.data, " & _
                "t_PEDIDO.pedido"
        strFabricanteAnterior = "XXXXX"
        strProdutoAnterior = "XXXXXXXXXX"
        vl_aux = 0
        s_erro = ""
        n = 0
        If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
        t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
        Do While Not t_PEDIDO_ITEM.EOF
            If (strFabricanteAnterior = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And (strProdutoAnterior = Trim$("" & t_PEDIDO_ITEM("produto"))) Then
                If vl_aux <> t_PEDIDO_ITEM("preco_NF") Then
                    n = n + 1
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Produto " & Trim$("" & t_PEDIDO_ITEM("produto")) & " do fabricante " & Trim$("" & t_PEDIDO_ITEM("fabricante")) & ":   " & Trim$("" & t_PEDIDO_ITEM("pedido")) & " = " & Format$(t_PEDIDO_ITEM("preco_NF"), FORMATO_MOEDA) & "   " & strPedidoAnterior & " = " & Format$(vl_aux, FORMATO_MOEDA)
                    End If
                End If
            
            strFabricanteAnterior = Trim$("" & t_PEDIDO_ITEM("fabricante"))
            strProdutoAnterior = Trim$("" & t_PEDIDO_ITEM("produto"))
            strPedidoAnterior = Trim$("" & t_PEDIDO_ITEM("pedido"))
            vl_aux = t_PEDIDO_ITEM("preco_NF")
            
            t_PEDIDO_ITEM.MoveNext
            Loop
        
        If s_erro <> "" Then
            If n = 1 Then
                s = "O seguinte produto aparece em mais de um pedido com preços de venda diferentes!!"
            Else
                s = "Os seguintes produtos aparecem em mais de um pedido com preços de venda diferentes!!"
                End If
            s_erro = s & vbCrLf & _
                "Continua com a emissão da nota usando o valor médio do preço de venda?" & _
                vbCrLf & vbCrLf & s_erro
            If Not confirma(s_erro) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
    
'   OBTÉM OS PRODUTOS E AS QUANTIDADES P/ USAR NA CONFERÊNCIA
    ReDim v_nf_confere(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf_confere(UBound(v_nf_confere))
    
    s = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
    s = "SELECT" & _
            " t_PEDIDO.pedido," & _
            " t_PEDIDO.data," & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.qtde" & _
        " FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
        " WHERE" & _
            " (" & s & ")" & _
            " AND (preco_NF > 0)" & _
        " ORDER BY " & _
            "produto, " & _
            "t_PEDIDO.data, " & _
            "t_PEDIDO.pedido"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    Do While Not t_PEDIDO_ITEM.EOF
        blnAchou = False
        For i = LBound(v_nf_confere) To UBound(v_nf_confere)
            With v_nf_confere(i)
                If (.fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And (.produto = Trim$("" & t_PEDIDO_ITEM("produto"))) Then
                    blnAchou = True
                    idx = i
                    Exit For
                    End If
                End With
            Next
        
        If Not blnAchou Then
            If v_nf_confere(UBound(v_nf_confere)).produto <> "" Then
                ReDim Preserve v_nf_confere(UBound(v_nf_confere) + 1)
                limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf_confere(UBound(v_nf_confere))
                End If
            idx = UBound(v_nf_confere)
            With v_nf_confere(UBound(v_nf_confere))
                .fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))
                .produto = Trim$("" & t_PEDIDO_ITEM("produto"))
                End With
            End If
        
        With v_nf_confere(idx)
        '   QUANTIDADE
            qtde = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde")) Then qtde = CLng(t_PEDIDO_ITEM("qtde"))
            .qtde_total = .qtde_total + qtde
            End With
        
        t_PEDIDO_ITEM.MoveNext
        Loop


'   OBTÉM OS DADOS DOS PRODUTOS
'   A QUANTIDADE DE PRODUTOS (IDENTIFICADO PELO CÓDIGO NCM) QUE DEU ENTRADA DEVE
'   COINCIDIR COM A QUANTIDADE QUE DEU SAÍDA. SENDO QUE O CÓDIGO NCM E/OU O CST
'   DE UM PRODUTO PODE SER ALTERADO PELO SEU FABRICANTE.
'   PORTANTO, A PARTIR DA VERSÃO 1.48 DESTE MÓDULO, O CÓDIGO NCM E O CST PASSAM
'   A SER REGISTRADOS NO MOMENTO DA ENTRADA DAS MERCADORIAS NO ESTOQUE E ESSES
'   CÓDIGOS É QUE SERÃO USADOS NA EMISSÃO DA NFe.
    ReDim v_nf(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf(UBound(v_nf))
    
'   A ORDENAÇÃO É FEITA SOMENTE PELO CÓDIGO DO PRODUTO PORQUE NA NOTA FISCAL NÃO HÁ COLUNA PARA O CÓDIGO DO FABRICANTE
    qtde_linhas_nf = 0
    s_aux = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
    s = "SELECT" & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.descricao," & _
            " t_PEDIDO_ITEM.ean," & _
            " t_PEDIDO_ITEM.preco_NF," & _
            " t_PEDIDO_ITEM.qtde_volumes," & _
            " t_PEDIDO_ITEM.peso," & _
            " t_PEDIDO_ITEM.cubagem," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst," & _
            " Coalesce(t_PRODUTO.perc_MVA_ST, 0) AS perc_MVA_ST," & _
            " Coalesce(t_PRODUTO.ean, '') AS tP_ean," & _
            " Coalesce(t_PRODUTO.peso, 0) AS tP_peso," & _
            " Coalesce(t_PRODUTO.cubagem, 0) AS tP_cubagem," & _
            " Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde"
    s = s & _
        " FROM t_PEDIDO_ITEM" & _
            " INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
            " LEFT JOIN t_PRODUTO ON (t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto)" & _
            " INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_PEDIDO_ITEM.pedido=t_ESTOQUE_MOVIMENTO.pedido) AND (t_PEDIDO_ITEM.fabricante=t_ESTOQUE_MOVIMENTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_ESTOQUE_MOVIMENTO.produto)" & _
            " INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto)"
    s = s & _
        " WHERE" & _
            " (" & s_aux & ")" & _
            " AND (anulado_status=0)" & _
            " AND (estoque <> '" & ID_ESTOQUE_DEVOLUCAO & "')" & _
            " AND (preco_NF > 0)"
    s = s & _
        " GROUP BY" & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.descricao," & _
            " t_PEDIDO_ITEM.ean," & _
            " t_PEDIDO_ITEM.preco_NF," & _
            " t_PEDIDO_ITEM.qtde_volumes," & _
            " t_PEDIDO_ITEM.peso," & _
            " t_PEDIDO_ITEM.cubagem," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst," & _
            " t_PRODUTO.perc_MVA_ST," & _
            " t_PRODUTO.ean," & _
            " t_PRODUTO.peso," & _
            " t_PRODUTO.cubagem"
    s = s & _
        " ORDER BY" & _
            " t_PEDIDO_ITEM.produto," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst"
    
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    Do While Not t_PEDIDO_ITEM.EOF
        blnAchou = False
        For i = LBound(v_nf) To UBound(v_nf)
            With v_nf(i)
                If (.fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And _
                   (.produto = Trim$("" & t_PEDIDO_ITEM("produto"))) And _
                   (.ncm = Trim$("" & t_PEDIDO_ITEM("ncm"))) And _
                   (.cst = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PEDIDO_ITEM("cst")))) Then
                    blnAchou = True
                    idx = i
                    Exit For
                    End If
                End With
            Next
            
        If Not blnAchou Then
            qtde_linhas_nf = qtde_linhas_nf + 1
            If v_nf(UBound(v_nf)).produto <> "" Then
                ReDim Preserve v_nf(UBound(v_nf) + 1)
                limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf(UBound(v_nf))
                End If
            idx = UBound(v_nf)
            With v_nf(UBound(v_nf))
                .fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))
                .produto = Trim$("" & t_PEDIDO_ITEM("produto"))
                .descricao = Trim$("" & t_PEDIDO_ITEM("descricao"))
                .EAN = Trim("" & t_PEDIDO_ITEM("ean"))
                .ncm = Trim("" & t_PEDIDO_ITEM("ncm"))
                .NCM_bd = Trim("" & t_PEDIDO_ITEM("ncm"))
                .cst = cst_converte_codigo_entrada_para_saida(Trim("" & t_PEDIDO_ITEM("cst")))
                .CST_bd = cst_converte_codigo_entrada_para_saida(Trim("" & t_PEDIDO_ITEM("cst")))
                End With
            End If
            
        With v_nf(idx)
        '   QUANTIDADE
            qtde = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde")) Then qtde = CLng(t_PEDIDO_ITEM("qtde"))
            .qtde_total = .qtde_total + qtde
        
        '   VALOR
            vl_unitario = 0
            If IsNumeric(t_PEDIDO_ITEM("preco_NF")) Then vl_unitario = t_PEDIDO_ITEM("preco_NF")
            .valor_total = .valor_total + (qtde * vl_unitario)
        
        '   QTDE DE VOLUMES
            n = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde_volumes")) Then n = CLng(t_PEDIDO_ITEM("qtde_volumes"))
            .qtde_volumes_total = .qtde_volumes_total + (qtde * n)
        
        '   PESO
            peso_aux = 0
            If IsNumeric(t_PEDIDO_ITEM("peso")) Then peso_aux = CSng(t_PEDIDO_ITEM("peso"))
            .peso_total = .peso_total + (qtde * peso_aux)
            
        '   CUBAGEM
            cubagem_aux = 0
            If IsNumeric(t_PEDIDO_ITEM("cubagem")) Then cubagem_aux = CSng(t_PEDIDO_ITEM("cubagem"))
            .cubagem_total = .cubagem_total + (qtde * cubagem_aux)
            
        '   PERCENTUAL DE MVA ST
            .perc_MVA_ST = t_PEDIDO_ITEM("perc_MVA_ST")
            
        '   EAN (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If Trim("" & t_PEDIDO_ITEM("ean")) = "" Then .EAN = Trim("" & t_PEDIDO_ITEM("tP_ean"))
        
        '   PESO (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If t_PEDIDO_ITEM("peso") = 0 Then
                peso_aux = 0
                If IsNumeric(t_PEDIDO_ITEM("tP_peso")) Then peso_aux = CSng(t_PEDIDO_ITEM("tP_peso"))
                .peso_total = .peso_total + (qtde * peso_aux)
                End If
            
        '   CUBAGEM (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If t_PEDIDO_ITEM("cubagem") = 0 Then
                cubagem_aux = 0
                If IsNumeric(t_PEDIDO_ITEM("tP_cubagem")) Then cubagem_aux = CSng(t_PEDIDO_ITEM("tP_cubagem"))
                .cubagem_total = .cubagem_total + (qtde * cubagem_aux)
                End If
            End With
            
        t_PEDIDO_ITEM.MoveNext
        Loop


'   FAZ A CONFERÊNCIA DA QUANTIDADE (APENAS P/ SE CERTIFICAR QUE A LÓGICA ESTÁ CORRETA)
    s_msg = ""
    For i = LBound(v_nf_confere) To UBound(v_nf_confere)
        If Trim$(v_nf_confere(i).produto) <> "" Then
            n = 0
            For j = LBound(v_nf) To UBound(v_nf)
                If (Trim$(v_nf_confere(i).fabricante) = Trim$(v_nf(j).fabricante)) And _
                    (Trim$(v_nf_confere(i).produto) = Trim$(v_nf(j).produto)) Then
                    n = n + v_nf(j).qtde_total
                    End If
                Next
            If CLng(v_nf_confere(i).qtde_total) <> CLng(n) Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "Houve divergência na quantidade do produto (" & v_nf_confere(i).fabricante & ")" & v_nf_confere(i).produto & ": quantidade esperada=" & CStr(v_nf_confere(i).qtde_total) & ", quantidade calculada=" & CStr(n)
                End If
            End If
        Next
    
    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
'   DADOS DA TELA: INFORMAÇÕES ADICIONAIS DO PRODUTO, CST, NCM, CFOP E ICMS
'   IMPORTANTE: O MESMO CÓDIGO DE PRODUTO PODE APARECER EM MAIS DE UMA LINHA DEVIDO AO
'   =========== CONSUMO DE DIFERENTES LOTES DO ESTOQUE QUE TENHAM DADO ENTRADA C/ CÓDIGOS
'               DIFERENTES DE NCM E/OU CST. PORTANTO, DEVE SER FEITO UM CONTROLE P/ OBTER
'               OS DADOS DA TELA EDITADOS DA OCORRÊNCIA CORRETA.
    ReDim v_flagDadosTelaJaLido(c_produto.LBound To c_produto.UBound)
    For i = LBound(v_flagDadosTelaJaLido) To UBound(v_flagDadosTelaJaLido)
        v_flagDadosTelaJaLido(i) = False
        Next
    
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            For j = c_produto.LBound To c_produto.UBound
                If Trim$(v_nf(i).fabricante) = Trim$(c_fabricante(j)) And _
                   Trim$(v_nf(i).produto) = Trim$(c_produto(j)) And _
                   Trim$(v_nf(i).ncm) = Trim$(c_NCM(j)) Then
                    If Not v_flagDadosTelaJaLido(j) Then
                        v_flagDadosTelaJaLido(j) = True
                        v_nf(i).vl_outras_despesas_acessorias = converte_para_currency(Trim$(c_vl_outras_despesas_acessorias(j)))
                        v_nf(i).infAdProd = Trim$(c_produto_obs(j))
                        v_nf(i).xPed = Trim$(c_xPed(j))
                        v_nf(i).nItemPed = Trim$(c_nItemPed(j))
                        v_nf(i).fcp = Trim$(c_fcp(j))
                        v_nf(i).CST_tela = Trim$(c_CST(j))
                        v_nf(i).NCM_tela = Trim$(c_NCM(j))
                        If cb_CFOP(j).ListIndex <> -1 Then
                            If Trim$(cb_CFOP(j)) <> "" Then
                                s = Trim$(cb_CFOP(j))
                                For k = 1 To Len(s)
                                    c = Mid$(s, k, 1)
                                    If c = " " Then Exit For
                                    v_nf(i).CFOP_tela_formatado = v_nf(i).CFOP_tela_formatado & c
                                    Next
                                v_nf(i).CFOP_tela = retorna_so_digitos(v_nf(i).CFOP_tela_formatado)
                                End If
                            End If
                        If Trim$(cb_ICMS_item(j)) <> "" Then
                            v_nf(i).ICMS_tela = Trim$(cb_ICMS_item(j))
                            End If
                        Exit For
                        End If
                    End If
                Next
            End If
        Next
    

'   CST => VERIFICA SE HOUVE ALTERAÇÃO NO CST DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).CST_tela) <> "" Then
                If Trim$(v_nf(i).CST_bd) <> Trim$(v_nf(i).CST_tela) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CST alterado de " & v_nf(i).CST_bd & " para " & v_nf(i).CST_tela
                    End If
                End If
            End If
        Next
    
    If s_msg <> "" Then
        s_msg = "Houve alteração no CST do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If

'   PREPARA O CAMPO QUE ARMAZENA O CST A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).cst = v_nf(i).CST_bd
            If Trim$(v_nf(i).CST_tela) <> "" Then v_nf(i).cst = Trim$(v_nf(i).CST_tela)
            End If
        Next
    
'   NCM => VERIFICA SE HOUVE ALTERAÇÃO NO NCM DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).NCM_tela) <> "" Then
                If Trim$(v_nf(i).NCM_bd) <> Trim$(v_nf(i).NCM_tela) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": NCM alterado de " & v_nf(i).NCM_bd & " para " & v_nf(i).NCM_tela
                    End If
                End If
            End If
        Next
    
    If s_msg <> "" Then
        s_msg = "Houve alteração no NCM do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If

'   PREPARA O CAMPO QUE ARMAZENA O NCM A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).ncm = v_nf(i).NCM_bd
            If Trim$(v_nf(i).NCM_tela) <> "" Then v_nf(i).ncm = Trim$(v_nf(i).NCM_tela)
            End If
        Next
    
'   CFOP => VERIFICA SE HOUVE ALTERAÇÃO NO CFOP DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).CFOP_tela) <> "" Then
                If Trim$(v_nf(i).CFOP_tela) <> strCfopCodigo Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CFOP alterado para " & v_nf(i).CFOP_tela_formatado
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "Houve alteração no CFOP do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   PREPARA O CAMPO QUE ARMAZENA O CFOP A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).cfop = strCfopCodigo
            v_nf(i).CFOP_formatado = strCfopCodigoFormatado
            If Trim$(v_nf(i).CFOP_tela) <> "" Then
                v_nf(i).cfop = Trim$(v_nf(i).CFOP_tela)
                v_nf(i).CFOP_formatado = Trim$(v_nf(i).CFOP_tela_formatado)
                End If
            End If
        Next

'   VERIFICA SE O CFOP A SER USADO É CONFLITANTE COM O LOCAL DE DESTINO DA OPERAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).cfop) <> "" Then
                If existe_divergencia_loc_dest_x_cpof(v_nf(i).cfop, rNFeImg.ide__idDest) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CFOP " & v_nf(i).cfop
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "O local de destino da operação é conflitante com o CFOP do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   ICMS => VERIFICA SE HOUVE ALTERAÇÃO NO ICMS DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).ICMS_tela) <> "" Then
                If Trim$(v_nf(i).ICMS_tela) <> Trim$(cb_icms) Then
                    If is_venda_interestadual_de_mercadoria_importada(v_nf(i).cfop, v_nf(i).cst) And _
                        (Trim$(v_nf(i).ICMS_tela) = CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA)) Then
                    '   NOP: EM VENDA INTERESTADUAL DE MERCADORIA IMPORTADA É OBRIGATÓRIO USAR A ALÍQUOTA DE ICMS ESPECÍFICA
                    Else
                        If s_msg <> "" Then s_msg = s_msg & vbCrLf
                        s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": ICMS alterado para " & v_nf(i).ICMS_tela & "%"
                        End If
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "Houve alteração no ICMS do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   PREPARA O CAMPO QUE ARMAZENA O ICMS A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).ICMS = cb_icms
            If Trim$(v_nf(i).ICMS_tela) <> "" Then
                v_nf(i).ICMS = Trim$(v_nf(i).ICMS_tela)
                End If
            End If
        Next


'   QUANTIDADE DE LINHAS EXCEDE O TAMANHO DA PÁGINA?
    MAX_LINHAS_NOTA_FISCAL = MAX_LINHAS_NOTA_FISCAL_DEFAULT
    If (Not blnTemPagtoPorBoleto) Then MAX_LINHAS_NOTA_FISCAL = MAX_LINHAS_NOTA_FISCAL_DEFAULT + 2
    
    If qtde_linhas_nf > MAX_LINHAS_NOTA_FISCAL Then
        s = "Não é possível imprimir a nota fiscal porque os " & CStr(qtde_linhas_nf) & _
            " itens excedem o máximo de " & CStr(MAX_LINHAS_NOTA_FISCAL) & _
            " linhas que podem ser impressas!!"
        aviso_erro s
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
'>  ARREDONDAMENTOS
    For ic = LBound(v_nf) To UBound(v_nf)
        With v_nf(ic)
            If Trim$(.produto) <> "" Then
                vl_unitario = .valor_total / .qtde_total
                .valor_total = CCur(Format$(vl_unitario, FORMATO_MOEDA)) * .qtde_total
                End If
            End With
        Next

        
'   CONSISTE DADOS
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).ncm) = "" Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO possui o código NCM!!"
            ElseIf Len(Trim$(v_nf(i).cst)) = 0 Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO possui a informação do CST!!"
            ElseIf Len(Trim$(v_nf(i).cst)) <> 3 Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " possui o campo CST preenchido com valor inválido!!"
                End If
            End If
        Next

    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
        
'   SE FOR NOTA DE ENTRADA, VERIFICA SE A DEVOLUÇÃO DE MERCADORIAS FOI INTEGRAL
'   0-Entrada  1-Saída
    s_msg = ""
    If rNFeImg.ide__tpNF = "0" Then
        For i = LBound(v_nf) To UBound(v_nf)
            If Trim$(v_nf(i).produto) <> "" Then
                s = "SELECT" & _
                        " Coalesce(Sum(qtde),0) AS qtde_total_devolvida" & _
                    " FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
                    " WHERE" & _
                        " (" & sql_monta_criterio_texto_or(v_pedido(), "pedido", True) & ")" & _
                        " AND (fabricante = '" & v_nf(i).fabricante & "')" & _
                        " AND (produto = '" & v_nf(i).produto & "')"
                If t_PEDIDO_ITEM_DEVOLVIDO.State <> adStateClosed Then t_PEDIDO_ITEM_DEVOLVIDO.Close
                t_PEDIDO_ITEM_DEVOLVIDO.Open s, dbc, , , adCmdText
                If t_PEDIDO_ITEM_DEVOLVIDO.EOF Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO teve nenhuma unidade devolvida de um total de " & CStr(v_nf(i).qtde_total)
                Else
                    If CLng(t_PEDIDO_ITEM_DEVOLVIDO("qtde_total_devolvida")) <> v_nf(i).qtde_total Then
                        If s_msg <> "" Then s_msg = s_msg & vbCrLf
                        s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " teve " & Trim$("" & t_PEDIDO_ITEM_DEVOLVIDO("qtde_total_devolvida")) & " unidade(s) devolvida(s) de um total de " & CStr(v_nf(i).qtde_total)
                        End If
                    End If
                End If
            Next
        
        If s_msg <> "" Then
            s_msg = "Não é possível emitir esta NFe de entrada através do painel de emissão automática porque o pedido não teve os produtos devolvidos integralmente:" & _
                    vbCrLf & _
                    s_msg
            End If
        End If
    
    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    

'   OBTÉM DADOS DA TRANSPORTADORA
    strTransportadoraCnpj = ""
    strTransportadoraRazaoSocial = ""
    strTransportadoraIE = ""
    strTransportadoraUF = ""
    strTransportadoraEmail = ""
    If strTransportadoraId <> "" Then
        s = "SELECT * FROM t_TRANSPORTADORA WHERE id = '" & strTransportadoraId & "'"
        t_TRANSPORTADORA.Open s, dbc, , , adCmdText
        If t_TRANSPORTADORA.EOF Then
            s = "Transportadora '" & strTransportadoraId & "' não está cadastrada!!"
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            strTransportadoraCnpj = retorna_so_digitos(Trim$("" & t_TRANSPORTADORA("cnpj")))
            strTransportadoraRazaoSocial = UCase$(Trim$("" & t_TRANSPORTADORA("razao_social")))
            strTransportadoraIE = Trim$("" & t_TRANSPORTADORA("ie"))
            strTransportadoraUF = Trim$("" & t_TRANSPORTADORA("uf"))
            strTransportadoraEmail = Trim$("" & t_TRANSPORTADORA("email"))
            End If
        
        If (strTransportadoraCnpj = "") Or (strTransportadoraRazaoSocial = "") Then
            s = ""
            If strTransportadoraCnpj = "" Then
                If s <> "" Then s = s & vbCrLf
                s = s & "A transportadora '" & strTransportadoraId & "' não possui CNPJ cadastrado!!"
                End If
                
            If strTransportadoraRazaoSocial = "" Then
                If s <> "" Then s = s & vbCrLf
                s = s & "A transportadora '" & strTransportadoraId & "' não possui razão social cadastrada!!"
                End If
            
            If s <> "" Then
                s = s & vbCrLf & "Continua mesmo assim?"
                End If
            
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
            
''   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
'    s = "SELECT * FROM t_CLIENTE WHERE (id='" & Trim$("" & t_PEDIDO("id_cliente")) & "')"
'    t_DESTINATARIO.Open s, dbc, , , adCmdText
'    If t_DESTINATARIO.EOF Then
'        s = "Cliente com nº registro " & Trim$("" & t_PEDIDO("id_cliente")) & " não foi encontrado!!"
'        aviso_erro s
'        GoSub NFE_EMITE_FECHA_TABELAS
'        aguarde INFO_NORMAL, m_id
'        Exit Sub
'        End If
        
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    'PRIMEIRO CASO: A MEMORIZAÇÃO DO ENDEREÇO DO CLIENTE NA TABELA DE PEDIDOS ESTÁ OK
    blnExisteMemorizacaoEndereco = False
    If param_pedidomemorizacaoenderecos.campo_inteiro = 1 Then
        s = "SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, endereco_uf as uf, endereco_cnpj_cpf as cnpj_cpf, " & _
                " endereco_logradouro as endereco, endereco_bairro as bairro, endereco_cidade as cidade, endereco_cep as cep, endereco_numero, endereco_complemento, " & _
                " endereco_logradouro as endereco_end_nota, " & _
                " endereco_bairro as bairro_end_nota, " & _
                " endereco_cidade as cidade_end_nota, " & _
                " endereco_cep as cep_end_nota, " & _
                " endereco_numero as numero_end_nota, " & _
                " endereco_complemento as complemento_end_nota, " & _
                " endereco_uf as uf_end_nota, " & _
                " endereco_email as email, endereco_email_xml as email_xml, " & _
                " endereco_nome as nome, " & _
                " endereco_ddd_res as ddd_res, endereco_tel_res as tel_res, " & _
                " endereco_ddd_com as ddd_com, endereco_tel_com as tel_com, endereco_ramal_com as ramal_com, " & _
                " endereco_ddd_cel as ddd_cel, endereco_tel_cel as tel_cel, " & _
                " endereco_ddd_com_2 as ddd_com_2, endereco_tel_com_2 as tel_com_2, endereco_ramal_com_2 as ramal_com_2, " & _
                " endereco_tipo_pessoa as tipo, " & _
                " endereco_contribuinte_icms_status as contribuinte_icms_status, " & _
                " endereco_produtor_rural_status as produtor_rural_status, " & _
                " endereco_ie as ie, " & _
                " endereco_rg as rg, " & _
                " endereco_contato as contato " & _
            " FROM t_PEDIDO" & _
            " WHERE (pedido = '" & Trim$("" & t_PEDIDO("pedido")) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PJ & "')"
        s = s & " UNION" & _
            " SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, endereco_uf as uf, endereco_cnpj_cpf as cnpj_cpf, " & _
                " endereco_logradouro as endereco, endereco_bairro as bairro, endereco_cidade as cidade, endereco_cep as cep, endereco_numero, endereco_complemento, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_logradouro else EndEtg_endereco end as endereco_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_bairro else EndEtg_bairro end as bairro_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cidade else EndEtg_cidade end as cidade_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cep else EndEtg_cep end as cep_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_numero else EndEtg_endereco_numero end as numero_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_complemento else EndEtg_endereco_complemento end as complemento_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_uf else EndEtg_uf end as uf_end_nota, " & _
                " endereco_email as email, endereco_email_xml as email_xml, " & _
                " endereco_nome as nome, " & _
                " endereco_ddd_res as ddd_res, endereco_tel_res as tel_res, " & _
                " endereco_ddd_com as ddd_com, endereco_tel_com as tel_com, endereco_ramal_com as ramal_com, " & _
                " endereco_ddd_cel as ddd_cel, endereco_tel_cel as tel_cel, " & _
                " endereco_ddd_com_2 as ddd_com_2, endereco_tel_com_2 as tel_com_2, endereco_ramal_com_2 as ramal_com_2, " & _
                " endereco_tipo_pessoa as tipo, " & _
                " endereco_contribuinte_icms_status as contribuinte_icms_status, " & _
                " endereco_produtor_rural_status as produtor_rural_status, " & _
                " endereco_ie as ie, " & _
                " endereco_rg as rg, " & _
                " endereco_contato as contato " & _
            " FROM t_PEDIDO" & _
            " WHERE (pedido = '" & Trim$("" & t_PEDIDO("pedido")) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PF & "')"
        t_DESTINATARIO.Open s, dbc, , , adCmdText
        If t_DESTINATARIO.EOF Then
            s = "Problemas na localização do endereço memorizado no pedido " & Trim$("" & t_PEDIDO("pedido")) & "!!"
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        If t_DESTINATARIO("st_memorizacao_completa_enderecos") > 0 Then blnExisteMemorizacaoEndereco = True
        End If
        
    'SEGUNDO CASO: A MEMORIZAÇÃO DO ENDEREÇO DO CLIENTE NA TABELA DE PEDIDOS NÃO ESTÁ OK
    If Not blnExisteMemorizacaoEndereco Then
        If t_DESTINATARIO.State <> adStateClosed Then t_DESTINATARIO.Close
    '   (se não houver memorização no pedido)
        s = "SELECT * FROM t_CLIENTE WHERE (id='" & Trim$("" & t_PEDIDO("id_cliente")) & "')"
        t_DESTINATARIO.Open s, dbc, , , adCmdText
        If t_DESTINATARIO.EOF Then
            s = "Cliente com nº registro " & Trim$("" & t_PEDIDO("id_cliente")) & " não foi encontrado!!"
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
        
        
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
    
'   CONFIRMA ALÍQUOTA DO ICMS
'    If usuario.emit_uf = "ES" Then
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "ES": strIcms = "17"
'            Case "RJ", "SP", "PR", "SC", "RS", "MG", "GO", "TO", "MT", "MS", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "12"
'            Case Else: strIcms = ""
'            End Select
'    ElseIf usuario.emit_uf = "MG" Then
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "MG": strIcms = "18"
'            Case "RJ", "SP", "PR", "SC", "RS": strIcms = "12"
'            Case "ES", "GO", "TO", "MT", "MS", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "7"
'            Case Else: strIcms = ""
'            End Select
'    ElseIf usuario.emit_uf = "MS" Then
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "MS": strIcms = "17"
'            Case "RJ", "MG", "PR", "SC", "RS", "ES", "GO", "TO", "MT", "SP", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "12"
'            Case Else: strIcms = ""
'            End Select
'    ElseIf usuario.emit_uf = "RJ" Then
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "RJ": strIcms = "19"
'            Case "MG", "SP", "PR", "SC", "RS": strIcms = "12"
'            Case "ES", "GO", "TO", "MT", "MS", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "7"
'            Case Else: strIcms = ""
'            End Select
'    ElseIf usuario.emit_uf = "TO" Then
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "TO": strIcms = "17"
'            Case "RJ", "MG", "PR", "SC", "RS", "ES", "GO", "MS", "MT", "SP", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "12"
'            Case Else: strIcms = ""
'            End Select
'    Else
'        Select Case UCase$(Trim$("" & t_DESTINATARIO("uf")))
'            Case "SP": strIcms = "18"
'            Case "RJ", "MG", "PR", "SC", "RS": strIcms = "12"
'            Case "ES", "GO", "TO", "MT", "MS", "AM", "AP", "RR", "RO", "AC", "PB", "MA", "PI", "CE", "RN", "BA", "PE", "AL", "SE", "DF", "PA": strIcms = "7"
'            Case Else: strIcms = ""
'            End Select
'        End If
    If obtem_aliquota_ICMS(usuario.emit_uf, UCase$(Trim$("" & t_DESTINATARIO("uf"))), aliquota_icms_interestadual) Then
        strIcms = Trim$(CStr(aliquota_icms_interestadual))
    Else
        strIcms = ""
        End If
    
    If (strIcms <> "") And (cb_icms <> "") Then
        If (CSng(strIcms) <> CSng(cb_icms)) Then
            s = "O destinatário é do estado de " & UCase$(Trim$("" & t_DESTINATARIO("uf"))) & " cuja alíquota de ICMS é de " & strIcms & "%" & _
                vbCrLf & "Confirma a emissão da NFe usando a alíquota de " & cb_icms & "%?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
            End If
        End If
        
'   MERCADORIA IMPORTADA EM VENDA INTERESTADUAL: VERIFICA SE ESTÁ C/ ALÍQUOTA DE ICMS ESPECÍFICA
'   NÃO EXIBIR ALERTA P/ PESSOA FÍSICA (EXCETO PRODUTOR RURAL CONTRIBUINTE DO ICMS) OU SE FOR PJ ISENTA DE I.E.
    If ((Len(retorna_so_digitos(Trim("" & t_DESTINATARIO("cnpj_cpf")))) = 14) And _
        (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) Or _
       ((Len(retorna_so_digitos(Trim("" & t_DESTINATARIO("cnpj_cpf")))) = 14) And _
       (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL) And _
        (InStr(UCase$(Trim$("" & t_DESTINATARIO("ie"))), "ISEN") = 0)) Or _
       ((t_DESTINATARIO("produtor_rural_status") = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) And _
        (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) Then
        s_confirma = ""
        For i = LBound(v_nf) To UBound(v_nf)
            If Trim$(v_nf(i).produto) <> "" Then
                If is_venda_interestadual_de_mercadoria_importada(v_nf(i).cfop, v_nf(i).cst) Then
                    If Trim$(v_nf(i).ICMS) <> CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA) Then
                        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                        s_confirma = s_confirma & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " está com ICMS de " & v_nf(i).ICMS & "% ao invés de " & CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA) & "%"
                        End If
                    End If
                End If
            Next
        
        If s_confirma <> "" Then
            s_confirma = "Foram encontradas possíveis incoerências na alíquota do ICMS na venda interestadual de mercadoria importada:" & _
                    vbCrLf & _
                    s_confirma & _
                    vbCrLf & vbCrLf & _
                    "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
            End If
        End If
    
    
'   SE HÁ PEDIDO ESPECIFICANDO PAGAMENTO VIA BOLETO BANCÁRIO, CALCULA QUANTIDADE DE PARCELAS, DATAS E VALORES
'   DOS BOLETOS. ESSES DADOS SERÃO IMPRESSOS NA NF E TAMBÉM SALVOS NO BD, POIS SERVIRÃO DE BASE PARA A GERAÇÃO
'   DOS BOLETOS NO ARQUIVO DE REMESSA.
    If (param_geracaoboletos.campo_texto = "Manual") And blnExisteParcelamentoBoleto Then
        ReDim v_parcela_pagto(UBound(v_parcela_manual_boleto))
        v_parcela_pagto = v_parcela_manual_boleto
    Else
        ReDim v_parcela_pagto(0)
        If Not geraDadosParcelasPagto(v_pedido(), v_parcela_pagto(), s_erro) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            If s_erro <> "" Then s_erro = Chr(13) & Chr(13) & s_erro
            s_erro = "Falha ao tentar processar os dados de pagamento!!" & s_erro
            aviso_erro s_erro
            Exit Sub
            End If
        End If
        
'   Tipo de NFe: 0-Entrada  1-Saída
    If rNFeImg.ide__tpNF = "1" Then
        s = ""
        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
            If v_parcela_pagto(i).intNumDestaParcela <> 0 Then
                blnImprimeDadosFatura = True
                If s <> "" Then s = s & Chr(13)
                s = s & "Parcela:  " & v_parcela_pagto(i).intNumDestaParcela & "/" & v_parcela_pagto(i).intNumTotalParcelas & " para " & Format$(v_parcela_pagto(i).dtVencto, FORMATO_DATA) & " de " & SIMBOLO_MONETARIO & " " & Format$(v_parcela_pagto(i).vlValor, FORMATO_MOEDA) & " (" & descricao_opcao_forma_pagamento(v_parcela_pagto(i).id_forma_pagto) & ")"
                End If
            Next
            
        If (s <> "") And Not blnRemessaEntregaFutura Then
            s = "Serão emitidas na NFe as seguintes informações de pagamento:" & Chr(13) & Chr(13) & s
            If DESENVOLVIMENTO Then
                aviso s
                End If
            End If
        End If
    
' na emissão da nota de venda, deve ser verificado o CFOP?
''   VERIFICA SE O CFOP ESTÁ COERENTE COM O CST DO ICMS
'    s_confirma = ""
'    For i = LBound(v_nf) To UBound(v_nf)
'        If Trim$(v_nf(i).produto) <> "" Then
'            strNFeCst = Trim$(right$(v_nf(i).cst, 2))
'            strCfopCodigoAux = Trim$(v_nf(i).cfop)
'            strCfopCodigoFormatadoAux = Trim$(v_nf(i).CFOP_formatado)
'            s = "O produto " & v_nf(i).produto & " possui CST = " & strNFeCst & ", mas o CFOP selecionado é " & strCfopCodigoFormatadoAux
'            If strNFeCst = "00" Then
'                If (strCfopCodigoAux = "5102") Or (strCfopCodigoAux = "6102") Then s = ""
'            ElseIf strNFeCst = "60" Then
'                If (strCfopCodigoAux = "5405") Or (strCfopCodigoAux = "6404") Then s = ""
'            Else
'                If (strCfopCodigoAux <> "5102") And (strCfopCodigoAux <> "6102") And _
'                   (strCfopCodigoAux <> "5405") And (strCfopCodigoAux <> "6404") Then s = ""
'                End If
'
'            If s <> "" Then
'                If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
'                s_confirma = s_confirma & s
'                End If
'            End If
'        Next
'
'    If s_confirma <> "" Then
'        s_confirma = "Foram encontradas possíveis incoerências entre CFOP e CST:" & _
'                     vbCrLf & _
'                     s_confirma & _
'                     vbCrLf & vbCrLf & _
'                     "Continua mesmo assim?"
'        If Not confirma(s_confirma) Then
'            GoSub NFE_EMITE_FECHA_TABELAS
'            aguarde INFO_NORMAL, m_id
'            Exit Sub
'            End If
'        End If


'   VERIFICAR SE É NOTA DE COMPROMISSO
    blnNotadeCompromisso = False
    If ((strCfopCodigo = "5922") Or (strCfopCodigo = "6922")) Then
        blnNotadeCompromisso = True
        End If
        
'   VERIFICAR SE É NOTA DE REMESSA DE ENTREGA FUTURA
    blnRemessaEntregaFutura = False
    If ((strCfopCodigo = "5117") Or (strCfopCodigo = "6117")) Then
        blnRemessaEntregaFutura = True
        End If
    
'   NÃO PERMITIR EMISSÃO DE NOTA DE COMPROMISSO, SENÃO A OPERAÇÃO TRIANGULAR NÃO PERMITIRÁ A NOTA DE VENDA
    If blnNotadeCompromisso Then
        aviso "Não é possível emitir NOTAS FUTURAS no painel triangular!!!" & vbCrLf & _
                "Emitir nos painéis automático ou manual"
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
'   CASO SEJA NOTA DE COMPROMISSO, VERIFICAR SE O CST É 041
    If blnNotadeCompromisso Then
        s_confirma = ""
        For i = LBound(v_nf) To UBound(v_nf)
            If Trim$(v_nf(i).produto) <> "" Then
                strNFeCst = Trim$(right$(v_nf(i).cst, 2))
                If strNFeCst <> "41" Then
                    s = "O o produto " & v_nf(i).produto & " possui CST diferente de 41"
                    End If
                If s <> "" Then
                    If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                    s_confirma = s_confirma & s
                    End If
                End If
            Next
        If s_confirma <> "" Then
            s_confirma = "PROBLEMAS COM CST EM PEDIDO DE VENDA FUTURA:" & _
                         vbCrLf & _
                         s_confirma & _
                         vbCrLf & vbCrLf & _
                         "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If

'   CASO O PEDIDO PAI SEJA PARA PAGAMENTO ANTECIPADO, VERIFICA SE O PEDIDO FILHO ESTÁ QUITADO
'   (não permitir emissão se não for nota de compromisso)
    If (strPagtoAntecipadoStatus = "1") And (strPagtoAntecipadoQuitadoStatus <> "1") Then
        If Not blnNotadeCompromisso Then
            s = "Pedido " & Trim$(v_pedido(i)) & " se refere a venda futura não quitada!"
            aviso s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
        

'   ZERAR PIS/COFINS?
    s_confirma = ""
    If Trim$(cb_zerar_PIS) <> "" Then
        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
        s_confirma = s_confirma & "Alíquota do PIS será zerada usando CST = " & cb_zerar_PIS
        End If
    
    If Trim$(cb_zerar_COFINS) <> "" Then
        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
        s_confirma = s_confirma & "Alíquota do COFINS será zerada usando CST = " & cb_zerar_COFINS
        End If
    
    If s_confirma <> "" Then
        s_confirma = s_confirma & _
                     vbCrLf & vbCrLf & _
                     "Continua mesmo assim?"
        If Not confirma(s_confirma) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    

'   CALCULA TOTAL ESTIMADO DOS TRIBUTOS USANDO DADOS DO IBPT?
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    s_confirma = ""
    If is_venda_consumidor_final(strCfopCodigo) Then
        blnExibirTotalTributos = True
    '   OBTÉM DADOS DO IBPT P/ CALCULAR TOTAL ESTIMADO DOS TRIBUTOS
        For i = LBound(v_nf) To UBound(v_nf)
            With v_nf(i)
                If Trim$(.produto) <> "" Then
                    s = "SELECT " & _
                            "*" & _
                        " FROM t_IBPT" & _
                        " WHERE" & _
                            " (codigo = '" & Trim$(.ncm) & "')" & _
                            " AND (tabela = '0')" & _
                        " ORDER BY" & _
                            " codigo," & _
                            " ex"
                    If t_IBPT.State <> adStateClosed Then t_IBPT.Close
                    t_IBPT.Open s, dbc, , , adCmdText
                    If t_IBPT.EOF Then
                        blnHaProdutoSemDadosIbpt = True
                        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                        s_confirma = s_confirma & "O NCM '" & Trim$(.ncm) & "' NÃO está cadastrado na tabela do IBPT!!"
                    Else
                        .tem_dados_IBPT = True
                        .percAliqNac = t_IBPT("percAliqNac")
                        .percAliqImp = t_IBPT("percAliqImp")
                        End If
                    End If
                End With
            Next
        
        If s_confirma <> "" Then
            s_confirma = s_confirma & _
                         "A nota fiscal será emitida sem a informação do total estimado dos tributos conforme exige a lei 12.741/2012!!" & _
                         vbCrLf & vbCrLf & _
                         "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
'   VERIFICAR DIVERGÊNCIA DE LOCAL DE DESTINO DA OPERAÇÃO
    If rNFeImg.ide__tpNF <> "0" Then
        s_confirma = ""
        If strEndEtgUf <> "" Then
            strDestinoUF = strEndEtgUf
        Else
            strDestinoUF = strEndClienteUf
            End If
        'primeira situação: UFs diferentes e Local de Destino  <> Interestadual
        If (Trim$(rNFeImg.ide__idDest) <> "2") And (strOrigemUF <> strDestinoUF) Then
            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
            s_confirma = s_confirma & "UF de origem e destino da Nota são diferentes, porém local de operação selecionado é " & vbCrLf & vbCrLf
            s_confirma = s_confirma & cb_loc_dest
            End If
        
        If (Trim$(rNFeImg.ide__idDest) <> "1") And (strOrigemUF = strDestinoUF) Then
            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
            s_confirma = s_confirma & "UF de origem e destino da Nota são iguais, porém local de operação selecionado é " & vbCrLf & vbCrLf
            s_confirma = s_confirma & cb_loc_dest
            End If
        
        If s_confirma <> "" Then
            s_confirma = s_confirma & _
                         vbCrLf & vbCrLf & _
                         "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
            End If
        End If
    

'   PREPARA DADOS DA NFe
    aguarde INFO_EXECUTANDO, "preparando emissão da NFe"
    
'   TAG OPERACIONAL
'   ~~~~~~~~~~~~~~~
    strNFeTagOperacional = "operacional;" & vbCrLf

'   EMAIL DO DESTINATÁRIO DA NFe
    'para a loja 201, caso o campo pedido_bs_x_marketplace indique ser um pedido de marketplace, desconsiderar o e-mail do cliente
    If (strLoja = "201") And (strPedidoBSMarketplace <> "") Then
        rNFeImg.operacional__email = ""
    Else
        rNFeImg.operacional__email = Trim("" & t_DESTINATARIO("email"))
        End If
    If (Trim$(rNFeImg.operacional__email) <> "") And (Trim$(strTransportadoraEmail) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
    rNFeImg.operacional__email = rNFeImg.operacional__email & strTransportadoraEmail
    strEmailXML = Trim("" & t_DESTINATARIO("email_xml"))
    If Trim$(strEmailXML) <> "" Then
        If (Trim$(rNFeImg.operacional__email) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
        rNFeImg.operacional__email = rNFeImg.operacional__email & strEmailXML
        End If

    If rNFeImg.operacional__email <> "" Then
        strNFeTagOperacional = strNFeTagOperacional & _
                               vbTab & NFeFormataCampo("email", rNFeImg.operacional__email)
        End If
    
'   TAG DEST (DADOS DO DESTINATÁRIO)
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNFeTagDestinatario = "dest;" & vbCrLf
    
'   CNPJ/CPF
    strDestinatarioCnpjCpf = retorna_so_digitos(Trim("" & t_DESTINATARIO("cnpj_cpf")))
    If strDestinatarioCnpjCpf = "" Then
        s_erro = "CNPJ/CPF do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Not cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
        s_erro = "CNPJ/CPF do cliente está cadastrado com informação inválida!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    
    If Len(strDestinatarioCnpjCpf) = 11 Then
        blnIsDestinatarioPJ = False
        rNFeImg.dest__CPF = strDestinatarioCnpjCpf
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CPF", rNFeImg.dest__CPF)
    ElseIf Len(strDestinatarioCnpjCpf) = 14 Then
        blnIsDestinatarioPJ = True
        rNFeImg.dest__CNPJ = strDestinatarioCnpjCpf
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CNPJ", rNFeImg.dest__CNPJ)
        End If
        
'   CAMPO: idEstrangeiro
    rNFeImg.dest__idEstrangeiro = ""
    If Trim(rNFeImg.dest__idEstrangeiro) <> "" Then
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("idEstrangeiro", rNFeImg.dest__idEstrangeiro)
        End If
    
'   NOME
    If NFE_AMBIENTE = NFE_AMBIENTE_HOMOLOGACAO Then
        strCampo = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
    Else
        strCampo = Trim("" & t_DESTINATARIO("nome"))
        End If
    If strCampo = "" Then
        s_erro = "O nome do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O nome do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xNome = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xNome", rNFeImg.dest__xNome)
    
'   LOGRADOURO
    If blnExisteMemorizacaoEndereco Then
        strCampo = Trim("" & t_DESTINATARIO("endereco_end_nota"))
    Else
        strCampo = Trim("" & t_DESTINATARIO("endereco"))
        End If
    If strCampo = "" Then
        s_erro = "O endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xLgr = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xLgr", rNFeImg.dest__xLgr)
    
'   ENDEREÇO: NÚMERO
    strCampo = Trim$("" & t_DESTINATARIO("endereco_numero"))
    If strCampo = "" Then
        s_erro = "O endereço no cadastro do cliente deve ser preenchido corretamente para poder emitir a NFe!!" & vbCrLf & _
                 "As informações de número e complemento do endereço devem ser preenchidas nos campos adequados!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O número do endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__nro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("nro", rNFeImg.dest__nro)
        
'   ENDEREÇO: COMPLEMENTO
    strCampo = Trim$("" & t_DESTINATARIO("endereco_complemento"))
    If Len(strCampo) > 60 Then
        s_erro = "O campo complemento do endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xCpl = strCampo
    If Len(strCampo) > 0 Then strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xCpl", rNFeImg.dest__xCpl)
    
'   BAIRRO
    strCampo = Trim$("" & t_DESTINATARIO("bairro"))
    If strCampo = "" Then
        s_erro = "O campo bairro no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O campo bairro no endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xBairro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xBairro", rNFeImg.dest__xBairro)
    
'   MUNICIPIO
    strCampo = Trim$("" & t_DESTINATARIO("cidade"))
    s_aux = Trim$("" & t_DESTINATARIO("uf"))
    If (strCampo <> "") And (s_aux <> "") Then strCampo = strCampo & "/"
    strCampo = strCampo & s_aux
    rNFeImg.dest__cMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cMun", rNFeImg.dest__cMun)
    
    strCampo = Trim$("" & t_DESTINATARIO("cidade"))
    If Len(strCampo) > 60 Then
        s_erro = "O campo cidade no endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xMun", rNFeImg.dest__xMun)
    
'   UF
    strCampo = Trim$("" & t_DESTINATARIO("uf"))
    If strCampo = "" Then
        s_erro = "O campo UF no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__UF = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("UF", rNFeImg.dest__UF)
    
'   MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
    If Not consiste_municipio_IBGE_ok(dbcNFe, rNFeImg.dest__xMun, rNFeImg.dest__UF, strListaSugeridaMunicipiosIBGE, s_erro_aux) Then
        If s_erro_aux <> "" Then
            s_erro = s_erro_aux
        Else
            s_erro = "Município '" & rNFeImg.dest__xMun & "' não consta na relação de municípios do IBGE para a UF de '" & rNFeImg.dest__UF & "'!!"
            End If
            
        If s_erro <> "" Then s_erro = s_erro & Chr(13)
        s_erro = s_erro & "Será necessário corrigir o município no cadastro do cliente antes de prosseguir!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If

'   CEP
    strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep")))
    If strCampo = "" Then
        s_erro = "O campo CEP no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__CEP = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CEP", rNFeImg.dest__CEP)
    
'   PAÍS
    rNFeImg.dest__cPais = "1058"
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cPais", rNFeImg.dest__cPais)
    rNFeImg.dest__xPais = "BRASIL"
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xPais", rNFeImg.dest__xPais)
    
'   FONE
    strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_cel")))
    If strCampo <> "" Then
        If Len(strCampo) > 9 Then
            s_erro = "O telefone celular no cadastro do destinatário excede o tamanho máximo permitido!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
            
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_cel")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        
        If strDDD = "" Then
            s_erro = "O DDD do telefone celular no cadastro do destinatário não está preenchido!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        ElseIf Len(strDDD) > 2 Then
            s_erro = "O DDD do telefone celular no cadastro do destinatário excede o tamanho máximo permitido!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
        strCampo = strDDD & strCampo
        strTelCel = strCampo
        End If
    
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O telefone residencial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
                
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do telefone residencial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone residencial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelRes = strCampo
            End If
        End If
        
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
        
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelCom = strCampo
            End If
        End If
        
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O segundo telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
        
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do segundo telefone comercial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelCom2 = strCampo
            End If
        End If
    If strCampo <> "" Then
        rNFeImg.dest__fone = strCampo
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("fone", rNFeImg.dest__fone)
        End If
        
    'preencher os campos de telefone que possam estar vazios
    If strTelRes = "" Then strTelRes = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res"))))
    If strTelCom = "" Then strTelCom = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com"))))
    If strTelCom2 = "" Then strTelCom2 = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2"))))
    If strTelRes <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelRes = "(" & strDDD & ")" & strTelRes
        End If
    If strTelCom <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com")))
        If (Len(strDDD) = 2) Then strTelCom = "(" & strDDD & ") " & strTelCom
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom = strTelCom & " R: " & strRamal
        End If
    If strTelCom2 <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com_2")))
        If (Len(strDDD) = 2) Then strTelCom2 = "(" & strDDD & ") " & strTelCom2
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom2 = strTelCom2 & " R: " & strRamal
        End If
        
        
'   CAMPO: indIEDest
    intContribuinteICMS = t_DESTINATARIO("contribuinte_icms_status")
    
    'Conforme orientação da Bueno Consultoria e Assessoria Contábil, em e-mail encaminhado em 22/06/2016,
    'deve-se informar a identificação da IE do destinatário como "Contribuinte do ICMS" ou "Não Contribuinte"
    If intContribuinteICMS = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO Then intContribuinteICMS = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO
    
    strCampo = Trim$("" & t_DESTINATARIO("ie"))
    If intContribuinteICMS = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM Then
        'Primeira situação: o campo Contribuinte ICMS está preenchido com Sim
        If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
        If ConsisteInscricaoEstadual(strCampo, rNFeImg.dest__UF) <> 0 Then
        '   Retorno = 0 -> IE válida
        '   Retorno = 1 -> IE inválida
            s_erro = "A Inscrição Estadual no cadastro do cliente (" & strCampo & ") é inválida para a UF de '" & rNFeImg.dest__UF & "'!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        ElseIf InStr(UCase$(strCampo), "ISEN") > 0 Then
            s_erro = "Cliente está marcado como Contribuinte, porém Inscrição Estadual apresenta valor (" & strCampo & ")!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        Else
        '   1 = CONTRIBUINTE ICMS
                rNFeImg.dest__indIEDest = "1"
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
            End If
    ElseIf intContribuinteICMS = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO Then
        'Segunda situação: o campo Contribuinte ICMS está preenchido com Não
        '   9 = NÃO-CONTRIBUINTE
        If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
        If (Trim$(strCampo) <> "") And (ConsisteInscricaoEstadual(strCampo, rNFeImg.dest__UF) <> 0) Then
        '   Retorno = 0 -> IE válida
        '   Retorno = 1 -> IE inválida
            s_erro = "A Inscrição Estadual no cadastro do cliente (" & strCampo & ") é inválida para a UF de '" & rNFeImg.dest__UF & "'!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
        rNFeImg.dest__indIEDest = "9"
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
    ElseIf intContribuinteICMS = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO Then
        'Terceira situação: o campo Contribuinte ICMS está preenchido com Isento
        '   2 = CONTRIBUINTE ISENTO
        rNFeImg.dest__indIEDest = "2"
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
    Else
        'Quarta situação: o campo Contribuinte ICMS não está preenchido
        If blnIsDestinatarioPJ Then
            If InStr(UCase$(strCampo), "ISEN") > 0 Then
                strCampo = "ISENTO"
                End If
            If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
            If strCampo = "" Then
                s_erro = "A Inscrição Estadual no cadastro do cliente está vazia ou está preenchida com conteúdo inválido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf (Len(strCampo) < 2) Or (Len(strCampo) > 14) Then
                s_erro = "A Inscrição Estadual no cadastro do cliente está preenchida com conteúdo inválido!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf ConsisteInscricaoEstadual(strCampo, rNFeImg.dest__UF) <> 0 Then
            '   Retorno = 0 -> IE válida
            '   Retorno = 1 -> IE inválida
                s_erro = "A Inscrição Estadual no cadastro do cliente (" & strCampo & ") é inválida para a UF de '" & rNFeImg.dest__UF & "'!!"
                GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            
            If strCampo = "ISENTO" Then
            '   2 = CONTRIBUINTE ISENTO
                rNFeImg.dest__indIEDest = "2"
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
            Else
            '   1 = CONTRIBUINTE ICMS
                rNFeImg.dest__indIEDest = "1"
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
                End If
        Else
        '   9 = NÃO-CONTRIBUINTE
            rNFeImg.dest__indIEDest = "9"
            strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
            End If
        End If
        
'   IE
    strCampo = Trim$("" & t_DESTINATARIO("ie"))
    If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
    If rNFeImg.dest__indIEDest = "1" Then
        'Primeira situação: o cliente é contribuinte do ICMS
        If InStr(UCase$(strCampo), "ISEN") > 0 Then
            s_erro = "Cliente está marcado como Contribuinte, porém Inscrição Estadual apresenta valor (" & strCampo & ")!!"
            GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
        rNFeImg.dest__IE = strCampo
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("IE", rNFeImg.dest__IE)
    ElseIf rNFeImg.dest__indIEDest = "9" Then
        'Segunda situação: o cliente não é contribuinte do ICMS
        If InStr(UCase$(strCampo), "ISEN") > 0 Then strCampo = ""
        If strCampo <> "" Then
            rNFeImg.dest__IE = strCampo
            strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("IE", rNFeImg.dest__IE)
            End If
        'Terceira situação: o cliente é isento
        'Não enviar a inscrição estadual
        End If
    
''>  DADOS DA FATURA
'    If blnImprimeDadosFatura Then
'        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'            With v_parcela_pagto(i)
'                If .intNumDestaParcela <> 0 Then
'                    If Trim$(vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc) <> "" Then
'                        ReDim Preserve vNFeImgTagDup(UBound(vNFeImgTagDup) + 1)
'                        End If
'
'                '   FORMA DE PAGTO
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup = abreviacao_opcao_forma_pagamento(.id_forma_pagto)
'                    s = vbTab & NFeFormataCampo("nDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup)
'                '   VENCTO
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc = NFeFormataData(.dtVencto)
'                    s = s & vbTab & NFeFormataCampo("dVenc", vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc)
'                '   VALOR
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup = NFeFormataMoeda2Dec(.vlValor)
'                    s = s & vbTab & NFeFormataCampo("vDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup)
'                '   ADICIONA PARCELA À TAG
'                    strNFeTagDup = strNFeTagDup & "dup;" & vbCrLf & s
'                    End If
'                End With
'            Next
'        End If

'>  DADOS DA FATURA
    If blnImprimeDadosFatura Then
        vl_aux = 0
        strInfoAdicParc = ""
        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
            With v_parcela_pagto(i)
                If .intNumDestaParcela <> 0 Then
                    If Trim$(vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc) <> "" Then
                        ReDim Preserve vNFeImgTagDup(UBound(vNFeImgTagDup) + 1)
                        End If
                        
                '   FORMA DE PAGTO
                    If blnInfoAdicParc Then
                        vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup = NFeFormataSerieNF(i + 1)
                        If strInfoAdicParc <> "" Then strInfoAdicParc = strInfoAdicParc & " / "
                        strInfoAdicParc = strInfoAdicParc & "Parcela " & NFeFormataSerieNF(i + 1) & " - " & _
                                            abreviacao_opcao_forma_pagamento(.id_forma_pagto) & " - " & _
                                            "Vencto: " & .dtVencto & " - " & _
                                            "Valor: " & NFeFormataMoeda2Dec(.vlValor)
                    Else
                        vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup = NFeFormataSerieNF(i + 1)
                        End If
                    s = vbTab & NFeFormataCampo("nDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup)
                '   VENCTO
                    vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc = NFeFormataData(.dtVencto)
                    s = s & vbTab & NFeFormataCampo("dVenc", vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc)
                '   VALOR
                    vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup = NFeFormataMoeda2Dec(.vlValor)
                    s = s & vbTab & NFeFormataCampo("vDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup)
                '   ADICIONA PARCELA À TAG
                    strNFeTagDup = strNFeTagDup & "dup;" & vbCrLf & s
                    vl_aux = vl_aux + .vlValor
                    End If
                End With
            Next
        strNFeTagFat = strNFeTagFat & "fat;" & vbCrLf & vbTab & NFeFormataCampo("nFat", "001") _
                                    & vbTab & NFeFormataCampo("vOrig", NFeFormataMoeda2Dec(vl_aux)) _
                                    & vbTab & NFeFormataCampo("vDesc", "0.00") _
                                    & vbTab & NFeFormataCampo("vLiq", NFeFormataMoeda2Dec(vl_aux))
        
        'se as faturas já foram gravadas na nota de compromisso, zerar as tags de parcelamento
        If ExisteDadosParcelasPagto(rNFeImg.pedido, s_erro) Then
            strNFeTagFat = ""
            strNFeTagDup = ""
            End If
                
        End If
        
        
    
'>  LISTA DE PRODUTOS
    vl_total_ICMS = 0
    vl_total_ICMSDeson = 0
    vl_total_ICMS_ST = 0
    vl_total_IPI = 0
    vl_total_produtos = 0
    vl_total_BC_ICMS = 0
    vl_total_BC_ICMS_ST = 0
    vl_total_PIS = 0
    vl_total_COFINS = 0
    vl_total_outras_despesas_acessorias = 0
    total_volumes = 0
    total_peso_bruto = 0
    total_peso_liquido = 0
    cubagem_bruto = 0
    intNumItem = 0
    vl_total_FCPUFDest = 0
    vl_total_ICMSUFDest = 0
    vl_total_ICMSUFRemet = 0
    vl_total_vFCP = 0
    vl_total_vFCPST = 0
    vl_total_vFCPSTRet = 0
    vl_total_vIPIDevol = 0

        
    'detectada necessidade de informar percentual de partilha do ano anterior, no caso de emisão de
    'nota de entrada referente a uma saída do ano anterior; restringir opção de utilização para
    'as notas de entrada com chave referenciada
    intAnoPartilha = Year(Date)
    If (rNFeImg.ide__tpNF = "0") And (Trim(c_chave_nfe_ref) <> "") Then
        s = "Utilizar percentual de partilha do ano anterior?"
        If confirma(s) Then
            intAnoPartilha = intAnoPartilha - 1
            End If
        End If
        
    
    For ic = LBound(v_nf) To UBound(v_nf)
        With v_nf(ic)
            If Trim$(.produto) <> "" Then
                intNumItem = intNumItem + 1
                
                If Trim$(vNFeImgItem(UBound(vNFeImgItem)).det__nItem) <> "" Then
                    ReDim Preserve vNFeImgItem(UBound(vNFeImgItem) + 1)
                    End If
                    
                vNFeImgItem(UBound(vNFeImgItem)).fabricante = .fabricante
                vNFeImgItem(UBound(vNFeImgItem)).produto = .produto
                
            '   TAG DET
            '   ~~~~~~~
            '   NÚMERO DO ITEM
                vNFeImgItem(UBound(vNFeImgItem)).det__nItem = CStr(intNumItem)
                strNFeTagDet = vbTab & NFeFormataCampo("nItem", vNFeImgItem(UBound(vNFeImgItem)).det__nItem)
                
            '   CÓDIGO DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__cProd = .produto
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cProd", vNFeImgItem(UBound(vNFeImgItem)).det__cProd)
                
            '   EAN
                vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = .EAN
                'NFE 4.0 - EM BRANCO, INFORMAR SEM GTIN
                If vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = "" Then vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = "SEM GTIN"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cEAN", vNFeImgItem(UBound(vNFeImgItem)).det__cEAN)
            
            '   DESCRIÇÃO DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__xProd = UCase$(.descricao)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("xProd", vNFeImgItem(UBound(vNFeImgItem)).det__xProd)
                
            '   NCM
                vNFeImgItem(UBound(vNFeImgItem)).det__NCM = .ncm
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("NCM", vNFeImgItem(UBound(vNFeImgItem)).det__NCM)
                
            '=== aqui: campo NVE (não será usado)
            
            '   CEST
                vNFeImgItem(UBound(vNFeImgItem)).det__CEST = retorna_CEST(.ncm)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("CEST", vNFeImgItem(UBound(vNFeImgItem)).det__CEST)
            
            '   Indicador de Escala Relevante
                'CONVÊNIO ICMS 52, DE 7 DE ABRIL DE 2017
                'Cláusula vigésima terceira Os bens e mercadorias relacionados no Anexo XXVII serão considerados fabricados em escala industrial não relevante quando produzidos por contribuinte que atender, cumulativamente, as seguintes condições:
                'I - ser optante pelo Simples Nacional;
                'II - auferir, no exercício anterior, receita bruta igual ou inferior a R$ 180.000,00 (cento e oitenta mil reais);
                'III - possuir estabelecimento único;
                'IV - ser credenciado pela administração tributária da unidade federada de destino dos bens e mercadorias, quando assim exigido.
                vNFeImgItem(UBound(vNFeImgItem)).det__indEscala = "S"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("indEscala", "S")
                
            '   CFOP
                vNFeImgItem(UBound(vNFeImgItem)).det__CFOP = .cfop
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("CFOP", vNFeImgItem(UBound(vNFeImgItem)).det__CFOP)
            
            '   UNIDADE COMERCIAL
                vNFeImgItem(UBound(vNFeImgItem)).det__uCom = "PC"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uCom", vNFeImgItem(UBound(vNFeImgItem)).det__uCom)
                
            '   QUANTIDADE
                vNFeImgItem(UBound(vNFeImgItem)).det__qCom = NFeFormataNumero4Dec(.qtde_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qCom", vNFeImgItem(UBound(vNFeImgItem)).det__qCom)
                
            '   VALOR UNITÁRIO
                vl_unitario = .valor_total / .qtde_total
                vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom = NFeFormataNumero4Dec(vl_unitario)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vUnCom", vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom)
                
            '   VALOR TOTAL
                vNFeImgItem(UBound(vNFeImgItem)).det__vProd = NFeFormataMoeda2Dec(.valor_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vProd", vNFeImgItem(UBound(vNFeImgItem)).det__vProd)
                
            '   cEANTrib - GTIN (Global Trade Item Number) da unidade tributável
                vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = .EAN
                'NFE 4.0 - EM BRANCO, INFORMAR SEM GTIN
                If vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "" Then vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "SEM GTIN"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cEANTrib", vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib)
            
            '   UNIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__uTrib = "PC"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uTrib", vNFeImgItem(UBound(vNFeImgItem)).det__uTrib)
                
            '   QUANTIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__qTrib = NFeFormataNumero4Dec(.qtde_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qTrib", vNFeImgItem(UBound(vNFeImgItem)).det__qTrib)
                
            '   VALOR UNITÁRIO DE TRIBUTAÇÃO
                vl_unitario = .valor_total / .qtde_total
                vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib = NFeFormataNumero4Dec(vl_unitario)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vUnTrib", vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib)
                
            '   OUTRAS DESPESAS ACESSÓRIAS
                If .vl_outras_despesas_acessorias > 0 Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__vOutro = NFeFormataMoeda2Dec(.vl_outras_despesas_acessorias)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vOutro", vNFeImgItem(UBound(vNFeImgItem)).det__vOutro)
                    End If
                
            '   INDICA SE VALOR DO ITEM (vProd) ENTRA NO VALOR TOTAL DA NF-e (vProd)
            '       0 – o valor do item (vProd) não compõe o valor total da NF-e (vProd)
            '       1 – o valor do item (vProd) compõe o valor total da NF-e (vProd) (v2.0)
                vNFeImgItem(UBound(vNFeImgItem)).det__indTot = "1"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("indTot", vNFeImgItem(UBound(vNFeImgItem)).det__indTot)
                
            '   xPed (número do pedido de compra)
                If Trim$(.xPed) <> "" Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__xPed = Trim$(.xPed)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("xPed", vNFeImgItem(UBound(vNFeImgItem)).det__xPed)
                    End If
                
            '   nItemPed (número do pedido de compra)
                If Trim$(.nItemPed) <> "" Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__nItemPed = Trim$(.nItemPed)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("nItemPed", vNFeImgItem(UBound(vNFeImgItem)).det__nItemPed)
                    End If
                
            '   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
                If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) Then
                    perc_IBPT = ibpt_aliquota_aplicavel(.cst, .percAliqNac, .percAliqImp)
                    vl_estimado_tributos = arredonda_para_monetario(.valor_total * (perc_IBPT / 100))
                    vNFeImgItem(UBound(vNFeImgItem)).det__vTotTrib = NFeFormataMoeda2Dec(vl_estimado_tributos)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vTotTrib", vNFeImgItem(UBound(vNFeImgItem)).det__vTotTrib)
                    vl_total_estimado_tributos = vl_total_estimado_tributos + vl_estimado_tributos
                    End If
                
                
            '   TAG ICMS
            '   ~~~~~~~~
                If IsNumeric(.ICMS) Then
                    perc_ICMS = CSng(.ICMS)
                Else
                    perc_ICMS = 0
                    End If
                
                vl_ICMS = 0
                vl_BC_ICMS = .valor_total
            
                vl_ICMSDeson = 0
                
                vl_ICMS_ST = 0
                vl_BC_ICMS_ST = 0
                
                vl_ICMS_ST_Ret = 0
                vl_BC_ICMS_ST_Ret = 0
                
                If Len(Trim$(.cst)) = 0 Then
                    s_erro = "O produto " & .produto & " - " & .descricao & " não possui a informação do CST!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Len(Trim$(.cst)) <> 3 Then
                    s_erro = "O produto " & .produto & " - " & .descricao & " possui o campo CST preenchido com valor inválido!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
            
            '   ORIGEM DA MERCADORIA
            '   LEMBRANDO QUE OS CAMPOS 'ORIG' E 'CST' ESTÃO CONCATENADOS NA PLANILHA DE PRODUTOS,
            '   MAS PODEM TER SIDO ALTERADOS ATRAVÉS DO CAMPO 'CST' NA TELA.
                vNFeImgItem(UBound(vNFeImgItem)).ICMS__orig = Trim$(left$(.cst, 1))
                strNFeTagIcms = vbTab & NFeFormataCampo("orig", vNFeImgItem(UBound(vNFeImgItem)).ICMS__orig)
                
            '   TAG ICMS
            '   ~~~~~~~~
                strNFeCst = Trim$(right$(.cst, 2))
                vNFeImgItem(UBound(vNFeImgItem)).ICMS__CST = strNFeCst
                strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__CST)
                                
            '   ICMS (CST=00): TRIBUTADO INTEGRALMENTE
                If strNFeCst = "00" Then
                    vl_ICMS = .valor_total * (perc_ICMS / 100)
                    vl_ICMS = CCur(Format$(vl_ICMS, FORMATO_MOEDA))
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS
                '   0: MARGEM VALOR AGREGADO (%); 1: PAUTA (VALOR); 2: PREÇO TABELADO MÁX. (VALOR); 3: VALOR DA OPERAÇÃO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC = "3"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC)
                    
                '   VALOR DA BC DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC)
                    
                '   ALÍQUOTA DO IMPOSTO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS)
                    
                '   VALOR DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS = NFeFormataMoeda2Dec(vl_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS)
                
                '   VALOR DO ICMS DESONERADO (ZERO, ATÉ RESOLUÇÃO EM CONTRÁRIO)
                    If vl_ICMSDeson <> 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson = NFeFormataMoeda2Dec(vl_ICMSDeson)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSDeson", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson)
                        End If
                
            '   ICMS (CST=10): TRIBUTADA E COM COBRANÇA DO ICMS POR SUBSTITUIÇÃO TRIBUTÁRIA
                ElseIf strNFeCst = "10" Then
                    vl_ICMS = .valor_total * (perc_ICMS / 100)
                    vl_ICMS = CCur(Format$(vl_ICMS, FORMATO_MOEDA))
                
                    If Not obtem_aliquota_ICMS_ST(rNFeImg.dest__UF, perc_ICMS_ST_aux, s_erro_aux) Then
                        s_erro = "Falha ao tentar obter a alíquota do ICMS ST para a UF: '" & rNFeImg.dest__UF & "'"
                        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                        End If
                    perc_ICMS_ST = perc_ICMS_ST_aux
                    
                    vl_BC_ICMS_ST = calcula_BC_ICMS_ST(.valor_total, .perc_MVA_ST)
                    vl_ICMS_ST = calcula_ICMS_ST(vl_BC_ICMS_ST, perc_ICMS_ST, vl_ICMS)
                    vl_ICMS_ST = CCur(Format$(vl_ICMS_ST, FORMATO_MOEDA))
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS
                '   0: MARGEM VALOR AGREGADO (%); 1: PAUTA (VALOR); 2: PREÇO TABELADO MÁX. (VALOR); 3: VALOR DA OPERAÇÃO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC = "3"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC)
                    
                '   VALOR DA BC DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC)
                    
                '   ALÍQUOTA DO IMPOSTO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS)
                
                '   VALOR DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS = NFeFormataMoeda2Dec(vl_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS)
                
                '   VALOR DO ICMS DESONERADO (ZERO, ATÉ RESOLUÇÃO EM CONTRÁRIO)
                    If vl_ICMSDeson <> 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson = NFeFormataMoeda2Dec(vl_ICMSDeson)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSDeson", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson)
                        End If
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS ST
                '   0: PREÇO TABELADO OU MÁXIMO SUGERIDO; 1: LISTA NEGATIVA (VALOR); 2: LISTA POSITIVA (VALOR); 3: LISTA NEUTRA (VALOR)
                '   4: MARGEM VALOR AGREGADO (%); 5: PAUTA (VALOR)
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBCST = "4"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBCST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBCST)
                    
                '   PERCENTUAL DA MARGEM DE VALOR ADICIONADO DO ICMS ST
                    If .perc_MVA_ST > 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__pMVAST = NFeFormataPercentual2Dec(.perc_MVA_ST)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pMVAST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pMVAST)
                        End If
                    
                '   VALOR DA BC DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCST = NFeFormataMoeda2Dec(vl_BC_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBCST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCST)
                    
                '   ALÍQUOTA DO IMPOSTO DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMSST = NFeFormataPercentual2Dec(perc_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMSST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMSST)
                    
                '   VALOR DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSST = NFeFormataMoeda2Dec(vl_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSST)
                    
            '   ICMS (CST=40,41,50): ISENTA, NÃO TRIBUTADA OU SUSPENSÃO (40=ISENTA, 41=NÃO TRIBUTADA, 50=SUSPENSÃO)
                ElseIf (strNFeCst = "40") Or (strNFeCst = "41") Or (strNFeCst = "50") Then
                '   NOP: DEMAIS CAMPOS SÃO OPCIONAIS E NÃO SE APLICAM
                    vl_ICMS = 0
                    vl_BC_ICMS = 0
                
            '   ICMS (CST=60): ICMS COBRADO ANTERIORMENTE POR SUBSTITUIÇÃO TRIBUTÁRIA
                ElseIf strNFeCst = "60" Then
                    blnHaProdutoCstIcms60 = True
                    
                    vl_ICMS = 0
                    vl_BC_ICMS = 0

                '   VALOR DA BC DO ICMS ST
                    vl_BC_ICMS_ST_Ret = 0
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCSTRet = NFeFormataMoeda2Dec(vl_BC_ICMS_ST_Ret)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBCSTRet", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCSTRet)
                    
                '   VALOR DO ICMS ST
                    vl_ICMS_ST_Ret = 0
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSSTRet = NFeFormataMoeda2Dec(vl_ICMS_ST_Ret)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSSTRet", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSSTRet)
                
            '   ICMS: CÓDIGO DE CST NÃO TRATADO PELO SISTEMA!!
                Else
                    s_erro = "Código de CST sem tratamento definido no sistema (CST=" & strNFeCst & ")!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
                    
            '   VERIFICAR SE A UF DO DESTINATÁRIO TEM LIMINAR PARA NÃO RECOLHER O DIFAL
                
                blnIgnorarDIFAL = False
                s_Texto_DIFAL_UF = ""
                
                s = "SELECT " & _
                    "st_ignorar_difal, " & _
                    "texto_adicional" & _
                    " FROM t_NFe_UF_PARAMETRO" & _
                    " WHERE" & _
                    " (UF='" & Trim$(strEndClienteUf) & "')"
                If t_NFe_UF_PARAMETRO.State <> adStateClosed Then t_NFe_UF_PARAMETRO.Close
                t_NFe_UF_PARAMETRO.Open s, dbc, , , adCmdText
                If Not t_NFe_UF_PARAMETRO.EOF Then
                    blnIgnorarDIFAL = t_NFe_UF_PARAMETRO("st_ignorar_difal") = 1
                    s_Texto_DIFAL_UF = Trim$("" & t_NFe_UF_PARAMETRO("texto_adicional"))
                    End If
            
                    
            '   OS CÁLCULOS DE PARTILHA FORAM MOVIDOS PARA CÁ DEVIDO À EXCLUSÃO DE ICMS E DIFAL DAS BASES DE CÁLCULO
            '   DE PIS E COFINS, CONFORME DECISÃO DO STF
            
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    Not blnIgnorarDIFAL And _
                    (vl_ICMS > 0) Then
                    
                    If IsNumeric(.fcp) Then
                        perc_fcp = CSng(.fcp)
                    Else
                        perc_fcp = 0
                        End If
                    
                    If Not obtem_aliquota_ICMS_UF_destino(rNFeImg.dest__UF, perc_ICMS_interna_UF_dest, s_erro_aux) Then
                        s_erro = "Falha ao tentar obter a alíquota interna do ICMS para a UF: '" & rNFeImg.dest__UF & "'"
                        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                        End If
                    
                    If intAnoPartilha < 2016 Then
                        perc_ICMS_UF_dest = 0
                        perc_ICMS_UF_remet = 100
                    ElseIf intAnoPartilha = 2016 Then
                        perc_ICMS_UF_dest = 40
                        perc_ICMS_UF_remet = 60
                    ElseIf intAnoPartilha = 2017 Then
                        perc_ICMS_UF_dest = 60
                        perc_ICMS_UF_remet = 40
                    ElseIf intAnoPartilha = 2018 Then
                        perc_ICMS_UF_dest = 80
                        perc_ICMS_UF_remet = 20
                    Else
                        perc_ICMS_UF_dest = 100
                        perc_ICMS_UF_remet = 0
                        End If
                    
                    'os cálculos abaixo se baseiam em um vídeo publicado pela Inventti Soluções
                    '(https://www.youtube.com/watch?v=MEoI88y-qNs)
                    perc_ICMS_diferencial_interestadual = perc_ICMS_interna_UF_dest + perc_fcp - perc_ICMS
                    vl_ICMS_diferencial_interestadual = vl_BC_ICMS * (perc_ICMS_diferencial_interestadual / 100)
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_interestadual
                    vl_fcp = vl_BC_ICMS * perc_fcp / 100
                    vl_fcp = CCur(Format$(vl_fcp, FORMATO_MOEDA))
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_aux - vl_fcp
                    vl_ICMS_UF_dest = arredonda_para_monetario(vl_ICMS_diferencial_aux * perc_ICMS_UF_dest / 100)
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_aux - vl_ICMS_UF_dest
                    vl_ICMS_UF_remet = arredonda_para_monetario(vl_ICMS_diferencial_aux)
                    If vl_ICMS_UF_remet < 0 Then vl_ICMS_UF_remet = 0
                    
                    End If
                    
            '   TAG IPI
            '   ~~~~~~~
            '   OBS: EXISTE IPI APENAS NA EMISSÃO DE NFe PARA DEVOLUÇÃO AO FORNECEDOR
                If IsNumeric(c_ipi) Then
                    perc_IPI = CSng(c_ipi)
                Else
                    perc_IPI = 0
                    End If
                
            '   TRAVA DE PROTEÇÃO ENQUANTO NÃO HÁ A IMPLEMENTAÇÃO DO TRATAMENTO
                If perc_IPI <> 0 Then
                    s_erro = "Não há tratamento definido no sistema para a alíquota de IPI!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
            
                vl_IPI = .valor_total * (perc_IPI / 100)
                vl_IPI = CCur(Format$(vl_IPI, FORMATO_MOEDA))
                
            '   TAG PIS
            '   ~~~~~~~
                vl_PIS = 0
                vl_BC_PIS = 0
                
                strZerarPisCst = Trim$(left$(cb_zerar_PIS, 2))
                
                If strZerarPisCst = "" Then
                    vl_BC_PIS = .valor_total
                    

                    If param_bc_pis_cofins_icms.campo_inteiro = 1 Then
                        vl_BC_PIS = vl_BC_PIS - vl_ICMS
                        End If
                    
                    If param_bc_pis_cofins_difal.campo_inteiro = 1 Then
                        vl_BC_PIS = vl_BC_PIS - vl_ICMS_UF_remet - vl_ICMS_UF_dest
                        End If

                    perc_PIS = PERC_PIS_ALIQUOTA_NORMAL
                    vl_PIS = vl_BC_PIS * (perc_PIS / 100)
                    vl_PIS = CCur(Format$(vl_PIS, FORMATO_MOEDA))
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__CST = "01"
                    strNFeTagPis = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).PIS__CST)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC = NFeFormataMoeda2Dec(vl_BC_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS = NFeFormataPercentual2Dec(perc_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("pPIS", vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS = NFeFormataMoeda2Dec(vl_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("vPIS", vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__CST = strZerarPisCst
                    strNFeTagPis = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).PIS__CST)
                    End If
            
            '   TAG COFINS
            '   ~~~~~~~~~~
                vl_COFINS = 0
                vl_BC_COFINS = 0
                
                strZerarCofinsCst = Trim$(left$(cb_zerar_COFINS, 2))
                
                If strZerarCofinsCst = "" Then
                    vl_BC_COFINS = .valor_total
                    
                    If param_bc_pis_cofins_icms.campo_inteiro = 1 Then
                        vl_BC_COFINS = vl_BC_COFINS - vl_ICMS
                        End If
                        
                    If param_bc_pis_cofins_difal.campo_inteiro = 1 Then
                        vl_BC_COFINS = vl_BC_COFINS - vl_ICMS_UF_remet - vl_ICMS_UF_dest
                        End If
                    
                    perc_COFINS = PERC_COFINS_ALIQUOTA_NORMAL
                    vl_COFINS = vl_BC_COFINS * (perc_COFINS / 100)
                    vl_COFINS = CCur(Format$(vl_COFINS, FORMATO_MOEDA))
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST = "01"
                    strNFeTagCofins = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC = NFeFormataMoeda2Dec(vl_BC_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS = NFeFormataPercentual2Dec(perc_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("pCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS = NFeFormataMoeda2Dec(vl_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST = strZerarCofinsCst
                    strNFeTagCofins = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST)
                    End If
                
            '   TAG ICMSUFDest
            '   ~~~~~~~~~~~~~~
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not blnIgnorarDIFAL And _
                    Not cfop_eh_de_remessa(strCfopCodigo) Then
                
                    strNFeTagIcmsUFDest = ""
                    
                '   VALOR DA BC DO ICMS NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest)
                    
                    
                    'VALOR DA BASE DE CÁLCULO DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    '(lhgx) obs: manter esta linha comentada, pois podemos ter problema com o resultado no ambiente de produção
                    'strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCFCPUFDest", NFeFormataMoeda2Dec(vl_BC_ICMS))
                    
                '   PERCENTUAL DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest = NFeFormataPercentual2Dec(perc_fcp)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pFCPUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest)
                '   ALÍQUOTA INTERNA DA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSUFDest = NFeFormataPercentual2Dec(perc_ICMS_interna_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSUFDest)
                '   ALÍQUOTA INTERESTADUAL DAS UF ENVOLVIDAS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInter = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSInter", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInter)
                '   PERCENTUAL PROVISÓRIO DE PARTILHA DO ICMS INTERESTADUAL
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInterPart = NFeFormataPercentual2Dec(perc_ICMS_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSInterPart", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInterPart)
                '   VALOR DO ICMS RELATIVO AO FCP DA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vFCPUFDest = NFeFormataMoeda2Dec(vl_fcp)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vFCPUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vFCPUFDest)
                '   VALOR DO ICMS INTERESTADUAL PARA A UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFDest = NFeFormataMoeda2Dec(vl_ICMS_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vICMSUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFDest)
                '   VALOR DO ICMS INTERESTADUAL PARA A UF DO REMETENTE
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFRemet = NFeFormataMoeda2Dec(vl_ICMS_UF_remet)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vICMSUFRemet", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFRemet)
    
                    vl_total_FCPUFDest = vl_total_FCPUFDest + vl_fcp
                    vl_total_ICMSUFDest = vl_total_ICMSUFDest + vl_ICMS_UF_dest
                    vl_total_ICMSUFRemet = vl_total_ICMSUFRemet + vl_ICMS_UF_remet
                    End If
                    
            
            
            '   MONTA BLOCO POR PRODUTO
            '   ~~~~~~~~~~~~~~~~~~~~~~~
                strNFeTagBlocoProduto = strNFeTagBlocoProduto & _
                                        "det;" & vbCrLf & strNFeTagDet & _
                                        "ICMS;" & vbCrLf & strNFeTagIcms & _
                                        "PIS;" & vbCrLf & strNFeTagPis & _
                                        "COFINS;" & vbCrLf & strNFeTagCofins
                
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    Not blnIgnorarDIFAL And _
                    (vl_ICMS > 0) Then
                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & _
                                            "ICMSUFDest;" & vbCrLf & strNFeTagIcmsUFDest
                    End If
                
            '   INFORMAÇÕES ADICIONAIS DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd = .infAdProd
                If Trim$(vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd) <> "" Then
                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & vbTab & NFeFormataCampo("infAdProd", vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd)
                    End If
                
            '   QTDE DE VOLUMES
                total_volumes = total_volumes + .qtde_volumes_total
                
            '   PESO BRUTO
                total_peso_bruto = total_peso_bruto + .peso_total
                    
            '   PESO LIQUIDO
                total_peso_liquido = total_peso_liquido + .peso_total
                
            '   CUBAGEM TOTAL
                cubagem_bruto = cubagem_bruto + .cubagem_total
                
            '   TOTALIZAÇÃO
                vl_total_ICMS = vl_total_ICMS + vl_ICMS
                vl_total_ICMSDeson = vl_total_ICMSDeson + vl_ICMSDeson
                vl_total_ICMS_ST = vl_total_ICMS_ST + vl_ICMS_ST
                vl_total_produtos = vl_total_produtos + .valor_total
                vl_total_BC_ICMS = vl_total_BC_ICMS + vl_BC_ICMS
                vl_total_BC_ICMS_ST = vl_total_BC_ICMS_ST + vl_BC_ICMS_ST
                vl_total_IPI = vl_total_IPI + vl_IPI
                vl_total_PIS = vl_total_PIS + vl_PIS
                vl_total_COFINS = vl_total_COFINS + vl_COFINS
                vl_total_outras_despesas_acessorias = vl_total_outras_despesas_acessorias + .vl_outras_despesas_acessorias
                End If
            End With
        Next
    
    
'   QTDE TOTAL DE VOLUMES
'   ~~~~~~~~~~~~~~~~~~~~~
    If Trim$(c_total_volumes) <> "" Then
        If CLng(c_total_volumes) <> total_volumes Then
            s = "A quantidade total de volumes foi editada de " & CStr(total_volumes) & " para " & c_total_volumes & vbCrLf & _
                "Continua mesmo assim?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
    
'   TAG TOTAL
'   ~~~~~~~~~
    strNFeTagValoresTotais = "total;" & vbCrLf
    
'   BASE DE CÁLCULO DO ICMS
    rNFeImg.total__vBC = NFeFormataMoeda2Dec(vl_total_BC_ICMS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vBC", rNFeImg.total__vBC)
                            
'   VALOR TOTAL DO ICMS
    rNFeImg.total__vICMS = NFeFormataMoeda2Dec(vl_total_ICMS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vICMS", rNFeImg.total__vICMS)

'   novo campo vICMSDeson (layout 3.10)
    rNFeImg.total__vICMSDeson = NFeFormataMoeda2Dec(vl_total_ICMSDeson)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vICMSDeson", rNFeImg.total__vICMSDeson)
    
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
    If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
        (rNFeImg.dest__indIEDest = "9") And _
        Not blnIgnorarDIFAL And _
        Not cfop_eh_de_remessa(strCfopCodigo) Then
            rNFeImg.total__vFCPUFDest = NFeFormataMoeda2Dec(vl_total_FCPUFDest)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vFCPUFDest", rNFeImg.total__vFCPUFDest)
            rNFeImg.total__vICMSUFDest = NFeFormataMoeda2Dec(vl_total_ICMSUFDest)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vICMSUFDest", rNFeImg.total__vICMSUFDest)
            rNFeImg.total__vICMSUFRemet = NFeFormataMoeda2Dec(vl_total_ICMSUFRemet)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vICMSUFRemet", rNFeImg.total__vICMSUFRemet)
        
        End If
        
    'NFE 4.0 - vFCP
    ' quando for emitida uma NF-e (modelo 55) interestadual (Campo: idDest = 2) para Consumidor Final (Campo: indFinal = 1)
    ' não contribuinte (Campo: indIEDest = 9) e o valor do FCP for informado em um campo diferente de vFCPUFDest haverá esta rejeição
    '(e-mail do Márcio da Target em 01/11/18
    rNFeImg.total__vFCP = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCP", rNFeImg.total__vFCP)
        
'   vBCST
    rNFeImg.total__vBCST = NFeFormataMoeda2Dec(vl_total_BC_ICMS_ST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vBCST", rNFeImg.total__vBCST)
    
'   vST
    rNFeImg.total__vST = NFeFormataMoeda2Dec(vl_total_ICMS_ST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vST", rNFeImg.total__vST)
    
    'NFE 4.0 - vFCPST
    rNFeImg.total__vFCPST = NFeFormataMoeda2Dec(vl_total_vFCPST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCPST", rNFeImg.total__vFCPST)
    
    'NFE 4.0 - vFCPSTRet
    rNFeImg.total__vFCPSTRet = NFeFormataMoeda2Dec(vl_total_vFCPSTRet)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCPSTRet", rNFeImg.total__vFCPSTRet)
    
    
    
'   VALOR TOTAL DOS PRODUTOS
    rNFeImg.total__vProd = NFeFormataMoeda2Dec(vl_total_produtos)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vProd", rNFeImg.total__vProd)
                             
'   VALOR TOTAL DO FRETE
    rNFeImg.total__vFrete = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFrete", rNFeImg.total__vFrete)
    
'   VALOR TOTAL DO SEGURO
    rNFeImg.total__vSeg = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vSeg", rNFeImg.total__vSeg)
    
'   VALOR TOTAL DO DESCONTO
    rNFeImg.total__vDesc = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vDesc", rNFeImg.total__vDesc)
    
'   VALOR TOTAL DO II
    rNFeImg.total__vII = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vII", rNFeImg.total__vII)
    
'   VALOR TOTAL DO IPI
    rNFeImg.total__vIPI = NFeFormataMoeda2Dec(vl_total_IPI)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vIPI", rNFeImg.total__vIPI)
                             
    'NFE 4.0 vIPIDevol
    rNFeImg.total__vIPIDevol = NFeFormataMoeda2Dec(vl_total_vIPIDevol)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vIPIDevol", rNFeImg.total__vIPIDevol)
                             
'   VALOR DO PIS
    rNFeImg.total__vPIS = NFeFormataMoeda2Dec(vl_total_PIS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vPIS", rNFeImg.total__vPIS)
    
'   VALOR DO COFINS
    rNFeImg.total__vCOFINS = NFeFormataMoeda2Dec(vl_total_COFINS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vCOFINS", rNFeImg.total__vCOFINS)
    
'   VALOR DESPESAS ACESSÓRIAS
    rNFeImg.total__vOutro = NFeFormataMoeda2Dec(vl_total_outras_despesas_acessorias)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vOutro", rNFeImg.total__vOutro)
    
'   VALOR TOTAL DA NOTA
    vl_total_NF = vl_total_produtos
    If vl_total_IPI > 0 Then vl_total_NF = vl_total_NF + vl_total_IPI
    If vl_total_outras_despesas_acessorias > 0 Then vl_total_NF = vl_total_NF + vl_total_outras_despesas_acessorias
    rNFeImg.total__vNF = NFeFormataMoeda2Dec(vl_total_NF)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vNF", rNFeImg.total__vNF)
                             
'   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    strInfoAdicIbpt = ""
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) Then
        rNFeImg.total__vTotTrib = NFeFormataMoeda2Dec(vl_total_estimado_tributos)
        strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vTotTrib", rNFeImg.total__vTotTrib)
        perc_aux = 100 * (vl_total_estimado_tributos / vl_total_NF)
        strInfoAdicIbpt = "Valor Aprox. dos Tributos: " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_estimado_tributos) & " (" & formata_numero_2dec(perc_aux) & "%) Fonte: IBPT"
        End If
    
'   TAG TRANSP
'   ~~~~~~~~~~
'   MODALIDADE DO FRETE
    strNFeTagTransp = "transp;" & vbCrLf
    rNFeImg.transp__modFrete = left$(Trim$(cb_frete), 1)
    strNFeTagTransp = strNFeTagTransp & _
                      vbTab & NFeFormataCampo("modFrete", rNFeImg.transp__modFrete)
                              
'   TAG TRANSPORTA
'   ~~~~~~~~~~~~~~
'   DADOS DA TRANSPORTADORA
    If strTransportadoraId <> "" Then
        If Len(strTransportadoraCnpj) = 14 Then
            rNFeImg.transporta__CNPJ = strTransportadoraCnpj
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("CNPJ", rNFeImg.transporta__CNPJ)
        ElseIf Len(strTransportadoraCnpj) = 11 Then
            rNFeImg.transporta__CPF = strTransportadoraCnpj
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("CPF", rNFeImg.transporta__CPF)
            End If
        
        If strTransportadoraRazaoSocial <> "" Then
            rNFeImg.transporta__xNome = strTransportadoraRazaoSocial
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("xNome", rNFeImg.transporta__xNome)
            End If
        
        If (Len(strTransportadoraCnpj) = 14) Then
            strCampo = strTransportadoraIE
            If InStr(UCase$(strCampo), "ISEN") > 0 Then strCampo = "ISENTO"
            If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
            strTransportadoraIE = strCampo
           
            If (Len(strTransportadoraIE) > 0) Then
                If (Len(strTransportadoraIE) < 2) Or (Len(strTransportadoraIE) > 14) Then
                    s_erro = "A Inscrição Estadual no cadastro da transportadora '" & strTransportadoraId & "' está preenchida com conteúdo inválido!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Len(strTransportadoraUF) = 0 Then
                    s_erro = "A UF no cadastro da transportadora '" & strTransportadoraId & "' não está preenchida!!" & vbCrLf & "Essa informação é necessária devido ao campo IE!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Not UF_ok(strTransportadoraUF) Then
                    s_erro = "A UF no cadastro da transportadora '" & strTransportadoraId & "' está preenchida com conteúdo inválido!!" & vbCrLf & "Essa informação é necessária devido ao campo IE!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf ConsisteInscricaoEstadual(strTransportadoraIE, strTransportadoraUF) <> 0 Then
                '   Retorno = 0 -> IE válida
                '   Retorno = 1 -> IE inválida
                    s_erro = "A Inscrição Estadual no cadastro da transportadora '" & strTransportadoraId & "' é inválida para a UF de '" & strTransportadoraUF & "'!!"
                    GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
                    
                rNFeImg.transporta__IE = strTransportadoraIE
                strNFeTagTransporta = strNFeTagTransporta & _
                                      vbTab & NFeFormataCampo("IE", rNFeImg.transporta__IE)
                
                rNFeImg.transporta__UF = strTransportadoraUF
                strNFeTagTransporta = strNFeTagTransporta & _
                                      vbTab & NFeFormataCampo("UF", rNFeImg.transporta__UF)
                End If
            End If
            
        If strNFeTagTransporta <> "" Then
            strNFeTagTransporta = "transporta;" & vbCrLf & strNFeTagTransporta
            End If
        End If
    
'   TAG VOL
'   ~~~~~~~
    strNFeTagVol = "vol;" & vbCrLf
    
'   QUANTIDADE DE VOLUMES TRANSPORTADOS
    If Trim$(c_total_volumes) <> "" Then
        rNFeImg.vol__qVol = retorna_so_digitos(CStr(CLng(c_total_volumes)))
    Else
        rNFeImg.vol__qVol = retorna_so_digitos(CStr(total_volumes))
        End If
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("qVol", rNFeImg.vol__qVol)
    
'   ESPÉCIE DOS VOLUMES TRANSPORTADOS
    rNFeImg.vol__esp = "VOLUME"
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("esp", rNFeImg.vol__esp)
    
'   PESO LÍQUIDO
    rNFeImg.vol__pesoL = NFeFormataNumero3Dec(total_peso_liquido)
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("pesoL", rNFeImg.vol__pesoL)
    
'   PESO BRUTO
    rNFeImg.vol__pesoB = NFeFormataNumero3Dec(total_peso_bruto)
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("pesoB", rNFeImg.vol__pesoB)
    
    
    'NFE 4.0 - tag pag
    strNFeTagPag = "pag;" & vbCrLf
    If Trim$(vNFeImgPag(UBound(vNFeImgPag)).pag__indPag) <> "" Then
        ReDim Preserve vNFeImgPag(UBound(vNFeImgPag) + 1)
        End If
    vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = ""
    'Segundo informado pelo Valter (Target) em e-mail de 27/06/2017, não deve ser informada no arquivo de integração,
    'ela é inserida automaticamente pelo sistema
    'strNFeTagPag = strNFeTagPag & "detpag;" & vbCrLf

    'Segundo informado pelo Valter (Target) em e-mail de 27/06/2017, não deve ser informada no arquivo de integração,
    'ela é inserida automaticamente pelo sistema
    'strNFeTagPag = strNFeTagPag & "detpag;" & vbCrLf
    'Os códigos de pagamento usados abaixo estão presente na nota técnica da SEFAZ
    'NT2020.006 v1.10 de Fevereiro de 2021:
    '   01=Dinheiro
    '   02=Cheque
    '   03=Cartão de Crédito
    '   04=Cartão de Débito
    '   05=Crédito Loja
    '   10=Vale Alimentação
    '   11=Vale Refeição
    '   12=Vale Presente
    '   13=Vale Combustível
    '   15=Boleto Bancário
    '   16=Depósito Bancário
    '   17=Pagamento Instantâneo (PIX)
    '   18=Transferência bancária, Carteira Digital
    '   19=Programa de fidelidade, Cashback, Crédito Virtual
    '   90=Sem pagamento
    '   99=Outros

    s_aux = param_nftipopag.campo_texto
    s = ""
    'Se a nota é de entrada ou ajuste/devolução - sem pagamento
    If rNFeImg.ide__tpNF = "0" Or _
        strNFeCodFinalidade = "3" Or _
        strNFeCodFinalidade = "4" Then
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = "90"
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(0)
    'Se o pagamento é à vista
    ElseIf strTipoParcelamento = COD_FORMA_PAGTO_A_VISTA Then
        'Para cada meio de pagamento abaixo:
        '   - Se for obrigatório informar um meio de pagamento diferente de "99-Outros" sem descrição:
        '       - Se o sistema estiver operando em contingência, informa "99-Outros" e fornece uma descrição
        '       - Se não estiver operando em contingência, informa o código da lista acima
        '   - Se não for obrigatório informar um meio de pagamento, informa "99-Outros" sem descrição
        Select Case t_PEDIDO("av_forma_pagto")
            Case ID_FORMA_PAGTO_DINHEIRO
                    If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                        s_aux = "99"
                        s = "Dinheiro"
                    Else
                        s_aux = "01"
                        End If
            Case ID_FORMA_PAGTO_CHEQUE
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Cheque"
                Else
                    s_aux = "02"
                    End If
            Case ID_FORMA_PAGTO_BOLETO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_BOLETO_AV
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_DEPOSITO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "16"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Depósito"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case Else
                If (param_nftipopag.campo_inteiro = 1) Then
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Meio de pagamento não identificado"
                    Else
                        s_aux = param_nftipopag.campo_texto
                        End If
                    Else
                        s_aux = "99" 'Outros
                        End If
            End Select
        
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = rNFeImg.total__vNF
    'Se o pagamento é à prazo
    ElseIf (strTipoParcelamento = COD_FORMA_PAGTO_PARCELADO_CARTAO) Or _
           (strTipoParcelamento = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) Then
        If (param_nftipopag.campo_inteiro = 1) Then
            s_aux = "03"
            If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                s_aux = "99"
                s = "Cartão"
                End If
        Else
            s_aux = "99"
            End If
        'obtém o total a prazo (retira o valor da entrada,se houver)
        vl_aux = vl_total_NF - vl_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "1"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(vl_aux)
    Else
        vl_aux = 0
        Select Case t_PEDIDO("pce_forma_pagto_prestacao")
            Case ID_FORMA_PAGTO_DINHEIRO
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Dinheiro"
                Else
                    s_aux = "01"
                    End If
            Case ID_FORMA_PAGTO_CHEQUE
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Cheque"
                Else
                    s_aux = "02"
                    End If
            Case ID_FORMA_PAGTO_BOLETO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_BOLETO_AV
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_DEPOSITO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "16"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Depósito"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case Else
                If (param_nftipopag.campo_inteiro = 1) Then
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Meio de pagamento não identificado"
                    Else
                        s_aux = param_nftipopag.campo_texto
                        End If
                    Else
                        s_aux = "99" 'Outros
                        End If
            End Select
        'obtém o total a prazo (retira o valor da entrada,se houver)
        vl_aux = vl_total_NF - vl_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "1"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(vl_aux)
        End If
    
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("indPag", vNFeImgPag(UBound(vNFeImgPag)).pag__indPag)
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tPag", vNFeImgPag(UBound(vNFeImgPag)).pag__tPag)
    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
        If s <> "" Then strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("xPag", s)
        End If
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vPag", vNFeImgPag(UBound(vNFeImgPag)).pag__vPag)
    'Segundo informado pelo Valter (Target) em e-mail de 27/07/2017, o grupo vcard não deve ser informado no arquivo texto,
    'ele é preenchido pelo sistema
    'informações do intermediador
    If (param_nfintermediador.campo_inteiro = 1) And (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") Then
        'If (strMarketplaceCodOrigem <> "") Then
        If ((strMarketPlaceCNPJ <> "") And (strMarketPlaceCadIntTran <> "")) Then
            strNFeTagPag = strNFeTagPag & vbTab & "infIntermed;"
            strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("CNPJ", strMarketPlaceCNPJ)
            strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("idCadIntTran", strMarketPlaceCadIntTran)
            End If
        End If
                    

'   TAG INFADIC
'   ~~~~~~~~~~~
'   TEXTO FIXO SOBRE RESPONSABILIDADE DA INSTALAÇÃO
    If blnTemPagtoPorBoleto Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Não efetue qualquer pagamento desta nota fiscal a terceiros, pois a quitação da mesma só terá validade após o pagamento do(s) título(s) bancário(s) emitidos por esta empresa. Caso não receba o(s) título(s) até a data(s) do(s) vencimento(s) favor contatar (11)4858-2431."
        End If
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
    strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "A responsabilidade pelo serviço de instalação e/ou manutenção dos produtos acima é única e exclusivamente da empresa e/ou técnico autônomo contratado pelo destinatário desta."
    
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
    strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Fabricante não cobre avarias de peças plásticas, portanto, é necessário avaliar o equipamento no ato da entrega."
    
'   TEXTO FIXO SOBRE REGIME ESPECIAL
    If txtFixoEspecifico <> "" Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & txtFixoEspecifico
        End If

'   OUTROS TELEFONES DE CONTATO (INF ADICIONAIS)
    s_aux = ""
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PF Then
        strSufixoRes = "Tel Res: "
        strSufixoCom = "Tel Com: "
    Else
        strSufixoRes = "Tel: "
        strSufixoCom = "Tel: "
        End If
    If (strTelCel <> "") And (strTelRes <> "") Then s_aux = strSufixoRes & strTelRes
    If ((strTelCel <> "") Or (strTelRes <> "")) And (strTelCom <> "") Then
        If s_aux <> "" Then s_aux = s_aux & " / "
        s_aux = s_aux & strSufixoCom & strTelCom
        End If
    If ((strTelCel <> "") Or (strTelRes <> "") Or (strTelCom <> "")) And (strTelCom2 <> "") Then
        If s_aux <> "" Then s_aux = s_aux & " / "
        s_aux = s_aux & strSufixoCom & strTelCom2
        End If
    If s_aux <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & s_aux
        End If
    
'   TEXTO DIGITADO
    If Trim$(c_dados_adicionais_venda) <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & Trim$(c_dados_adicionais_venda)
        End If
    
    If blnHaProdutoCstIcms60 Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = TEXTO_LEI_CST_ICMS_60 & strNFeInfAdicQuadroProdutos
        End If
    
'   BEM DE USO E CONSUMO
    If blnTemPedidoComStBemUsoConsumo And (Not blnTemPedidoSemStBemUsoConsumo) Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = "BEM DE USO E CONSUMO" & strNFeInfAdicQuadroProdutos
        End If

'   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) And (strInfoAdicIbpt <> "") Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = strInfoAdicIbpt & strNFeInfAdicQuadroProdutos
        End If
    
'   Nº PEDIDO (NA 1ª LINHA) + CUBAGEM
    strTextoCubagem = ""
    If cubagem_bruto > 0 Then strTextoCubagem = Space$(20) & "CUB: " & formata_numero_2dec(cubagem_bruto) & " m3"
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
    strNFeInfAdicQuadroProdutos = Join(v_pedido, ", ") & strTextoCubagem & strNFeInfAdicQuadroProdutos
    
'   INFORMAÇÕES SOBRE PARTILHA DO ICMS
    If PARTILHA_ICMS_ATIVA And Not blnIgnorarDIFAL Then
        'DIFAL- suprimir texto em notas de entrada/devolução
        If (rNFeImg.ide__tpNF <> "0") And _
            (strNFeCodFinalidade <> "3") And _
            (strNFeCodFinalidade <> "4") And _
                Not tem_instricao_virtual(usuario.emit_id, rNFeImg.dest__UF) Then
            If (vl_total_ICMSUFDest > 0) Then
                If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
                strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Valores totais do ICMS Interestadual: partilha da UF Destino " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_ICMSUFDest)
                If (vl_total_FCPUFDest > 0) Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & " + FCP " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_FCPUFDest)
                strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "; partilha da UF Origem " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_ICMSUFRemet) & "."
                End If
            End If
        End If
        
'   SE UF TEM LIMINAR PARA NÃO RECOLHIMENTO DO DIFAL, INFORMAR
    If PARTILHA_ICMS_ATIVA And blnIgnorarDIFAL And _
        (rNFeImg.ide__idDest = "2") And _
        (rNFeImg.dest__indIEDest = "9") Then
        If s_Texto_DIFAL_UF <> "" Then
            If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
            strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & s_Texto_DIFAL_UF
            End If
        End If
                

'   INFORMAÇÕES SOBRE MEIO DE PAGAMENTO DAS PARCELAS
    If blnImprimeDadosFatura And _
        strInfoAdicParc <> "" Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & strInfoAdicParc
        End If
        
'   INFORMAR QUANDO SE TRATA DE PEDIDO QUITADO (PAGAMENTO ANTECIPADO)
    If (strPagtoAntecipadoStatus = "1") And (strPagtoAntecipadoQuitadoStatus = "1") Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Pedido com pagamento antecipado (Quitado)"
        End If

    rNFeImg.infAdic__infCpl = strNFeInfAdicQuadroInfAdic & "|" & strNFeInfAdicQuadroProdutos
    strNFeTagInfAdicionais = "infAdic;" & vbCrLf & _
                             vbTab & NFeFormataCampo("infCpl", rNFeImg.infAdic__infCpl)
                             
    
'   LHGX - rever (a tag "entrega" só vale nesta situação?)
'   TAG ENTREGA
'   ~~~~~~~~~~~
    strDestinatarioCnpjCpf = retorna_so_digitos(c_cnpj_cpf_dest)
    strEndEtgEndereco = l_end_recebedor_logradouro
    strEndEtgEnderecoNumero = l_end_recebedor_numero
    strEndEtgEnderecoComplemento = l_end_recebedor_complemento
    strEndEtgBairro = l_end_recebedor_bairro
    strEndEtgCidade = l_end_recebedor_cidade
    strEndEtgUf = l_end_recebedor_uf
'   NO MOMENTO, A SEFAZ ACEITA ENDEREÇO DE ENTREGA DIFERENTE DO ENDEREÇO DE CADASTRO SOMENTE P/ PJ
    If (UCase$(strEndEtgUf) <> UCase$(strEndClienteUf)) And _
        cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
        strNFeTagEndEntrega = "entrega;" & vbCrLf
        If (Len(strDestinatarioCnpjCpf) = 14) Then
            rNFeImg.entrega__CNPJ = strDestinatarioCnpjCpf
            strNFeTagEndEntrega = strNFeTagEndEntrega & vbTab & NFeFormataCampo("CNPJ", rNFeImg.entrega__CNPJ)
        Else
            rNFeImg.entrega__CPF = strDestinatarioCnpjCpf
            strNFeTagEndEntrega = strNFeTagEndEntrega & vbTab & NFeFormataCampo("CPF", rNFeImg.entrega__CPF)
            End If
            
        rNFeImg.entrega__xLgr = strEndEtgEndereco
        rNFeImg.entrega__nro = strEndEtgEnderecoNumero
        rNFeImg.entrega__xCpl = strEndEtgEnderecoComplemento
        rNFeImg.entrega__xBairro = strEndEtgBairro
        rNFeImg.entrega__cMun = strEndEtgCidade & "/" & strEndEtgUf
        rNFeImg.entrega__xMun = strEndEtgCidade
        rNFeImg.entrega__UF = strEndEtgUf
        
        strNFeTagEndEntrega = strNFeTagEndEntrega & _
                              vbTab & NFeFormataCampo("xLgr", rNFeImg.entrega__xLgr) & _
                              vbTab & NFeFormataCampo("nro", rNFeImg.entrega__nro)
                              
        If Len(rNFeImg.entrega__xCpl) > 0 Then
            strNFeTagEndEntrega = strNFeTagEndEntrega & _
                              vbTab & NFeFormataCampo("xCpl", rNFeImg.entrega__xCpl)
            End If
        
        strNFeTagEndEntrega = strNFeTagEndEntrega & _
                              vbTab & NFeFormataCampo("xBairro", rNFeImg.entrega__xBairro) & _
                              vbTab & NFeFormataCampo("cMun", rNFeImg.entrega__cMun) & _
                              vbTab & NFeFormataCampo("xMun", rNFeImg.entrega__xMun) & _
                              vbTab & NFeFormataCampo("UF", rNFeImg.entrega__UF)
        End If
        
        
'   SÓ AUTORIZA EMISSÃO SEM INTERMEDIADOR SE intImprimeIntermediadorAusente FOR 1
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If (param_nfintermediador.campo_inteiro = 1) Then
        If (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") And _
            ((strMarketPlaceCNPJ = "") Or (strMarketPlaceCadIntTran = "")) And _
            (intImprimeIntermediadorAusente = 0) Then
            s = "Não é possível prosseguir com a emissão, pois o intermediador do pedido não está identificado!!"
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If

        

'   SE HOUVER MAIS DE UMA CONFIRMAÇÃO DE EMISSÃO QUE PODEM GERAR NFe PARA UM EMITENTE INDEVIDO, CONFIRMAR NOVAMENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If iQtdConfirmaDuvidaEmit > 1 Then
        s = "Algumas confirmações efetuadas indicam que a NFe pode ser gerada em um Emitente indevido." & vbCrLf & _
            "Confirma a emissão no Emitente " & usuario.emit & "?"
        If Not confirma(s) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
'   CONFIRMAÇÃO FINAL
'   ~~~~~~~~~~~~~~~~~
    s = Join(v_pedido(), ", ")
    If qtde_pedidos = 1 Then
        s = " para o pedido " & s & "?"
    Else
        s = " para os pedidos " & s & "?"
        End If
    
    s = "Emite a NFe " & s
    
    If Not confirma(s) Then
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    
'   OBTÉM NSU P/ GRAVAR OS DADOS DA NFe P/ FINS DE HISTÓRICO, CONTROLE E CONSULTA DA DANFE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If Not geraNsu(NSU_T_NFe_EMISSAO, lngNsuNFeEmissao, s_erro_aux) Then
        s = "Falha ao tentar gerar o NSU para a tabela " & NSU_T_NFe_EMISSAO & "!!"
        If s_erro_aux <> "" Then s = s & vbCrLf
        s = s & s_erro_aux
        aviso_erro s
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
  
    
'   OBTÉM Nº SÉRIE E PRÓXIMO Nº PARA ATRIBUIR À NFe
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If blnEsperaNFTriangular Then
        strSerieNf = CStr(lngSerieNFeTriangular)
        strNumeroNf = CStr(lngNumVendaNFeTriangular)
        strSerieNfTriangular = CStr(lngNumRemessaNFeTriangular)
    Else
        aguarde INFO_EXECUTANDO, "obtendo próximo número de NF"
        If Not NFeObtemProximoNumero(rNFeImg.id_nfe_emitente, strSerieNf, strNumeroNf, s_erro_aux) Or _
            Not NFeObtemProximoNumero(rNFeImg.id_nfe_emitente, strSerieNfTriangular, strNumeroNfTriangular, s_erro_aux) Then
            s = "Falha ao tentar gerar o número para a NFe!!"
            If s_erro_aux <> "" Then s = s & vbCrLf
            s = s & s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            lngSerieNFeTriangular = CLng(strSerieNf)
            lngNumVendaNFeTriangular = CLng(strNumeroNf)
            lngNumRemessaNFeTriangular = CLng(strNumeroNfTriangular)
            'agora que temos a numeração, substituir os "???" dos dados adicionais pelos respectivos números
            c_dados_adicionais_venda = Replace(c_dados_adicionais_venda, "???", strNumeroNfTriangular)
            strNFeTagInfAdicionais = Replace(strNFeTagInfAdicionais, "???", strNumeroNfTriangular)
            c_dados_adicionais_remessa = Replace(c_dados_adicionais_remessa, "???", strNumeroNf)
            End If
        End If


'   VERIFICA SE O Nº DA NFE A SER EMITIDA ENCONTRA-SE INUTILIZADO (A OPERAÇÃO DE INUTILIZAÇÃO DE FAIXAS DE NÚMEROS DA NFe É
'   REALIZADA NO SISTEMA DA TARGET ONE)
    s = "SELECT " & _
            "*" & _
        " FROM NFE_INUTILIZA" & _
        " WHERE" & _
            " (Serie = '" & NFeFormataSerieNF(strSerieNf) & "')" & _
            " AND (NumIni >= '" & NFeFormataNumeroNF(strNumeroNf) & "')" & _
            " AND (NumFim <= '" & NFeFormataNumeroNF(strNumeroNfTriangular) & "')"
    If t_T1_NFE_INUTILIZA.State <> adStateClosed Then t_T1_NFE_INUTILIZA.Close
    t_T1_NFE_INUTILIZA.Open s, dbcNFe, , , adCmdText
    If Not t_T1_NFE_INUTILIZA.EOF Then
    '   CÓDIGOS: 1=Em Processamento; 2=Falha; 3=Homologado
        strCodStatusInutilizacao = Trim$("" & t_T1_NFE_INUTILIZA("Status"))
        s_erro_aux = "Data: " & Format$(t_T1_NFE_INUTILIZA("DataHora"), FORMATO_DATA_HORA) & vbCrLf & _
                     "Nº inicial: " & Trim$("" & t_T1_NFE_INUTILIZA("NumIni")) & vbCrLf & _
                     "Nº final: " & Trim$("" & t_T1_NFE_INUTILIZA("NumFim")) & vbCrLf & _
                     "Série: " & Trim$("" & t_T1_NFE_INUTILIZA("Serie")) & vbCrLf & _
                     "Motivo: " & Trim$("" & t_T1_NFE_INUTILIZA("Motivo")) & vbCrLf & _
                     "Usuário: " & Trim$("" & t_T1_NFE_INUTILIZA("Usuario")) & vbCrLf & _
                     "Status: " & strCodStatusInutilizacao & " - " & decodifica_NFe_inutilizacao_status(strCodStatusInutilizacao) & _
                     "Código: " & Trim$("" & t_T1_NFE_INUTILIZA("PendSta")) & vbCrLf & _
                     "Mensagem: " & Trim$("" & t_T1_NFE_INUTILIZA("PendDes"))
        If strCodStatusInutilizacao = "3" Then
            s = "Não é possível prosseguir com a emissão, pois o número de NFe informado foi inutilizado!!" & vbCrLf & _
                s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        ElseIf strCodStatusInutilizacao = "1" Then
            s = "Não é possível prosseguir com a emissão, pois o número de NFe informado consta em uma operação de inutilização de números de NFe que está em andamento!!" & vbCrLf & _
                s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If


'   SE O PEDIDO ESTIVER NA FILA DE SOLICITAÇÃO DE EMISSÃO DE NFE, SINALIZA QUE JÁ FOI TRATADO
'   LHGX - executar a sinalização na emissão da nota de remessa
'    For i = LBound(v_pedido) To UBound(v_pedido)
'        If Trim$(v_pedido(i)) <> "" Then
'            If Not marca_status_atendido_fila_solicitacoes_emissao_NFe(Trim$(v_pedido(i)), rNFeImg.id_nfe_emitente, CLng(strSerieNf), CLng(strNumeroNf), s_erro_aux) Then
'                s = "Não é possível prosseguir com a emissão, pois houve falha ao atualizar os dados da fila de solicitações de emissão de NFe!!" & vbCrLf & _
'                    s_erro_aux
'                aviso_erro s
'                GoSub NFE_EMITE_FECHA_TABELAS
'                aguarde INFO_NORMAL, m_id
'                Exit Sub
'                End If
'            End If
'        Next


'   MONTA TAG IDENTIFICAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~
    rNFeImg.ide__natOp = strCfopDescricao
    rNFeImg.ide__serie = strSerieNf
    rNFeImg.ide__nNF = strNumeroNf
    rNFeImg.ide__dEmi = NFeFormataData(Date)
    rNFeImg.ide__dEmiUTC = NFeFormataDataHoraUTC(Now, blnHorarioVerao)
    rNFeImg.ide__cMunFG = strEmitenteCidade & "/" & strEmitenteUf
    rNFeImg.ide__tpAmb = NFE_AMBIENTE
    rNFeImg.ide__finNFe = strNFeCodFinalidade
    rNFeImg.ide__indFinal = NFE_INDFINAL_CONSUMIDOR_FINAL
    rNFeImg.ide__indPres = strPresComprador
    
    strNFeTagIdentificacao = "ide;" & vbCrLf
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("natOp", rNFeImg.ide__natOp)
    'NFE 4.0 - não enviar indPag (Este campo agora se encontra na tag "pag"
    'strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indPag", rNFeImg.ide__indPag)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("serie", rNFeImg.ide__serie)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("nNF", rNFeImg.ide__nNF)
    '=== Substituindo campo de acordo com layout 3.10
    'strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("dEmi", rNFeImg.ide__dEmi)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("dhEmi", rNFeImg.ide__dEmiUTC)
    '=== aqui: campo dhSaiEnt
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("tpNF", rNFeImg.ide__tpNF) '0-Entrada  1-Saída
    '=== Novo campo idDest
    '=== (1-Operação Interna; 2-Operação Interestadual; 3-Operação com o Exterior)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("idDest", rNFeImg.ide__idDest)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("cMunFG", rNFeImg.ide__cMunFG)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("tpAmb", rNFeImg.ide__tpAmb) '1-Produção  2-Homologação
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("finNFe", rNFeImg.ide__finNFe) '1-Normal  2-Complementar  3-Ajuste
    '=== Novo campo indFinal
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indFinal", rNFeImg.ide__indFinal) '0-Normal  1-Consumidor Final
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indPres", rNFeImg.ide__indPres) '2-Internet  3-Teleatendimento
        '=== Campo indIntermed  (0-Sem intermediador 1-Operação em site ou plataforma de terceiros)
    If (param_nfintermediador.campo_inteiro = 1) Then
        If ((strMarketPlaceCNPJ <> "") And (strMarketPlaceCadIntTran <> "")) Then
            strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indIntermed", "1")
        Else
            strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indIntermed", "0")
            End If
        End If
    '=== aqui: campo IEST
    
    '=== Grupo NFref
    strNFeChaveAcessoNotaReferenciada = Trim$(c_chave_nfe_ref)
    If strNFeChaveAcessoNotaReferenciada <> "" Then
        vListaNFeRef = Split(strNFeChaveAcessoNotaReferenciada, vbCrLf)
        For i = LBound(vListaNFeRef) To UBound(vListaNFeRef)
            strNFeRef = Trim$(vListaNFeRef(i))
            If strNFeRef <> "" Then
                strNFeTagIdentificacao = strNFeTagIdentificacao & _
                                        "NFref;" & vbCrLf & _
                                        vbTab & NFeFormataCampo("refNFe", strNFeRef)
                If Trim$(vNFeImgNFeRef(UBound(vNFeImgNFeRef)).refNFe) <> "" Then
                    ReDim Preserve vNFeImgNFeRef(UBound(vNFeImgNFeRef) + 1)
                    End If
                vNFeImgNFeRef(UBound(vNFeImgNFeRef)).refNFe = strNFeRef
                End If
            Next
        End If

'   MONTA O ARQUIVO DE INTEGRAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNFeArquivo = strNFeTagOperacional & _
                   strNFeTagIdentificacao & _
                   strNFeTagDestinatario & _
                   strNFeTagEndEntrega & _
                   strNFeTagBlocoProduto & _
                   strNFeTagValoresTotais & _
                   strNFeTagTransp & _
                   strNFeTagTransporta & _
                   strNFeTagVol & _
                   strNFeTagFat & _
                   strNFeTagDup & _
                   strNFeTagPag & _
                   strNFeTagInfAdicionais
    
    
'   REGISTRA DADOS DA NFE P/ FINS DE HISTÓRICO, CONTROLE E CONSULTA DA DANFE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "gravando histórico no sistema"
    
    If Not grava_NFe_imagem(usuario.id, CLng(strSerieNf), CLng(strNumeroNf), rNFeImg, vNFeImgItem(), vNFeImgTagDup(), vNFeImgNFeRef(), vNFeImgPag(), lngNsuNFeImagem, s_erro_aux) Then
        s = "Falha ao tentar gravar os dados da NFe (tabela imagem)!!"
        If s_erro_aux <> "" Then s = s & vbCrLf
        s = s & s_erro_aux
        aviso_erro s
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
            
'   LEMBRANDO QUE OS CAMPOS 'dt_emissao' E 'dt_hr_emissao' SÃO PREENCHIDOS AUTOMATICAMENTE POR UM "CONSTRAINT DEFAULT"
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (id = -1)"
    If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    t_NFe_EMISSAO.AddNew
    t_NFe_EMISSAO("id") = lngNsuNFeEmissao
    t_NFe_EMISSAO("id_nfe_emitente") = rNFeImg.id_nfe_emitente
    t_NFe_EMISSAO("NFe_serie_NF") = CLng(strSerieNf)
    t_NFe_EMISSAO("NFe_numero_NF") = CLng(strNumeroNf)
    t_NFe_EMISSAO("versao_layout_NFe") = ID_VERSAO_LAYOUT_NFe
    t_NFe_EMISSAO("usuario_emissao") = usuario.id
    t_NFe_EMISSAO("pedido") = rNFeImg.pedido
    t_NFe_EMISSAO("email_destinatario") = rNFeImg.operacional__email
    t_NFe_EMISSAO("nome_destinatario") = rNFeImg.dest__xNome
    t_NFe_EMISSAO("tipo_NF") = rNFeImg.ide__tpNF
    t_NFe_EMISSAO("tipo_ambiente") = NFE_AMBIENTE
    t_NFe_EMISSAO("finalidade_NF") = rNFeImg.ide__finNFe
    t_NFe_EMISSAO("natureza_operacao_codigo") = strCfopCodigoFormatado
    t_NFe_EMISSAO("natureza_operacao_descricao") = strCfopDescricao
    t_NFe_EMISSAO("aliquota_ICMS") = perc_ICMS
    t_NFe_EMISSAO("aliquota_IPI") = perc_IPI
    t_NFe_EMISSAO("frete_por_conta") = rNFeImg.transp__modFrete
    t_NFe_EMISSAO("volumes_qtde_total_sistema") = total_volumes
    t_NFe_EMISSAO("volumes_qtde_total_tela") = c_total_volumes
    
    s = RTrim$(c_dados_adicionais_venda)
    lngMax = 2000
    If Len(s) > lngMax Then
        s_aux = " (...)"
        s = left$(s, lngMax - Len(s_aux)) & s_aux
        End If
    t_NFe_EMISSAO("dados_adicionais_digitado") = s
    
    s = strNFeArquivo
    lngMax = 6000
    If Len(s) > lngMax Then
        s_aux = " (...)"
        s = left$(s, lngMax - Len(s_aux)) & s_aux
        End If
    t_NFe_EMISSAO("arquivo_integracao_NFe_T1") = s
    t_NFe_EMISSAO.Update
    
            
'   TRANSFERE O ARQUIVO DE INTEGRAÇÃO PARA O SISTEMA DE NFe DA TARGET ONE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNumeroNfNormalizado = NFeFormataNumeroNF(strNumeroNf)
    strSerieNfNormalizado = NFeFormataSerieNF(strSerieNf)

  ' COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    aguarde INFO_EXECUTANDO, "emitindo NFe"
    Set cmdNFeEmite.ActiveConnection = dbcNFe
    cmdNFeEmite.CommandType = adCmdStoredProc
    cmdNFeEmite.CommandText = "Proc_NFe_Integracao_Emite"
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("NFe", adChar, adParamInput, 9, strNumeroNfNormalizado)
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("Serie", adChar, adParamInput, 3, strSerieNfNormalizado)
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("Arquivo", adVarChar, adParamInput, 16000, strNFeArquivo)
    Set rsNFeRetornoSPEmite = cmdNFeEmite.Execute
    intNfeRetornoSPEmite = rsNFeRetornoSPEmite("Retorno")
    strNFeMsgRetornoSPEmite = Trim$("" & rsNFeRetornoSPEmite("Mensagem"))
    
'   GRAVA O RESULTADO DA CHAMADA DA STORED PROCEDURE
    strNFeMsgRetornoSPEmiteTamAjustadoBD = strNFeMsgRetornoSPEmite
    lngMax = 2000
    If Len(strNFeMsgRetornoSPEmiteTamAjustadoBD) > lngMax Then
        s_aux = " (...)"
        strNFeMsgRetornoSPEmiteTamAjustadoBD = left$(strNFeMsgRetornoSPEmiteTamAjustadoBD, lngMax - Len(s_aux)) & s_aux
        End If
    
    Call atualiza_NFe_imagem_com_retorno_NFe_T1(lngNsuNFeImagem, CStr(intNfeRetornoSPEmite), strNFeMsgRetornoSPEmiteTamAjustadoBD, s_erro_aux)
    
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (id = " & lngNsuNFeEmissao & ")"
    If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If Not t_NFe_EMISSAO.EOF Then
        t_NFe_EMISSAO("codigo_retorno_NFe_T1") = CStr(intNfeRetornoSPEmite)
        t_NFe_EMISSAO("msg_retorno_NFe_T1") = strNFeMsgRetornoSPEmiteTamAjustadoBD
        t_NFe_EMISSAO.Update
        End If
        
'   ATUALIZA AS INFORMAÇÕES SOBRE A EMISSÃO TRIANGULAR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If blnEsperaNFTriangular Then
        If Not atualiza_nfe_triangular_venda(lngIdNFeTriangular, ST_NFT_EMITIDA, s_erro) Then
            s_erro = "Problemas no registro da operação triangular (nota de venda): " & s_erro
            aviso_erro s_erro
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
    Else
        If Not insere_registro_nfe_triangular(lngIdNFeTriangular, lngSerieNFeTriangular, lngNumVendaNFeTriangular, lngNumRemessaNFeTriangular, c_pedido, False, s_erro) Or _
            Not atualiza_nfe_triangular_venda(lngIdNFeTriangular, ST_NFT_EMITIDA, s_erro) Then
            s_erro = "Problemas no registro da operação triangular (nota de venda): " & s_erro
            aviso_erro s_erro
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    If Not atualiza_nfe_triangular_inf_adicionais(lngIdNFeTriangular, _
                                                    retorna_so_digitos(c_cnpj_cpf_dest), _
                                                    c_nome_dest, _
                                                    c_rg_dest, _
                                                    l_end_recebedor_logradouro, _
                                                    l_end_recebedor_numero, _
                                                    l_end_recebedor_complemento, _
                                                    l_end_recebedor_bairro, _
                                                    l_end_recebedor_cidade, _
                                                    l_end_recebedor_uf, _
                                                    l_end_recebedor_cep, _
                                                    c_dados_adicionais_venda, _
                                                    "", _
                                                    s_erro) Then
        s_erro = "Problemas na atualização da nota triangular de venda: " & s_erro
        aviso_erro s_erro
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
'   SE HOUVE BLOQUEIO DE ESPERA, ATUALIZAR O Nº NA t_NFE_EMITENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If blnEsperaNFTriangular Then
        s = "UPDATE n SET" & _
                " n.NFe_numero_NF = " & strSerieNfTriangular & _
            " FROM t_NFE_EMITENTE e" & _
            " INNER JOIN t_NFE_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" & _
            " WHERE" & _
                " (e.id = " & CStr(intIdNfeEmitente) & ")" & _
                " AND (n.NFe_serie_NF = " & CStr(strSerieNf) & ")"
        Call dbc.Execute(s, lngAffectedRecords)
        If lngAffectedRecords <> 1 Then
            s = "Falha ao atualizar a numeração sequencial para o emitente atual!!" & vbCrLf & s
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
'   GRAVA O LOG
'   ~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "gravando log"
    If LBound(v_pedido) = UBound(v_pedido) Then
        strLogPedido = v_pedido(UBound(v_pedido))
    Else
        strLogPedido = ""
        End If
    strLogComplemento = "Retorno SP=" & CStr(intNfeRetornoSPEmite) & " (" & IIf(intNfeRetornoSPEmite = 1, "Sucesso", "Falha") & ")" & _
                        "; Msg SP=" & strNFeMsgRetornoSPEmite & _
                        "; Série NFe=" & strSerieNf & _
                        "; Nº NFe=" & strNumeroNf & _
                        "; tipo=" & cb_tipo_NF & _
                        "; pedido=" & Join(v_pedido, ", ") & _
                        "; natureza operação=" & cb_natureza & _
                        "; ICMS=" & cb_icms & _
                        "; IPI=" & c_ipi & _
                        "; frete=" & cb_frete & _
                        "; zerar PIS=(" & Trim$(cb_zerar_PIS) & ")" & _
                        "; zerar COFINS=(" & Trim$(cb_zerar_COFINS) & ")" & _
                        "; finalidade=" & Trim$(cb_finalidade) & _
                        "; chave NFe referenciada=" & Trim$(c_chave_nfe_ref) & _
                        "; dados adicionais=" & Trim$(c_dados_adicionais_venda)
    Call grava_log(usuario.id, "", strLogPedido, "", OP_LOG_NFE_EMISSAO, strLogComplemento)
        
        
'   SUCESSO NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "processamento complementar"
    If intNfeRetornoSPEmite = 1 Then
        aguarde INFO_EXECUTANDO, "atualizando banco de dados"
    '   ATUALIZA O CAMPO "OBSERVAÇÕES II" COM O Nº DA NOTA FISCAL?
    '   A ATUALIZAÇÃO É FEITA SOMENTE P/ NOTAS DE SAÍDA, POIS EM NOTAS DE ENTRADA O Nº DA NFe NÃO É ANOTADO NO CAMPO
    '   OBS_2 DO PEDIDO, MAS SIM NOS ITENS DEVOLVIDOS, QUANDO APLICÁVEL.
    '   0-Entrada  1-Saída
        If rNFeImg.ide__tpNF = "1" Then
            If qtde_pedidos = 1 Then
              ' T_PEDIDO
                If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
                t_PEDIDO.CursorType = BD_CURSOR_EDICAO
                s = sql_monta_criterio_texto_or(v_pedido(), "pedido", True)
                s = "SELECT * FROM t_PEDIDO WHERE (" & s & ")"
                t_PEDIDO.Open s, dbc, , , adCmdText
                If Not t_PEDIDO.EOF Then
                    If blnNotadeCompromisso Then
                        If (Trim$("" & t_PEDIDO("obs_4")) = "") Or IsLetra(Trim$("" & t_PEDIDO("obs_4"))) Then
                            t_PEDIDO("obs_4") = strNumeroNf
                            t_PEDIDO.Update
                            End If
                    Else
                        If (Trim$("" & t_PEDIDO("obs_2")) = "") Or IsLetra(Trim$("" & t_PEDIDO("obs_2"))) Then
                            t_PEDIDO("obs_2") = strNumeroNf
                            t_PEDIDO.Update
                            End If
                        End If
                    End If
                End If
        ElseIf rNFeImg.ide__tpNF = "0" Then
            s = sql_monta_criterio_texto_or(v_pedido(), "pedido", True)
            If s <> "" Then
                s = "UPDATE t_PEDIDO_ITEM_DEVOLVIDO SET" & _
                        " id_nfe_emitente = " & CStr(rNFeImg.id_nfe_emitente) & "," & _
                        " NFe_serie_NF = " & strSerieNf & "," & _
                        " NFe_numero_NF = " & strNumeroNf & "," & _
                        " dt_hr_anotacao_numero_NF = getdate()," & _
                        " usuario_anotacao_numero_NF = '" & usuario.id & "'" & _
                    " WHERE" & _
                        " (" & s & ")" & _
                        " AND (NFe_numero_NF = 0)"
                dbc.Execute s, lngAffectedRecords
                End If
            End If
        
    '   Tipo de NFe: 0-Entrada  1-Saída
        If rNFeImg.ide__tpNF = "1" Then
        '   GRAVA OS DADOS DE BOLETOS NO BD!!
            If Not ExisteDadosParcelasPagto(rNFeImg.pedido, s_erro) Then
                If Not gravaDadosParcelaPagto(CLng(strNumeroNf), v_parcela_pagto(), s_erro) Then
                    If s_erro <> "" Then s_erro = Chr(13) & Chr(13) & s_erro
                    s_erro = "Falha ao gravar as informações dos boletos no banco de dados!!" & s_erro
                    aviso_erro s_erro
                    End If
                End If
            End If
            
'   FALHA NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        aviso_erro "Falha na emissão da NFe:" & vbCrLf & strNFeMsgRetornoSPEmite
        End If
        
        
  ' LIMPA FORMULÁRIO
    c_pedido_danfe = rNFeImg.pedido
    ' (retirada limpeza, para poder dar sequência na nota de remessa)
    'formulario_limpa
    '(vamos verificar o que acontece se recarregarmos o pedido - zerando a variável pedido_anterior para efetuar consulta)
    pedido_anterior = ""
    pedido_preenche_dados_tela c_pedido
        
  ' EXIBE DADOS DA ÚLTIMA NFe EMITIDA
    l_serie_NF = strSerieNfNormalizado
    l_num_NF = strNumeroNfNormalizado
    l_emitente_NF = strEmitenteNf
        
    GoSub NFE_EMITE_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
    sPedidoDANFETelaAnterior = rNFeImg.pedido
    sNFAnteriorSerie = l_serie_NF
    sNFAnteriorNumero = l_num_NF
    sNFAnteriorEmitente = l_emitente_NF
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA:
'=======================================
    aviso_erro s_erro
    GoSub NFE_EMITE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub NFE_EMITE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_FECHA_TABELAS:
'=======================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_PEDIDO_ITEM_DEVOLVIDO, True
    bd_desaloca_recordset t_DESTINATARIO, True
    bd_desaloca_recordset t_TRANSPORTADORA, True
    bd_desaloca_recordset t_IBPT, True
    bd_desaloca_recordset t_NFe_EMITENTE_X_LOJA, True
    'bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset t_NFe_IMAGEM, True
    bd_desaloca_recordset t_T1_NFE_INUTILIZA, True
    bd_desaloca_recordset t_CODIGO_DESCRICAO, True
    bd_desaloca_recordset t_NFe_UF_PARAMETRO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPEmite, True
  
  ' COMMAND
    bd_desaloca_command cmdNFeEmite
    bd_desaloca_command cmdNFeSituacao
    
  ' CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
    
    Return

End Sub

Sub NFe_emite_remessa()
' __________________________________________________________________________________________
'|
'|  EMITE A NOTA FISCAL ELETRÔNICA (NFe) DE REMESSA COM BASE NO PEDIDO
'|  ESPECIFICADO E NOS DEMAIS PARÂMETROS PREENCHIDOS MANUALMENTE.
'|
'|  OS PRODUTOS (T_PEDIDO_ITEM) COM PRECO_NF = R$ 0,00 SÃO
'|  RELATIVOS A BRINDES E DEVEM SER TOTALMENTE IGNORADOS.
'|  OS BRINDES ACOMPANHAM OS OUTROS PRODUTOS DENTRO DA MESMA CAIXA.
'|

' CONSTANTES
Const NomeDestaRotina = "NFe_emite_remessa()"
Const MAX_LINHAS_NOTA_FISCAL_DEFAULT = 19
Const NFE_AMBIENTE_PRODUCAO = "1" '1-Produção  2-Homologação
Const NFE_AMBIENTE_HOMOLOGACAO = "2" '1-Produção  2-Homologação
'Const NFE_FINALIDADE_NFE = "1" '1-Normal  2-Complementar  3-Ajuste
Const NFE_INDFINAL_CONSUMIDOR_NORMAL = "0"
Const NFE_INDFINAL_CONSUMIDOR_FINAL = "1"


' STRINGS
Dim NFE_AMBIENTE As String
Dim c As String
Dim s As String
Dim s_confirma As String
Dim s_aux As String
Dim s_msg As String
Dim s_serie_NF_aux As String
Dim s_numero_NF_aux As String
Dim s_erro As String
Dim s_erro_aux As String
Dim strCampo As String
Dim strCnpjCpfAux As String
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim strRamal As String
Dim strConfirmacaoEtgImediata As String
Dim strIcms As String
Dim strSerieNf As String
Dim strSerieNfNormalizado As String
Dim strNumeroNf As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfTriangular As String
Dim strNumeroNfTriangular As String
Dim strEmitenteNf As String
Dim strIdCliente As String
Dim strFabricanteAnterior As String
Dim strProdutoAnterior As String
Dim strPedidoAnterior As String
Dim strLoja As String
Dim strOrigemUF As String
Dim strDestinoUF As String
Dim strPresComprador As String
Dim strConfirmacaoObs3 As String
Dim strTransportadoraId As String
Dim strTransportadoraCnpj As String
Dim strTransportadoraRazaoSocial As String
Dim strTransportadoraIE As String
Dim strTransportadoraUF As String
Dim strTransportadoraEmail As String
Dim strListaPedidosSemTransportadora As String
Dim strListaPedidosComTransportadora As String
Dim strTipoParcelamento As String
Dim strLogPedido As String
Dim strLogComplemento As String
Dim strNFeCodFinalidade As String
Dim strNFeCodFinalidadeAux As String
Dim strNFeChaveAcessoNotaReferenciada As String
Dim strNFeArquivo As String
Dim strNFeTagOperacional As String
Dim strNFeTagIdentificacao As String
Dim strNFeTagDestinatario As String
Dim strNFeTagEndEntrega As String
Dim strNFeTagBlocoProduto As String
Dim strNFeTagDet As String
Dim strNFeTagIcms As String
Dim strNFeCst As String
Dim strNFeTagPis As String
Dim strNFeTagCofins As String
Dim strNFeTagIcmsUFDest As String
Dim strNFeTagValoresTotais As String
Dim strNFeTagTransp As String
Dim strNFeTagTransporta As String
Dim strNFeTagVol As String
Dim strNFeTagDup As String
Dim strNFeTagInfAdicionais As String
Dim strNFeTagPag As String
Dim strNFeInfAdicQuadroProdutos As String
Dim strNFeInfAdicQuadroInfAdic As String
Dim strCfopCodigo As String
Dim strCfopCodigoFormatado As String
Dim strCfopDescricao As String
Dim strCfopCodigoAux As String
Dim strCfopCodigoFormatadoAux As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim strDestinatarioCnpjCpf As String
Dim strEndEtgEndereco As String
Dim strEndEtgEnderecoNumero As String
Dim strEndEtgEnderecoComplemento As String
Dim strEndEtgBairro As String
Dim strEndEtgCidade As String
Dim strEndEtgUf As String
Dim strEndEtgCep As String
Dim strEndEtgEnderecoCompletoFormatado As String
Dim strEndClienteUf As String
Dim strEmitenteCidade As String
Dim strEmitenteUf As String
Dim strNFeMsgRetornoSPSituacao As String
Dim strNFeMsgRetornoSPEmite As String
Dim strNFeMsgRetornoSPEmiteTamAjustadoBD As String
Dim strCodStatusInutilizacao As String
Dim strListaSugeridaMunicipiosIBGE As String
Dim strTextoCubagem As String
Dim strZerarPisCst As String
Dim strZerarCofinsCst As String
Dim strInfoAdicIbpt As String
Dim strEmailXML As String
Dim strNFeRef As String
Dim strPedidoBSMarketplace As String
Dim strMarketplaceCodOrigem As String
Dim strMarketPlaceCNPJ As String
Dim strMarketPlaceCadIntTran As String


' FLAGS
Dim blnAchou As Boolean
Dim blnTemPedidoComTransportadora As Boolean
Dim blnTemPedidoSemTransportadora As Boolean
Dim blnTemPedidoComStBemUsoConsumo As Boolean
Dim blnTemPedidoSemStBemUsoConsumo As Boolean
Dim blnTemPagtoPorBoleto As Boolean
Dim blnImprimeDadosFatura As Boolean
Dim blnIsDestinatarioPJ As Boolean
Dim blnTemEndEtg As Boolean
Dim blnHaProdutoCstIcms60 As Boolean
Dim blnErro As Boolean
Dim blnExibirTotalTributos As Boolean
Dim blnHaProdutoSemDadosIbpt As Boolean
Dim blnExisteMemorizacaoEndereco As Boolean
Dim blnNotadeCompromisso As Boolean

' CONTADORES
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim n As Long
Dim ic As Integer
Dim intNumItem As Integer
Dim intIdNfeEmitente As Integer
Dim iQtdConfirmaDuvidaEmit As Integer

' QUANTIDADES
Dim qtde As Long
Dim total_volumes As Long
Dim qtde_pedidos As Integer
Dim qtde_linhas_nf As Integer
Dim idx As Integer
Dim lngMax As Long
Dim lngAffectedRecords As Long
Dim MAX_LINHAS_NOTA_FISCAL As Integer

' CÓDIGOS E NSU
Dim intNfeRetornoSPSituacao As Integer
Dim intNfeRetornoSPEmite As Integer
Dim lngNsuNFeEmissao As Long
Dim lngNsuNFeImagem As Long
Dim lngNFeUltNumeroNfEmitido As Long
Dim lngNFeUltSerieEmitida As Long
Dim lngNFeSerieManual As Long
Dim lngNFeNumeroNfManual As Long
Dim intContribuinteICMS As Integer
Dim intAnoPartilha As Integer
Dim intImprimeIntermediadorAusente As Integer

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_PEDIDO_ITEM_DEVOLVIDO As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset
Dim t_TRANSPORTADORA As ADODB.Recordset
Dim t_IBPT As ADODB.Recordset
Dim t_NFe_EMITENTE_X_LOJA As ADODB.Recordset
'Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim t_T1_NFE_INUTILIZA As ADODB.Recordset
Dim t_CODIGO_DESCRICAO As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPEmite As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeEmite As New ADODB.Command
Dim dbcNFe As ADODB.Connection

' MOEDA
Dim vl_unitario As Currency
Dim vl_total_produtos As Currency
Dim vl_total_BC_ICMS As Currency
Dim vl_total_BC_ICMS_ST As Currency
Dim vl_BC_ICMS As Currency
Dim vl_BC_ICMS_ST As Currency
Dim vl_BC_ICMS_ST_Ret As Currency
Dim vl_ICMS As Currency
Dim vl_ICMSDeson As Currency
Dim vl_ICMS_ST As Currency
Dim vl_ICMS_ST_Ret As Currency
Dim vl_IPI As Currency
Dim vl_total_ICMS As Currency
Dim vl_total_ICMSDeson As Currency
Dim vl_total_ICMS_ST As Currency
Dim vl_total_IPI As Currency
Dim vl_aux As Currency
Dim vl_total_outras_despesas_acessorias As Currency
Dim vl_BC_PIS As Currency
Dim vl_PIS As Currency
Dim vl_total_PIS As Currency
Dim vl_BC_COFINS As Currency
Dim vl_COFINS As Currency
Dim vl_total_COFINS As Currency
Dim vl_estimado_tributos As Currency
Dim vl_total_estimado_tributos As Currency
Dim vl_total_NF As Currency
Dim vl_fcp As Currency
Dim vl_ICMS_UF_dest As Currency
Dim vl_ICMS_UF_remet As Currency
Dim vl_ICMS_diferencial_interestadual As Currency
Dim vl_ICMS_diferencial_aux As Currency
Dim vl_total_FCPUFDest As Currency
Dim vl_total_ICMSUFDest As Currency
Dim vl_total_ICMSUFRemet As Currency
Dim vl_total_vFCP As Currency
Dim vl_total_vFCPST As Currency
Dim vl_total_vFCPSTRet As Currency
Dim vl_total_vIPIDevol As Currency

' PERCENTUAL
Dim perc_ICMS As Single
Dim perc_ICMS_ST As Single
Dim perc_ICMS_ST_aux As Single
Dim perc_IPI As Single
Dim perc_PIS As Single
Dim perc_COFINS As Single
Dim perc_IBPT As Single
Dim perc_aux As Single
Dim perc_ICMS_interna_UF_dest As Single
Dim perc_ICMS_UF_dest As Single
Dim perc_ICMS_UF_remet As Single
Dim perc_fcp As Single
Dim perc_ICMS_diferencial_interestadual As Single

' REAL
Dim peso_aux As Single
Dim total_peso_bruto As Single
Dim total_peso_liquido As Single
Dim cubagem_aux As Single
Dim cubagem_bruto As Single
Dim aliquota_icms_interestadual As Single

' VETORES
Dim v() As String
Dim v_pedido() As String
Dim v_nf() As TIPO_LINHA_NOTA_FISCAL
Dim v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim v_nf_confere() As TIPO_LINHA_NOTA_FISCAL
Dim v_flagDadosTelaJaLido() As Boolean
Dim vListaNFeRef() As String

' DADOS DE IMAGEM DA NFE
Dim rNFeImg As TIPO_NFe_IMG
Dim vNFeImgItem() As TIPO_NFe_IMG_ITEM
Dim vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP
Dim vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA
Dim vNFeImgPag() As TIPO_NFe_IMG_PAG

    On Error GoTo NFE_EMITE_REMESSA_TRATA_ERRO
            
    c_pedido = normaliza_lista_pedidos(c_pedido)
    
    If Not pedido_eh_do_emitente_atual(c_pedido) Then Exit Sub
    
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_produto(i)) <> "" Then
            If converte_para_currency(c_vl_outras_despesas_acessorias(i)) < 0 Then
                aviso_erro "O valor das outras despesas acessórias do produto " & Trim$(c_produto(i)) & " não pode ser negativo!!"
                c_vl_outras_despesas_acessorias(i).SetFocus
                Exit Sub
                End If
            End If
        Next
    
    
    If DESENVOLVIMENTO Then
        NFE_AMBIENTE = NFE_AMBIENTE_HOMOLOGACAO
    Else
        NFE_AMBIENTE = NFE_AMBIENTE_PRODUCAO
        End If
        
    ReDim v_pedido(0)
    v_pedido(UBound(v_pedido)) = ""
    
    ReDim vNFeImgItem(0)
    ReDim vNFeImgTagDup(0)
    ReDim vNFeImgNFeRef(0)
    ReDim vNFeImgPag(0)
    
    qtde_pedidos = 0
    iQtdConfirmaDuvidaEmit = 0
    
    strNFeArquivo = ""
    strNFeTagOperacional = ""
    strNFeTagIdentificacao = ""
    strNFeTagDestinatario = ""
    strNFeTagEndEntrega = ""
    strNFeTagBlocoProduto = ""
    strNFeTagValoresTotais = ""
    strNFeTagTransp = ""
    strNFeTagTransporta = ""
    strNFeTagInfAdicionais = ""
    strNFeInfAdicQuadroProdutos = ""
    strNFeInfAdicQuadroInfAdic = ""
    strNFeTagDup = ""
    
    blnTemPedidoComStBemUsoConsumo = False
    blnTemPedidoSemStBemUsoConsumo = False
    blnTemPedidoComTransportadora = False
    blnTemPedidoSemTransportadora = False
    blnTemPagtoPorBoleto = False
    blnImprimeDadosFatura = False
    strListaPedidosSemTransportadora = ""
    strListaPedidosComTransportadora = ""
    
    v = Split(c_pedido, vbCrLf)
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
        '   REPETIDO?
            For j = LBound(v_pedido) To UBound(v_pedido)
                If Trim$(v(i)) = v_pedido(j) Then
                    aviso_erro "Pedido " & Trim$(v(i)) & " está repetido na lista!!"
                    c_pedido.SetFocus
                    Exit Sub
                    End If
                Next
                
            If v_pedido(UBound(v_pedido)) <> "" Then ReDim Preserve v_pedido(UBound(v_pedido) + 1)
            v_pedido(UBound(v_pedido)) = Trim$(v(i))
            qtde_pedidos = qtde_pedidos + 1
            End If
        Next
    
    If qtde_pedidos = 0 Then
        aviso_erro "Informe o número do pedido!!"
        c_pedido.SetFocus
        Exit Sub
        End If
        
    If qtde_pedidos > 1 Then
        aviso_erro "É possível emitir a NFe de apenas 1 pedido por vez!!"
        c_pedido.SetFocus
        Exit Sub
        End If
    
    rNFeImg.pedido = c_pedido
    
    
'   tenta buscar informações da nota fiscal de venda
'   (variáveis: s = destino da operação, s_aux = quem responde pelo frete, s_confirma = natureza da operação)
'   caso não existam (variáveis voltarão vazias), a nota de venda ainda não foi emitida
'   neste caso, abortar emissão da nota de remessa
    If obtem_dados_nf_venda(lngNumVendaNFeTriangular, lngSerieNFeTriangular, s, s_aux, s_confirma) Then
        'nop
        End If
    If (s = "") Or (s_aux = "") Or (s_confirma = "") Then
        s = ""
        s_aux = ""
        s_confirma = ""
        aviso_erro "A emissão da nota de remessa só é possível depois que a nota de venda for emitida!!"
        Exit Sub
        End If
    
'   OBTÉM TIPO DO DOCUMENTO FISCAL
'   LHGX - fixando tipo do documento em SAÍDA [no form, cb_tipo_NF estará desabilitado](verificar com Bonshop se está correto)
    rNFeImg.ide__tpNF = left$(Trim$(cb_tipo_NF), 1)
    If rNFeImg.ide__tpNF = "" Then
        aviso_erro "Selecione o tipo de documento fiscal (entrada ou saída)!!"
        Exit Sub
        End If
        
    If rNFeImg.ide__tpNF = "0" Then
        s = "A NFe que será emitida será de ENTRADA!!" & vbCrLf & "Continua com a emissão da NFe?"
        If Not confirma(s) Then
            Exit Sub
            End If
        End If
        
        
'>  NATUREZA DA OPERAÇÃO
    s = UCase$(cb_natureza_recebedor)
    strCfopCodigoFormatado = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = " " Then Exit For
        strCfopCodigoFormatado = strCfopCodigoFormatado & c
        Next
        
    strCfopCodigo = retorna_so_digitos(strCfopCodigoFormatado)
    strCfopDescricao = Trim$(Mid$(s, Len(strCfopCodigoFormatado) + 1, Len(s) - Len(strCfopCodigoFormatado)))
        
'>  LOCAL DE DESTINO DA OPERAÇÃO
    rNFeImg.ide__idDest = left$(Trim$(cb_loc_dest_remessa), 1)
        
'>  FINALIDADE DE EMISSÃO
    strNFeCodFinalidade = left$(Trim$(cb_finalidade), 1)
    If strNFeCodFinalidade = "" Then
        aviso_erro "Selecione a finalidade da NFe!!"
        Exit Sub
        End If
    
    strNFeCodFinalidadeAux = retorna_finalidade_nfe(strCfopCodigo)
    If strNFeCodFinalidade <> strNFeCodFinalidadeAux Then
        s = "Possível divergência encontrada na finalidade da NFe:" & vbCrLf & _
            "Finalidade selecionada: " & strNFeCodFinalidade & " - " & descricao_finalidade_nfe(strNFeCodFinalidade) & vbCrLf & _
            "Finalidade recomendada para o CFOP " & strCfopCodigoFormatado & ": " & strNFeCodFinalidadeAux & " - " & descricao_finalidade_nfe(strNFeCodFinalidadeAux) & _
            vbCrLf & vbCrLf & _
            "Continua com a emissão da NFe?"
        If Not confirma(s) Then
            Exit Sub
            End If
        End If
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM_DEVOLVIDO
    Set t_PEDIDO_ITEM_DEVOLVIDO = New ADODB.Recordset
    With t_PEDIDO_ITEM_DEVOLVIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_DESTINATARIO
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_TRANSPORTADORA
    Set t_TRANSPORTADORA = New ADODB.Recordset
    With t_TRANSPORTADORA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_IBPT
    Set t_IBPT = New ADODB.Recordset
    With t_IBPT
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

  ' T_NFE_EMITENTE_X_LOJA
    Set t_NFe_EMITENTE_X_LOJA = New ADODB.Recordset
    With t_NFe_EMITENTE_X_LOJA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'  ' T_FIN_BOLETO_CEDENTE
'    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
'    With t_FIN_BOLETO_CEDENTE
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
  
  ' T_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
   
   ' T_NFE_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
  
  ' T_NFe_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
  
  ' T_T1_NFE_INUTILIZA
    Set t_T1_NFE_INUTILIZA = New ADODB.Recordset
    With t_T1_NFE_INUTILIZA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_CODIGO_DESCRICAO
    Set t_CODIGO_DESCRICAO = New ADODB.Recordset
    With t_CODIGO_DESCRICAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  

'   VERIFICA CADA UM DOS PEDIDOS
    strIdCliente = ""
    strPedidoAnterior = ""
    strLoja = ""
    s_erro = ""
    strConfirmacaoObs3 = ""
    strConfirmacaoEtgImediata = ""
    strTransportadoraId = ""
    rNFeImg.ide__indPag = "2" ' Forma de pagamento: outros
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            s = "SELECT" & _
                    " t_PEDIDO.pedido," & _
                    " t_PEDIDO.loja," & _
                    " t_PEDIDO.st_entrega," & _
                    " t_PEDIDO.id_cliente," & _
                    " t_PEDIDO.obs_3," & _
                    " t_PEDIDO.transportadora_id," & _
                    " t_PEDIDO.StBemUsoConsumo," & _
                    " t_PEDIDO.st_etg_imediata," & _
                    " t_PEDIDO__BASE.tipo_parcelamento," & _
                    " t_PEDIDO__BASE.av_forma_pagto," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_entrada," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_prestacao," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_demais_prest," & _
                    " t_PEDIDO__BASE.pu_forma_pagto" & _
                " FROM t_PEDIDO" & _
                    " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
                        " ON (SUBSTRING(t_PEDIDO.pedido,1," & CStr(TAM_MIN_ID_PEDIDO) & ")=t_PEDIDO__BASE.pedido)" & _
                " WHERE" & _
                    " (t_PEDIDO.pedido='" & Trim$(v_pedido(i)) & "')"
            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
            t_PEDIDO.Open s, dbc, , , adCmdText
            If t_PEDIDO.EOF Then
                If s_erro <> "" Then s_erro = s_erro & vbCrLf
                s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " não está cadastrado!!"
            Else
                strLoja = Trim$("" & t_PEDIDO("loja"))
                
                If CLng(t_PEDIDO("StBemUsoConsumo")) = 1 Then
                    blnTemPedidoComStBemUsoConsumo = True
                Else
                    blnTemPedidoSemStBemUsoConsumo = True
                    End If
                    
                If (Trim$("" & t_PEDIDO("obs_3")) <> "") And (Not IsLetra(Trim$("" & t_PEDIDO("obs_3")))) Then
                    If strConfirmacaoObs3 <> "" Then strConfirmacaoObs3 = strConfirmacaoObs3 & vbCrLf
                    strConfirmacaoObs3 = strConfirmacaoObs3 & Trim$("" & t_PEDIDO("pedido")) & " preenchido com: " & Trim$("" & t_PEDIDO("obs_3"))
                    End If
                    
                If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " está cancelado!!"
                    End If
                    
                If CLng(t_PEDIDO("st_etg_imediata")) <> 2 Then
                    If strConfirmacaoEtgImediata <> "" Then strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & vbCrLf
                    strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & "Pedido " & Trim$(v_pedido(i)) & " NÃO está definido para 'Entrega Imediata'!!"
                    End If
                
                strTipoParcelamento = Trim$("" & t_PEDIDO("tipo_parcelamento"))
                If strTipoParcelamento = CStr(COD_FORMA_PAGTO_A_VISTA) Then
                    rNFeImg.ide__indPag = "0"  ' A vista
                    If Trim$("" & t_PEDIDO("av_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
                    rNFeImg.ide__indPag = "1"  ' A prazo
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_entrada")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_prestacao")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
                    rNFeImg.ide__indPag = "1"  ' A prazo
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_prim_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_demais_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
                    rNFeImg.ide__indPag = "2"  ' Outros
                    If Trim$("" & t_PEDIDO("pu_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnTemPagtoPorBoleto = True
                    End If
                
                If Trim$("" & t_PEDIDO("transportadora_id")) = "" Then
                    blnTemPedidoSemTransportadora = True
                    If strListaPedidosSemTransportadora <> "" Then strListaPedidosSemTransportadora = strListaPedidosSemTransportadora & ", "
                    strListaPedidosSemTransportadora = strListaPedidosSemTransportadora & Trim$(v_pedido(i))
                Else
                    blnTemPedidoComTransportadora = True
                    If strListaPedidosComTransportadora <> "" Then strListaPedidosComTransportadora = strListaPedidosComTransportadora & ", "
                    strListaPedidosComTransportadora = strListaPedidosComTransportadora & Trim$(v_pedido(i))
                    
                    If strTransportadoraId = "" Then
                        strTransportadoraId = Trim$("" & t_PEDIDO("transportadora_id"))
                    Else
                        If strTransportadoraId <> Trim$("" & t_PEDIDO("transportadora_id")) Then
                            If s_erro <> "" Then s_erro = s_erro & vbCrLf
                            s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " informa uma transportadora diferente!!"
                            End If
                        End If
                    End If
                    
            '   TODOS OS PEDIDOS DEVEM PERTENCER AO MESMO CLIENTE
                If strIdCliente = "" Then
                    strIdCliente = Trim$("" & t_PEDIDO("id_cliente"))
                    strPedidoAnterior = Trim$("" & t_PEDIDO("pedido"))
                    End If
                If strIdCliente <> Trim$("" & t_PEDIDO("id_cliente")) Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " pertence a um cliente diferente que o pedido " & strPedidoAnterior & "!!"
                    End If
                End If
            
            s = "SELECT " & _
                    "pedido, " & _
                    "fabricante, " & _
                    "produto" & _
                " FROM t_PEDIDO_ITEM" & _
                " WHERE" & _
                    " (pedido='" & Trim$(v_pedido(i)) & "')" & _
                    " AND (preco_NF > 0)"
            If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
            t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
            If t_PEDIDO_ITEM.EOF Then
                If s_erro <> "" Then s_erro = s_erro & vbCrLf
                s_erro = s_erro & "Não foi encontrado nenhum produto relacionado ao pedido " & Trim$(v_pedido(i)) & "!!"
                End If
            End If
            
            'obter as informações de marketplace
            If (param_nfintermediador.campo_inteiro = 1) And (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") Then
                s = "SELECT o.codigo, o.descricao, og.parametro_campo_texto, og.parametro_2_campo_texto, og.parametro_3_campo_flag  " & _
                    "FROM (select * from t_CODIGO_DESCRICAO where grupo = 'PedidoECommerce_Origem') o  " & _
                        "INNER JOIN (select * from t_CODIGO_DESCRICAO where grupo = 'PedidoECommerce_Origem_Grupo') og  " & _
                        "on o.codigo_pai = og.codigo " & _
                    "WHERE o.codigo = '" & strMarketplaceCodOrigem & "'"
                If t_CODIGO_DESCRICAO.State <> adStateClosed Then t_CODIGO_DESCRICAO.Close
                t_CODIGO_DESCRICAO.Open s, dbc, , , adCmdText
                If t_CODIGO_DESCRICAO.EOF Then
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Problema na identificação do marketplace do pedido " & Trim$(v_pedido(i)) & "!!"
                Else
                    strMarketPlaceCNPJ = Trim$("" & t_CODIGO_DESCRICAO("parametro_campo_texto"))
                    strMarketPlaceCadIntTran = Trim$("" & t_CODIGO_DESCRICAO("parametro_2_campo_texto"))
                    intImprimeIntermediadorAusente = t_CODIGO_DESCRICAO("parametro_3_campo_flag")
                    End If
                    
                End If
                        
        Next
        
    If s_erro = "" Then
        If blnTemPedidoComTransportadora And blnTemPedidoSemTransportadora Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Há pedido(s) com transportadora cadastrada (" & strListaPedidosComTransportadora & ") e há pedido(s) sem transportadora cadastrada (" & strListaPedidosSemTransportadora & ")!!"
            End If
        End If
        
'   ENCONTROU ERRO?
    If s_erro <> "" Then
        aviso_erro s_erro
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
        
'   OBTÉM OS DADOS DO EMITENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~
    If strLoja = "" Then
        aviso_erro "Falha ao obter o nº da loja do pedido!!"
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
                
    If usuario.emit_id <> "" Then
        intIdNfeEmitente = CInt(usuario.emit_id)
            
        s = "SELECT" & _
                " id," & _
                " razao_social," & _
                " cidade," & _
                " uf," & _
                " NFe_T1_servidor_BD," & _
                " NFe_T1_nome_BD," & _
                " NFe_T1_usuario_BD," & _
                " NFe_T1_senha_BD" & _
            " FROM t_NFE_EMITENTE" & _
            " WHERE" & _
                " (id = " & CStr(intIdNfeEmitente) & ")"
        
        t_NFE_EMITENTE.Open s, dbc, , , adCmdText
        If t_NFE_EMITENTE.EOF Then
            aviso_erro "Dados do emitente não foram localizados no BD (id=" & CStr(intIdNfeEmitente) & ")!!"
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            strEmitenteNf = Trim$("" & t_NFE_EMITENTE("razao_social"))
            strEmitenteCidade = Trim$("" & t_NFE_EMITENTE("cidade"))
            strEmitenteUf = Trim$("" & t_NFE_EMITENTE("uf"))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            End If
    Else
    '   OBTÉM O EMITENTE PADRÃO
        s = "SELECT" & _
                " id," & _
                " razao_social," & _
                " cidade," & _
                " uf," & _
                " NFe_T1_servidor_BD," & _
                " NFe_T1_nome_BD," & _
                " NFe_T1_usuario_BD," & _
                " NFe_T1_senha_BD" & _
            " FROM t_NFE_EMITENTE" & _
            " WHERE" & _
                " (NFe_st_emitente_padrao = 1)"
        
        t_NFE_EMITENTE.Open s, dbc, , , adCmdText
        If t_NFE_EMITENTE.EOF Then
            aviso_erro "Não há emitente padrão definido no sistema!!"
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        ElseIf t_NFE_EMITENTE.RecordCount > 1 Then
            aviso_erro "Há mais de 1 emitente padrão definido no sistema!!"
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            intIdNfeEmitente = t_NFE_EMITENTE("id")
            strEmitenteNf = Trim$("" & t_NFE_EMITENTE("razao_social"))
            strEmitenteCidade = Trim$("" & t_NFE_EMITENTE("cidade"))
            strEmitenteUf = Trim$("" & t_NFE_EMITENTE("uf"))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            End If
        End If
   
    rNFeImg.id_nfe_emitente = intIdNfeEmitente
   
    
    'OBTÉM O INDICADOR DE PRESENÇA DO COMPRADOR NO ESTABELECIMENTO COMERCIAL NO MOMENTO DA OPERAÇÃO
    'se loja for 201 (E-Commerce), indicador será 2 (Internet); senão, indicador será 3 (Teleatendimento)
    strPresComprador = ""
    If strLoja = "201" Then
        strPresComprador = "2"
    Else
        strPresComprador = "3"
        End If

    ' OBTÉM UF DO EMITENTE (pegar UF do emitente padrão, conforme conversa entre Hamilton e Luiz em 21/10/2014)
    strOrigemUF = strEmitenteUf
        
        
'   CONEXÃO AO BD NFE
'   ~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "conectando ao banco dados de NFe"
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
    decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
    s = "Provider=" & BD_OLEDB_PROVIDER & _
        ";Data Source=" & strNfeT1ServidorBd & _
        ";Initial Catalog=" & strNfeT1NomeBd & _
        ";User Id=" & strNfeT1UsuarioBd & _
        ";Password=" & s_aux
    dbcNFe.Open s
    
        
'LHGX - como o pedido certamente terá a nota de venda emitida, desconsiderar as verificações abaixo
'(comentando duplamente o trecho em questão para identificação)
'''   VERIFICA SE O PEDIDO JÁ TEM NFe EMITIDA
''   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''    Set cmdNFeSituacao.ActiveConnection = dbcNFe
''    cmdNFeSituacao.CommandType = adCmdStoredProc
''    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
''    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
''    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
''
''    For i = LBound(v_pedido) To UBound(v_pedido)
''        If Trim$(v_pedido(i)) <> "" Then
''            s = "SELECT DISTINCT" & _
''                    " NFe_serie_NF," & _
''                    " NFe_numero_NF" & _
''                " FROM t_NFe_EMISSAO" & _
''                " WHERE" & _
''                    " (pedido = '" & Trim$(v_pedido(i)) & "')" & _
''                " ORDER BY" & _
''                    " NFe_serie_NF," & _
''                    " NFe_numero_NF"
''            If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
''            t_NFe_EMISSAO.Open s, dbc, , , adCmdText
''
''            s_msg = ""
''            j = 0
''            Do While Not t_NFe_EMISSAO.EOF
''                j = j + 1
''                s_serie_NF_aux = NFeFormataSerieNF(Trim$("" & t_NFe_EMISSAO("NFe_serie_NF")))
''                s_numero_NF_aux = NFeFormataNumeroNF(Trim$("" & t_NFe_EMISSAO("NFe_numero_NF")))
''
''                cmdNFeSituacao.Parameters("NFe") = s_numero_NF_aux
''                cmdNFeSituacao.Parameters("Serie") = s_serie_NF_aux
''                Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
''                intNfeRetornoSPSituacao = rsNFeRetornoSPSituacao("Retorno")
''                strNFeMsgRetornoSPSituacao = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
''
''                If s_msg <> "" Then s_msg = s_msg & vbCrLf
''                s_msg = s_msg & CStr(j) & ") " & _
''                    "Série: " & s_serie_NF_aux & _
''                    ", Nº: " & s_numero_NF_aux & _
''                    ", Situação: " & intNfeRetornoSPSituacao & " - " & strNFeMsgRetornoSPSituacao
''                t_NFe_EMISSAO.MoveNext
''                Loop
''
''            If s_msg <> "" Then
''                s_msg = "O pedido " & Trim$(v_pedido(i)) & " já possui NFe que se encontra na seguinte situação:" & vbCrLf & s_msg
''                s_msg = s_msg & vbCrLf & vbCrLf & "Continua com a emissão desta NFe?"
''                If Not confirma(s_msg) Then
''                    GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
''                    aguarde INFO_NORMAL, m_id
''                    Exit Sub
''                    End If
''                End If
''            End If
''        Next
'
'
'''   O(S) PEDIDO(S) ESTÁ COM 'ENTREGA IMEDIATA' IGUAL A 'NÃO'?
''    If strConfirmacaoEtgImediata <> "" Then
''        strConfirmacaoEtgImediata = strConfirmacaoEtgImediata & _
''                                    vbCrLf & vbCrLf & "Continua com a emissão da NFe?"
''        If Not confirma(strConfirmacaoEtgImediata) Then
''            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
''            aguarde INFO_NORMAL, m_id
''            Exit Sub
''            End If
''        End If
        
'   SE HÁ PEDIDO COM O CAMPO "OBSERVAÇÕES III" JÁ PREENCHIDO, DEVE AVISAR E PEDIR CONFIRMAÇÃO ANTES DE PROSSEGUIR
'   A CONFIRMAÇÃO É FEITA SOMENTE P/ NOTAS DE SAÍDA, POIS EM NOTAS DE ENTRADA O Nº DA NFe NÃO É ANOTADO NO CAMPO
'   OBS_3 DO PEDIDO, MAS SIM NOS ITENS DEVOLVIDOS, QUANDO APLICÁVEL.
'   0-Entrada  1-Saída
    If rNFeImg.ide__tpNF = "1" Then
        If strConfirmacaoObs3 <> "" Then
            strConfirmacaoObs3 = "O campo " & Chr$(34) & "Observações III" & Chr$(34) & " já está preenchido nos seguintes pedidos:" & _
                                 vbCrLf & strConfirmacaoObs3 & _
                                 vbCrLf & vbCrLf & "Continua com a emissão da NFe?"
            If Not confirma(strConfirmacaoObs3) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
        
        
'   NO CASO DE UM PRODUTO APARECER EM VÁRIOS PEDIDOS E O PREÇO DE VENDA FOR DIFERENTE,
'   DEVE PEDIR UMA CONFIRMAÇÃO AO OPERADOR ANTES DE USAR A MÉDIA DO PREÇO DE VENDA
    If qtde_pedidos > 1 Then
        s = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
        s = "SELECT " & _
                "fabricante, " & _
                "produto, " & _
                "preco_NF, " & _
                "t_PEDIDO_ITEM.pedido, " & _
                "descricao" & _
            " FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
            " WHERE" & _
                " (" & s & ")" & _
                " AND (preco_NF > 0)" & _
            " ORDER BY " & _
                "fabricante, " & _
                "produto, " & _
                "t_PEDIDO.data, " & _
                "t_PEDIDO.pedido"
        strFabricanteAnterior = "XXXXX"
        strProdutoAnterior = "XXXXXXXXXX"
        vl_aux = 0
        s_erro = ""
        n = 0
        If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
        t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
        Do While Not t_PEDIDO_ITEM.EOF
            If (strFabricanteAnterior = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And (strProdutoAnterior = Trim$("" & t_PEDIDO_ITEM("produto"))) Then
                If vl_aux <> t_PEDIDO_ITEM("preco_NF") Then
                    n = n + 1
                    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                    s_erro = s_erro & "Produto " & Trim$("" & t_PEDIDO_ITEM("produto")) & " do fabricante " & Trim$("" & t_PEDIDO_ITEM("fabricante")) & ":   " & Trim$("" & t_PEDIDO_ITEM("pedido")) & " = " & Format$(t_PEDIDO_ITEM("preco_NF"), FORMATO_MOEDA) & "   " & strPedidoAnterior & " = " & Format$(vl_aux, FORMATO_MOEDA)
                    End If
                End If
            
            strFabricanteAnterior = Trim$("" & t_PEDIDO_ITEM("fabricante"))
            strProdutoAnterior = Trim$("" & t_PEDIDO_ITEM("produto"))
            strPedidoAnterior = Trim$("" & t_PEDIDO_ITEM("pedido"))
            vl_aux = t_PEDIDO_ITEM("preco_NF")
            
            t_PEDIDO_ITEM.MoveNext
            Loop
        
        If s_erro <> "" Then
            If n = 1 Then
                s = "O seguinte produto aparece em mais de um pedido com preços de venda diferentes!!"
            Else
                s = "Os seguintes produtos aparecem em mais de um pedido com preços de venda diferentes!!"
                End If
            s_erro = s & vbCrLf & _
                "Continua com a emissão da nota usando o valor médio do preço de venda?" & _
                vbCrLf & vbCrLf & s_erro
            If Not confirma(s_erro) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
    
'   OBTÉM OS PRODUTOS E AS QUANTIDADES P/ USAR NA CONFERÊNCIA
    ReDim v_nf_confere(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf_confere(UBound(v_nf_confere))
    
    s = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
    s = "SELECT" & _
            " t_PEDIDO.pedido," & _
            " t_PEDIDO.data," & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.qtde" & _
        " FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
        " WHERE" & _
            " (" & s & ")" & _
            " AND (preco_NF > 0)" & _
        " ORDER BY " & _
            "produto, " & _
            "t_PEDIDO.data, " & _
            "t_PEDIDO.pedido"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    Do While Not t_PEDIDO_ITEM.EOF
        blnAchou = False
        For i = LBound(v_nf_confere) To UBound(v_nf_confere)
            With v_nf_confere(i)
                If (.fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And (.produto = Trim$("" & t_PEDIDO_ITEM("produto"))) Then
                    blnAchou = True
                    idx = i
                    Exit For
                    End If
                End With
            Next
        
        If Not blnAchou Then
            If v_nf_confere(UBound(v_nf_confere)).produto <> "" Then
                ReDim Preserve v_nf_confere(UBound(v_nf_confere) + 1)
                limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf_confere(UBound(v_nf_confere))
                End If
            idx = UBound(v_nf_confere)
            With v_nf_confere(UBound(v_nf_confere))
                .fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))
                .produto = Trim$("" & t_PEDIDO_ITEM("produto"))
                End With
            End If
        
        With v_nf_confere(idx)
        '   QUANTIDADE
            qtde = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde")) Then qtde = CLng(t_PEDIDO_ITEM("qtde"))
            .qtde_total = .qtde_total + qtde
            End With
        
        t_PEDIDO_ITEM.MoveNext
        Loop


'   OBTÉM OS DADOS DOS PRODUTOS
'   A QUANTIDADE DE PRODUTOS (IDENTIFICADO PELO CÓDIGO NCM) QUE DEU ENTRADA DEVE
'   COINCIDIR COM A QUANTIDADE QUE DEU SAÍDA. SENDO QUE O CÓDIGO NCM E/OU O CST
'   DE UM PRODUTO PODE SER ALTERADO PELO SEU FABRICANTE.
'   PORTANTO, A PARTIR DA VERSÃO 1.48 DESTE MÓDULO, O CÓDIGO NCM E O CST PASSAM
'   A SER REGISTRADOS NO MOMENTO DA ENTRADA DAS MERCADORIAS NO ESTOQUE E ESSES
'   CÓDIGOS É QUE SERÃO USADOS NA EMISSÃO DA NFe.
    ReDim v_nf(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf(UBound(v_nf))
    
'   A ORDENAÇÃO É FEITA SOMENTE PELO CÓDIGO DO PRODUTO PORQUE NA NOTA FISCAL NÃO HÁ COLUNA PARA O CÓDIGO DO FABRICANTE
    qtde_linhas_nf = 0
    s_aux = sql_monta_criterio_texto_or(v_pedido(), "t_PEDIDO_ITEM.pedido", True)
    s = "SELECT" & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.descricao," & _
            " t_PEDIDO_ITEM.ean," & _
            " t_PEDIDO_ITEM.preco_NF," & _
            " t_PEDIDO_ITEM.qtde_volumes," & _
            " t_PEDIDO_ITEM.peso," & _
            " t_PEDIDO_ITEM.cubagem," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst," & _
            " Coalesce(t_PRODUTO.perc_MVA_ST, 0) AS perc_MVA_ST," & _
            " Coalesce(t_PRODUTO.ean, '') AS tP_ean," & _
            " Coalesce(t_PRODUTO.peso, 0) AS tP_peso," & _
            " Coalesce(t_PRODUTO.cubagem, 0) AS tP_cubagem," & _
            " Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde"
    s = s & _
        " FROM t_PEDIDO_ITEM" & _
            " INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
            " LEFT JOIN t_PRODUTO ON (t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto)" & _
            " INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_PEDIDO_ITEM.pedido=t_ESTOQUE_MOVIMENTO.pedido) AND (t_PEDIDO_ITEM.fabricante=t_ESTOQUE_MOVIMENTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_ESTOQUE_MOVIMENTO.produto)" & _
            " INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto)"
    s = s & _
        " WHERE" & _
            " (" & s_aux & ")" & _
            " AND (anulado_status=0)" & _
            " AND (estoque <> '" & ID_ESTOQUE_DEVOLUCAO & "')" & _
            " AND (preco_NF > 0)"
    s = s & _
        " GROUP BY" & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.descricao," & _
            " t_PEDIDO_ITEM.ean," & _
            " t_PEDIDO_ITEM.preco_NF," & _
            " t_PEDIDO_ITEM.qtde_volumes," & _
            " t_PEDIDO_ITEM.peso," & _
            " t_PEDIDO_ITEM.cubagem," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst," & _
            " t_PRODUTO.perc_MVA_ST," & _
            " t_PRODUTO.ean," & _
            " t_PRODUTO.peso," & _
            " t_PRODUTO.cubagem"
    s = s & _
        " ORDER BY" & _
            " t_PEDIDO_ITEM.produto," & _
            " t_ESTOQUE_ITEM.ncm," & _
            " t_ESTOQUE_ITEM.cst"
    
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    Do While Not t_PEDIDO_ITEM.EOF
        blnAchou = False
        For i = LBound(v_nf) To UBound(v_nf)
            With v_nf(i)
                If (.fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And _
                   (.produto = Trim$("" & t_PEDIDO_ITEM("produto"))) And _
                   (.ncm = Trim$("" & t_PEDIDO_ITEM("ncm"))) And _
                   (.cst = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PEDIDO_ITEM("cst")))) Then
                    blnAchou = True
                    idx = i
                    Exit For
                    End If
                End With
            Next
            
        If Not blnAchou Then
            qtde_linhas_nf = qtde_linhas_nf + 1
            If v_nf(UBound(v_nf)).produto <> "" Then
                ReDim Preserve v_nf(UBound(v_nf) + 1)
                limpa_item_TIPO_LINHA_NOTA_FISCAL v_nf(UBound(v_nf))
                End If
            idx = UBound(v_nf)
            With v_nf(UBound(v_nf))
                .fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))
                .produto = Trim$("" & t_PEDIDO_ITEM("produto"))
                .descricao = Trim$("" & t_PEDIDO_ITEM("descricao"))
                .EAN = Trim("" & t_PEDIDO_ITEM("ean"))
                .ncm = Trim("" & t_PEDIDO_ITEM("ncm"))
                .NCM_bd = Trim("" & t_PEDIDO_ITEM("ncm"))
                .cst = cst_converte_codigo_entrada_para_saida(Trim("" & t_PEDIDO_ITEM("cst")))
                .CST_bd = cst_converte_codigo_entrada_para_saida(Trim("" & t_PEDIDO_ITEM("cst")))
                End With
            End If
            
        With v_nf(idx)
        '   QUANTIDADE
            qtde = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde")) Then qtde = CLng(t_PEDIDO_ITEM("qtde"))
            .qtde_total = .qtde_total + qtde
        
        '   VALOR
            vl_unitario = 0
            If IsNumeric(t_PEDIDO_ITEM("preco_NF")) Then vl_unitario = t_PEDIDO_ITEM("preco_NF")
            .valor_total = .valor_total + (qtde * vl_unitario)
        
        '   QTDE DE VOLUMES
            n = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde_volumes")) Then n = CLng(t_PEDIDO_ITEM("qtde_volumes"))
            .qtde_volumes_total = .qtde_volumes_total + (qtde * n)
        
        '   PESO
            peso_aux = 0
            If IsNumeric(t_PEDIDO_ITEM("peso")) Then peso_aux = CSng(t_PEDIDO_ITEM("peso"))
            .peso_total = .peso_total + (qtde * peso_aux)
            
        '   CUBAGEM
            cubagem_aux = 0
            If IsNumeric(t_PEDIDO_ITEM("cubagem")) Then cubagem_aux = CSng(t_PEDIDO_ITEM("cubagem"))
            .cubagem_total = .cubagem_total + (qtde * cubagem_aux)
            
        '   PERCENTUAL DE MVA ST
            .perc_MVA_ST = t_PEDIDO_ITEM("perc_MVA_ST")
            
        '   EAN (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If Trim("" & t_PEDIDO_ITEM("ean")) = "" Then .EAN = Trim("" & t_PEDIDO_ITEM("tP_ean"))
        
        '   PESO (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If t_PEDIDO_ITEM("peso") = 0 Then
                peso_aux = 0
                If IsNumeric(t_PEDIDO_ITEM("tP_peso")) Then peso_aux = CSng(t_PEDIDO_ITEM("tP_peso"))
                .peso_total = .peso_total + (qtde * peso_aux)
                End If
            
        '   CUBAGEM (SE NÃO HÁ A INFORMAÇÃO ARMAZENADA JUNTO C/ O PEDIDO, USA O QUE ESTÁ GRAVADO NA TABELA DE PRODUTOS)
            If t_PEDIDO_ITEM("cubagem") = 0 Then
                cubagem_aux = 0
                If IsNumeric(t_PEDIDO_ITEM("tP_cubagem")) Then cubagem_aux = CSng(t_PEDIDO_ITEM("tP_cubagem"))
                .cubagem_total = .cubagem_total + (qtde * cubagem_aux)
                End If
            End With
            
        t_PEDIDO_ITEM.MoveNext
        Loop


'   FAZ A CONFERÊNCIA DA QUANTIDADE (APENAS P/ SE CERTIFICAR QUE A LÓGICA ESTÁ CORRETA)
    s_msg = ""
    For i = LBound(v_nf_confere) To UBound(v_nf_confere)
        If Trim$(v_nf_confere(i).produto) <> "" Then
            n = 0
            For j = LBound(v_nf) To UBound(v_nf)
                If (Trim$(v_nf_confere(i).fabricante) = Trim$(v_nf(j).fabricante)) And _
                    (Trim$(v_nf_confere(i).produto) = Trim$(v_nf(j).produto)) Then
                    n = n + v_nf(j).qtde_total
                    End If
                Next
            If CLng(v_nf_confere(i).qtde_total) <> CLng(n) Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "Houve divergência na quantidade do produto (" & v_nf_confere(i).fabricante & ")" & v_nf_confere(i).produto & ": quantidade esperada=" & CStr(v_nf_confere(i).qtde_total) & ", quantidade calculada=" & CStr(n)
                End If
            End If
        Next
    
    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
'   DADOS DA TELA: INFORMAÇÕES ADICIONAIS DO PRODUTO, CST, NCM, CFOP E ICMS
'   IMPORTANTE: O MESMO CÓDIGO DE PRODUTO PODE APARECER EM MAIS DE UMA LINHA DEVIDO AO
'   =========== CONSUMO DE DIFERENTES LOTES DO ESTOQUE QUE TENHAM DADO ENTRADA C/ CÓDIGOS
'               DIFERENTES DE NCM E/OU CST. PORTANTO, DEVE SER FEITO UM CONTROLE P/ OBTER
'               OS DADOS DA TELA EDITADOS DA OCORRÊNCIA CORRETA.
    ReDim v_flagDadosTelaJaLido(c_produto.LBound To c_produto.UBound)
    For i = LBound(v_flagDadosTelaJaLido) To UBound(v_flagDadosTelaJaLido)
        v_flagDadosTelaJaLido(i) = False
        Next
    
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            For j = c_produto.LBound To c_produto.UBound
                If Trim$(v_nf(i).fabricante) = Trim$(c_fabricante(j)) And _
                   Trim$(v_nf(i).produto) = Trim$(c_produto(j)) And _
                   Trim$(v_nf(i).ncm) = Trim$(c_NCM(j)) Then
                    If Not v_flagDadosTelaJaLido(j) Then
                        v_flagDadosTelaJaLido(j) = True
                        v_nf(i).vl_outras_despesas_acessorias = converte_para_currency(Trim$(c_vl_outras_despesas_acessorias(j)))
                        v_nf(i).infAdProd = Trim$(c_produto_obs(j))
                        v_nf(i).xPed = Trim$(c_xPed(j))
                        v_nf(i).nItemPed = Trim$(c_nItemPed(j))
                        v_nf(i).fcp = Trim$(c_fcp(j))
                        v_nf(i).CST_tela = Trim$(c_CST(j))
                        v_nf(i).NCM_tela = Trim$(c_NCM(j))
                        If cb_CFOP(j).ListIndex <> -1 Then
                            If Trim$(cb_CFOP(j)) <> "" Then
                                s = Trim$(cb_CFOP(j))
                                For k = 1 To Len(s)
                                    c = Mid$(s, k, 1)
                                    If c = " " Then Exit For
                                    v_nf(i).CFOP_tela_formatado = v_nf(i).CFOP_tela_formatado & c
                                    Next
                                v_nf(i).CFOP_tela = retorna_so_digitos(v_nf(i).CFOP_tela_formatado)
                                End If
                            End If
                        If Trim$(cb_ICMS_item(j)) <> "" Then
                            v_nf(i).ICMS_tela = Trim$(cb_ICMS_item(j))
                            End If
                        Exit For
                        End If
                    End If
                Next
            End If
        Next
    

'   CST => VERIFICA SE HOUVE ALTERAÇÃO NO CST DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).CST_tela) <> "" Then
                If Trim$(v_nf(i).CST_bd) <> Trim$(v_nf(i).CST_tela) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CST alterado de " & v_nf(i).CST_bd & " para " & v_nf(i).CST_tela
                    End If
                End If
            End If
        Next
    
    If s_msg <> "" Then
        s_msg = "Houve alteração no CST do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If

'   PREPARA O CAMPO QUE ARMAZENA O CST A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).cst = v_nf(i).CST_bd
            If Trim$(v_nf(i).CST_tela) <> "" Then v_nf(i).cst = Trim$(v_nf(i).CST_tela)
            End If
        Next
    
'   NCM => VERIFICA SE HOUVE ALTERAÇÃO NO NCM DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).NCM_tela) <> "" Then
                If Trim$(v_nf(i).NCM_bd) <> Trim$(v_nf(i).NCM_tela) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": NCM alterado de " & v_nf(i).NCM_bd & " para " & v_nf(i).NCM_tela
                    End If
                End If
            End If
        Next
    
    If s_msg <> "" Then
        s_msg = "Houve alteração no NCM do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If

'   PREPARA O CAMPO QUE ARMAZENA O NCM A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).ncm = v_nf(i).NCM_bd
            If Trim$(v_nf(i).NCM_tela) <> "" Then v_nf(i).ncm = Trim$(v_nf(i).NCM_tela)
            End If
        Next
    
'   CFOP => VERIFICA SE HOUVE ALTERAÇÃO NO CFOP DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).CFOP_tela) <> "" Then
                If Trim$(v_nf(i).CFOP_tela) <> strCfopCodigo Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CFOP alterado para " & v_nf(i).CFOP_tela_formatado
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "Houve alteração no CFOP do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   PREPARA O CAMPO QUE ARMAZENA O CFOP A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).cfop = strCfopCodigo
            v_nf(i).CFOP_formatado = strCfopCodigoFormatado
            If Trim$(v_nf(i).CFOP_tela) <> "" Then
                v_nf(i).cfop = Trim$(v_nf(i).CFOP_tela)
                v_nf(i).CFOP_formatado = Trim$(v_nf(i).CFOP_tela_formatado)
                End If
            End If
        Next

'   VERIFICA SE O CFOP A SER USADO É CONFLITANTE COM O LOCAL DE DESTINO DA OPERAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).cfop) <> "" Then
                If existe_divergencia_loc_dest_x_cpof(v_nf(i).cfop, rNFeImg.ide__idDest) Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": CFOP " & v_nf(i).cfop
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "O local de destino da operação é conflitante com o CFOP do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   ICMS => VERIFICA SE HOUVE ALTERAÇÃO NO ICMS DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).ICMS_tela) <> "" Then
                If Trim$(v_nf(i).ICMS_tela) <> Trim$(cb_icms_remessa) Then
                    If is_venda_interestadual_de_mercadoria_importada(v_nf(i).cfop, v_nf(i).cst) And _
                        (Trim$(v_nf(i).ICMS_tela) = CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA)) Then
                    '   NOP: EM VENDA INTERESTADUAL DE MERCADORIA IMPORTADA É OBRIGATÓRIO USAR A ALÍQUOTA DE ICMS ESPECÍFICA
                    Else
                        If s_msg <> "" Then s_msg = s_msg & vbCrLf
                        s_msg = s_msg & "Produto " & v_nf(i).produto & " - " & v_nf(i).descricao & ": ICMS alterado para " & v_nf(i).ICMS_tela & "%"
                        End If
                    End If
                End If
            End If
        Next

    If s_msg <> "" Then
        s_msg = "Houve alteração no ICMS do(s) seguinte(s) produto(s):" & _
                vbCrLf & _
                s_msg & _
                vbCrLf & _
                vbCrLf & _
                "Continua mesmo assim?"
        If Not confirma(s_msg) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If

'   PREPARA O CAMPO QUE ARMAZENA O ICMS A SER USADO
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).ICMS = cb_icms_remessa
            If Trim$(v_nf(i).ICMS_tela) <> "" Then
                v_nf(i).ICMS = Trim$(v_nf(i).ICMS_tela)
                End If
            End If
        Next


'   QUANTIDADE DE LINHAS EXCEDE O TAMANHO DA PÁGINA?
    MAX_LINHAS_NOTA_FISCAL = MAX_LINHAS_NOTA_FISCAL_DEFAULT
    If (Not blnTemPagtoPorBoleto) Then MAX_LINHAS_NOTA_FISCAL = MAX_LINHAS_NOTA_FISCAL_DEFAULT + 2
    
    If qtde_linhas_nf > MAX_LINHAS_NOTA_FISCAL Then
        s = "Não é possível imprimir a nota fiscal porque os " & CStr(qtde_linhas_nf) & _
            " itens excedem o máximo de " & CStr(MAX_LINHAS_NOTA_FISCAL) & _
            " linhas que podem ser impressas!!"
        aviso_erro s
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
'>  ARREDONDAMENTOS
    For ic = LBound(v_nf) To UBound(v_nf)
        With v_nf(ic)
            If Trim$(.produto) <> "" Then
                vl_unitario = .valor_total / .qtde_total
                .valor_total = CCur(Format$(vl_unitario, FORMATO_MOEDA)) * .qtde_total
                End If
            End With
        Next

        
'   CONSISTE DADOS
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).ncm) = "" Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO possui o código NCM!!"
            ElseIf Len(Trim$(v_nf(i).cst)) = 0 Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO possui a informação do CST!!"
            ElseIf Len(Trim$(v_nf(i).cst)) <> 3 Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " possui o campo CST preenchido com valor inválido!!"
                End If
            End If
        Next

    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
        
'   SE FOR NOTA DE ENTRADA, VERIFICA SE A DEVOLUÇÃO DE MERCADORIAS FOI INTEGRAL
'   0-Entrada  1-Saída
    s_msg = ""
    If rNFeImg.ide__tpNF = "0" Then
        For i = LBound(v_nf) To UBound(v_nf)
            If Trim$(v_nf(i).produto) <> "" Then
                s = "SELECT" & _
                        " Coalesce(Sum(qtde),0) AS qtde_total_devolvida" & _
                    " FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
                    " WHERE" & _
                        " (" & sql_monta_criterio_texto_or(v_pedido(), "pedido", True) & ")" & _
                        " AND (fabricante = '" & v_nf(i).fabricante & "')" & _
                        " AND (produto = '" & v_nf(i).produto & "')"
                If t_PEDIDO_ITEM_DEVOLVIDO.State <> adStateClosed Then t_PEDIDO_ITEM_DEVOLVIDO.Close
                t_PEDIDO_ITEM_DEVOLVIDO.Open s, dbc, , , adCmdText
                If t_PEDIDO_ITEM_DEVOLVIDO.EOF Then
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO teve nenhuma unidade devolvida de um total de " & CStr(v_nf(i).qtde_total)
                Else
                    If CLng(t_PEDIDO_ITEM_DEVOLVIDO("qtde_total_devolvida")) <> v_nf(i).qtde_total Then
                        If s_msg <> "" Then s_msg = s_msg & vbCrLf
                        s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " teve " & Trim$("" & t_PEDIDO_ITEM_DEVOLVIDO("qtde_total_devolvida")) & " unidade(s) devolvida(s) de um total de " & CStr(v_nf(i).qtde_total)
                        End If
                    End If
                End If
            Next
        
        If s_msg <> "" Then
            s_msg = "Não é possível emitir esta NFe de entrada através do painel de emissão automática porque o pedido não teve os produtos devolvidos integralmente:" & _
                    vbCrLf & _
                    s_msg
            End If
        End If
    
    If s_msg <> "" Then
        aviso_erro s_msg
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    

'   OBTÉM DADOS DA TRANSPORTADORA
    strTransportadoraCnpj = ""
    strTransportadoraRazaoSocial = ""
    strTransportadoraIE = ""
    strTransportadoraUF = ""
    strTransportadoraEmail = ""
    If strTransportadoraId <> "" Then
        s = "SELECT * FROM t_TRANSPORTADORA WHERE id = '" & strTransportadoraId & "'"
        t_TRANSPORTADORA.Open s, dbc, , , adCmdText
        If t_TRANSPORTADORA.EOF Then
            s = "Transportadora '" & strTransportadoraId & "' não está cadastrada!!"
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        Else
            strTransportadoraCnpj = retorna_so_digitos(Trim$("" & t_TRANSPORTADORA("cnpj")))
            strTransportadoraRazaoSocial = UCase$(Trim$("" & t_TRANSPORTADORA("razao_social")))
            strTransportadoraIE = Trim$("" & t_TRANSPORTADORA("ie"))
            strTransportadoraUF = Trim$("" & t_TRANSPORTADORA("uf"))
            strTransportadoraEmail = Trim$("" & t_TRANSPORTADORA("email"))
            End If
        
        If (strTransportadoraCnpj = "") Or (strTransportadoraRazaoSocial = "") Then
            s = ""
            If strTransportadoraCnpj = "" Then
                If s <> "" Then s = s & vbCrLf
                s = s & "A transportadora '" & strTransportadoraId & "' não possui CNPJ cadastrado!!"
                End If
                
            If strTransportadoraRazaoSocial = "" Then
                If s <> "" Then s = s & vbCrLf
                s = s & "A transportadora '" & strTransportadoraId & "' não possui razão social cadastrada!!"
                End If
            
            If s <> "" Then
                s = s & vbCrLf & "Continua mesmo assim?"
                End If
            
            If Not confirma(s) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
            
''   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
'    s = "SELECT * FROM t_CLIENTE WHERE (id='" & Trim$("" & t_PEDIDO("id_cliente")) & "')"
'    t_DESTINATARIO.Open s, dbc, , , adCmdText
'    If t_DESTINATARIO.EOF Then
'        s = "Cliente com nº registro " & Trim$("" & t_PEDIDO("id_cliente")) & " não foi encontrado!!"
'        aviso_erro s
'        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
'        aguarde INFO_NORMAL, m_id
'        Exit Sub
'        End If
        
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    'PRIMEIRO CASO: A MEMORIZAÇÃO DO ENDEREÇO DO CLIENTE NA TABELA DE PEDIDOS ESTÁ OK
    blnExisteMemorizacaoEndereco = False
    If param_pedidomemorizacaoenderecos.campo_inteiro = 1 Then
        s = "SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, endereco_uf as uf, endereco_cnpj_cpf as cnpj_cpf, " & _
                " endereco_logradouro as endereco, endereco_bairro as bairro, endereco_cidade as cidade, endereco_cep as cep, endereco_numero, endereco_complemento, " & _
                " endereco_logradouro as endereco_end_nota, " & _
                " endereco_bairro as bairro_end_nota, " & _
                " endereco_cidade as cidade_end_nota, " & _
                " endereco_cep as cep_end_nota, " & _
                " endereco_numero as numero_end_nota, " & _
                " endereco_complemento as complemento_end_nota, " & _
                " endereco_uf as uf_end_nota, " & _
                " endereco_email as email, endereco_email_xml as email_xml, " & _
                " endereco_nome as nome, " & _
                " endereco_ddd_res as ddd_res, endereco_tel_res as tel_res, " & _
                " endereco_ddd_com as ddd_com, endereco_tel_com as tel_com, endereco_ramal_com as ramal_com, " & _
                " endereco_ddd_cel as ddd_cel, endereco_tel_cel as tel_cel, " & _
                " endereco_ddd_com_2 as ddd_com_2, endereco_tel_com_2 as tel_com_2, endereco_ramal_com_2 as ramal_com_2, " & _
                " endereco_tipo_pessoa as tipo, " & _
                " endereco_contribuinte_icms_status as contribuinte_icms_status, " & _
                " endereco_produtor_rural_status as produtor_rural_status, " & _
                " endereco_ie as ie, " & _
                " endereco_rg as rg, " & _
                " endereco_contato as contato " & _
            " FROM t_PEDIDO" & _
            " WHERE (pedido = '" & Trim$("" & t_PEDIDO("pedido")) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PJ & "')"
        s = s & " UNION" & _
            " SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, endereco_uf as uf, endereco_cnpj_cpf as cnpj_cpf, " & _
                " endereco_logradouro as endereco, endereco_bairro as bairro, endereco_cidade as cidade, endereco_cep as cep, endereco_numero, endereco_complemento, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_logradouro else EndEtg_endereco end as endereco_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_bairro else EndEtg_bairro end as bairro_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cidade else EndEtg_cidade end as cidade_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cep else EndEtg_cep end as cep_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_numero else EndEtg_endereco_numero end as numero_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_complemento else EndEtg_endereco_complemento end as complemento_end_nota, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_uf else EndEtg_uf end as uf_end_nota, " & _
                " endereco_email as email, endereco_email_xml as email_xml, " & _
                " endereco_nome as nome, " & _
                " endereco_ddd_res as ddd_res, endereco_tel_res as tel_res, " & _
                " endereco_ddd_com as ddd_com, endereco_tel_com as tel_com, endereco_ramal_com as ramal_com, " & _
                " endereco_ddd_cel as ddd_cel, endereco_tel_cel as tel_cel, " & _
                " endereco_ddd_com_2 as ddd_com_2, endereco_tel_com_2 as tel_com_2, endereco_ramal_com_2 as ramal_com_2, " & _
                " endereco_tipo_pessoa as tipo, " & _
                " endereco_contribuinte_icms_status as contribuinte_icms_status, " & _
                " endereco_produtor_rural_status as produtor_rural_status, " & _
                " endereco_ie as ie, " & _
                " endereco_rg as rg, " & _
                " endereco_contato as contato " & _
            " FROM t_PEDIDO" & _
            " WHERE (pedido = '" & Trim$("" & t_PEDIDO("pedido")) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PF & "')"
        t_DESTINATARIO.Open s, dbc, , , adCmdText
        If t_DESTINATARIO.EOF Then
            s = "Problemas na localização do endereço memorizado no pedido " & Trim$("" & t_PEDIDO("pedido")) & "!!"
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        If t_DESTINATARIO("st_memorizacao_completa_enderecos") > 0 Then blnExisteMemorizacaoEndereco = True
        End If
        
    'SEGUNDO CASO: A MEMORIZAÇÃO DO ENDEREÇO DO CLIENTE NA TABELA DE PEDIDOS NÃO ESTÁ OK
    If Not blnExisteMemorizacaoEndereco Then
        If t_DESTINATARIO.State <> adStateClosed Then t_DESTINATARIO.Close
    '   (se não houver memorização no pedido)
        s = "SELECT * FROM t_CLIENTE WHERE (id='" & Trim$("" & t_PEDIDO("id_cliente")) & "')"
        t_DESTINATARIO.Open s, dbc, , , adCmdText
        If t_DESTINATARIO.EOF Then
            s = "Cliente com nº registro " & Trim$("" & t_PEDIDO("id_cliente")) & " não foi encontrado!!"
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
        
        
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
    
'   CONFIRMA ALÍQUOTA DO ICMS
'   ALERTAR QUE A ALÍQUOTA DEVERIA SER ZERO PARA A NOTA DE REMESSA
'    If obtem_aliquota_ICMS(usuario.emit_uf, UCase$(Trim$("" & t_DESTINATARIO("uf"))), aliquota_icms_interestadual) Then
'        strIcms = Trim$(CStr(aliquota_icms_interestadual))
'    Else
'        strIcms = ""
'        End If
'
'    If (strIcms <> "") And (cb_icms <> "") Then
'        If (CSng(strIcms) <> CSng(cb_icms)) Then
'            s = "O destinatário é do estado de " & UCase$(Trim$("" & t_DESTINATARIO("uf"))) & " cuja alíquota de ICMS é de " & strIcms & "%" & _
'                vbCrLf & "Confirma a emissão da NFe usando a alíquota de " & cb_icms & "%?"
'            If Not confirma(s) Then
'                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
'                aguarde INFO_NORMAL, m_id
'                Exit Sub
'                End If
'            iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
'            End If
'        End If
    If (CSng(cb_icms_remessa) <> 0) Then
        s = "Notas fiscais de remessa geralmente são emitidas com alíquota de ICMS igual a 0%" & _
            vbCrLf & "Confirma a emissão da NFe usando a alíquota de " & cb_icms_remessa & "%?"
        If Not confirma(s) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
        End If
        
        
'   MERCADORIA IMPORTADA EM VENDA INTERESTADUAL: SE A ALÍQUOTA FOR DIFERENTE DE ZERO, VERIFICA SE ESTÁ C/ ALÍQUOTA DE ICMS ESPECÍFICA
'   NÃO EXIBIR ALERTA P/ PESSOA FÍSICA (EXCETO PRODUTOR RURAL CONTRIBUINTE DO ICMS) OU SE FOR PJ ISENTA DE I.E.
    If (CSng(cb_icms_remessa) <> 0) Then
        If ((Len(retorna_so_digitos(Trim("" & t_DESTINATARIO("cnpj_cpf")))) = 14) And _
            (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) Or _
           ((Len(retorna_so_digitos(Trim("" & t_DESTINATARIO("cnpj_cpf")))) = 14) And _
           (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL) And _
            (InStr(UCase$(Trim$("" & t_DESTINATARIO("ie"))), "ISEN") = 0)) Or _
           ((t_DESTINATARIO("produtor_rural_status") = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) And _
            (t_DESTINATARIO("contribuinte_icms_status") = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) Then
            s_confirma = ""
            For i = LBound(v_nf) To UBound(v_nf)
                If Trim$(v_nf(i).produto) <> "" Then
                    If is_venda_interestadual_de_mercadoria_importada(v_nf(i).cfop, v_nf(i).cst) Then
                        If Trim$(v_nf(i).ICMS) <> CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA) Then
                            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                            s_confirma = s_confirma & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " está com ICMS de " & v_nf(i).ICMS & "% ao invés de " & CStr(PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA) & "%"
                            End If
                        End If
                    End If
                Next
            
            If s_confirma <> "" Then
                s_confirma = "Foram encontradas possíveis incoerências na alíquota do ICMS na venda interestadual de mercadoria importada:" & _
                        vbCrLf & _
                        s_confirma & _
                        vbCrLf & vbCrLf & _
                        "Continua mesmo assim?"
                If Not confirma(s_confirma) Then
                    GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                    aguarde INFO_NORMAL, m_id
                    Exit Sub
                    End If
                iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
                End If
            End If
        End If
    
    
'   IGNORAR CÁLCULO DE BOLETOS PARA A NOTA DE REMESSA
''   SE HÁ PEDIDO ESPECIFICANDO PAGAMENTO VIA BOLETO BANCÁRIO, CALCULA QUANTIDADE DE PARCELAS, DATAS E VALORES
''   DOS BOLETOS. ESSES DADOS SERÃO IMPRESSOS NA NF E TAMBÉM SALVOS NO BD, POIS SERVIRÃO DE BASE PARA A GERAÇÃO
''   DOS BOLETOS NO ARQUIVO DE REMESSA.
'    If (param_geracaoboletos.campo_texto = "Manual") And blnExisteParcelamentoBoleto Then
'        ReDim v_parcela_pagto(UBound(v_parcela_manual_boleto))
'        v_parcela_pagto = v_parcela_manual_boleto
'    Else
'        ReDim v_parcela_pagto(0)
'        If Not geraDadosParcelasPagto(v_pedido(), v_parcela_pagto(), s_erro) Then
'            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
'            aguarde INFO_NORMAL, m_id
'            If s_erro <> "" Then s_erro = Chr(13) & Chr(13) & s_erro
'            s_erro = "Falha ao tentar processar os dados de pagamento!!" & s_erro
'            aviso_erro s_erro
'            Exit Sub
'            End If
'        End If
'
''   Tipo de NFe: 0-Entrada  1-Saída
'    If rNFeImg.ide__tpNF = "1" Then
'        s = ""
'        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'            If v_parcela_pagto(i).intNumDestaParcela <> 0 Then
'                blnImprimeDadosFatura = True
'                If s <> "" Then s = s & Chr(13)
'                s = s & "Parcela:  " & v_parcela_pagto(i).intNumDestaParcela & "/" & v_parcela_pagto(i).intNumTotalParcelas & " para " & Format$(v_parcela_pagto(i).dtVencto, FORMATO_DATA) & " de " & SIMBOLO_MONETARIO & " " & Format$(v_parcela_pagto(i).vlValor, FORMATO_MOEDA) & " (" & descricao_opcao_forma_pagamento(v_parcela_pagto(i).id_forma_pagto) & ")"
'                End If
'            Next
'
'        If s <> "" Then
'            s = "Serão emitidas na NFe as seguintes informações de pagamento:" & Chr(13) & Chr(13) & s
'            If DESENVOLVIMENTO Then
'                aviso s
'                End If
'            End If
'        End If
    
' na emissão da nota de remessa, deve ser verificado o CFOP?
''   VERIFICA SE O CFOP ESTÁ COERENTE COM O CST DO ICMS
'    s_confirma = ""
'    For i = LBound(v_nf) To UBound(v_nf)
'        If Trim$(v_nf(i).produto) <> "" Then
'            strNFeCst = Trim$(right$(v_nf(i).cst, 2))
'            strCfopCodigoAux = Trim$(v_nf(i).cfop)
'            strCfopCodigoFormatadoAux = Trim$(v_nf(i).CFOP_formatado)
'            s = "O produto " & v_nf(i).produto & " possui CST = " & strNFeCst & ", mas o CFOP selecionado é " & strCfopCodigoFormatadoAux
'            If strNFeCst = "00" Then
'                If (strCfopCodigoAux = "5102") Or (strCfopCodigoAux = "6102") Then s = ""
'            ElseIf strNFeCst = "60" Then
'                If (strCfopCodigoAux = "5405") Or (strCfopCodigoAux = "6404") Then s = ""
'            Else
'                If (strCfopCodigoAux <> "5102") And (strCfopCodigoAux <> "6102") And _
'                   (strCfopCodigoAux <> "5405") And (strCfopCodigoAux <> "6404") Then s = ""
'                End If
'
'            If s <> "" Then
'                If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
'                s_confirma = s_confirma & s
'                End If
'            End If
'        Next
'
'    If s_confirma <> "" Then
'        s_confirma = "Foram encontradas possíveis incoerências entre CFOP e CST:" & _
'                     vbCrLf & _
'                     s_confirma & _
'                     vbCrLf & vbCrLf & _
'                     "Continua mesmo assim?"
'        If Not confirma(s_confirma) Then
'            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
'            aguarde INFO_NORMAL, m_id
'            Exit Sub
'            End If
'        End If


'   ZERAR PIS/COFINS?
    s_confirma = ""
    If Trim$(cb_zerar_PIS) <> "" Then
        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
        s_confirma = s_confirma & "Alíquota do PIS será zerada usando CST = " & cb_zerar_PIS
        End If
    
    If Trim$(cb_zerar_COFINS) <> "" Then
        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
        s_confirma = s_confirma & "Alíquota do COFINS será zerada usando CST = " & cb_zerar_COFINS
        End If
    
    If s_confirma <> "" Then
        s_confirma = s_confirma & _
                     vbCrLf & vbCrLf & _
                     "Continua mesmo assim?"
        If Not confirma(s_confirma) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    

'   CALCULA TOTAL ESTIMADO DOS TRIBUTOS USANDO DADOS DO IBPT?
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    s_confirma = ""
    If is_venda_consumidor_final(strCfopCodigo) Then
        blnExibirTotalTributos = True
    '   OBTÉM DADOS DO IBPT P/ CALCULAR TOTAL ESTIMADO DOS TRIBUTOS
        For i = LBound(v_nf) To UBound(v_nf)
            With v_nf(i)
                If Trim$(.produto) <> "" Then
                    s = "SELECT " & _
                            "*" & _
                        " FROM t_IBPT" & _
                        " WHERE" & _
                            " (codigo = '" & Trim$(.ncm) & "')" & _
                            " AND (tabela = '0')" & _
                        " ORDER BY" & _
                            " codigo," & _
                            " ex"
                    If t_IBPT.State <> adStateClosed Then t_IBPT.Close
                    t_IBPT.Open s, dbc, , , adCmdText
                    If t_IBPT.EOF Then
                        blnHaProdutoSemDadosIbpt = True
                        If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                        s_confirma = s_confirma & "O NCM '" & Trim$(.ncm) & "' NÃO está cadastrado na tabela do IBPT!!"
                    Else
                        .tem_dados_IBPT = True
                        .percAliqNac = t_IBPT("percAliqNac")
                        .percAliqImp = t_IBPT("percAliqImp")
                        End If
                    End If
                End With
            Next
        
        If s_confirma <> "" Then
            s_confirma = s_confirma & _
                         "A nota fiscal será emitida sem a informação do total estimado dos tributos conforme exige a lei 12.741/2012!!" & _
                         vbCrLf & vbCrLf & _
                         "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
'   VERIFICAR DIVERGÊNCIA DE LOCAL DE DESTINO DA OPERAÇÃO
    If rNFeImg.ide__tpNF <> "0" Then
        s_confirma = ""
        strDestinoUF = l_end_recebedor_uf
        'primeira situação: UFs diferentes e Local de Destino  <> Interestadual
        If (Trim$(rNFeImg.ide__idDest) <> "2") And (strOrigemUF <> strDestinoUF) Then
            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
            s_confirma = s_confirma & "UF de origem e destino da Nota são diferentes, porém local de operação selecionado é " & vbCrLf & vbCrLf
            s_confirma = s_confirma & cb_loc_dest_remessa
            End If
        
        If (Trim$(rNFeImg.ide__idDest) <> "1") And (strOrigemUF = strDestinoUF) Then
            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
            s_confirma = s_confirma & "UF de origem e destino da Nota são iguais, porém local de operação selecionado é " & vbCrLf & vbCrLf
            s_confirma = s_confirma & cb_loc_dest_remessa
            End If
        
        If s_confirma <> "" Then
            s_confirma = s_confirma & _
                         vbCrLf & vbCrLf & _
                         "Continua mesmo assim?"
            If Not confirma(s_confirma) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            iQtdConfirmaDuvidaEmit = iQtdConfirmaDuvidaEmit + 1
            End If
        End If
    

'   PREPARA DADOS DA NFe
    aguarde INFO_EXECUTANDO, "preparando emissão da NFe"
    
'   TAG OPERACIONAL
'   ~~~~~~~~~~~~~~~
    strNFeTagOperacional = "operacional;" & vbCrLf

'   EMAIL DO DESTINATÁRIO DA NFe
    rNFeImg.operacional__email = Trim("" & t_DESTINATARIO("email"))
    If (Trim$(rNFeImg.operacional__email) <> "") And (Trim$(strTransportadoraEmail) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
    rNFeImg.operacional__email = rNFeImg.operacional__email & strTransportadoraEmail
    strEmailXML = Trim("" & t_DESTINATARIO("email_xml"))
    If Trim$(strEmailXML) <> "" Then
        If (Trim$(rNFeImg.operacional__email) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
        rNFeImg.operacional__email = rNFeImg.operacional__email & strEmailXML
        End If

    If rNFeImg.operacional__email <> "" Then
        strNFeTagOperacional = strNFeTagOperacional & _
                               vbTab & NFeFormataCampo("email", rNFeImg.operacional__email)
        End If
    
'   TAG DEST (DADOS DO DESTINATÁRIO)
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNFeTagDestinatario = "dest;" & vbCrLf
    
'   CNPJ/CPF
    strDestinatarioCnpjCpf = retorna_so_digitos(Trim(c_cnpj_cpf_dest))
    If strDestinatarioCnpjCpf = "" Then
        s_erro = "CNPJ/CPF do recebedor não está preenchido na tela!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Not cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
        s_erro = "CNPJ/CPF do recebedor está preenchido com informação inválida!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    
    If Len(strDestinatarioCnpjCpf) = 11 Then
        blnIsDestinatarioPJ = False
        rNFeImg.dest__CPF = strDestinatarioCnpjCpf
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CPF", rNFeImg.dest__CPF)
    ElseIf Len(strDestinatarioCnpjCpf) = 14 Then
        blnIsDestinatarioPJ = True
        rNFeImg.dest__CNPJ = strDestinatarioCnpjCpf
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CNPJ", rNFeImg.dest__CNPJ)
        End If
        
'   CAMPO: idEstrangeiro
    rNFeImg.dest__idEstrangeiro = ""
    If Trim(rNFeImg.dest__idEstrangeiro) <> "" Then
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("idEstrangeiro", rNFeImg.dest__idEstrangeiro)
        End If
    
'   NOME
    If NFE_AMBIENTE = NFE_AMBIENTE_HOMOLOGACAO Then
        strCampo = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
    Else
        strCampo = Trim(c_nome_dest)
        End If
    If strCampo = "" Then
        s_erro = "O nome do recebedor não está preenchido na tela!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O nome do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xNome = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xNome", rNFeImg.dest__xNome)
    
'   LOGRADOURO
    strCampo = l_end_recebedor_logradouro
    If strCampo = "" Then
        s_erro = "O endereço do recebedor não está preenchido na tela!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O endereço do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xLgr = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xLgr", rNFeImg.dest__xLgr)
    
'   ENDEREÇO: NÚMERO
    strCampo = l_end_recebedor_numero
    If strCampo = "" Then
        s_erro = "O endereço no cadastro do recebedor deve ser preenchido corretamente para poder emitir a NFe!!" & vbCrLf & _
                 "As informações de número e complemento do endereço devem ser preenchidas nos campos adequados!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O número do endereço do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__nro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("nro", rNFeImg.dest__nro)
        
'   ENDEREÇO: COMPLEMENTO
    strCampo = l_end_recebedor_complemento
    If Len(strCampo) > 60 Then
        s_erro = "O campo complemento do endereço do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xCpl = strCampo
    If Len(strCampo) > 0 Then strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xCpl", rNFeImg.dest__xCpl)
    
'   BAIRRO
    strCampo = l_end_recebedor_bairro
    If strCampo = "" Then
        s_erro = "O campo bairro no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O campo bairro no endereço do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xBairro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xBairro", rNFeImg.dest__xBairro)
    
'   MUNICIPIO
    strCampo = l_end_recebedor_cidade
    s_aux = l_end_recebedor_uf
    If (strCampo <> "") And (s_aux <> "") Then strCampo = strCampo & "/"
    strCampo = strCampo & s_aux
    rNFeImg.dest__cMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cMun", rNFeImg.dest__cMun)
    
    strCampo = l_end_recebedor_cidade
    If Len(strCampo) > 60 Then
        s_erro = "O campo cidade no endereço do recebedor excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xMun", rNFeImg.dest__xMun)
    
'   UF
    strCampo = l_end_recebedor_uf
    If strCampo = "" Then
        s_erro = "O campo UF no endereço do recebedor não está preenchido no cadastro!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__UF = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("UF", rNFeImg.dest__UF)
    
'   MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
    If Not consiste_municipio_IBGE_ok(dbcNFe, rNFeImg.dest__xMun, rNFeImg.dest__UF, strListaSugeridaMunicipiosIBGE, s_erro_aux) Then
        If s_erro_aux <> "" Then
            s_erro = s_erro_aux
        Else
            s_erro = "Município '" & rNFeImg.dest__xMun & "' não consta na relação de municípios do IBGE para a UF de '" & rNFeImg.dest__UF & "'!!"
            End If
            
        If s_erro <> "" Then s_erro = s_erro & Chr(13)
        s_erro = s_erro & "Será necessário corrigir o município no cadastro do cliente antes de prosseguir!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If

'   CEP
    strCampo = retorna_so_digitos(l_end_recebedor_cep)
    If strCampo = "" Then
        s_erro = "O campo CEP no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__CEP = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("CEP", rNFeImg.dest__CEP)
    
'   PAÍS
    rNFeImg.dest__cPais = "1058"
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cPais", rNFeImg.dest__cPais)
    rNFeImg.dest__xPais = "BRASIL"
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xPais", rNFeImg.dest__xPais)
    
'   FONE
    strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_cel")))
    If strCampo <> "" Then
        If Len(strCampo) > 9 Then
            s_erro = "O telefone celular no cadastro do destinatário excede o tamanho máximo permitido!!"
            GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
            
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_cel")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        
        If strDDD = "" Then
            s_erro = "O DDD do telefone celular no cadastro do destinatário não está preenchido!!"
            GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        ElseIf Len(strDDD) > 2 Then
            s_erro = "O DDD do telefone celular no cadastro do destinatário excede o tamanho máximo permitido!!"
            GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
        strCampo = strDDD & strCampo
        strTelCel = strCampo
        End If
    
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O telefone residencial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
                
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do telefone residencial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone residencial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelRes = strCampo
            End If
        End If
        
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
        
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelCom = strCampo
            End If
        End If
        
    If strCampo = "" Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2")))
        If strCampo <> "" Then
            If Len(strCampo) > 9 Then
                s_erro = "O segundo telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
        
            strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
            If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
            
            If strDDD = "" Then
                s_erro = "O DDD do segundo telefone comercial no cadastro do destinatário não está preenchido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            ElseIf Len(strDDD) > 2 Then
                s_erro = "O DDD do telefone comercial no cadastro do destinatário excede o tamanho máximo permitido!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            strCampo = strDDD & strCampo
            strTelCom2 = strCampo
            End If
        End If
    If strCampo <> "" Then
        rNFeImg.dest__fone = strCampo
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("fone", rNFeImg.dest__fone)
        End If
        
    'preencher os campos de telefone que possam estar vazios
    If strTelRes = "" Then strTelRes = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_res"))))
    If strTelCom = "" Then strTelCom = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com"))))
    If strTelCom2 = "" Then strTelCom2 = telefone_formata_2(retorna_so_digitos(Trim$("" & t_DESTINATARIO("tel_com_2"))))
    If strTelRes <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_res")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        If (Len(strDDD) = 2) Then strTelRes = "(" & strDDD & ")" & strTelRes
        End If
    If strTelCom <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com")))
        If (Len(strDDD) = 2) Then strTelCom = "(" & strDDD & ") " & strTelCom
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom = strTelCom & " R: " & strRamal
        End If
    If strTelCom2 <> "" Then
        strDDD = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ddd_com_2")))
        If (Len(strDDD) = 3) And (left$(strDDD, 1) = "0") Then strDDD = right$(strDDD, 2)
        strRamal = retorna_so_digitos(Trim$("" & t_DESTINATARIO("ramal_com_2")))
        If (Len(strDDD) = 2) Then strTelCom2 = "(" & strDDD & ") " & strTelCom2
        If (strRamal <> "") And (converte_para_single(strRamal) <> 0) Then strTelCom2 = strTelCom2 & " R: " & strRamal
        End If
        
        
'   CAMPO: indIEDest

    'Conforme orientação da Bueno Consultoria e Assessoria Contábil, em e-mail encaminhado em 22/06/2016,
    'deve-se informar a identificação da IE do destinatário como "Contribuinte do ICMS" ou "Não Contribuinte"
    strCampo = Trim$(c_rg_dest)
    s_aux = Trim$(l_end_recebedor_uf)
    If blnIsDestinatarioPJ Then
        If InStr(UCase$(strCampo), "ISEN") > 0 Then
            strCampo = "ISENTO"
            End If
        If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
        'se estiver vazia, considerar não contribuinte
        If strCampo = "" Then
            rNFeImg.dest__indIEDest = "9"
            strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
        ElseIf (Len(strCampo) < 2) Or (Len(strCampo) > 14) Then
            s_erro = "A Inscrição Estadual no cadastro do cliente está preenchida com conteúdo inválido!!"
            GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
        ElseIf ConsisteInscricaoEstadual(strCampo, s_aux) <> 0 Then
        '   Retorno = 0 -> IE válida
        '   Retorno = 1 -> IE inválida
            s_erro = "A Inscrição Estadual no cadastro do cliente (" & strCampo & ") é inválida para a UF de '" & s_aux & "'!!"
            GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
            End If
        
        'se ainda não foi preenchido o valor do campo indIEDest, preencher
        If Trim(rNFeImg.dest__indIEDest) = "" Then
            If strCampo = "ISENTO" Then
            '   2 = CONTRIBUINTE ISENTO, PREENCHER COM NÃO CONTRIBUINTE
                rNFeImg.dest__indIEDest = "9"
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
            Else
            '   1 = CONTRIBUINTE ICMS
                rNFeImg.dest__indIEDest = "1"
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
                End If
            End If
            
    Else
    '   9 = NÃO-CONTRIBUINTE
        rNFeImg.dest__indIEDest = "9"
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("indIEDest", rNFeImg.dest__indIEDest)
        End If
        
'   IE
    If blnIsDestinatarioPJ Then
        strCampo = Trim$(c_rg_dest)
        If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
        If rNFeImg.dest__indIEDest = "1" Then
            'Primeira situação: o cliente é contribuinte do ICMS
            If InStr(UCase$(strCampo), "ISEN") > 0 Then
                s_erro = "Cliente está marcado como Contribuinte, porém Inscrição Estadual apresenta valor (" & strCampo & ")!!"
                GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                End If
            rNFeImg.dest__IE = strCampo
            strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("IE", rNFeImg.dest__IE)
        ElseIf rNFeImg.dest__indIEDest = "9" Then
            'Segunda situação: o cliente não é contribuinte do ICMS
            If InStr(UCase$(strCampo), "ISEN") > 0 Then strCampo = ""
            If strCampo <> "" Then
                rNFeImg.dest__IE = strCampo
                strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("IE", rNFeImg.dest__IE)
                End If
            'Terceira situação: o cliente é isento
            'Não enviar a inscrição estadual
            End If
        End If
            
            
'>  DADOS DA FATURA
''(já foram impressos na nota de venda)
'    If blnImprimeDadosFatura Then
'        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'            With v_parcela_pagto(i)
'                If .intNumDestaParcela <> 0 Then
'                    If Trim$(vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc) <> "" Then
'                        ReDim Preserve vNFeImgTagDup(UBound(vNFeImgTagDup) + 1)
'                        End If
'
'                '   FORMA DE PAGTO
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup = abreviacao_opcao_forma_pagamento(.id_forma_pagto)
'                    s = vbTab & NFeFormataCampo("nDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup)
'                '   VENCTO
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc = NFeFormataData(.dtVencto)
'                    s = s & vbTab & NFeFormataCampo("dVenc", vNFeImgTagDup(UBound(vNFeImgTagDup)).dVenc)
'                '   VALOR
'                    vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup = NFeFormataMoeda2Dec(.vlValor)
'                    s = s & vbTab & NFeFormataCampo("vDup", vNFeImgTagDup(UBound(vNFeImgTagDup)).vDup)
'                '   ADICIONA PARCELA À TAG
'                    strNFeTagDup = strNFeTagDup & "dup;" & vbCrLf & s
'                    End If
'                End With
'            Next
'        End If
    

'>  LISTA DE PRODUTOS
    vl_total_ICMS = 0
    vl_total_ICMSDeson = 0
    vl_total_ICMS_ST = 0
    vl_total_IPI = 0
    vl_total_produtos = 0
    vl_total_BC_ICMS = 0
    vl_total_BC_ICMS_ST = 0
    vl_total_PIS = 0
    vl_total_COFINS = 0
    vl_total_outras_despesas_acessorias = 0
    total_volumes = 0
    total_peso_bruto = 0
    total_peso_liquido = 0
    cubagem_bruto = 0
    intNumItem = 0
    vl_total_FCPUFDest = 0
    vl_total_ICMSUFDest = 0
    vl_total_ICMSUFRemet = 0
    vl_total_vFCP = 0
    vl_total_vFCPST = 0
    vl_total_vFCPSTRet = 0
    vl_total_vIPIDevol = 0


    'detectada necessidade de informar percentual de partilha do ano anterior, no caso de emisão de
    'nota de entrada referente a uma saída do ano anterior; restringir opção de utilização para
    'as notas de entrada com chave referenciada
    intAnoPartilha = Year(Date)
    If (rNFeImg.ide__tpNF = "0") And (Trim(c_chave_nfe_ref) <> "") Then
        s = "Utilizar percentual de partilha do ano anterior?"
        If confirma(s) Then
            intAnoPartilha = intAnoPartilha - 1
            End If
        End If
        
    
    For ic = LBound(v_nf) To UBound(v_nf)
        With v_nf(ic)
            If Trim$(.produto) <> "" Then
                intNumItem = intNumItem + 1
                
                If Trim$(vNFeImgItem(UBound(vNFeImgItem)).det__nItem) <> "" Then
                    ReDim Preserve vNFeImgItem(UBound(vNFeImgItem) + 1)
                    End If
                    
                vNFeImgItem(UBound(vNFeImgItem)).fabricante = .fabricante
                vNFeImgItem(UBound(vNFeImgItem)).produto = .produto
                
            '   TAG DET
            '   ~~~~~~~
            '   NÚMERO DO ITEM
                vNFeImgItem(UBound(vNFeImgItem)).det__nItem = CStr(intNumItem)
                strNFeTagDet = vbTab & NFeFormataCampo("nItem", vNFeImgItem(UBound(vNFeImgItem)).det__nItem)
                
            '   CÓDIGO DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__cProd = .produto
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cProd", vNFeImgItem(UBound(vNFeImgItem)).det__cProd)
                
            '   EAN
                vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = .EAN
                'NFE 4.0 - EM BRANCO, INFORMAR SEM GTIN
                If vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = "" Then vNFeImgItem(UBound(vNFeImgItem)).det__cEAN = "SEM GTIN"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cEAN", vNFeImgItem(UBound(vNFeImgItem)).det__cEAN)
            
            '   DESCRIÇÃO DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__xProd = UCase$(.descricao)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("xProd", vNFeImgItem(UBound(vNFeImgItem)).det__xProd)
                
            '   NCM
                vNFeImgItem(UBound(vNFeImgItem)).det__NCM = .ncm
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("NCM", vNFeImgItem(UBound(vNFeImgItem)).det__NCM)
                
            '=== aqui: campo NVE (não será usado)
            
            '   CEST
                vNFeImgItem(UBound(vNFeImgItem)).det__CEST = retorna_CEST(.ncm)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("CEST", vNFeImgItem(UBound(vNFeImgItem)).det__CEST)
            
            '   Indicador de Escala Relevante
                'CONVÊNIO ICMS 52, DE 7 DE ABRIL DE 2017
                'Cláusula vigésima terceira Os bens e mercadorias relacionados no Anexo XXVII serão considerados fabricados em escala industrial não relevante quando produzidos por contribuinte que atender, cumulativamente, as seguintes condições:
                'I - ser optante pelo Simples Nacional;
                'II - auferir, no exercício anterior, receita bruta igual ou inferior a R$ 180.000,00 (cento e oitenta mil reais);
                'III - possuir estabelecimento único;
                'IV - ser credenciado pela administração tributária da unidade federada de destino dos bens e mercadorias, quando assim exigido.
                vNFeImgItem(UBound(vNFeImgItem)).det__indEscala = "S"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("indEscala", "S")
                
            '   CFOP
                vNFeImgItem(UBound(vNFeImgItem)).det__CFOP = .cfop
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("CFOP", vNFeImgItem(UBound(vNFeImgItem)).det__CFOP)
            
            '   UNIDADE COMERCIAL
                vNFeImgItem(UBound(vNFeImgItem)).det__uCom = "PC"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uCom", vNFeImgItem(UBound(vNFeImgItem)).det__uCom)
                
            '   QUANTIDADE
                vNFeImgItem(UBound(vNFeImgItem)).det__qCom = NFeFormataNumero4Dec(.qtde_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qCom", vNFeImgItem(UBound(vNFeImgItem)).det__qCom)
                
            '   VALOR UNITÁRIO
                vl_unitario = .valor_total / .qtde_total
                vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom = NFeFormataNumero4Dec(vl_unitario)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vUnCom", vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom)
                
            '   VALOR TOTAL
                vNFeImgItem(UBound(vNFeImgItem)).det__vProd = NFeFormataMoeda2Dec(.valor_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vProd", vNFeImgItem(UBound(vNFeImgItem)).det__vProd)
                
            '   cEANTrib - GTIN (Global Trade Item Number) da unidade tributável
                vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = .EAN
                'NFE 4.0 - EM BRANCO, INFORMAR SEM GTIN
                If vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "" Then vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "SEM GTIN"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cEANTrib", vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib)
            
            '   UNIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__uTrib = "PC"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uTrib", vNFeImgItem(UBound(vNFeImgItem)).det__uTrib)
                
            '   QUANTIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__qTrib = NFeFormataNumero4Dec(.qtde_total)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qTrib", vNFeImgItem(UBound(vNFeImgItem)).det__qTrib)
                
            '   VALOR UNITÁRIO DE TRIBUTAÇÃO
                vl_unitario = .valor_total / .qtde_total
                vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib = NFeFormataNumero4Dec(vl_unitario)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vUnTrib", vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib)
                
            '   OUTRAS DESPESAS ACESSÓRIAS
                If .vl_outras_despesas_acessorias > 0 Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__vOutro = NFeFormataMoeda2Dec(.vl_outras_despesas_acessorias)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vOutro", vNFeImgItem(UBound(vNFeImgItem)).det__vOutro)
                    End If
                
            '   INDICA SE VALOR DO ITEM (vProd) ENTRA NO VALOR TOTAL DA NF-e (vProd)
            '       0 – o valor do item (vProd) não compõe o valor total da NF-e (vProd)
            '       1 – o valor do item (vProd) compõe o valor total da NF-e (vProd) (v2.0)
                vNFeImgItem(UBound(vNFeImgItem)).det__indTot = "1"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("indTot", vNFeImgItem(UBound(vNFeImgItem)).det__indTot)
                
            '   xPed (número do pedido de compra)
                If Trim$(.xPed) <> "" Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__xPed = Trim$(.xPed)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("xPed", vNFeImgItem(UBound(vNFeImgItem)).det__xPed)
                    End If
                
            '   nItemPed (item do pedido de compra)
                If Trim$(.nItemPed) <> "" Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__nItemPed = Trim$(.nItemPed)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("nItemPed", vNFeImgItem(UBound(vNFeImgItem)).det__nItemPed)
                    End If
                
            '   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
                If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) Then
                    perc_IBPT = ibpt_aliquota_aplicavel(.cst, .percAliqNac, .percAliqImp)
                    vl_estimado_tributos = arredonda_para_monetario(.valor_total * (perc_IBPT / 100))
                    vNFeImgItem(UBound(vNFeImgItem)).det__vTotTrib = NFeFormataMoeda2Dec(vl_estimado_tributos)
                    strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vTotTrib", vNFeImgItem(UBound(vNFeImgItem)).det__vTotTrib)
                    vl_total_estimado_tributos = vl_total_estimado_tributos + vl_estimado_tributos
                    End If
                
                
            '   TAG ICMS
            '   ~~~~~~~~
                If IsNumeric(.ICMS) Then
                    perc_ICMS = CSng(.ICMS)
                Else
                    perc_ICMS = 0
                    End If
                
                vl_ICMS = 0
                vl_BC_ICMS = .valor_total
            
                vl_ICMSDeson = 0
                
                vl_ICMS_ST = 0
                vl_BC_ICMS_ST = 0
                
                vl_ICMS_ST_Ret = 0
                vl_BC_ICMS_ST_Ret = 0
                
                If Len(Trim$(.cst)) = 0 Then
                    s_erro = "O produto " & .produto & " - " & .descricao & " não possui a informação do CST!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Len(Trim$(.cst)) <> 3 Then
                    s_erro = "O produto " & .produto & " - " & .descricao & " possui o campo CST preenchido com valor inválido!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
            
            '   ORIGEM DA MERCADORIA
            '   LEMBRANDO QUE OS CAMPOS 'ORIG' E 'CST' ESTÃO CONCATENADOS NA PLANILHA DE PRODUTOS,
            '   MAS PODEM TER SIDO ALTERADOS ATRAVÉS DO CAMPO 'CST' NA TELA.
                vNFeImgItem(UBound(vNFeImgItem)).ICMS__orig = Trim$(left$(.cst, 1))
                strNFeTagIcms = vbTab & NFeFormataCampo("orig", vNFeImgItem(UBound(vNFeImgItem)).ICMS__orig)
                
            '   TAG ICMS
            '   ~~~~~~~~
                'para a nota de remessa da operação triangular, não tributada (CST="41") / ORIENTAÇÃO CONTABILIDADE
                strNFeCst = "41"
                vNFeImgItem(UBound(vNFeImgItem)).ICMS__CST = strNFeCst
                strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__CST)
                                
            '   ICMS (CST=00): TRIBUTADO INTEGRALMENTE
                If strNFeCst = "00" Then
                    vl_ICMS = .valor_total * (perc_ICMS / 100)
                    vl_ICMS = CCur(Format$(vl_ICMS, FORMATO_MOEDA))
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS
                '   0: MARGEM VALOR AGREGADO (%); 1: PAUTA (VALOR); 2: PREÇO TABELADO MÁX. (VALOR); 3: VALOR DA OPERAÇÃO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC = "3"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC)
                    
                '   VALOR DA BC DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC)
                    
                '   ALÍQUOTA DO IMPOSTO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS)
                    
                '   VALOR DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS = NFeFormataMoeda2Dec(vl_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS)
                
                '   VALOR DO ICMS DESONERADO (ZERO, ATÉ RESOLUÇÃO EM CONTRÁRIO)
                    If vl_ICMSDeson <> 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson = NFeFormataMoeda2Dec(vl_ICMSDeson)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSDeson", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson)
                        End If
                
            '   ICMS (CST=10): TRIBUTADA E COM COBRANÇA DO ICMS POR SUBSTITUIÇÃO TRIBUTÁRIA
                ElseIf strNFeCst = "10" Then
                    vl_ICMS = .valor_total * (perc_ICMS / 100)
                    vl_ICMS = CCur(Format$(vl_ICMS, FORMATO_MOEDA))
                
                    If Not obtem_aliquota_ICMS_ST(rNFeImg.dest__UF, perc_ICMS_ST_aux, s_erro_aux) Then
                        s_erro = "Falha ao tentar obter a alíquota do ICMS ST para a UF: '" & rNFeImg.dest__UF & "'"
                        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                        End If
                    perc_ICMS_ST = perc_ICMS_ST_aux
                    
                    vl_BC_ICMS_ST = calcula_BC_ICMS_ST(.valor_total, .perc_MVA_ST)
                    vl_ICMS_ST = calcula_ICMS_ST(vl_BC_ICMS_ST, perc_ICMS_ST, vl_ICMS)
                    vl_ICMS_ST = CCur(Format$(vl_ICMS_ST, FORMATO_MOEDA))
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS
                '   0: MARGEM VALOR AGREGADO (%); 1: PAUTA (VALOR); 2: PREÇO TABELADO MÁX. (VALOR); 3: VALOR DA OPERAÇÃO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC = "3"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBC)
                    
                '   VALOR DA BC DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBC)
                    
                '   ALÍQUOTA DO IMPOSTO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMS)
                
                '   VALOR DO ICMS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS = NFeFormataMoeda2Dec(vl_ICMS)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMS", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMS)
                
                '   VALOR DO ICMS DESONERADO (ZERO, ATÉ RESOLUÇÃO EM CONTRÁRIO)
                    If vl_ICMSDeson <> 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson = NFeFormataMoeda2Dec(vl_ICMSDeson)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSDeson", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson)
                        End If
                
                '   MODALIDADE DE DETERMINAÇÃO DA BC DO ICMS ST
                '   0: PREÇO TABELADO OU MÁXIMO SUGERIDO; 1: LISTA NEGATIVA (VALOR); 2: LISTA POSITIVA (VALOR); 3: LISTA NEUTRA (VALOR)
                '   4: MARGEM VALOR AGREGADO (%); 5: PAUTA (VALOR)
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBCST = "4"
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("modBCST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__modBCST)
                    
                '   PERCENTUAL DA MARGEM DE VALOR ADICIONADO DO ICMS ST
                    If .perc_MVA_ST > 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__pMVAST = NFeFormataPercentual2Dec(.perc_MVA_ST)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pMVAST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pMVAST)
                        End If
                    
                '   VALOR DA BC DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCST = NFeFormataMoeda2Dec(vl_BC_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBCST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCST)
                    
                '   ALÍQUOTA DO IMPOSTO DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMSST = NFeFormataPercentual2Dec(perc_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pICMSST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__pICMSST)
                    
                '   VALOR DO ICMS ST
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSST = NFeFormataMoeda2Dec(vl_ICMS_ST)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSST", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSST)
                    
            '   ICMS (CST=40,41,50): ISENTA, NÃO TRIBUTADA OU SUSPENSÃO (40=ISENTA, 41=NÃO TRIBUTADA, 50=SUSPENSÃO)
                ElseIf (strNFeCst = "40") Or (strNFeCst = "41") Or (strNFeCst = "50") Then
                '   NOP: DEMAIS CAMPOS SÃO OPCIONAIS E NÃO SE APLICAM
                    vl_ICMS = 0
                    vl_BC_ICMS = 0
                
            '   ICMS (CST=60): ICMS COBRADO ANTERIORMENTE POR SUBSTITUIÇÃO TRIBUTÁRIA
                ElseIf strNFeCst = "60" Then
                    blnHaProdutoCstIcms60 = True
                    
                    vl_ICMS = 0
                    vl_BC_ICMS = 0

                '   VALOR DA BC DO ICMS ST
                    vl_BC_ICMS_ST_Ret = 0
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCSTRet = NFeFormataMoeda2Dec(vl_BC_ICMS_ST_Ret)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vBCSTRet", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vBCSTRet)
                    
                '   VALOR DO ICMS ST
                    vl_ICMS_ST_Ret = 0
                    vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSSTRet = NFeFormataMoeda2Dec(vl_ICMS_ST_Ret)
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSSTRet", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSSTRet)
                
            '   ICMS: CÓDIGO DE CST NÃO TRATADO PELO SISTEMA!!
                Else
                    s_erro = "Código de CST sem tratamento definido no sistema (CST=" & strNFeCst & ")!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
                    
            '   OS CÁLCULOS DE PARTILHA FORAM MOVIDOS PARA CÁ DEVIDO À EXCLUSÃO DE ICMS E DIFAL DAS BASES DE CÁLCULO
            '   DE PIS E COFINS, CONFORME DECISÃO DO STF
            
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    (vl_ICMS > 0) Then
                    
                    If IsNumeric(.fcp) Then
                        perc_fcp = CSng(.fcp)
                    Else
                        perc_fcp = 0
                        End If
                    
                    If Not obtem_aliquota_ICMS_UF_destino(rNFeImg.dest__UF, perc_ICMS_interna_UF_dest, s_erro_aux) Then
                        s_erro = "Falha ao tentar obter a alíquota interna do ICMS para a UF: '" & rNFeImg.dest__UF & "'"
                        GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                        End If
                    
                    If intAnoPartilha < 2016 Then
                        perc_ICMS_UF_dest = 0
                        perc_ICMS_UF_remet = 100
                    ElseIf intAnoPartilha = 2016 Then
                        perc_ICMS_UF_dest = 40
                        perc_ICMS_UF_remet = 60
                    ElseIf intAnoPartilha = 2017 Then
                        perc_ICMS_UF_dest = 60
                        perc_ICMS_UF_remet = 40
                    ElseIf intAnoPartilha = 2018 Then
                        perc_ICMS_UF_dest = 80
                        perc_ICMS_UF_remet = 20
                    Else
                        perc_ICMS_UF_dest = 100
                        perc_ICMS_UF_remet = 0
                        End If
                    
                    'os cálculos abaixo se baseiam em um vídeo publicado pela Inventti Soluções
                    '(https://www.youtube.com/watch?v=MEoI88y-qNs)
                    perc_ICMS_diferencial_interestadual = perc_ICMS_interna_UF_dest + perc_fcp - perc_ICMS
                    vl_ICMS_diferencial_interestadual = vl_BC_ICMS * (perc_ICMS_diferencial_interestadual / 100)
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_interestadual
                    vl_fcp = vl_BC_ICMS * perc_fcp / 100
                    vl_fcp = CCur(Format$(vl_fcp, FORMATO_MOEDA))
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_aux - vl_fcp
                    vl_ICMS_UF_dest = arredonda_para_monetario(vl_ICMS_diferencial_aux * perc_ICMS_UF_dest / 100)
                    vl_ICMS_diferencial_aux = vl_ICMS_diferencial_aux - vl_ICMS_UF_dest
                    vl_ICMS_UF_remet = arredonda_para_monetario(vl_ICMS_diferencial_aux)
                    If vl_ICMS_UF_remet < 0 Then vl_ICMS_UF_remet = 0
                    
                    End If
                    
            '   TAG IPI
            '   ~~~~~~~
            '   OBS: EXISTE IPI APENAS NA EMISSÃO DE NFe PARA DEVOLUÇÃO AO FORNECEDOR
                If IsNumeric(c_ipi) Then
                    perc_IPI = CSng(c_ipi)
                Else
                    perc_IPI = 0
                    End If
                
            '   TRAVA DE PROTEÇÃO ENQUANTO NÃO HÁ A IMPLEMENTAÇÃO DO TRATAMENTO
                If perc_IPI <> 0 Then
                    s_erro = "Não há tratamento definido no sistema para a alíquota de IPI!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
            
                vl_IPI = .valor_total * (perc_IPI / 100)
                vl_IPI = CCur(Format$(vl_IPI, FORMATO_MOEDA))
                
            '   TAG PIS
            '   ~~~~~~~
                vl_PIS = 0
                vl_BC_PIS = 0
                
                strZerarPisCst = Trim$(left$(cb_zerar_PIS, 2))
                
                If strZerarPisCst = "" Then
                    vl_BC_PIS = .valor_total
                    

                    If param_bc_pis_cofins_icms.campo_inteiro = 1 Then
                        vl_BC_PIS = vl_BC_PIS - vl_ICMS
                        End If
                    
                    If param_bc_pis_cofins_difal.campo_inteiro = 1 Then
                        vl_BC_PIS = vl_BC_PIS - vl_ICMS_UF_remet - vl_ICMS_UF_dest
                        End If

                    perc_PIS = PERC_PIS_ALIQUOTA_NORMAL
                    vl_PIS = vl_BC_PIS * (perc_PIS / 100)
                    vl_PIS = CCur(Format$(vl_PIS, FORMATO_MOEDA))
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__CST = "01"
                    strNFeTagPis = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).PIS__CST)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC = NFeFormataMoeda2Dec(vl_BC_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS = NFeFormataPercentual2Dec(perc_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("pPIS", vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS)
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS = NFeFormataMoeda2Dec(vl_PIS)
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("vPIS", vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).PIS__CST = strZerarPisCst
                    strNFeTagPis = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).PIS__CST)
                    End If
            
            '   TAG COFINS
            '   ~~~~~~~~~~
                vl_COFINS = 0
                vl_BC_COFINS = 0
                
                strZerarCofinsCst = Trim$(left$(cb_zerar_COFINS, 2))
                
                If strZerarCofinsCst = "" Then
                    vl_BC_COFINS = .valor_total
                    
                    If param_bc_pis_cofins_icms.campo_inteiro = 1 Then
                        vl_BC_COFINS = vl_BC_COFINS - vl_ICMS
                        End If
                        
                    If param_bc_pis_cofins_difal.campo_inteiro = 1 Then
                        vl_BC_COFINS = vl_BC_COFINS - vl_ICMS_UF_remet - vl_ICMS_UF_dest
                        End If
                    
                    perc_COFINS = PERC_COFINS_ALIQUOTA_NORMAL
                    vl_COFINS = vl_BC_COFINS * (perc_COFINS / 100)
                    vl_COFINS = CCur(Format$(vl_COFINS, FORMATO_MOEDA))
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST = "01"
                    strNFeTagCofins = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC = NFeFormataMoeda2Dec(vl_BC_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS = NFeFormataPercentual2Dec(perc_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("pCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS)
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS = NFeFormataMoeda2Dec(vl_COFINS)
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST = strZerarCofinsCst
                    strNFeTagCofins = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST)
                    End If
                
            '   TAG ICMSUFDest
            '   ~~~~~~~~~~~~~~
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    (vl_ICMS > 0) Then
                
                    strNFeTagIcmsUFDest = ""
                    
                '   VALOR DA BC DO ICMS NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest)
                    
                    'VALOR DA BASE DE CÁLCULO DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    '(lhgx) obs: manter esta linha comentada, pois podemos ter problema com o resultado no ambiente de produção
                    'strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCFCPUFDest", NFeFormataMoeda2Dec(vl_BC_ICMS))
                    
                '   PERCENTUAL DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest = NFeFormataPercentual2Dec(perc_fcp)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pFCPUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest)
                '   ALÍQUOTA INTERNA DA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSUFDest = NFeFormataPercentual2Dec(perc_ICMS_interna_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSUFDest)
                '   ALÍQUOTA INTERESTADUAL DAS UF ENVOLVIDAS
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInter = NFeFormataPercentual2Dec(perc_ICMS)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSInter", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInter)
                '   PERCENTUAL PROVISÓRIO DE PARTILHA DO ICMS INTERESTADUAL
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInterPart = NFeFormataPercentual2Dec(perc_ICMS_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pICMSInterPart", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pICMSInterPart)
                '   VALOR DO ICMS RELATIVO AO FCP DA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vFCPUFDest = NFeFormataMoeda2Dec(vl_fcp)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vFCPUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vFCPUFDest)
                '   VALOR DO ICMS INTERESTADUAL PARA A UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFDest = NFeFormataMoeda2Dec(vl_ICMS_UF_dest)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vICMSUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFDest)
                '   VALOR DO ICMS INTERESTADUAL PARA A UF DO REMETENTE
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFRemet = NFeFormataMoeda2Dec(vl_ICMS_UF_remet)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vICMSUFRemet", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vICMSUFRemet)
    
                    vl_total_FCPUFDest = vl_total_FCPUFDest + vl_fcp
                    vl_total_ICMSUFDest = vl_total_ICMSUFDest + vl_ICMS_UF_dest
                    vl_total_ICMSUFRemet = vl_total_ICMSUFRemet + vl_ICMS_UF_remet
                    End If
                    
            
            
            '   MONTA BLOCO POR PRODUTO
            '   ~~~~~~~~~~~~~~~~~~~~~~~
                strNFeTagBlocoProduto = strNFeTagBlocoProduto & _
                                        "det;" & vbCrLf & strNFeTagDet & _
                                        "ICMS;" & vbCrLf & strNFeTagIcms & _
                                        "PIS;" & vbCrLf & strNFeTagPis & _
                                        "COFINS;" & vbCrLf & strNFeTagCofins
                
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    (vl_ICMS > 0) Then
                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & _
                                            "ICMSUFDest;" & vbCrLf & strNFeTagIcmsUFDest
                    End If
                
            '   INFORMAÇÕES ADICIONAIS DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd = .infAdProd
                If Trim$(vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd) <> "" Then
                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & vbTab & NFeFormataCampo("infAdProd", vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd)
                    End If
                
            '   QTDE DE VOLUMES
                total_volumes = total_volumes + .qtde_volumes_total
                
            '   PESO BRUTO
                total_peso_bruto = total_peso_bruto + .peso_total
                    
            '   PESO LIQUIDO
                total_peso_liquido = total_peso_liquido + .peso_total
                
            '   CUBAGEM TOTAL
                cubagem_bruto = cubagem_bruto + .cubagem_total
                
            '   TOTALIZAÇÃO
                vl_total_ICMS = vl_total_ICMS + vl_ICMS
                vl_total_ICMSDeson = vl_total_ICMSDeson + vl_ICMSDeson
                vl_total_ICMS_ST = vl_total_ICMS_ST + vl_ICMS_ST
                vl_total_produtos = vl_total_produtos + .valor_total
                vl_total_BC_ICMS = vl_total_BC_ICMS + vl_BC_ICMS
                vl_total_BC_ICMS_ST = vl_total_BC_ICMS_ST + vl_BC_ICMS_ST
                vl_total_IPI = vl_total_IPI + vl_IPI
                vl_total_PIS = vl_total_PIS + vl_PIS
                vl_total_COFINS = vl_total_COFINS + vl_COFINS
                vl_total_outras_despesas_acessorias = vl_total_outras_despesas_acessorias + .vl_outras_despesas_acessorias
                End If
            End With
        Next
    
    
'   QTDE TOTAL DE VOLUMES
'   ~~~~~~~~~~~~~~~~~~~~~
    If Trim$(c_total_volumes) <> "" Then
        If CLng(c_total_volumes) <> total_volumes Then
            s = "A quantidade total de volumes foi editada de " & CStr(total_volumes) & " para " & c_total_volumes & vbCrLf & _
                "Continua mesmo assim?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If
    
    
'   TAG TOTAL
'   ~~~~~~~~~
    strNFeTagValoresTotais = "total;" & vbCrLf
    
'   BASE DE CÁLCULO DO ICMS
    rNFeImg.total__vBC = NFeFormataMoeda2Dec(vl_total_BC_ICMS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vBC", rNFeImg.total__vBC)
                            
'   VALOR TOTAL DO ICMS
    rNFeImg.total__vICMS = NFeFormataMoeda2Dec(vl_total_ICMS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vICMS", rNFeImg.total__vICMS)

'   novo campo vICMSDeson (layout 3.10)
    rNFeImg.total__vICMSDeson = NFeFormataMoeda2Dec(vl_total_ICMSDeson)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vICMSDeson", rNFeImg.total__vICMSDeson)
    
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    ((rNFeImg.dest__indIEDest = "9") Or _
'                     ((rNFeImg.dest__indIEDest = "2") And (rNFeImg.dest__IE = ""))) Then
    If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
        (rNFeImg.dest__indIEDest = "9") And _
        Not cfop_eh_de_remessa(strCfopCodigo) Then
            rNFeImg.total__vFCPUFDest = NFeFormataMoeda2Dec(vl_total_FCPUFDest)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vFCPUFDest", rNFeImg.total__vFCPUFDest)
            rNFeImg.total__vICMSUFDest = NFeFormataMoeda2Dec(vl_total_ICMSUFDest)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vICMSUFDest", rNFeImg.total__vICMSUFDest)
            rNFeImg.total__vICMSUFRemet = NFeFormataMoeda2Dec(vl_total_ICMSUFRemet)
            strNFeTagValoresTotais = strNFeTagValoresTotais & _
                                     vbTab & NFeFormataCampo("vICMSUFRemet", rNFeImg.total__vICMSUFRemet)
        
        End If
        
    'NFE 4.0 - vFCP
    ' quando for emitida uma NF-e (modelo 55) interestadual (Campo: idDest = 2) para Consumidor Final (Campo: indFinal = 1)
    ' não contribuinte (Campo: indIEDest = 9) e o valor do FCP for informado em um campo diferente de vFCPUFDest haverá esta rejeição
    '(e-mail do Márcio da Target em 01/11/18
    rNFeImg.total__vFCP = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCP", rNFeImg.total__vFCP)


'   vBCST
    rNFeImg.total__vBCST = NFeFormataMoeda2Dec(vl_total_BC_ICMS_ST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vBCST", rNFeImg.total__vBCST)
    
'   vST
    rNFeImg.total__vST = NFeFormataMoeda2Dec(vl_total_ICMS_ST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vST", rNFeImg.total__vST)
    
    'NFE 4.0 - vFCPST
    rNFeImg.total__vFCPST = NFeFormataMoeda2Dec(vl_total_vFCPST)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCPST", rNFeImg.total__vFCPST)
    
    'NFE 4.0 - vFCPSTRet
    rNFeImg.total__vFCPSTRet = NFeFormataMoeda2Dec(vl_total_vFCPSTRet)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFCPSTRet", rNFeImg.total__vFCPSTRet)
    
    
'   VALOR TOTAL DOS PRODUTOS
    rNFeImg.total__vProd = NFeFormataMoeda2Dec(vl_total_produtos)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vProd", rNFeImg.total__vProd)
                             
'   VALOR TOTAL DO FRETE
    rNFeImg.total__vFrete = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vFrete", rNFeImg.total__vFrete)
    
'   VALOR TOTAL DO SEGURO
    rNFeImg.total__vSeg = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vSeg", rNFeImg.total__vSeg)
    
'   VALOR TOTAL DO DESCONTO
    rNFeImg.total__vDesc = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vDesc", rNFeImg.total__vDesc)
    
'   VALOR TOTAL DO II
    rNFeImg.total__vII = NFeFormataMoeda2Dec(0)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vII", rNFeImg.total__vII)
    
'   VALOR TOTAL DO IPI
    rNFeImg.total__vIPI = NFeFormataMoeda2Dec(vl_total_IPI)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vIPI", rNFeImg.total__vIPI)
                             
    'NFE 4.0 vIPIDevol
    rNFeImg.total__vIPIDevol = NFeFormataMoeda2Dec(vl_total_vIPIDevol)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vIPIDevol", rNFeImg.total__vIPIDevol)
                             
'   VALOR DO PIS
    rNFeImg.total__vPIS = NFeFormataMoeda2Dec(vl_total_PIS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vPIS", rNFeImg.total__vPIS)
    
'   VALOR DO COFINS
    rNFeImg.total__vCOFINS = NFeFormataMoeda2Dec(vl_total_COFINS)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vCOFINS", rNFeImg.total__vCOFINS)
    
'   VALOR DESPESAS ACESSÓRIAS
    rNFeImg.total__vOutro = NFeFormataMoeda2Dec(vl_total_outras_despesas_acessorias)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vOutro", rNFeImg.total__vOutro)
    
'   VALOR TOTAL DA NOTA
    vl_total_NF = vl_total_produtos
    If vl_total_IPI > 0 Then vl_total_NF = vl_total_NF + vl_total_IPI
    If vl_total_outras_despesas_acessorias > 0 Then vl_total_NF = vl_total_NF + vl_total_outras_despesas_acessorias
    rNFeImg.total__vNF = NFeFormataMoeda2Dec(vl_total_NF)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vNF", rNFeImg.total__vNF)
                             
'   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    strInfoAdicIbpt = ""
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) Then
        rNFeImg.total__vTotTrib = NFeFormataMoeda2Dec(vl_total_estimado_tributos)
        strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vTotTrib", rNFeImg.total__vTotTrib)
        perc_aux = 100 * (vl_total_estimado_tributos / vl_total_NF)
        strInfoAdicIbpt = "Valor Aprox. dos Tributos: " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_estimado_tributos) & " (" & formata_numero_2dec(perc_aux) & "%) Fonte: IBPT"
        End If
    
'   TAG TRANSP
'   ~~~~~~~~~~
'   MODALIDADE DO FRETE
    strNFeTagTransp = "transp;" & vbCrLf
    rNFeImg.transp__modFrete = left$(Trim$(cb_frete), 1)
    strNFeTagTransp = strNFeTagTransp & _
                      vbTab & NFeFormataCampo("modFrete", rNFeImg.transp__modFrete)
                              
'   TAG TRANSPORTA
'   ~~~~~~~~~~~~~~
'   DADOS DA TRANSPORTADORA
    If strTransportadoraId <> "" Then
        If Len(strTransportadoraCnpj) = 14 Then
            rNFeImg.transporta__CNPJ = strTransportadoraCnpj
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("CNPJ", rNFeImg.transporta__CNPJ)
        ElseIf Len(strTransportadoraCnpj) = 11 Then
            rNFeImg.transporta__CPF = strTransportadoraCnpj
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("CPF", rNFeImg.transporta__CPF)
            End If
        
        If strTransportadoraRazaoSocial <> "" Then
            rNFeImg.transporta__xNome = strTransportadoraRazaoSocial
            strNFeTagTransporta = strNFeTagTransporta & _
                                  vbTab & NFeFormataCampo("xNome", rNFeImg.transporta__xNome)
            End If
        
        If (Len(strTransportadoraCnpj) = 14) Then
            strCampo = strTransportadoraIE
            If InStr(UCase$(strCampo), "ISEN") > 0 Then strCampo = "ISENTO"
            If UCase$(strCampo) <> "ISENTO" Then strCampo = retorna_so_digitos(strCampo)
            strTransportadoraIE = strCampo
           
            If (Len(strTransportadoraIE) > 0) Then
                If (Len(strTransportadoraIE) < 2) Or (Len(strTransportadoraIE) > 14) Then
                    s_erro = "A Inscrição Estadual no cadastro da transportadora '" & strTransportadoraId & "' está preenchida com conteúdo inválido!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Len(strTransportadoraUF) = 0 Then
                    s_erro = "A UF no cadastro da transportadora '" & strTransportadoraId & "' não está preenchida!!" & vbCrLf & "Essa informação é necessária devido ao campo IE!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf Not UF_ok(strTransportadoraUF) Then
                    s_erro = "A UF no cadastro da transportadora '" & strTransportadoraId & "' está preenchida com conteúdo inválido!!" & vbCrLf & "Essa informação é necessária devido ao campo IE!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                ElseIf ConsisteInscricaoEstadual(strTransportadoraIE, strTransportadoraUF) <> 0 Then
                '   Retorno = 0 -> IE válida
                '   Retorno = 1 -> IE inválida
                    s_erro = "A Inscrição Estadual no cadastro da transportadora '" & strTransportadoraId & "' é inválida para a UF de '" & strTransportadoraUF & "'!!"
                    GoTo NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA
                    End If
                    
                rNFeImg.transporta__IE = strTransportadoraIE
                strNFeTagTransporta = strNFeTagTransporta & _
                                      vbTab & NFeFormataCampo("IE", rNFeImg.transporta__IE)
                
                rNFeImg.transporta__UF = strTransportadoraUF
                strNFeTagTransporta = strNFeTagTransporta & _
                                      vbTab & NFeFormataCampo("UF", rNFeImg.transporta__UF)
                End If
            End If
            
        If strNFeTagTransporta <> "" Then
            strNFeTagTransporta = "transporta;" & vbCrLf & strNFeTagTransporta
            End If
        End If
    
'   TAG VOL
'   ~~~~~~~
    strNFeTagVol = "vol;" & vbCrLf
    
'   QUANTIDADE DE VOLUMES TRANSPORTADOS
    If Trim$(c_total_volumes) <> "" Then
        rNFeImg.vol__qVol = retorna_so_digitos(CStr(CLng(c_total_volumes)))
    Else
        rNFeImg.vol__qVol = retorna_so_digitos(CStr(total_volumes))
        End If
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("qVol", rNFeImg.vol__qVol)
    
'   ESPÉCIE DOS VOLUMES TRANSPORTADOS
    rNFeImg.vol__esp = "VOLUME"
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("esp", rNFeImg.vol__esp)
    
'   PESO LÍQUIDO
    rNFeImg.vol__pesoL = NFeFormataNumero3Dec(total_peso_liquido)
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("pesoL", rNFeImg.vol__pesoL)
    
'   PESO BRUTO
    rNFeImg.vol__pesoB = NFeFormataNumero3Dec(total_peso_bruto)
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("pesoB", rNFeImg.vol__pesoB)
    
    
    'NFE 4.0 - tag pag
    strNFeTagPag = "pag;" & vbCrLf
    If Trim$(vNFeImgPag(UBound(vNFeImgPag)).pag__indPag) <> "" Then
        ReDim Preserve vNFeImgPag(UBound(vNFeImgPag) + 1)
        End If
    vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = ""
    'Segundo informado pelo Valter (Target) em e-mail de 27/06/2017, não deve ser informada no arquivo de integração,
    'ela é inserida automaticamente pelo sistema
    'strNFeTagPag = strNFeTagPag & "detpag;" & vbCrLf

    'Segundo informado pelo Valter (Target) em e-mail de 27/06/2017, não deve ser informada no arquivo de integração,
    'ela é inserida automaticamente pelo sistema
    'strNFeTagPag = strNFeTagPag & "detpag;" & vbCrLf
    'Os códigos de pagamento usados abaixo estão presente na nota técnica da SEFAZ
    'NT2020.006 v1.10 de Fevereiro de 2021:
    '   01=Dinheiro
    '   02=Cheque
    '   03=Cartão de Crédito
    '   04=Cartão de Débito
    '   05=Crédito Loja
    '   10=Vale Alimentação
    '   11=Vale Refeição
    '   12=Vale Presente
    '   13=Vale Combustível
    '   15=Boleto Bancário
    '   16=Depósito Bancário
    '   17=Pagamento Instantâneo (PIX)
    '   18=Transferência bancária, Carteira Digital
    '   19=Programa de fidelidade, Cashback, Crédito Virtual
    '   90=Sem pagamento
    '   99=Outros

    s_aux = param_nftipopag.campo_texto
    s = ""
    'Se a nota é de entrada ou ajuste/devolução - sem pagamento
    If rNFeImg.ide__tpNF = "0" Or _
        strNFeCodFinalidade = "3" Or _
        strNFeCodFinalidade = "4" Then
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = "90"
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(0)
    'Se o pagamento é à vista
    ElseIf strTipoParcelamento = COD_FORMA_PAGTO_A_VISTA Then
        'Para cada meio de pagamento abaixo:
        '   - Se for obrigatório informar um meio de pagamento diferente de "99-Outros" sem descrição:
        '       - Se o sistema estiver operando em contingência, informa "99-Outros" e fornece uma descrição
        '       - Se não estiver operando em contingência, informa o código da lista acima
        '   - Se não for obrigatório informar um meio de pagamento, informa "99-Outros" sem descrição
        Select Case t_PEDIDO("av_forma_pagto")
            Case ID_FORMA_PAGTO_DINHEIRO
                    If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                        s_aux = "99"
                        s = "Dinheiro"
                    Else
                        s_aux = "01"
                        End If
            Case ID_FORMA_PAGTO_CHEQUE
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Cheque"
                Else
                    s_aux = "02"
                    End If
            Case ID_FORMA_PAGTO_BOLETO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_BOLETO_AV
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_DEPOSITO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "16"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Depósito"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case Else
                If (param_nftipopag.campo_inteiro = 1) Then
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Meio de pagamento não identificado"
                    Else
                        s_aux = param_nftipopag.campo_texto
                        End If
                    Else
                        s_aux = "99" 'Outros
                        End If
            End Select
        
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = rNFeImg.total__vNF
    'Se o pagamento é à prazo
    ElseIf (strTipoParcelamento = COD_FORMA_PAGTO_PARCELADO_CARTAO) Or _
           (strTipoParcelamento = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) Then
        If (param_nftipopag.campo_inteiro = 1) Then
            s_aux = "03"
            If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                s_aux = "99"
                s = "Cartão"
                End If
        Else
            s_aux = "99"
            End If
        'obtém o total a prazo (retira o valor da entrada,se houver)
        vl_aux = vl_total_NF - vl_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "1"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(vl_aux)
    Else
        vl_aux = 0
        Select Case t_PEDIDO("pce_forma_pagto_prestacao")
            Case ID_FORMA_PAGTO_DINHEIRO
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Dinheiro"
                Else
                    s_aux = "01"
                    End If
            Case ID_FORMA_PAGTO_CHEQUE
                If param_contingencia_meio_pagamento_geral.campo_inteiro = 1 Then
                    s_aux = "99"
                    s = "Cheque"
                Else
                    s_aux = "02"
                    End If
            Case ID_FORMA_PAGTO_BOLETO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_BOLETO_AV
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "15"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Boleto"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "03"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
                        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Cartão"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case ID_FORMA_PAGTO_DEPOSITO
                If (param_nftipopag.campo_inteiro = 1) Then
                    s_aux = "16"
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Depósito"
                        End If
                Else
                    s_aux = "99"
                    End If
            Case Else
                If (param_nftipopag.campo_inteiro = 1) Then
                    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Then
                        s_aux = "99"
                        s = "Meio de pagamento não identificado"
                    Else
                        s_aux = param_nftipopag.campo_texto
                        End If
                    Else
                        s_aux = "99" 'Outros
                        End If
            End Select
        'obtém o total a prazo (retira o valor da entrada,se houver)
        vl_aux = vl_total_NF - vl_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "1"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = s_aux
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(vl_aux)
        End If
    
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("indPag", vNFeImgPag(UBound(vNFeImgPag)).pag__indPag)
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tPag", vNFeImgPag(UBound(vNFeImgPag)).pag__tPag)
    If (param_contingencia_meio_pagamento_geral.campo_inteiro = 1) Or _
        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Then
        If s <> "" Then strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("xPag", s)
        End If
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vPag", vNFeImgPag(UBound(vNFeImgPag)).pag__vPag)
    'Segundo informado pelo Valter (Target) em e-mail de 27/07/2017, o grupo vcard não deve ser informado no arquivo texto,
    'ele é preenchido pelo sistema
    'informações do intermediador
    If (param_nfintermediador.campo_inteiro = 1) And (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") Then
        'If (strMarketplaceCodOrigem <> "") Then
        If ((strMarketPlaceCNPJ <> "") And (strMarketPlaceCadIntTran <> "")) Then
            strNFeTagPag = strNFeTagPag & vbTab & "infIntermed;"
            strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("CNPJ", strMarketPlaceCNPJ)
            strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("idCadIntTran", strMarketPlaceCadIntTran)
            End If
        End If

                               
'   TAG INFADIC
'   ~~~~~~~~~~~
'   TEXTO FIXO SOBRE RESPONSABILIDADE DA INSTALAÇÃO
    If blnTemPagtoPorBoleto Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Não efetue qualquer pagamento desta nota fiscal a terceiros, pois a quitação da mesma só terá validade após o pagamento do(s) título(s) bancário(s) emitidos por esta empresa. Caso não receba o(s) título(s) até a data(s) do(s) vencimento(s) favor contatar (11)4858-2431."
        End If
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
    strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "A responsabilidade pelo serviço de instalação e/ou manutenção dos produtos acima é única e exclusivamente da empresa e/ou técnico autônomo contratado pelo destinatário desta."
    
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
    strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Fabricante não cobre avarias de peças plásticas, portanto, é necessário avaliar o equipamento no ato da entrega."
    
'   TEXTO FIXO SOBRE REGIME ESPECIAL
    If txtFixoEspecifico <> "" Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & txtFixoEspecifico
        End If

'   OUTROS TELEFONES DE CONTATO (INF ADICIONAIS)
    s_aux = ""
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PF Then
        strSufixoRes = "Tel Res: "
        strSufixoCom = "Tel Com: "
    Else
        strSufixoRes = "Tel: "
        strSufixoCom = "Tel: "
        End If
    If (strTelCel <> "") And (strTelRes <> "") Then s_aux = strSufixoRes & strTelRes
    If ((strTelCel <> "") Or (strTelRes <> "")) And (strTelCom <> "") Then
        If s_aux <> "" Then s_aux = s_aux & " / "
        s_aux = s_aux & strSufixoCom & strTelCom
        End If
    If ((strTelCel <> "") Or (strTelRes <> "") Or (strTelCom <> "")) And (strTelCom2 <> "") Then
        If s_aux <> "" Then s_aux = s_aux & " / "
        s_aux = s_aux & strSufixoCom & strTelCom2
        End If
    If s_aux <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & s_aux
        End If
    
'   TEXTO DIGITADO
    If Trim$(c_dados_adicionais_remessa) <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & Trim$(c_dados_adicionais_remessa)
        End If
    
    If blnHaProdutoCstIcms60 Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = TEXTO_LEI_CST_ICMS_60 & strNFeInfAdicQuadroProdutos
        End If
    
'   BEM DE USO E CONSUMO
    If blnTemPedidoComStBemUsoConsumo And (Not blnTemPedidoSemStBemUsoConsumo) Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = "BEM DE USO E CONSUMO" & strNFeInfAdicQuadroProdutos
        End If

'   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) And (strInfoAdicIbpt <> "") Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = strInfoAdicIbpt & strNFeInfAdicQuadroProdutos
        End If
    
'   Nº PEDIDO (NA 1ª LINHA) + CUBAGEM
    strTextoCubagem = ""
    If cubagem_bruto > 0 Then strTextoCubagem = Space$(20) & "CUB: " & formata_numero_2dec(cubagem_bruto) & " m3"
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
    strNFeInfAdicQuadroProdutos = Join(v_pedido, ", ") & strTextoCubagem & strNFeInfAdicQuadroProdutos
    
'   INFORMAÇÕES SOBRE PARTILHA DO ICMS
    If PARTILHA_ICMS_ATIVA Then
        'DIFAL- suprimir texto em notas de entrada/devolução
        If (rNFeImg.ide__tpNF <> "0") And _
            (strNFeCodFinalidade <> "3") And _
            (strNFeCodFinalidade <> "4") And _
                Not tem_instricao_virtual(usuario.emit_id, rNFeImg.dest__UF) Then
            If (vl_total_ICMSUFDest > 0) Then
                If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
                strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Valores totais do ICMS Interestadual: partilha da UF Destino " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_ICMSUFDest)
                If (vl_total_FCPUFDest > 0) Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & " + FCP " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_FCPUFDest)
                strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "; partilha da UF Origem " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_ICMSUFRemet) & "."
                End If
            End If
        End If

    rNFeImg.infAdic__infCpl = strNFeInfAdicQuadroInfAdic & "|" & strNFeInfAdicQuadroProdutos
    strNFeTagInfAdicionais = "infAdic;" & vbCrLf & _
                             vbTab & NFeFormataCampo("infCpl", rNFeImg.infAdic__infCpl)
    
'   LHGX - ENTENDO QUE A TAG 'entrega' SÓ FAZ SENTIDO NA NOTA FISCAL DE VENDA; DESPREZAR POR ENQUANTO
''   TAG ENTREGA
''   ~~~~~~~~~~~
'    strDestinatarioCnpjCpf = c_cnpj_cpf_dest
'    strEndEtgEndereco = l_end_recebedor_logradouro
'    strEndEtgEnderecoNumero = l_end_recebedor_numero
'    strEndEtgEnderecoComplemento = l_end_recebedor_complemento
'    strEndEtgBairro = l_end_recebedor_bairro
'    strEndEtgCidade = l_end_recebedor_cidade
'    strEndEtgUf = l_end_recebedor_uf
''   NO MOMENTO, A SEFAZ ACEITA ENDEREÇO DE ENTREGA DIFERENTE DO ENDEREÇO DE CADASTRO SOMENTE P/ PJ
'    If (UCase$(strEndEtgUf) <> UCase$(strEndClienteUf)) And _
'        cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
'        strNFeTagEndEntrega = "entrega;" & vbCrLf
'        If (Len(strDestinatarioCnpjCpf) = 14) Then
'            rNFeImg.entrega__CNPJ = strDestinatarioCnpjCpf
'            strNFeTagEndEntrega = strNFeTagEndEntrega & vbTab & NFeFormataCampo("CNPJ", rNFeImg.entrega__CNPJ)
'        Else
'            rNFeImg.entrega__CPF = strDestinatarioCnpjCpf
'            strNFeTagEndEntrega = strNFeTagEndEntrega & vbTab & NFeFormataCampo("CPF", rNFeImg.entrega__CPF)
'            End If
'
'        rNFeImg.entrega__xLgr = strEndEtgEndereco
'        rNFeImg.entrega__nro = strEndEtgEnderecoNumero
'        rNFeImg.entrega__xCpl = strEndEtgEnderecoComplemento
'        rNFeImg.entrega__xBairro = strEndEtgBairro
'        rNFeImg.entrega__cMun = strEndEtgCidade & "/" & strEndEtgUf
'        rNFeImg.entrega__xMun = strEndEtgCidade
'        rNFeImg.entrega__UF = strEndEtgUf
'
'        strNFeTagEndEntrega = strNFeTagEndEntrega & _
'                              vbTab & NFeFormataCampo("xLgr", rNFeImg.entrega__xLgr) & _
'                              vbTab & NFeFormataCampo("nro", rNFeImg.entrega__nro)
'
'        If Len(rNFeImg.entrega__xCpl) > 0 Then
'            strNFeTagEndEntrega = strNFeTagEndEntrega & _
'                              vbTab & NFeFormataCampo("xCpl", rNFeImg.entrega__xCpl)
'            End If
'
'        strNFeTagEndEntrega = strNFeTagEndEntrega & _
'                              vbTab & NFeFormataCampo("xBairro", rNFeImg.entrega__xBairro) & _
'                              vbTab & NFeFormataCampo("cMun", rNFeImg.entrega__cMun) & _
'                              vbTab & NFeFormataCampo("xMun", rNFeImg.entrega__xMun) & _
'                              vbTab & NFeFormataCampo("UF", rNFeImg.entrega__UF)
'        End If


'   SÓ AUTORIZA EMISSÃO SEM INTERMEDIADOR SE intImprimeIntermediadorAusente FOR 1
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If (param_nfintermediador.campo_inteiro = 1) Then
        If (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") And _
            ((strMarketPlaceCNPJ = "") Or (strMarketPlaceCadIntTran = "")) And _
            (intImprimeIntermediadorAusente = 0) Then
            s = "Não é possível prosseguir com a emissão, pois o intermediador do pedido não está identificado!!"
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If


'   SE HOUVER MAIS DE UMA CONFIRMAÇÃO DE EMISSÃO QUE PODEM GERAR NFe PARA UM EMITENTE INDEVIDO, CONFIRMAR NOVAMENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If iQtdConfirmaDuvidaEmit > 1 Then
        s = "Algumas confirmações efetuadas indicam que a NFe pode ser gerada em um Emitente indevido." & vbCrLf & _
            "Confirma a emissão no Emitente " & usuario.emit & "?"
        If Not confirma(s) Then
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
'   CONFIRMAÇÃO FINAL
'   ~~~~~~~~~~~~~~~~~
    s = Join(v_pedido(), ", ")
    If qtde_pedidos = 1 Then
        s = " para o pedido " & s & "?"
    Else
        s = " para os pedidos " & s & "?"
        End If
    
    s = "Emite a NFe " & s
    
    If Not confirma(s) Then
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    
'   OBTÉM NSU P/ GRAVAR OS DADOS DA NFe P/ FINS DE HISTÓRICO, CONTROLE E CONSULTA DA DANFE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If Not geraNsu(NSU_T_NFe_EMISSAO, lngNsuNFeEmissao, s_erro_aux) Then
        s = "Falha ao tentar gerar o NSU para a tabela " & NSU_T_NFe_EMISSAO & "!!"
        If s_erro_aux <> "" Then s = s & vbCrLf
        s = s & s_erro_aux
        aviso_erro s
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
  
    
'   OBTÉM Nº SÉRIE E PRÓXIMO Nº PARA ATRIBUIR À NFe
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    If blnEsperaNFTriangular Then
'        strSerieNf = CStr(lngSerieNFeTriangular)
'        strNumeroNf = CStr(lngNumRemessaNFeTriangular)
'    Else
'        aguarde INFO_EXECUTANDO, "obtendo próximo número de NF"
'        If Not NFeObtemProximoNumero(rNFeImg.id_nfe_emitente, strSerieNf, strNumeroNf, s_erro_aux) Or _
'            Not NFeObtemProximoNumero(rNFeImg.id_nfe_emitente, strSerieNfTriangular, strNumeroNfTriangular, s_erro_aux) Then
'            s = "Falha ao tentar gerar o número para a NFe!!"
'            If s_erro_aux <> "" Then s = s & vbCrLf
'            s = s & s_erro_aux
'            aviso_erro s
'            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
'            aguarde INFO_NORMAL, m_id
'            Exit Sub
'        Else
'            lngSerieNFeTriangular = CLng(strSerieNf)
'            lngNumVendaNFeTriangular = CLng(strNumeroNf)
'            lngNumRemessaNFeTriangular = CLng(strNumeroNfTriangular)
'            End If
'        End If
    strNumeroNf = CStr(lngNumRemessaNFeTriangular)
    strSerieNf = CStr(lngSerieNFeTriangular)

'   VERIFICA SE O Nº DA NFE A SER EMITIDA ENCONTRA-SE INUTILIZADO (A OPERAÇÃO DE INUTILIZAÇÃO DE FAIXAS DE NÚMEROS DA NFe É
'   REALIZADA NO SISTEMA DA TARGET ONE)
    s = "SELECT " & _
            "*" & _
        " FROM NFE_INUTILIZA" & _
        " WHERE" & _
            " (Serie = '" & NFeFormataSerieNF(strSerieNf) & "')" & _
            " AND (NumIni >= '" & NFeFormataNumeroNF(strNumeroNf) & "')" & _
            " AND (NumFim <= '" & NFeFormataNumeroNF(strNumeroNf) & "')"
    If t_T1_NFE_INUTILIZA.State <> adStateClosed Then t_T1_NFE_INUTILIZA.Close
    t_T1_NFE_INUTILIZA.Open s, dbcNFe, , , adCmdText
    If Not t_T1_NFE_INUTILIZA.EOF Then
    '   CÓDIGOS: 1=Em Processamento; 2=Falha; 3=Homologado
        strCodStatusInutilizacao = Trim$("" & t_T1_NFE_INUTILIZA("Status"))
        s_erro_aux = "Data: " & Format$(t_T1_NFE_INUTILIZA("DataHora"), FORMATO_DATA_HORA) & vbCrLf & _
                     "Nº inicial: " & Trim$("" & t_T1_NFE_INUTILIZA("NumIni")) & vbCrLf & _
                     "Nº final: " & Trim$("" & t_T1_NFE_INUTILIZA("NumFim")) & vbCrLf & _
                     "Série: " & Trim$("" & t_T1_NFE_INUTILIZA("Serie")) & vbCrLf & _
                     "Motivo: " & Trim$("" & t_T1_NFE_INUTILIZA("Motivo")) & vbCrLf & _
                     "Usuário: " & Trim$("" & t_T1_NFE_INUTILIZA("Usuario")) & vbCrLf & _
                     "Status: " & strCodStatusInutilizacao & " - " & decodifica_NFe_inutilizacao_status(strCodStatusInutilizacao) & _
                     "Código: " & Trim$("" & t_T1_NFE_INUTILIZA("PendSta")) & vbCrLf & _
                     "Mensagem: " & Trim$("" & t_T1_NFE_INUTILIZA("PendDes"))
        If strCodStatusInutilizacao = "3" Then
            s = "Não é possível prosseguir com a emissão, pois o número de NFe informado foi inutilizado!!" & vbCrLf & _
                s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
        ElseIf strCodStatusInutilizacao = "1" Then
            s = "Não é possível prosseguir com a emissão, pois o número de NFe informado consta em uma operação de inutilização de números de NFe que está em andamento!!" & vbCrLf & _
                s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If


'   SE O PEDIDO ESTIVER NA FILA DE SOLICITAÇÃO DE EMISSÃO DE NFE, SINALIZA QUE JÁ FOI TRATADO
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            If Not marca_status_atendido_fila_solicitacoes_emissao_NFe(Trim$(v_pedido(i)), rNFeImg.id_nfe_emitente, CLng(strSerieNf), CLng(strNumeroNf), s_erro_aux) Then
                s = "Não é possível prosseguir com a emissão, pois houve falha ao atualizar os dados da fila de solicitações de emissão de NFe!!" & vbCrLf & _
                    s_erro_aux
                aviso_erro s
                GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        Next


'   MONTA TAG IDENTIFICAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~
    rNFeImg.ide__natOp = strCfopDescricao
    rNFeImg.ide__serie = strSerieNf
    rNFeImg.ide__nNF = strNumeroNf
    rNFeImg.ide__dEmi = NFeFormataData(Date)
    rNFeImg.ide__dEmiUTC = NFeFormataDataHoraUTC(Now, blnHorarioVerao)
    rNFeImg.ide__cMunFG = strEmitenteCidade & "/" & strEmitenteUf
    rNFeImg.ide__tpAmb = NFE_AMBIENTE
    rNFeImg.ide__finNFe = strNFeCodFinalidade
    rNFeImg.ide__indFinal = NFE_INDFINAL_CONSUMIDOR_FINAL
    rNFeImg.ide__indPres = strPresComprador
    
    strNFeTagIdentificacao = "ide;" & vbCrLf
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("natOp", rNFeImg.ide__natOp)
    'NFE 4.0 - não enviar indPag (Este campo agora se encontra na tag "pag"
    'strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indPag", rNFeImg.ide__indPag)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("serie", rNFeImg.ide__serie)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("nNF", rNFeImg.ide__nNF)
    '=== Substituindo campo de acordo com layout 3.10
    'strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("dEmi", rNFeImg.ide__dEmi)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("dhEmi", rNFeImg.ide__dEmiUTC)
    '=== aqui: campo dhSaiEnt
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("tpNF", rNFeImg.ide__tpNF) '0-Entrada  1-Saída
    '=== Novo campo idDest
    '=== (1-Operação Interna; 2-Operação Interestadual; 3-Operação com o Exterior)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("idDest", rNFeImg.ide__idDest)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("cMunFG", rNFeImg.ide__cMunFG)
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("tpAmb", rNFeImg.ide__tpAmb) '1-Produção  2-Homologação
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("finNFe", rNFeImg.ide__finNFe) '1-Normal  2-Complementar  3-Ajuste
    '=== Novo campo indFinal
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indFinal", rNFeImg.ide__indFinal) '0-Normal  1-Consumidor Final
    strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indPres", rNFeImg.ide__indPres) '2-Internet  3-Teleatendimento
    '=== Campo indIntermed  (0-Sem intermediador 1-Operação em site ou plataforma de terceiros)
    If (param_nfintermediador.campo_inteiro = 1) Then
        If ((strMarketPlaceCNPJ <> "") And (strMarketPlaceCadIntTran <> "")) Then
            strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indIntermed", "1")
        Else
            strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indIntermed", "0")
            End If
        End If
    '=== aqui: campo IEST
    
    '=== Grupo NFref
    strNFeChaveAcessoNotaReferenciada = Trim$(c_chave_nfe_ref)
    If strNFeChaveAcessoNotaReferenciada <> "" Then
        vListaNFeRef = Split(strNFeChaveAcessoNotaReferenciada, vbCrLf)
        For i = LBound(vListaNFeRef) To UBound(vListaNFeRef)
            strNFeRef = Trim$(vListaNFeRef(i))
            If strNFeRef <> "" Then
                strNFeTagIdentificacao = strNFeTagIdentificacao & _
                                        "NFref;" & vbCrLf & _
                                        vbTab & NFeFormataCampo("refNFe", strNFeRef)
                If Trim$(vNFeImgNFeRef(UBound(vNFeImgNFeRef)).refNFe) <> "" Then
                    ReDim Preserve vNFeImgNFeRef(UBound(vNFeImgNFeRef) + 1)
                    End If
                vNFeImgNFeRef(UBound(vNFeImgNFeRef)).refNFe = strNFeRef
                End If
            Next
        End If

'   MONTA O ARQUIVO DE INTEGRAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNFeArquivo = strNFeTagOperacional & _
                   strNFeTagIdentificacao & _
                   strNFeTagDestinatario & _
                   strNFeTagEndEntrega & _
                   strNFeTagBlocoProduto & _
                   strNFeTagValoresTotais & _
                   strNFeTagTransp & _
                   strNFeTagTransporta & _
                   strNFeTagVol & _
                   strNFeTagDup & _
                   strNFeTagPag & _
                   strNFeTagInfAdicionais
    
    
'   REGISTRA DADOS DA NFE P/ FINS DE HISTÓRICO, CONTROLE E CONSULTA DA DANFE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "gravando histórico no sistema"
    
    If Not grava_NFe_imagem(usuario.id, CLng(strSerieNf), CLng(strNumeroNf), rNFeImg, vNFeImgItem(), vNFeImgTagDup(), vNFeImgNFeRef(), vNFeImgPag(), lngNsuNFeImagem, s_erro_aux) Then
        s = "Falha ao tentar gravar os dados da NFe (tabela imagem)!!"
        If s_erro_aux <> "" Then s = s & vbCrLf
        s = s & s_erro_aux
        aviso_erro s
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
            
'   LEMBRANDO QUE OS CAMPOS 'dt_emissao' E 'dt_hr_emissao' SÃO PREENCHIDOS AUTOMATICAMENTE POR UM "CONSTRAINT DEFAULT"
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (id = -1)"
    If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    t_NFe_EMISSAO.AddNew
    t_NFe_EMISSAO("id") = lngNsuNFeEmissao
    t_NFe_EMISSAO("id_nfe_emitente") = rNFeImg.id_nfe_emitente
    t_NFe_EMISSAO("NFe_serie_NF") = CLng(strSerieNf)
    t_NFe_EMISSAO("NFe_numero_NF") = CLng(strNumeroNf)
    t_NFe_EMISSAO("versao_layout_NFe") = ID_VERSAO_LAYOUT_NFe
    t_NFe_EMISSAO("usuario_emissao") = usuario.id
    t_NFe_EMISSAO("pedido") = rNFeImg.pedido
    t_NFe_EMISSAO("email_destinatario") = rNFeImg.operacional__email
    t_NFe_EMISSAO("nome_destinatario") = rNFeImg.dest__xNome
    t_NFe_EMISSAO("tipo_NF") = rNFeImg.ide__tpNF
    t_NFe_EMISSAO("tipo_ambiente") = NFE_AMBIENTE
    t_NFe_EMISSAO("finalidade_NF") = rNFeImg.ide__finNFe
    t_NFe_EMISSAO("natureza_operacao_codigo") = strCfopCodigoFormatado
    t_NFe_EMISSAO("natureza_operacao_descricao") = strCfopDescricao
    t_NFe_EMISSAO("aliquota_ICMS") = perc_ICMS
    t_NFe_EMISSAO("aliquota_IPI") = perc_IPI
    t_NFe_EMISSAO("frete_por_conta") = rNFeImg.transp__modFrete
    t_NFe_EMISSAO("volumes_qtde_total_sistema") = total_volumes
    t_NFe_EMISSAO("volumes_qtde_total_tela") = c_total_volumes
    
    s = RTrim$(c_dados_adicionais_remessa)
    lngMax = 2000
    If Len(s) > lngMax Then
        s_aux = " (...)"
        s = left$(s, lngMax - Len(s_aux)) & s_aux
        End If
    t_NFe_EMISSAO("dados_adicionais_digitado") = s
    
    s = strNFeArquivo
    lngMax = 6000
    If Len(s) > lngMax Then
        s_aux = " (...)"
        s = left$(s, lngMax - Len(s_aux)) & s_aux
        End If
    t_NFe_EMISSAO("arquivo_integracao_NFe_T1") = s
    t_NFe_EMISSAO.Update
    
            
'   TRANSFERE O ARQUIVO DE INTEGRAÇÃO PARA O SISTEMA DE NFe DA TARGET ONE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    strNumeroNfNormalizado = NFeFormataNumeroNF(strNumeroNf)
    strSerieNfNormalizado = NFeFormataSerieNF(strSerieNf)

  ' COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    aguarde INFO_EXECUTANDO, "emitindo NFe"
    Set cmdNFeEmite.ActiveConnection = dbcNFe
    cmdNFeEmite.CommandType = adCmdStoredProc
    cmdNFeEmite.CommandText = "Proc_NFe_Integracao_Emite"
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("NFe", adChar, adParamInput, 9, strNumeroNfNormalizado)
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("Serie", adChar, adParamInput, 3, strSerieNfNormalizado)
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("Arquivo", adVarChar, adParamInput, 16000, strNFeArquivo)
    Set rsNFeRetornoSPEmite = cmdNFeEmite.Execute
    intNfeRetornoSPEmite = rsNFeRetornoSPEmite("Retorno")
    strNFeMsgRetornoSPEmite = Trim$("" & rsNFeRetornoSPEmite("Mensagem"))
    
'   GRAVA O RESULTADO DA CHAMADA DA STORED PROCEDURE
    strNFeMsgRetornoSPEmiteTamAjustadoBD = strNFeMsgRetornoSPEmite
    lngMax = 2000
    If Len(strNFeMsgRetornoSPEmiteTamAjustadoBD) > lngMax Then
        s_aux = " (...)"
        strNFeMsgRetornoSPEmiteTamAjustadoBD = left$(strNFeMsgRetornoSPEmiteTamAjustadoBD, lngMax - Len(s_aux)) & s_aux
        End If
    
    Call atualiza_NFe_imagem_com_retorno_NFe_T1(lngNsuNFeImagem, CStr(intNfeRetornoSPEmite), strNFeMsgRetornoSPEmiteTamAjustadoBD, s_erro_aux)
    
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (id = " & lngNsuNFeEmissao & ")"
    If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If Not t_NFe_EMISSAO.EOF Then
        t_NFe_EMISSAO("codigo_retorno_NFe_T1") = CStr(intNfeRetornoSPEmite)
        t_NFe_EMISSAO("msg_retorno_NFe_T1") = strNFeMsgRetornoSPEmiteTamAjustadoBD
        t_NFe_EMISSAO.Update
        End If
        
        
'   ATUALIZA AS INFORMAÇÕES SOBRE A EMISSÃO TRIANGULAR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If Not atualiza_nfe_triangular_remessa(lngIdNFeTriangular, ST_NFT_EMITIDA, s_erro) Then
        s_erro = "Problemas na atualização da operação triangular (nota de remessa): " & s_erro
        aviso_erro s_erro
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    If Not atualiza_nfe_triangular_geral(lngIdNFeTriangular, ST_NFT_EMITIDA, s_erro) Then
        s_erro = "Problemas na atualização de status da operação triangular: " & s_erro
        aviso_erro s_erro
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    If Not atualiza_nfe_triangular_inf_adicionais(lngIdNFeTriangular, _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    "", _
                                                    c_dados_adicionais_remessa, _
                                                    s_erro) Then
        s_erro = "Problemas na atualização da nota triangular de remessa: " & s_erro
        aviso_erro s_erro
        GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
'   GRAVA O LOG
'   ~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "gravando log"
    If LBound(v_pedido) = UBound(v_pedido) Then
        strLogPedido = v_pedido(UBound(v_pedido))
    Else
        strLogPedido = ""
        End If
    strLogComplemento = "Retorno SP=" & CStr(intNfeRetornoSPEmite) & " (" & IIf(intNfeRetornoSPEmite = 1, "Sucesso", "Falha") & ")" & _
                        "; Msg SP=" & strNFeMsgRetornoSPEmite & _
                        "; Série NFe=" & strSerieNf & _
                        "; Nº NFe=" & strNumeroNf & _
                        "; tipo=" & cb_tipo_NF & _
                        "; pedido=" & Join(v_pedido, ", ") & _
                        "; natureza operação=" & cb_natureza_recebedor & _
                        "; ICMS=" & cb_icms_remessa & _
                        "; IPI=" & c_ipi & _
                        "; frete=" & cb_frete & _
                        "; zerar PIS=(" & Trim$(cb_zerar_PIS) & ")" & _
                        "; zerar COFINS=(" & Trim$(cb_zerar_COFINS) & ")" & _
                        "; finalidade=" & Trim$(cb_finalidade) & _
                        "; chave NFe referenciada=" & Trim$(c_chave_nfe_ref) & _
                        "; dados adicionais=" & Trim$(c_dados_adicionais_venda)
    Call grava_log(usuario.id, "", strLogPedido, "", OP_LOG_NFE_EMISSAO, strLogComplemento)
        
        
'   SUCESSO NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "processamento complementar"
    If intNfeRetornoSPEmite = 1 Then
        aguarde INFO_EXECUTANDO, "atualizando banco de dados"
    '   ATUALIZA O CAMPO "OBSERVAÇÕES III" COM O Nº DA NOTA FISCAL?
    '   A ATUALIZAÇÃO É FEITA SOMENTE P/ NOTAS DE SAÍDA, POIS EM NOTAS DE ENTRADA O Nº DA NFe NÃO É ANOTADO NO CAMPO
    '   OBS_3 DO PEDIDO, MAS SIM NOS ITENS DEVOLVIDOS, QUANDO APLICÁVEL.
    '   0-Entrada  1-Saída
        If rNFeImg.ide__tpNF = "1" Then
            If qtde_pedidos = 1 Then
              ' T_PEDIDO
                If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
                t_PEDIDO.CursorType = BD_CURSOR_EDICAO
                s = sql_monta_criterio_texto_or(v_pedido(), "pedido", True)
                s = "SELECT * FROM t_PEDIDO WHERE (" & s & ")"
                t_PEDIDO.Open s, dbc, , , adCmdText
                If Not t_PEDIDO.EOF Then
                    If (Trim$("" & t_PEDIDO("obs_3")) = "") Or IsLetra(Trim$("" & t_PEDIDO("obs_3"))) Then
                        t_PEDIDO("obs_3") = strNumeroNf
                        t_PEDIDO.Update
                        End If
                    End If
                End If
        ElseIf rNFeImg.ide__tpNF = "0" Then
            s = sql_monta_criterio_texto_or(v_pedido(), "pedido", True)
            If s <> "" Then
                s = "UPDATE t_PEDIDO_ITEM_DEVOLVIDO SET" & _
                        " id_nfe_emitente = " & CStr(rNFeImg.id_nfe_emitente) & "," & _
                        " NFe_serie_NF = " & strSerieNf & "," & _
                        " NFe_numero_NF = " & strNumeroNf & "," & _
                        " dt_hr_anotacao_numero_NF = getdate()," & _
                        " usuario_anotacao_numero_NF = '" & usuario.id & "'" & _
                    " WHERE" & _
                        " (" & s & ")" & _
                        " AND (NFe_numero_NF = 0)"
                dbc.Execute s, lngAffectedRecords
                End If
            End If
        
'   FALHA NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        aviso_erro "Falha na emissão da NFe:" & vbCrLf & strNFeMsgRetornoSPEmite
        End If
        
        
  ' LIMPA FORMULÁRIO
    c_pedido_danfe = rNFeImg.pedido
    formulario_limpa
        
  ' EXIBE DADOS DA ÚLTIMA NFe EMITIDA
    l_serie_NF = strSerieNfNormalizado
    l_num_NF = strNumeroNfNormalizado
    l_emitente_NF = strEmitenteNf
        
    GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
    
'    If blnFilaSolicitacoesEmissaoNFeEmTratamento Then
'    '   AO PREENCHER C/ O PRÓXIMO PEDIDO DA FILA, A QTDE PENDENTE NA FILA É ATUALIZADA AUTOMATICAMENTE
'        preenche_prox_pedido_fila_solicitacoes_emissao_NFe
'    Else
'    '   ATUALIZA A QTDE PENDENTE NA FILA, POIS O PEDIDO INFORMADO MANUALMENTE PODE TER SIDO UM QUE CONSTAVA NA FILA
'        atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
'        End If
    
    aguarde INFO_NORMAL, m_id
    
    If sPedidoTriangular <> "" Then
        sPedidoTriangular = ""
        sPedidoDANFETelaAnterior = rNFeImg.pedido
        sNFAnteriorSerie = l_serie_NF
        sNFAnteriorNumero = l_num_NF
        sNFAnteriorEmitente = l_emitente_NF
        fechar_modo_emissao_nfe_triangular
        End If
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_REMESSA_ENCERRA_POR_ERRO_CONSISTENCIA:
'=======================================
    aviso_erro s_erro
    GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_REMESSA_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub NFE_EMITE_REMESSA_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMITE_REMESSA_FECHA_TABELAS:
'=======================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_PEDIDO_ITEM_DEVOLVIDO, True
    bd_desaloca_recordset t_DESTINATARIO, True
    bd_desaloca_recordset t_TRANSPORTADORA, True
    bd_desaloca_recordset t_IBPT, True
    bd_desaloca_recordset t_NFe_EMITENTE_X_LOJA, True
    'bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset t_NFe_IMAGEM, True
    bd_desaloca_recordset t_T1_NFE_INUTILIZA, True
    bd_desaloca_recordset t_CODIGO_DESCRICAO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPEmite, True
  
  ' COMMAND
    bd_desaloca_command cmdNFeEmite
    bd_desaloca_command cmdNFeSituacao
    
  ' CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
    
    Return

End Sub

Function marca_status_atendido_fila_solicitacoes_emissao_NFe(ByVal pedido As String, _
                                                        ByVal intIdNfeEmitente As Integer, _
                                                        ByVal lngSerieNFe As Long, _
                                                        ByVal lngNumeroNFe As Long, _
                                                        ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "marca_status_atendido_fila_solicitacoes_emissao_NFe()"
' DECLARAÇÕES
Dim s As String
Dim strId As String
Dim lngRecordsAffected As Long
' BANCO DE DADOS
Dim t As ADODB.Recordset

    On Error GoTo MSAFSEN_TRATA_ERRO
    
    marca_status_atendido_fila_solicitacoes_emissao_NFe = False
    strMsgErro = ""
    
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    s = "SELECT" & _
            " id" & _
        " FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA" & _
        " WHERE" & _
            " (pedido = '" & pedido & "')" & _
            " AND (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")"
    t.Open s, dbc, , , adCmdText
    If t.EOF Then
        marca_status_atendido_fila_solicitacoes_emissao_NFe = True
        GoSub MSAFSEN_FECHA_TABELAS
        Exit Function
        End If
        
    strId = Trim$("" & t("id"))
    
    s = "UPDATE t_PEDIDO_NFe_EMISSAO_SOLICITADA SET" & _
            " nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__ATENDIDA & ", " & _
            " nfe_emitida_usuario = '" & usuario.id & "', " & _
            " nfe_emitida_data = " & sqlMontaGetdateSomenteData() & ", " & _
            " nfe_emitida_data_hora = getdate(), " & _
            " id_nfe_emitente = " & CStr(intIdNfeEmitente) & ", " & _
            " NFe_serie_NF = " & CStr(lngSerieNFe) & ", " & _
            " NFe_numero_NF = " & CStr(lngNumeroNFe) & _
        " WHERE" & _
            " (id = " & strId & ")" & _
            " AND (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")"
    dbc.Execute s, lngRecordsAffected
    If lngRecordsAffected = 1 Then
        marca_status_atendido_fila_solicitacoes_emissao_NFe = True
    Else
        strMsgErro = "Falha ao tentar assinalar o pedido " & pedido & " como já tratado na fila de solicitações de emissão de NFe!!"
        End If
    
    GoSub MSAFSEN_FECHA_TABELAS
            
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MSAFSEN_TRATA_ERRO:
'==================
    strMsgErro = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub MSAFSEN_FECHA_TABELAS
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MSAFSEN_FECHA_TABELAS:
'=====================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function

Sub trata_botao_fila_remove()

' CONSTANTES
Const NomeDestaRotina = "trata_botao_fila_remove()"
' DECLARAÇÕES
Dim s As String
Dim strId As String
Dim lngRecordsAffected As Long
' BANCO DE DADOS
Dim t As ADODB.Recordset

    On Error GoTo TBFR_TRATA_ERRO
    
    c_pedido = Trim$(c_pedido)
    c_pedido = normaliza_num_pedido(c_pedido)
    If c_pedido = "" Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT" & _
            " id" & _
        " FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA" & _
        " WHERE" & _
            " (pedido = '" & c_pedido & "')" & _
            " AND (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")"
    t.Open s, dbc, , , adCmdText
    If t.EOF Then
        GoSub TBFR_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        aviso_erro "Pedido " & c_pedido & " NÃO está na fila de solicitações de emissão de NFe!!"
        Exit Sub
        End If
        
    strId = Trim$("" & t("id"))
    
    s = "Remove o pedido " & c_pedido & " da fila de solicitações de emissão de NFe?"
    f_CONFIRMACAO_VIA_SENHA.strMensagemInformativa = s
    f_CONFIRMACAO_VIA_SENHA.strSenhaCorreta = usuario.senha
    f_CONFIRMACAO_VIA_SENHA.Show vbModal, Me
    If Not f_CONFIRMACAO_VIA_SENHA.blnResultadoFormOk Then
        GoSub TBFR_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        aviso "Operação cancelada!!"
        Exit Sub
        End If
        
    s = "UPDATE t_PEDIDO_NFe_EMISSAO_SOLICITADA SET" & _
            " nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__CANCELADA & ", " & _
            " nfe_emitida_usuario = '" & usuario.id & "', " & _
            " nfe_emitida_data = " & sqlMontaGetdateSomenteData() & ", " & _
            " nfe_emitida_data_hora = getdate()" & _
        " WHERE" & _
            " (id = " & strId & ")" & _
            " AND (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")"
    dbc.Execute s, lngRecordsAffected
    If lngRecordsAffected = 0 Then
        GoSub TBFR_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        aviso_erro "Falha ao tentar remover o pedido " & c_pedido & " da fila de solicitações de emissão de NFe!!"
        Exit Sub
        End If
    
    GoSub TBFR_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
    aviso "Pedido " & c_pedido & " removido com sucesso da fila de solicitações de emissão de NFe!!"
    
    sPedidoTriangular = ""
    'fechar_modo_emissao_nfe_triangular
    Unload f_EMISSAO_NFE_TRIANGULAR
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBFR_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aviso_erro s
    GoSub TBFR_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBFR_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    

End Sub
