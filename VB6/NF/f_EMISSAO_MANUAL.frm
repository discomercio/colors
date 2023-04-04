VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form f_EMISSAO_MANUAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal"
   ClientHeight    =   12780
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   20655
   Icon            =   "f_EMISSAO_MANUAL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12780
   ScaleWidth      =   20655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pnDanfe 
      Caption         =   "DANFE"
      Height          =   915
      Left            =   120
      TabIndex        =   111
      Top             =   11670
      Width           =   14520
      Begin VB.ComboBox cb_emitente_danfe 
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
         ItemData        =   "f_EMISSAO_MANUAL.frx":0442
         Left            =   150
         List            =   "f_EMISSAO_MANUAL.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   435
         Width           =   7290
      End
      Begin VB.TextBox c_num_nfe_danfe 
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
         Left            =   9735
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Top             =   420
         Width           =   1650
      End
      Begin VB.CommandButton b_danfe 
         Caption         =   "D&ANFE"
         Height          =   450
         Left            =   11820
         TabIndex        =   99
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox c_num_serie_danfe 
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
         Left            =   8070
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   435
         Width           =   1170
      End
      Begin VB.Label l_tit_emitente_danfe 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   165
         TabIndex        =   125
         Top             =   225
         Width           =   615
      End
      Begin VB.Label l_tit_num_nfe_danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   9750
         TabIndex        =   114
         Top             =   225
         Width           =   525
      End
      Begin VB.Label l_tit_num_serie_Danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   8085
         TabIndex        =   112
         Top             =   225
         Width           =   585
      End
   End
   Begin VB.Frame pnNumeroNFe 
      Caption         =   "Última NFe emitida"
      Height          =   705
      Left            =   120
      TabIndex        =   107
      Top             =   10860
      Width           =   14520
      Begin VB.Label l_tit_emitente_NF 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   6060
         TabIndex        =   110
         Top             =   345
         Width           =   615
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
         Left            =   6840
         TabIndex        =   95
         Top             =   300
         Width           =   7485
      End
      Begin VB.Label l_tit_serie_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   480
         TabIndex        =   109
         Top             =   345
         Width           =   585
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
         Left            =   1230
         TabIndex        =   93
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label l_tit_num_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   3000
         TabIndex        =   108
         Top             =   345
         Width           =   525
      End
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
         Left            =   3690
         TabIndex        =   94
         Top             =   300
         Width           =   1830
      End
   End
   Begin VB.Frame pnParcelasEmBoletos 
      Caption         =   "Parcelas em Boletos"
      Height          =   4695
      Left            =   14760
      TabIndex        =   350
      Top             =   7920
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
         Left            =   360
         TabIndex        =   356
         Top             =   3120
         Width           =   945
      End
      Begin VB.CommandButton b_parc_edicao_ok 
         Height          =   390
         Left            =   360
         Picture         =   "f_EMISSAO_MANUAL.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   355
         Top             =   4005
         Width           =   690
      End
      Begin VB.CommandButton b_parc_edicao_cancela 
         Height          =   390
         Left            =   1560
         Picture         =   "f_EMISSAO_MANUAL.frx":0698
         Style           =   1  'Graphical
         TabIndex        =   354
         Top             =   4005
         Width           =   690
      End
      Begin VB.CommandButton b_recalculaparc 
         Caption         =   "&Reagendar Parcelas Seguintes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         TabIndex        =   353
         Top             =   3960
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
         TabIndex        =   352
         Top             =   3120
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
         TabIndex        =   351
         Top             =   3120
         Width           =   1260
      End
      Begin MSComctlLib.ListView lvParcBoletos 
         Height          =   2415
         Left            =   120
         TabIndex        =   357
         Top             =   360
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   4260
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
         TabIndex        =   360
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label l_tit_dataparc 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   1560
         TabIndex        =   359
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label l_tit_numparc 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   360
         TabIndex        =   358
         Top             =   2880
         Width           =   540
      End
   End
   Begin VB.Frame pn_info_pagamento 
      Caption         =   "Informações de Pagamento"
      Height          =   975
      Left            =   9960
      TabIndex        =   293
      Top             =   7680
      Width           =   4335
      Begin VB.ComboBox cb_meio_pagto 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   348
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cb_forma_pagto 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   346
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label l_meio_pato 
         Caption         =   "Meio de Pagamento"
         Height          =   255
         Left            =   2280
         TabIndex        =   347
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label l_forma_pato 
         Caption         =   "Forma de Pagamento"
         Height          =   255
         Left            =   240
         TabIndex        =   345
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame pn_pedido_nota 
      Height          =   1875
      Left            =   120
      TabIndex        =   287
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton opVendaFutura 
         Caption         =   "Como Venda Futura"
         Height          =   375
         Left            =   120
         TabIndex        =   349
         Top             =   1320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opVenda 
         Caption         =   "Como Venda"
         Height          =   255
         Left            =   120
         TabIndex        =   289
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton opRemessa 
         Caption         =   "Como Remessa"
         Height          =   255
         Left            =   120
         TabIndex        =   288
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl_pedido_nota 
         Caption         =   "Gravar NF no pedido"
         Height          =   375
         Left            =   120
         TabIndex        =   290
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox c_pedido_nota 
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
      Height          =   375
      Left            =   120
      MaxLength       =   9
      TabIndex        =   286
      Top             =   7320
      Width           =   1815
   End
   Begin VB.PictureBox picEndereco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   17160
      MousePointer    =   14  'Arrow and Question
      Picture         =   "f_EMISSAO_MANUAL.frx":0B0B
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   280
      ToolTipText     =   "Clique para visualizar o endereço editado"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton b_assistencia_tecnica 
      Caption         =   "Pedido Assist &Técnica"
      Height          =   495
      Left            =   12270
      TabIndex        =   238
      Top             =   9600
      Width           =   2115
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
      ItemData        =   "f_EMISSAO_MANUAL.frx":1241
      Left            =   120
      List            =   "f_EMISSAO_MANUAL.frx":1243
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2250
      Width           =   8301
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
      Left            =   11265
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "f_EMISSAO_MANUAL.frx":1245
      Top             =   2250
      Width           =   9015
   End
   Begin VB.ComboBox cb_indpres 
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
      Left            =   3120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   930
      Width           =   2820
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
      Left            =   12675
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1560
      Width           =   2460
   End
   Begin VB.CommandButton b_editar_endereco 
      Caption         =   "Editar E&ndereço"
      Height          =   450
      Left            =   12270
      TabIndex        =   233
      Top             =   9000
      Width           =   2115
   End
   Begin VB.Frame pnZerarAliquotas 
      Height          =   1830
      Left            =   15360
      TabIndex        =   215
      Top             =   120
      Width           =   4810
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
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   217
         Top             =   450
         Width           =   4555
      End
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
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   216
         Top             =   1305
         Width           =   4555
      End
      Begin VB.Label l_tit_zerar_PIS 
         AutoSize        =   -1  'True
         Caption         =   "Zerar PIS"
         Height          =   195
         Left            =   150
         TabIndex        =   219
         Top             =   240
         Width           =   675
      End
      Begin VB.Label l_tit_zerar_COFINS 
         AutoSize        =   -1  'True
         Caption         =   "Zerar COFINS"
         Height          =   195
         Left            =   150
         TabIndex        =   218
         Top             =   1095
         Width           =   1005
      End
   End
   Begin VB.CommandButton b_edicao 
      Caption         =   "Li&berar Edição"
      Height          =   450
      Left            =   9915
      TabIndex        =   184
      Top             =   9600
      Width           =   2115
   End
   Begin VB.CommandButton b_emite_numeracao_manual 
      Caption         =   "Emitir NFe (Nº &Manual)"
      Height          =   450
      Left            =   7560
      TabIndex        =   89
      Top             =   9600
      Width           =   2115
   End
   Begin VB.CommandButton b_emissao_automatica 
      Caption         =   "Painel Emissão &Automática"
      Height          =   450
      Left            =   9915
      TabIndex        =   91
      Top             =   9000
      Width           =   2115
   End
   Begin VB.ComboBox cb_transportadora 
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
      ItemData        =   "f_EMISSAO_MANUAL.frx":1272
      Left            =   7350
      List            =   "f_EMISSAO_MANUAL.frx":1274
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   5130
   End
   Begin VB.Frame pnItens 
      Caption         =   "Itens"
      Height          =   4290
      Left            =   120
      TabIndex        =   117
      Top             =   2760
      Width           =   20250
      Begin VB.TextBox c_vl_total_icms 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15660
         Locked          =   -1  'True
         TabIndex        =   278
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1425
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
         TabIndex        =   276
         Top             =   3600
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
         TabIndex        =   275
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
         Index           =   9
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   274
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
         Index           =   8
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   273
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
         Index           =   7
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   272
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
         Index           =   6
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   271
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
         Index           =   5
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   270
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
         Index           =   4
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   269
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
         Index           =   3
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   268
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
         Index           =   2
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   267
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
         Index           =   1
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   266
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
         Index           =   0
         Left            =   19530
         MaxLength       =   6
         TabIndex        =   265
         Top             =   465
         Width           =   525
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   0
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   263
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   1
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   262
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   2
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   261
         Top             =   1035
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   3
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   260
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   4
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   259
         Top             =   1605
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   5
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   258
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   6
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   257
         Top             =   2175
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   7
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   256
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   8
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   255
         Top             =   2745
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   9
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   254
         Top             =   3030
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   10
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   253
         Top             =   3315
         Width           =   735
      End
      Begin VB.TextBox c_nItemPed 
         Height          =   285
         Index           =   11
         Left            =   18050
         MaxLength       =   6
         TabIndex        =   252
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   11
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   251
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   10
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   250
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   9
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   249
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   8
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   248
         Top             =   2745
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   7
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   247
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   6
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   246
         Top             =   2175
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   5
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   245
         Top             =   1890
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   4
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   244
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   3
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   243
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   2
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   242
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   1
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   241
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox c_unidade 
         Height          =   285
         Index           =   0
         Left            =   18850
         MaxLength       =   6
         TabIndex        =   240
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   231
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   230
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   229
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   228
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   227
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   226
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   225
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   224
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   223
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   222
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   221
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_xPed 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   17125
         MaxLength       =   15
         TabIndex        =   220
         Top             =   480
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
         Index           =   11
         Left            =   16540
         TabIndex        =   202
         Top             =   3600
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
         Left            =   16540
         TabIndex        =   203
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
         Index           =   9
         Left            =   16540
         TabIndex        =   204
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
         Index           =   8
         Left            =   16540
         TabIndex        =   205
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
         Index           =   7
         Left            =   16540
         TabIndex        =   206
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
         Index           =   6
         Left            =   16540
         TabIndex        =   207
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
         Index           =   5
         Left            =   16540
         TabIndex        =   208
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
         Index           =   4
         Left            =   16540
         TabIndex        =   209
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
         Index           =   3
         Left            =   16540
         TabIndex        =   210
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
         Index           =   2
         Left            =   16540
         TabIndex        =   211
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
         Index           =   1
         Left            =   16540
         TabIndex        =   212
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
         Index           =   0
         Left            =   16540
         TabIndex        =   213
         Top             =   465
         Width           =   585
      End
      Begin VB.TextBox c_total_peso_liquido 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8185
         MaxLength       =   15
         TabIndex        =   200
         Top             =   3885
         Width           =   1095
      End
      Begin VB.TextBox c_total_peso_bruto 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5260
         MaxLength       =   15
         TabIndex        =   198
         Top             =   3885
         Width           =   1095
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   196
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   195
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   194
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   193
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   192
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   191
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   190
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   189
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   188
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   187
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   186
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_NCM 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   15655
         MaxLength       =   8
         TabIndex        =   185
         Top             =   465
         Width           =   885
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   11
         ItemData        =   "f_EMISSAO_MANUAL.frx":1276
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1278
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Top             =   3600
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   10
         ItemData        =   "f_EMISSAO_MANUAL.frx":127A
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":127C
         Style           =   2  'Dropdown List
         TabIndex        =   181
         Top             =   3315
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   9
         ItemData        =   "f_EMISSAO_MANUAL.frx":127E
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1280
         Style           =   2  'Dropdown List
         TabIndex        =   180
         Top             =   3030
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   8
         ItemData        =   "f_EMISSAO_MANUAL.frx":1282
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1284
         Style           =   2  'Dropdown List
         TabIndex        =   179
         Top             =   2745
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   7
         ItemData        =   "f_EMISSAO_MANUAL.frx":1286
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1288
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Top             =   2460
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   6
         ItemData        =   "f_EMISSAO_MANUAL.frx":128A
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":128C
         Style           =   2  'Dropdown List
         TabIndex        =   177
         Top             =   2175
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   5
         ItemData        =   "f_EMISSAO_MANUAL.frx":128E
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1290
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   1890
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   4
         ItemData        =   "f_EMISSAO_MANUAL.frx":1292
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1294
         Style           =   2  'Dropdown List
         TabIndex        =   175
         Top             =   1605
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   3
         ItemData        =   "f_EMISSAO_MANUAL.frx":1296
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":1298
         Style           =   2  'Dropdown List
         TabIndex        =   174
         Top             =   1320
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   2
         ItemData        =   "f_EMISSAO_MANUAL.frx":129A
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":129C
         Style           =   2  'Dropdown List
         TabIndex        =   173
         Top             =   1035
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   1
         ItemData        =   "f_EMISSAO_MANUAL.frx":129E
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":12A0
         Style           =   2  'Dropdown List
         TabIndex        =   172
         Top             =   750
         Width           =   1985
      End
      Begin VB.ComboBox cb_CFOP 
         Height          =   315
         Index           =   0
         ItemData        =   "f_EMISSAO_MANUAL.frx":12A2
         Left            =   13670
         List            =   "f_EMISSAO_MANUAL.frx":12A4
         Style           =   2  'Dropdown List
         TabIndex        =   171
         Top             =   465
         Width           =   1985
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   169
         Top             =   465
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   168
         Top             =   750
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   167
         Top             =   1035
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   166
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   165
         Top             =   1605
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   164
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   163
         Top             =   2175
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   162
         Top             =   2460
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   161
         Top             =   2745
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   160
         Top             =   3030
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   159
         Top             =   3315
         Width           =   525
      End
      Begin VB.TextBox c_CST 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   13145
         MaxLength       =   4
         TabIndex        =   158
         Top             =   3600
         Width           =   525
      End
      Begin VB.TextBox c_vl_total_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11840
         Locked          =   -1  'True
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   155
         Top             =   465
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   154
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   153
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   152
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   151
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   150
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   149
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   148
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   147
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   146
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   145
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_vl_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   11840
         MaxLength       =   18
         TabIndex        =   144
         Top             =   3600
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   11
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   137
         Top             =   3600
         Width           =   1705
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   139
         Top             =   3600
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   140
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   141
         Top             =   3600
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   705
         MaxLength       =   8
         TabIndex        =   142
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   180
         MaxLength       =   4
         TabIndex        =   143
         Top             =   3600
         Width           =   525
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   180
         MaxLength       =   4
         TabIndex        =   136
         Top             =   3315
         Width           =   525
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   705
         MaxLength       =   8
         TabIndex        =   135
         Top             =   3315
         Width           =   885
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   134
         Top             =   3315
         Width           =   5320
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   133
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   132
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   3315
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   10
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   130
         Top             =   3315
         Width           =   1705
      End
      Begin VB.TextBox c_total_volumes 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   128
         Top             =   3885
         Width           =   735
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   9
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   79
         Top             =   3030
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   8
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   72
         Top             =   2745
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   7
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   65
         Top             =   2460
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   6
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   58
         Top             =   2175
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   5
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   51
         Top             =   1890
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   4
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   44
         Top             =   1605
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   3
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   37
         Top             =   1320
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   2
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   30
         Top             =   1035
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   1
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   23
         Top             =   750
         Width           =   1705
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   0
         Left            =   6910
         MaxLength       =   500
         TabIndex        =   16
         Top             =   465
         Width           =   1705
      End
      Begin VB.TextBox c_vl_total_geral 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10535
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3885
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   81
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   80
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   78
         Top             =   3030
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   705
         MaxLength       =   8
         TabIndex        =   77
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   180
         MaxLength       =   4
         TabIndex        =   76
         Top             =   3030
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   74
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   73
         Top             =   2745
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   71
         Top             =   2745
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   705
         MaxLength       =   8
         TabIndex        =   70
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   180
         MaxLength       =   4
         TabIndex        =   69
         Top             =   2745
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   67
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   66
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   64
         Top             =   2460
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   705
         MaxLength       =   8
         TabIndex        =   63
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   180
         MaxLength       =   4
         TabIndex        =   62
         Top             =   2460
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   60
         Top             =   2175
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   59
         Top             =   2175
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   57
         Top             =   2175
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   705
         MaxLength       =   8
         TabIndex        =   56
         Top             =   2175
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   180
         MaxLength       =   4
         TabIndex        =   55
         Top             =   2175
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   53
         Top             =   1890
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   52
         Top             =   1890
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   50
         Top             =   1890
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   705
         MaxLength       =   8
         TabIndex        =   49
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   180
         MaxLength       =   4
         TabIndex        =   48
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   46
         Top             =   1605
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   45
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   43
         Top             =   1605
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   705
         MaxLength       =   8
         TabIndex        =   42
         Top             =   1605
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   180
         MaxLength       =   4
         TabIndex        =   41
         Top             =   1605
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   39
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   36
         Top             =   1320
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   705
         MaxLength       =   8
         TabIndex        =   35
         Top             =   1320
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   180
         MaxLength       =   4
         TabIndex        =   34
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   32
         Top             =   1035
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   31
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   29
         Top             =   1035
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   705
         MaxLength       =   8
         TabIndex        =   28
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   180
         MaxLength       =   4
         TabIndex        =   27
         Top             =   1035
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   25
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   24
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   22
         Top             =   750
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   705
         MaxLength       =   8
         TabIndex        =   21
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   180
         MaxLength       =   4
         TabIndex        =   20
         Top             =   750
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   10535
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   9230
         MaxLength       =   18
         TabIndex        =   18
         Top             =   465
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   8615
         MaxLength       =   5
         TabIndex        =   17
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   15
         Top             =   465
         Width           =   5320
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   705
         MaxLength       =   8
         TabIndex        =   14
         Top             =   465
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   180
         MaxLength       =   4
         TabIndex        =   13
         Top             =   465
         Width           =   525
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
         Left            =   14640
         TabIndex        =   279
         Top             =   3930
         Width           =   960
      End
      Begin VB.Label l_tit_FCP 
         AutoSize        =   -1  'True
         Caption         =   "%FCP"
         Height          =   195
         Left            =   19530
         TabIndex        =   277
         Top             =   240
         Width           =   420
      End
      Begin VB.Label l_tit_nItemPed 
         AutoSize        =   -1  'True
         Caption         =   "nItemPed"
         Height          =   195
         Left            =   18050
         TabIndex        =   264
         Top             =   255
         Width           =   675
      End
      Begin VB.Label l_tit_unidade 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Left            =   18850
         TabIndex        =   239
         Top             =   255
         Width           =   600
      End
      Begin VB.Label l_tit_xPed 
         AutoSize        =   -1  'True
         Caption         =   "xPed"
         Height          =   195
         Left            =   17140
         TabIndex        =   232
         Top             =   255
         Width           =   360
      End
      Begin VB.Label l_tit_ICMS_item 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
         Height          =   195
         Left            =   16555
         TabIndex        =   214
         Top             =   255
         Width           =   390
      End
      Begin VB.Label l_tit_total_peso_liquido 
         AutoSize        =   -1  'True
         Caption         =   "Peso Líq (kg)"
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
         Left            =   6910
         TabIndex        =   201
         Top             =   3930
         Width           =   1170
      End
      Begin VB.Label l_tit_total_peso_bruto 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto (kg)"
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
         Left            =   3820
         TabIndex        =   199
         Top             =   3930
         Width           =   1335
      End
      Begin VB.Label l_tit_NCM 
         AutoSize        =   -1  'True
         Caption         =   "NCM"
         Height          =   195
         Left            =   15670
         TabIndex        =   197
         Top             =   255
         Width           =   360
      End
      Begin VB.Label l_tit_CFOP 
         AutoSize        =   -1  'True
         Caption         =   "CFOP"
         Height          =   195
         Left            =   13685
         TabIndex        =   183
         Top             =   255
         Width           =   420
      End
      Begin VB.Label l_tit_CST 
         AutoSize        =   -1  'True
         Caption         =   "CST"
         Height          =   195
         Left            =   13250
         TabIndex        =   170
         Top             =   255
         Width           =   315
      End
      Begin VB.Label l_tit_vl_outras_despesas_acessorias 
         AutoSize        =   -1  'True
         Caption         =   "Desp Acessórias"
         Height          =   195
         Left            =   11945
         TabIndex        =   156
         Top             =   255
         Width           =   1185
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
         Left            =   765
         TabIndex        =   129
         Top             =   3930
         Width           =   720
      End
      Begin VB.Label l_tit_produto_obs 
         AutoSize        =   -1  'True
         Caption         =   "Informações Adicionais"
         Height          =   195
         Left            =   6925
         TabIndex        =   127
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
         Left            =   9980
         TabIndex        =   124
         Top             =   3930
         Width           =   450
      End
      Begin VB.Label l_tit_vl_total 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total"
         Height          =   195
         Left            =   11060
         TabIndex        =   123
         Top             =   255
         Width           =   765
      End
      Begin VB.Label l_tit_vl_unitario 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário"
         Height          =   195
         Left            =   9575
         TabIndex        =   122
         Top             =   255
         Width           =   945
      End
      Begin VB.Label l_tit_qtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   8750
         TabIndex        =   121
         Top             =   255
         Width           =   345
      End
      Begin VB.Label l_tit_descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1605
         TabIndex        =   120
         Top             =   255
         Width           =   720
      End
      Begin VB.Label l_tit_produto 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   195
         Left            =   720
         TabIndex        =   119
         Top             =   255
         Width           =   555
      End
      Begin VB.Label l_tit_fabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabric"
         Height          =   195
         Left            =   195
         TabIndex        =   118
         Top             =   255
         Width           =   435
      End
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
      Left            =   9120
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   330
      Width           =   5250
   End
   Begin VB.ComboBox cb_emitente 
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
      ItemData        =   "f_EMISSAO_MANUAL.frx":12A6
      Left            =   120
      List            =   "f_EMISSAO_MANUAL.frx":12A8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   5820
   End
   Begin VB.ComboBox cb_tipo_NF 
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
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   930
      Width           =   2820
   End
   Begin VB.CommandButton b_fechar 
      Caption         =   "&Fechar"
      Height          =   450
      Left            =   9915
      TabIndex        =   92
      Top             =   10215
      Width           =   2115
   End
   Begin VB.CommandButton b_destinatario 
      Caption         =   "&Dados do Destinatário"
      Height          =   450
      Left            =   7560
      TabIndex        =   90
      Top             =   10215
      Width           =   2115
   End
   Begin VB.Timer relogio 
      Interval        =   1000
      Left            =   3225
      Top             =   6945
   End
   Begin VB.TextBox c_dados_adicionais 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   84
      Top             =   7320
      Width           =   5103
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
      Left            =   2730
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   4410
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
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1560
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
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
      ItemData        =   "f_EMISSAO_MANUAL.frx":12AA
      Left            =   6630
      List            =   "f_EMISSAO_MANUAL.frx":12AC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   8490
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
      Left            =   6630
      MaxLength       =   18
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   330
      Width           =   2370
   End
   Begin VB.CommandButton b_imprime 
      Caption         =   "&Emitir NFe"
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
      Left            =   7560
      TabIndex        =   88
      Top             =   9000
      Width           =   2115
   End
   Begin VB.CommandButton b_dummy 
      Appearance      =   0  'Flat
      Caption         =   "b_dummy"
      Height          =   345
      Left            =   5565
      TabIndex        =   100
      Top             =   -525
      Width           =   1350
   End
   Begin VB.Frame pn_aviso_pedido_nota 
      Enabled         =   0   'False
      Height          =   1995
      Left            =   120
      TabIndex        =   291
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label lbl_aviso_pedido_nota 
         Caption         =   "O   número  da  Nota Fiscal   deverá    ser digitado   no  campo  'Nº  Nota  Fiscal' ou 'NF Simples Remessa' do   pedido"
         Height          =   1215
         Left            =   120
         TabIndex        =   292
         Top             =   240
         Width           =   1575
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame pn_endereco 
      Caption         =   "Endereço"
      Height          =   5580
      Left            =   105
      TabIndex        =   294
      Top             =   7125
      Width           =   20490
      Begin VB.Frame pn_endereco_cadastro 
         Caption         =   "Endereço do Cadastro"
         Height          =   900
         Left            =   690
         TabIndex        =   330
         Top             =   360
         Width           =   18420
         Begin VB.Label l_tit_end_cadastro_logradouro 
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
            TabIndex        =   344
            Top             =   285
            Width           =   885
         End
         Begin VB.Label l_end_cadastro_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Rua do João da Silva"
            Height          =   195
            Left            =   1155
            TabIndex        =   343
            Top             =   285
            Width           =   1530
         End
         Begin VB.Label l_tit_end_cadastro_numero 
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
            TabIndex        =   342
            Top             =   285
            Width           =   285
         End
         Begin VB.Label l_end_cadastro_numero 
            AutoSize        =   -1  'True
            Caption         =   "999"
            Height          =   195
            Left            =   10170
            TabIndex        =   341
            Top             =   285
            Width           =   270
         End
         Begin VB.Label l_tit_end_cadastro_complemento 
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
            TabIndex        =   340
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label l_end_cadastro_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Apartamento 99"
            Height          =   195
            Left            =   14220
            TabIndex        =   339
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label l_tit_end_cadastro_bairro 
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
            TabIndex        =   338
            Top             =   585
            Width           =   570
         End
         Begin VB.Label l_end_cadastro_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Vila dos Testadores"
            Height          =   195
            Left            =   1155
            TabIndex        =   337
            Top             =   585
            Width           =   1395
         End
         Begin VB.Label l_tit_end_cadastro_cidade 
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
            TabIndex        =   336
            Top             =   585
            Width           =   660
         End
         Begin VB.Label l_end_cadastro_cidade 
            AutoSize        =   -1  'True
            Caption         =   "São Paulo"
            Height          =   195
            Left            =   8625
            TabIndex        =   335
            Top             =   585
            Width           =   735
         End
         Begin VB.Label l_tit_end_cadastro_uf 
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
            TabIndex        =   334
            Top             =   585
            Width           =   315
         End
         Begin VB.Label l_end_cadastro_uf 
            AutoSize        =   -1  'True
            Caption         =   "SP"
            Height          =   195
            Left            =   14220
            TabIndex        =   333
            Top             =   585
            Width           =   210
         End
         Begin VB.Label l_tit_end_cadastro_cep 
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
            TabIndex        =   332
            Top             =   585
            Width           =   435
         End
         Begin VB.Label l_end_cadastro_cep 
            AutoSize        =   -1  'True
            Caption         =   "00000-000"
            Height          =   195
            Left            =   16050
            TabIndex        =   331
            Top             =   585
            Width           =   765
         End
      End
      Begin VB.Frame pn_endereco_editado 
         Caption         =   "Endereço Editado"
         Height          =   900
         Left            =   690
         TabIndex        =   314
         Top             =   1410
         Width           =   18420
         Begin VB.CommandButton b_end_editado_exclui 
            Height          =   390
            Left            =   17340
            Picture         =   "f_EMISSAO_MANUAL.frx":12AE
            Style           =   1  'Graphical
            TabIndex        =   315
            Top             =   465
            Width           =   810
         End
         Begin VB.Label l_end_editado_cep 
            AutoSize        =   -1  'True
            Caption         =   "00000-000"
            Height          =   195
            Left            =   16050
            TabIndex        =   329
            Top             =   585
            Width           =   765
         End
         Begin VB.Label l_tit_end_editado_cep 
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
            TabIndex        =   328
            Top             =   585
            Width           =   435
         End
         Begin VB.Label l_end_editado_uf 
            AutoSize        =   -1  'True
            Caption         =   "SP"
            Height          =   195
            Left            =   14220
            TabIndex        =   327
            Top             =   585
            Width           =   210
         End
         Begin VB.Label l_tit_end_editado_uf 
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
            TabIndex        =   326
            Top             =   585
            Width           =   315
         End
         Begin VB.Label l_end_editado_cidade 
            AutoSize        =   -1  'True
            Caption         =   "São Paulo"
            Height          =   195
            Left            =   8625
            TabIndex        =   325
            Top             =   585
            Width           =   735
         End
         Begin VB.Label l_tit_end_editado_cidade 
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
            TabIndex        =   324
            Top             =   585
            Width           =   660
         End
         Begin VB.Label l_end_editado_bairro 
            AutoSize        =   -1  'True
            Caption         =   "Vila dos Testadores"
            Height          =   195
            Left            =   1155
            TabIndex        =   323
            Top             =   585
            Width           =   1395
         End
         Begin VB.Label l_tit_end_editado_bairro 
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
            TabIndex        =   322
            Top             =   585
            Width           =   570
         End
         Begin VB.Label l_end_editado_complemento 
            AutoSize        =   -1  'True
            Caption         =   "Apartamento 99"
            Height          =   195
            Left            =   14220
            TabIndex        =   321
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label l_tit_end_editado_complemento 
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
            TabIndex        =   320
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label l_end_editado_numero 
            AutoSize        =   -1  'True
            Caption         =   "999"
            Height          =   195
            Left            =   10170
            TabIndex        =   319
            Top             =   285
            Width           =   270
         End
         Begin VB.Label l_tit_end_editado_numero 
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
            TabIndex        =   318
            Top             =   285
            Width           =   285
         End
         Begin VB.Label l_end_editado_logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Rua do João da Silva"
            Height          =   195
            Left            =   1155
            TabIndex        =   317
            Top             =   285
            Width           =   1530
         End
         Begin VB.Label l_tit_end_editado_logradouro 
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
            TabIndex        =   316
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.Frame pn_endereco_edicao 
         Caption         =   "Edição do Endereço"
         Height          =   1335
         Left            =   690
         TabIndex        =   295
         Top             =   2460
         Width           =   18420
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
            TabIndex        =   306
            Text            =   "00000-000"
            Top             =   300
            Width           =   1320
         End
         Begin VB.CommandButton b_cep_pesquisar 
            Height          =   390
            Left            =   2535
            Picture         =   "f_EMISSAO_MANUAL.frx":144B
            Style           =   1  'Graphical
            TabIndex        =   305
            Top             =   255
            Width           =   810
         End
         Begin VB.TextBox c_end_edicao_logradouro 
            Height          =   285
            Left            =   4800
            MaxLength       =   60
            TabIndex        =   304
            Text            =   "Rua do João da Silva"
            Top             =   300
            Width           =   5730
         End
         Begin VB.TextBox c_end_edicao_numero 
            Height          =   285
            Left            =   11520
            MaxLength       =   60
            TabIndex        =   303
            Text            =   "999"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox c_end_edicao_complemento 
            Height          =   285
            Left            =   14760
            MaxLength       =   60
            TabIndex        =   302
            Text            =   "Apartamento 99"
            Top             =   300
            Width           =   3390
         End
         Begin VB.TextBox c_end_edicao_bairro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   301
            Text            =   "Rua do João da Silva"
            Top             =   900
            Width           =   5730
         End
         Begin VB.TextBox c_end_edicao_cidade 
            Height          =   285
            Left            =   8115
            MaxLength       =   60
            TabIndex        =   300
            Text            =   "São Paulo"
            Top             =   900
            Width           =   5730
         End
         Begin VB.TextBox c_end_edicao_uf 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   14760
            MaxLength       =   2
            TabIndex        =   299
            Text            =   "SP"
            Top             =   900
            Width           =   585
         End
         Begin VB.CommandButton b_end_edicao_cancela 
            Height          =   390
            Left            =   17340
            Picture         =   "f_EMISSAO_MANUAL.frx":169D
            Style           =   1  'Graphical
            TabIndex        =   298
            Top             =   855
            Width           =   810
         End
         Begin VB.CommandButton b_end_edicao_ok 
            Height          =   390
            Left            =   16470
            Picture         =   "f_EMISSAO_MANUAL.frx":1B10
            Style           =   1  'Graphical
            TabIndex        =   297
            Top             =   855
            Width           =   810
         End
         Begin VB.CommandButton b_end_edicao_limpa 
            Height          =   390
            Left            =   15600
            Picture         =   "f_EMISSAO_MANUAL.frx":1D62
            Style           =   1  'Graphical
            TabIndex        =   296
            Top             =   855
            Width           =   810
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
            TabIndex        =   313
            Top             =   345
            Width           =   375
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
            TabIndex        =   312
            Top             =   345
            Width           =   825
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
            TabIndex        =   311
            Top             =   345
            Width           =   225
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
            TabIndex        =   310
            Top             =   345
            Width           =   1140
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
            TabIndex        =   309
            Top             =   945
            Width           =   510
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
            TabIndex        =   308
            Top             =   945
            Width           =   600
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
            TabIndex        =   307
            Top             =   945
            Width           =   255
         End
      End
   End
   Begin VB.Label l_pedido_nota 
      Caption         =   "Pedido da Nota"
      Height          =   255
      Left            =   120
      TabIndex        =   285
      Top             =   7110
      Width           =   1335
   End
   Begin VB.Label l_tit_IE 
      AutoSize        =   -1  'True
      Caption         =   "IE"
      Height          =   195
      Left            =   14760
      TabIndex        =   284
      Top             =   120
      Width           =   150
   End
   Begin VB.Label l_IE 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14520
      TabIndex        =   283
      Top             =   360
      Width           =   585
   End
   Begin VB.Label l_tit_emitente_uf 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UF do Emitente"
      Height          =   390
      Left            =   8760
      TabIndex        =   282
      Top             =   2160
      Width           =   855
      WordWrap        =   -1  'True
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
      Height          =   600
      Left            =   9720
      TabIndex        =   281
      Top             =   2020
      Width           =   825
   End
   Begin VB.Label l_tit_finalidade 
      AutoSize        =   -1  'True
      Caption         =   "Finalidade"
      Height          =   195
      Left            =   135
      TabIndex        =   237
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label l_tit_chave_nfe_ref 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso NFe Referenciada"
      Height          =   195
      Left            =   11280
      TabIndex        =   236
      Top             =   2040
      Width           =   2610
   End
   Begin VB.Label l_tit_indpres 
      AutoSize        =   -1  'True
      Caption         =   "Indicador de Presença"
      Height          =   195
      Left            =   3135
      TabIndex        =   235
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label l_tit_loc_dest 
      AutoSize        =   -1  'True
      Caption         =   "Local de Destino da Operação"
      Height          =   195
      Left            =   12675
      TabIndex        =   234
      Top             =   1350
      Width           =   2175
   End
   Begin VB.Label l_tit_transportadora 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora"
      Height          =   195
      Left            =   7365
      TabIndex        =   126
      Top             =   1350
      Width           =   1080
   End
   Begin VB.Label l_tit_nome_dest 
      AutoSize        =   -1  'True
      Caption         =   "Nome Destinatário"
      Height          =   195
      Left            =   9165
      TabIndex        =   116
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label l_tit_emitente 
      AutoSize        =   -1  'True
      Caption         =   "Emitente"
      Height          =   195
      Left            =   135
      TabIndex        =   115
      Top             =   90
      Width           =   615
   End
   Begin VB.Label l_tit_tipo_NF 
      AutoSize        =   -1  'True
      Caption         =   "Tipo do Documento Fiscal"
      Height          =   195
      Left            =   135
      TabIndex        =   113
      Top             =   720
      Width           =   1860
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
      Left            =   7620
      TabIndex        =   87
      Top             =   8925
      Width           =   1980
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
      Left            =   7620
      TabIndex        =   86
      Top             =   8520
      Width           =   1980
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
      Height          =   1080
      Left            =   7620
      TabIndex        =   85
      Top             =   7320
      Width           =   1980
      WordWrap        =   -1  'True
   End
   Begin VB.Label l_tit_dados_adicionais 
      AutoSize        =   -1  'True
      Caption         =   "Dados Adicionais"
      Height          =   195
      Left            =   2055
      TabIndex        =   106
      Top             =   7110
      Width           =   1230
   End
   Begin VB.Label l_tit_frete 
      AutoSize        =   -1  'True
      Caption         =   "Frete por Conta"
      Height          =   195
      Left            =   2745
      TabIndex        =   105
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label l_tit_aliquota_IPI 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota IPI"
      Height          =   195
      Left            =   1440
      TabIndex        =   104
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label l_tit_aliquota_icms 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota ICMS"
      Height          =   195
      Left            =   135
      TabIndex        =   103
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label l_tit_natureza 
      AutoSize        =   -1  'True
      Caption         =   "Natureza da Operação"
      Height          =   195
      Left            =   6645
      TabIndex        =   102
      Top             =   720
      Width           =   1620
   End
   Begin VB.Label l_tit_cnpj_cpf_dest 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ/CPF Destinatário"
      Height          =   195
      Left            =   6645
      TabIndex        =   101
      Top             =   120
      Width           =   1665
   End
   Begin VB.Menu mnu_ARQUIVO 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_emissao_automatica 
         Caption         =   "&Modo de Emissão Automática"
      End
      Begin VB.Menu mnu_download_pdf_danfe 
         Caption         =   "&Download de PDF's de DANFE (por data)"
      End
      Begin VB.Menu mnu_download_pdf_danfe_periodo 
         Caption         =   "Do&wnload de PDF's de DANFE (por período)"
      End
      Begin VB.Menu mnu_FECHAR 
         Caption         =   "&Fechar"
      End
   End
End
Attribute VB_Name = "f_EMISSAO_MANUAL"
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

Private Type TIPO_TOTALIZACAO_ITENS
    qtde_volumes As Long
    peso_bruto As Single
    peso_liquido As Single
    End Type
    
Dim v_totalizacao_itens() As TIPO_TOTALIZACAO_ITENS

Dim edicao_manual_liberada As Boolean
Dim usar_endereco_editado As Boolean
Dim endereco_editado__cep As String
Dim endereco_editado__logradouro As String
Dim endereco_editado__numero As String
Dim endereco_editado__complemento As String
Dim endereco_editado__bairro As String
Dim endereco_editado__cidade As String
Dim endereco_editado__uf As String

Dim blnAtualizaNFnoPedido As Boolean

Dim v_pedido_manual_boleto() As String
Dim v_parcela_manual_boleto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim blnExisteParcelamentoBoleto As Boolean

Sub atualiza_dados_endereco_cadastro()

Dim s As String
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo ADEC_TRATA_ERRO
    
  ' T_DESTINATARIO
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (cnpj_cpf='" & retorna_so_digitos(c_cnpj_cpf_dest) & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        s = "Destinatário com o CNPJ/CPF " & cnpj_cpf_formata(c_cnpj_cpf_dest) & " não foi encontrado no cadastro de clientes!!"
        aviso_erro s
        GoSub ADEC_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    l_end_cadastro_logradouro = Trim("" & t_DESTINATARIO("endereco"))
    l_end_cadastro_numero = Trim$("" & t_DESTINATARIO("endereco_numero"))
    l_end_cadastro_complemento = Trim$("" & t_DESTINATARIO("endereco_complemento"))
    l_end_cadastro_bairro = Trim$("" & t_DESTINATARIO("bairro"))
    l_end_cadastro_cidade = Trim$("" & t_DESTINATARIO("cidade"))
    l_end_cadastro_uf = Trim$("" & t_DESTINATARIO("uf"))
    l_end_cadastro_cep = cep_formata(retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep"))))

    GoSub ADEC_FECHA_TABELAS
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ADEC_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err)
    GoSub ADEC_FECHA_TABELAS
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ADEC_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_DESTINATARIO, True
    Return

End Sub

Sub trata_botao_editar_endereco()

Dim s As String

    On Error GoTo TBEE_TRATA_ERRO
    
    If Trim$(c_cnpj_cpf_dest) = "" Then
        aviso_erro "É necessário selecionar um cliente!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
    
    limpa_campos_endereco_edicao
    pn_endereco.Visible = True
    pn_endereco.ZOrder 0
        
    atualiza_dados_endereco_cadastro
    
    c_end_edicao_cep.SetFocus
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBEE_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    Exit Sub

End Sub


Sub fechar_programa()

'   FECHA BANCO DE DADOS
    BD_Fecha
    BD_CEP_Fecha
    BD_Assist_Fecha
    
'   ENCERRA PROGRAMA
    End

End Sub


Sub formulario_limpa()

Dim s As String
Dim i As Integer
Dim msg_erro As String
Dim aliquota_icms As Single

'   TRANSPORTADORA
'   ~~~~~~~~~~~~~~
    cb_transportadora.ListIndex = -1

'   CNPJ/CPF DESTINATÁRIO
'   ~~~~~~~~~~~~~~~~~~~~~
    c_cnpj_cpf_dest = ""
    c_nome_dest = ""
    
'   ITENS
'   ~~~~~
    c_vl_total_outras_despesas_acessorias = ""
    c_vl_total_geral = ""
    c_total_peso_liquido = ""
    c_total_peso_bruto = ""
    c_total_volumes = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_fcp(i) = ""
        c_unidade(i) = ""
        c_nItemPed(i) = ""
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
        With v_totalizacao_itens(i)
            .qtde_volumes = 0
            .peso_bruto = 0
            .peso_liquido = 0
            End With
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
    
'   LOCAL DE DESTINO DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "2 -"
    For i = 0 To cb_loc_dest.ListCount - 1
        If left$(cb_loc_dest.List(i), Len(s)) = s Then
            cb_loc_dest.ListIndex = i
            Exit For
            End If
        Next
        
'   NATUREZA DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "6.108"
    For i = 0 To cb_natureza.ListCount - 1
        If left$(cb_natureza.List(i), Len(s)) = s Then
            cb_natureza.ListIndex = i
            Exit For
            End If
        Next
        
'   ALÍQUOTA ICMS
'   ~~~~~~~~~~~~~
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
    
'   FORMA DE PAGAMENTO
'   ~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "0 -"
    For i = 0 To cb_forma_pagto.ListCount - 1
        If left$(cb_forma_pagto.List(i), Len(s)) = s Then
            cb_forma_pagto.ListIndex = i
            Exit For
            End If
        Next

'   MEIO DE PAGAMENTO
'   ~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "99 -"
    For i = 0 To cb_meio_pagto.ListCount - 1
        If left$(cb_meio_pagto.List(i), Len(s)) = s Then
            cb_meio_pagto.ListIndex = i
            Exit For
            End If
        Next


'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    c_dados_adicionais = ""
           
'   FRAME P/ EDIÇÃO DO ENDEREÇO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pn_endereco.Visible = False
    limpa_dados_endereco_cadastro
    limpa_dados_endereco_editado
    limpa_campos_endereco_edicao
    
    l_IE.Caption = ""
    
    c_pedido_nota = ""
    pn_pedido_nota.Visible = False
    pn_aviso_pedido_nota.Visible = False
    pnParcelasEmBoletos.Visible = False
    
'   FOCO INICIAL
'   ~~~~~~~~~~~~
    If cb_emitente.ListCount > 1 Then
        cb_emitente.SetFocus
    Else
        c_cnpj_cpf_dest.SetFocus
        End If

End Sub

Sub formulario_limpa_campos_itens_pedido()
Dim i As Integer
    
    c_vl_total_outras_despesas_acessorias = ""
    c_vl_total_geral = ""
    c_total_volumes = ""
    c_total_peso_bruto = ""
    c_total_peso_liquido = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_fcp(i) = ""
        c_unidade(i) = ""
        c_nItemPed(i) = ""
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


Sub obtem_info_itens_pedido_assistencia(ByVal pedido_selecionado As String)
Const NomeDestaRotina = "obtem_itens_pedido_assistencia()"
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


    On Error GoTo OIIPA_TRATA_ERRO
    
    If Trim$(pedido_selecionado) = "" Then Exit Sub
    
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
    t_PEDIDO.Open s, dbcAssist, , , adCmdText
    If t_PEDIDO.EOF Then
        aviso_erro "O pedido " & pedido_selecionado & " NÃO está cadastrado!!"
        GoSub OIIPA_FECHA_TABELAS
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
            " INNER JOIN t_PRODUTO tP ON (tPI.fabricante=tP.fabricante) AND (tPI.produto=tP.produto)" & _
            " INNER JOIN (" & _
                "SELECT fabricante, produto, max(id_estoque) as max_id_estoque FROM t_ESTOQUE_ITEM GROUP BY fabricante, produto" & _
            ") t_ULT_ESTOQUE ON (tP.fabricante = t_ULT_ESTOQUE.fabricante) AND (tP.produto = t_ULT_ESTOQUE.produto)" & _
            " INNER JOIN t_ESTOQUE_ITEM tEI ON (t_ULT_ESTOQUE.max_id_estoque=tEI.id_estoque) AND (t_ULT_ESTOQUE.fabricante=tEI.fabricante) AND (t_ULT_ESTOQUE.produto=tEI.produto)"
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
    t_PEDIDO_ITEM.Open s, dbcAssist, , , adCmdText
    intIndice = c_produto.LBound
    Do While Not t_PEDIDO_ITEM.EOF
    '   VERIFICA SE AINDA HÁ LINHAS DISPONÍVEIS
        If intIndice > c_produto.UBound Then
            aviso_erro "O pedido " & pedido_selecionado & " possui mais itens do que o permitido!!"
            GoSub OIIPA_FECHA_TABELAS
            Exit Sub
            End If
            
        c_fabricante(intIndice) = Trim$("" & t_PEDIDO_ITEM("fabricante"))
        c_produto(intIndice) = Trim$("" & t_PEDIDO_ITEM("produto"))
        c_descricao(intIndice) = Trim$("" & t_PEDIDO_ITEM("descricao"))
        
        c_CST(intIndice) = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PEDIDO_ITEM("cst")))
        c_NCM(intIndice) = Trim$("" & t_PEDIDO_ITEM("ncm"))
        c_unidade(intIndice) = "PC"
        
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
    
    GoSub OIIPA_FECHA_TABELAS
    
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIIPA_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIIPA_FECHA_TABELAS
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIIPA_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    Return
    
End Sub

Sub atualiza_dados_endereco_cadastro_assistencia()

Dim s As String
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo ADECA_TRATA_ERRO
    
  ' T_DESTINATARIO
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (cnpj_cpf='" & retorna_so_digitos(c_cnpj_cpf_dest) & "')"
    t_DESTINATARIO.Open s, dbcAssist, , , adCmdText
    If t_DESTINATARIO.EOF Then
        s = "Destinatário com o CNPJ/CPF " & cnpj_cpf_formata(c_cnpj_cpf_dest) & " não foi encontrado no cadastro de clientes!!"
        aviso_erro s
        GoSub ADECA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    l_end_cadastro_logradouro = Trim("" & t_DESTINATARIO("endereco"))
    l_end_cadastro_numero = Trim$("" & t_DESTINATARIO("endereco_numero"))
    l_end_cadastro_complemento = Trim$("" & t_DESTINATARIO("endereco_complemento"))
    l_end_cadastro_bairro = Trim$("" & t_DESTINATARIO("bairro"))
    l_end_cadastro_cidade = Trim$("" & t_DESTINATARIO("cidade"))
    l_end_cadastro_uf = Trim$("" & t_DESTINATARIO("uf"))
    l_end_cadastro_cep = cep_formata(retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep"))))

    GoSub ADECA_FECHA_TABELAS
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ADECA_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err)
    GoSub ADECA_FECHA_TABELAS
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ADECA_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_DESTINATARIO, True
    Return

End Sub

Sub DANFE_consulta(ByVal intEmitente As Integer, ByVal intSerieNFe As Integer, ByVal lngNumeroNFe As Long)

' CONSTANTES
Const NomeDestaRotina = "DANFE_consulta()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
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

Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

' BANCO DE DADOS
Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim dbcNFe As ADODB.Connection

    On Error GoTo DANFE_CONSULTA_TRATA_ERRO
    
    If intEmitente = 0 Then
        aviso_erro "Informe o emitente da NFe!!"
        cb_emitente_danfe.SetFocus
        Exit Sub
        End If
    
    If intSerieNFe = 0 Then
        aviso_erro "Informe a série da NFe!!"
        c_num_serie_danfe.SetFocus
        Exit Sub
        End If
        
    If lngNumeroNFe = 0 Then
        aviso_erro "Informe o número da NFe!!"
        c_num_nfe_danfe.SetFocus
        Exit Sub
        End If
        
  ' T_FIN_BOLETO_CEDENTE
    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
    With t_FIN_BOLETO_CEDENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
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
    
    aguarde INFO_EXECUTANDO, "consultando situação da NFe"
    
    s = "SELECT" & _
            " nome_empresa," & _
            " NFe_T1_servidor_BD," & _
            " NFe_T1_nome_BD," & _
            " NFe_T1_usuario_BD," & _
            " NFe_T1_senha_BD" & _
        " FROM t_FIN_BOLETO_CEDENTE" & _
        " WHERE" & _
            " (id = " & CStr(intEmitente) & ")"
    If t_FIN_BOLETO_CEDENTE.State <> adStateClosed Then t_FIN_BOLETO_CEDENTE.Close
    t_FIN_BOLETO_CEDENTE.Open s, dbc, , , adCmdText
    If t_FIN_BOLETO_CEDENTE.EOF Then
        s = "Falha ao localizar o registro em t_FIN_BOLETO_CEDENTE (id=" & CStr(intEmitente) & ")!!"
        aviso_erro s
        GoSub DANFE_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
    strNomeEmitente = UCase$(Trim$("" & t_FIN_BOLETO_CEDENTE("nome_empresa")))
    strNfeT1ServidorBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_servidor_BD"))
    strNfeT1NomeBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_nome_BD"))
    strNfeT1UsuarioBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_usuario_BD"))
    strNfeT1SenhaCriptografadaBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_senha_BD"))
    
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
    
    strNumeroNfNormalizado = NFeFormataNumeroNF(lngNumeroNFe)
    strSerieNfNormalizado = NFeFormataSerieNF(intSerieNFe)
    
'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
    cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
    Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
    intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
    strNFeMsgRetornoSP = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
    
    If intNfeRetornoSP <> 1 Then
        s = "Não é possível consultar a DANFE nº " & strNumeroNfNormalizado & " porque a NFe encontra-se na seguinte situação: " & CStr(intNfeRetornoSP) & " - " & strNFeMsgRetornoSP
        aviso_erro s
        GoSub DANFE_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
                    
    aguarde INFO_EXECUTANDO, "copiando PDF do DANFE"
    Set cmdNFeDanfe.ActiveConnection = dbcNFe
    cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
    cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
    Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
    If rsNFeRetornoSPDanfe.EOF Then
        s = "O conteúdo da DANFE nº " & strNumeroNfNormalizado & " não foi encontrado!!"
        aviso_erro s
        GoSub DANFE_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
    
    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
            aviso_erro s
            GoSub DANFE_CONSULTA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
            aviso_erro s
            GoSub DANFE_CONSULTA_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
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
    
    Close #lFileHandle
        
    GoSub DANFE_CONSULTA_FECHA_TABELAS
    
    aguarde INFO_EXECUTANDO, "exibindo PDF do DANFE"
    
    If Not start_doc(strNomeArqCompletoDanfe, s_erro) Then
        s = "Falha ao exibir o arquivo PDF do DANFE (" & strNomeArqCompletoDanfe & "): " & s_erro
        aviso_erro s
        End If
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DANFE_CONSULTA_FECHA_TABELAS:
'============================
  ' RECORDSETS
    bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
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
DANFE_CONSULTA_TRATA_ERRO:
'=========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub DANFE_CONSULTA_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Sub DANFE_CONSULTA_parametro_emitente(ByVal intEmitente As Integer, ByVal intSerieNFe As Integer, ByVal lngNumeroNFe As Long)

' CONSTANTES
Const NomeDestaRotina = "DANFE_CONSULTA_parametro_emitente()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strPastaEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim strNFeMsgRetornoSP As String

Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

' BANCO DE DADOS
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim dbcNFe As ADODB.Connection

    On Error GoTo DANFE_CONSULTA_PARAM_EMITENTE_TRATA_ERRO
    
    If intEmitente = 0 Then
        aviso_erro "Informe o emitente da NFe!!"
        cb_emitente_danfe.SetFocus
        Exit Sub
        End If
    
    If intSerieNFe = 0 Then
        aviso_erro "Informe a série da NFe!!"
        c_num_serie_danfe.SetFocus
        Exit Sub
        End If
        
    If lngNumeroNFe = 0 Then
        aviso_erro "Informe o número da NFe!!"
        c_num_nfe_danfe.SetFocus
        Exit Sub
        End If
        
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
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    aguarde INFO_EXECUTANDO, "consultando situação da NFe"
    
    s = "SELECT" & _
            " razao_social," & _
            " cnpj," & _
            " apelido," & _
            " NFe_T1_servidor_BD," & _
            " NFe_T1_nome_BD," & _
            " NFe_T1_usuario_BD," & _
            " NFe_T1_senha_BD" & _
        " FROM t_NFE_EMITENTE" & _
        " WHERE" & _
            " (id = " & CStr(intEmitente) & ")"
    If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    If t_NFE_EMITENTE.EOF Then
        s = "Falha ao localizar o registro em t_NFE_EMITENTE (id=" & CStr(intEmitente) & ")!!"
        aviso_erro s
        GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
    strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
    strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
    strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
    strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
    strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
    'novo padrão de nome da pasta para DANFEs: <cnpj>-<apelido_com_underlines_substituindo_barras>
    '(ex: 23209013000332-DIS_ES)
    strPastaEmitente = Trim$("" & t_NFE_EMITENTE("cnpj"))
    strPastaEmitente = retorna_so_digitos(strPastaEmitente)
    strPastaEmitente = strPastaEmitente & "-" & Trim$("" & t_NFE_EMITENTE("apelido"))
    strPastaEmitente = substitui_caracteres(strPastaEmitente, "/", "_")

    
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
    
    strNumeroNfNormalizado = NFeFormataNumeroNF(lngNumeroNFe)
    strSerieNfNormalizado = NFeFormataSerieNF(intSerieNFe)
    
'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
    cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
    cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
    Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
    intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
    strNFeMsgRetornoSP = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
    
    If intNfeRetornoSP <> 1 Then
        s = "Não é possível consultar a DANFE nº " & strNumeroNfNormalizado & " porque a NFe encontra-se na seguinte situação: " & CStr(intNfeRetornoSP) & " - " & strNFeMsgRetornoSP
        aviso_erro s
        GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
                    
    aguarde INFO_EXECUTANDO, "copiando PDF do DANFE"
    Set cmdNFeDanfe.ActiveConnection = dbcNFe
    cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
    cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
    Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
    If rsNFeRetornoSPDanfe.EOF Then
        s = "O conteúdo da DANFE nº " & strNumeroNfNormalizado & " não foi encontrado!!"
        aviso_erro s
        GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
    
    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
            aviso_erro s
            GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
            aviso_erro s
            GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
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
    
    Close #lFileHandle
        
    GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
    
    aguarde INFO_EXECUTANDO, "exibindo PDF do DANFE"
    
    If Not start_doc(strNomeArqCompletoDanfe, s_erro) Then
        s = "Falha ao exibir o arquivo PDF do DANFE (" & strNomeArqCompletoDanfe & "): " & s_erro
        aviso_erro s
        End If
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS:
'============================
  ' RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
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


Function ha_dados_preenchidos() As Boolean
Dim i As Integer

    ha_dados_preenchidos = True
    
    If Trim$(c_cnpj_cpf_dest) <> "" Then Exit Function
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_fabricante(i)) <> "" Then Exit Function
        If Trim$(c_produto(i)) <> "" Then Exit Function
        If Trim$(c_qtde(i)) <> "" Then Exit Function
        If converte_para_currency(c_vl_unitario(i)) <> 0 Then Exit Function
        Next
    
    If Trim$(c_dados_adicionais) <> "" Then Exit Function
    
    ha_dados_preenchidos = False
    
End Function

Sub limpa_campos_endereco_edicao()
    
    c_end_edicao_cep = ""
    c_end_edicao_logradouro = ""
    c_end_edicao_numero = ""
    c_end_edicao_complemento = ""
    c_end_edicao_bairro = ""
    c_end_edicao_cidade = ""
    c_end_edicao_uf = ""
    
End Sub

Sub limpa_dados_endereco_cadastro()

    l_end_cadastro_logradouro = ""
    l_end_cadastro_numero = ""
    l_end_cadastro_complemento = ""
    l_end_cadastro_bairro = ""
    l_end_cadastro_cidade = ""
    l_end_cadastro_uf = ""
    l_end_cadastro_cep = ""

End Sub

Sub limpa_dados_endereco_editado()

    usar_endereco_editado = False
    picEndereco.Visible = False
    picEndereco.ToolTipText = ""

    l_end_editado_logradouro = ""
    l_end_editado_numero = ""
    l_end_editado_complemento = ""
    l_end_editado_bairro = ""
    l_end_editado_cidade = ""
    l_end_editado_uf = ""
    l_end_editado_cep = ""

    endereco_editado__logradouro = ""
    endereco_editado__numero = ""
    endereco_editado__complemento = ""
    endereco_editado__bairro = ""
    endereco_editado__cidade = ""
    endereco_editado__uf = ""
    endereco_editado__cep = ""

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

Function tentativa_NF_anterior_OK(ByVal pedido_digitado As String) As Boolean
    Const NomeDestaRotina = "tentativa_NF_anterior_OK()"
    Dim s As String
    Dim s_cd As String
    Dim t_NFe_EMISSAO As ADODB.Recordset
    
    On Error GoTo TNFAOK_TRATA_ERRO
    
    tentativa_NF_anterior_OK = False
    
'   t_NFE_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   VERIFICA SE HOUVE TENTATIVA ANTERIOR DE EMITIR NF ANTERIOR REFERENTE AO PEDIDO
    s = "SELECT" & _
            " *" & _
        " FROM t_NFE_EMISSAO" & _
        " WHERE" & _
            " (pedido = '" & pedido_digitado & "')"
    If t_NFe_EMISSAO.State <> adStateClosed Then t_NFe_EMISSAO.Close
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If t_NFe_EMISSAO.EOF Then
        GoSub TNFAOK_FECHA_TABELAS
        Exit Function
        End If
    
    tentativa_NF_anterior_OK = True
    
    GoSub TNFAOK_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TNFAOK_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub TNFAOK_FECHA_TABELAS
    aviso_erro s
    Exit Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TNFAOK_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_NFe_EMISSAO, True
    Return
    
End Function


Function obtem_info_cliente_pedido(ByVal pedido As String, _
                            ByRef strNome As String, _
                            ByRef strCNPJCPF As String, _
                            ByRef strPgAntecStatus As String, _
                            ByRef strPgAntecQuitado As String, _
                            ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_cliente_pedido()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_id_cliente As String

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_CLIENTE As ADODB.Recordset

    On Error GoTo OICP_TRATA_ERRO
    
    obtem_info_cliente_pedido = False
    strNome = ""
    strCNPJCPF = ""
    strMsgErro = ""
    
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
        
  ' t_CLIENTE
    Set t_CLIENTE = New ADODB.Recordset
    With t_CLIENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'   VERIFICA O PEDIDO
    s_id_cliente = ""
    s_erro = ""
    s = "SELECT" & _
            " t_PEDIDO.pedido, t_PEDIDO.st_entrega, t_PEDIDO.id_cliente, " & _
            " t_PEDIDO.PagtoAntecipadoQuitadoStatus," & _
            " t_PEDIDO__BASE.PagtoAntecipadoStatus" & _
        " FROM t_PEDIDO" & _
        " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
        " ON (SUBSTRING(t_PEDIDO.pedido,1," & CStr(TAM_MIN_ID_PEDIDO) & ")=t_PEDIDO__BASE.pedido)" & _
        " WHERE" & _
            " (t_PEDIDO.pedido = '" & Trim$(pedido) & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Pedido " & Trim$(pedido) & " não está cadastrado !!"
    Else
    
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        s_id_cliente = Trim$("" & t_PEDIDO("id_cliente"))
        
        strPgAntecQuitado = Trim$("" & CStr(t_PEDIDO("PagtoAntecipadoQuitadoStatus")))
        strPgAntecStatus = Trim$("" & CStr(t_PEDIDO("PagtoAntecipadoStatus")))
        
        End If
    
'   ENCONTROU ERRO ?
    If s_erro <> "" Then
        strMsgErro = s_erro
        GoSub OICP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        

'   OBTÉM DADOS DO CLIENTE DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (id='" & s_id_cliente & "')"
    t_CLIENTE.Open s, dbc, , , adCmdText
    If t_CLIENTE.EOF Then
        strMsgErro = "Cliente com nº registro " & s_id_cliente & " não foi encontrado!!"
        GoSub OICP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If


'   NOME/RAZÃO SOCIAL DO CLIENTE
    strNome = Trim$("" & t_CLIENTE("nome"))

'   CNPJ/CPF
    strCNPJCPF = Trim$("" & t_CLIENTE("cnpj_cpf"))


'   TRATA A EXIBIÇÃO DE PARCELAMENTO CASO O PEDIDO POSSUA
    If (param_geracaoboletos.campo_texto = "Manual") Then
        If pedido <> "" Then
            ReDim v_pedido_manual_boleto(0)
            v_pedido_manual_boleto(UBound(v_pedido_manual_boleto)) = pedido
            blnExisteParcelamentoBoleto = False
            pnParcelasEmBoletos.Visible = False
            'If geraDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
            '    AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
            '    End If
            End If
        End If


    GoSub OICP_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

    obtem_info_cliente_pedido = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OICP_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OICP_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OICP_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_CLIENTE, True
    Return
    
End Function

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


Function lista_de_produtos_OK(ByVal pedido_digitado As String, ByRef mensagem As String) As Boolean
' a rotina fará a comparação dos itens do pedido com os itens em tela
' Retorno:
'   True - a lista bate com os itens do pedido
'   False - a lista não bate com os itens do pedido
    Const NomeDestaRotina = "lista_de_produtos_OK()"
    Dim s As String
    Dim v_tela_itens() As TIPO_LINHA_NOTA_FISCAL
    Dim v_pedido_itens() As TIPO_LINHA_NOTA_FISCAL
    Dim t_PEDIDO_ITEM As ADODB.Recordset
    Dim blnAchou As Boolean
    Dim i As Integer
    Dim idx As Integer
    Dim linha As Integer
    Dim qtde As Long
    
    On Error GoTo LDPOK_TRATA_ERRO
    
    lista_de_produtos_OK = False
    mensagem = ""
    
'   t_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'   MONTANDO VETOR COM OS ITENS DA TELA
    ReDim v_tela_itens(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_tela_itens(UBound(v_tela_itens))


    linha = 0
    Do While linha <= c_produto.UBound
        blnAchou = False
        If (c_fabricante(linha) <> "") And (c_produto(linha) <> "") Then
            For i = LBound(v_tela_itens) To UBound(v_tela_itens)
                With v_tela_itens(i)
                    If (.fabricante = c_fabricante(linha)) And (.produto = c_produto(linha)) Then
                        blnAchou = True
                        idx = i
                        Exit For
                        End If
                    End With
                Next
            
            If Not blnAchou Then
                If v_tela_itens(UBound(v_tela_itens)).produto <> "" Then
                    ReDim Preserve v_tela_itens(UBound(v_tela_itens) + 1)
                    limpa_item_TIPO_LINHA_NOTA_FISCAL v_tela_itens(UBound(v_tela_itens))
                    End If
                idx = UBound(v_tela_itens)
                With v_tela_itens(UBound(v_tela_itens))
                    .fabricante = c_fabricante(linha)
                    .produto = c_produto(linha)
                    End With
                End If
            
            With v_tela_itens(idx)
            '   QUANTIDADE
                qtde = 0
                If IsNumeric(c_qtde(linha)) Then qtde = CLng(c_qtde(linha))
                .qtde_total = .qtde_total + qtde
                End With
            End If
        linha = linha + 1
        Loop

'   MONTANDO VETOR COM OS ITENS DO PEDIDO
    ReDim v_pedido_itens(0)
    limpa_item_TIPO_LINHA_NOTA_FISCAL v_pedido_itens(UBound(v_pedido_itens))
    
    s = "SELECT" & _
            " t_PEDIDO_ITEM.fabricante," & _
            " t_PEDIDO_ITEM.produto," & _
            " t_PEDIDO_ITEM.qtde" & _
        " FROM t_PEDIDO_ITEM" & _
        " WHERE" & _
            " (t_PEDIDO_ITEM.pedido = '" & pedido_digitado & "')"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbc, , , adCmdText
    Do While Not t_PEDIDO_ITEM.EOF
        blnAchou = False
        For i = LBound(v_pedido_itens) To UBound(v_pedido_itens)
            With v_pedido_itens(i)
                If (.fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))) And (.produto = Trim$("" & t_PEDIDO_ITEM("produto"))) Then
                    blnAchou = True
                    idx = i
                    Exit For
                    End If
                End With
            Next
        
        If Not blnAchou Then
            If v_pedido_itens(UBound(v_pedido_itens)).produto <> "" Then
                ReDim Preserve v_pedido_itens(UBound(v_pedido_itens) + 1)
                limpa_item_TIPO_LINHA_NOTA_FISCAL v_pedido_itens(UBound(v_pedido_itens))
                End If
            idx = UBound(v_pedido_itens)
            With v_pedido_itens(UBound(v_pedido_itens))
                .fabricante = Trim$("" & t_PEDIDO_ITEM("fabricante"))
                .produto = Trim$("" & t_PEDIDO_ITEM("produto"))
                End With
            End If
        
        With v_pedido_itens(idx)
        '   QUANTIDADE
            qtde = 0
            If IsNumeric(t_PEDIDO_ITEM("qtde")) Then qtde = CLng(t_PEDIDO_ITEM("qtde"))
            .qtde_total = .qtde_total + qtde
            End With
        
        t_PEDIDO_ITEM.MoveNext
        Loop

    
'   VERIFICA SE OS ITENS DA TELA TEM CORRESPONDÊNCIA NOS ITENS DO PEDIDO
    If (UBound(v_pedido_itens) > UBound(v_tela_itens)) Then
        mensagem = "Número de itens do pedido é maior que o número de itens na tela"
        GoSub LDPOK_FECHA_TABELAS
        Exit Function
    ElseIf (UBound(v_pedido_itens) < UBound(v_tela_itens)) Then
        mensagem = "Número de itens na tela é maior que o número de itens do pedido"
        GoSub LDPOK_FECHA_TABELAS
        Exit Function
    Else
        i = 0
        Do While i <= UBound(v_pedido_itens)
            blnAchou = False
            idx = 0
            Do While (idx <= UBound(v_tela_itens)) And Not blnAchou
                If (v_pedido_itens(i).fabricante = v_tela_itens(idx).fabricante) And _
                    (v_pedido_itens(i).produto = v_tela_itens(idx).produto) Then
                    blnAchou = True
                    If (v_pedido_itens(i).qtde_total <> v_tela_itens(idx).qtde_total) Then
                        mensagem = "Quantidade do produto " & v_pedido_itens(i).produto & " divergente no pedido e na tela"
                        GoSub LDPOK_FECHA_TABELAS
                        Exit Function
                        End If
                    End If
                idx = idx + 1
                Loop
            If Not blnAchou Then
                mensagem = "O produto " & v_pedido_itens(i).produto & " está nos itens do pedido, mas não se encontra na tela."
                GoSub LDPOK_FECHA_TABELAS
                Exit Function
                End If
            i = i + 1
            Loop
        End If
   
    lista_de_produtos_OK = True
    
    GoSub LDPOK_FECHA_TABELAS
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LDPOK_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub LDPOK_FECHA_TABELAS
    aviso_erro s
    Exit Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LDPOK_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    Return
    
End Function


Sub NFe_emite(ByVal FLAG_NUMERACAO_MANUAL As Boolean)
' __________________________________________________________________________________________
'|
'|  EMITE A NOTA FISCAL ELETRÔNICA (NFe) COM BASE NOS DADOS PREENCHIDOS MANUALMENTE.
'|

' CONSTANTES
Const NomeDestaRotina = "NFe_emite()"
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
Dim s_erro As String
Dim s_erro_aux As String
Dim s_msg As String
Dim strCampo As String
Dim strCnpjCpfAux As String
Dim strDDD As String
Dim strIcms As String
Dim strSerieNf As String
Dim strSerieNfNormalizado As String
Dim strNumeroNf As String
Dim strNumeroNfNormalizado As String
Dim strEmitenteNf As String
Dim strTransportadoraId As String
Dim strTransportadoraCnpj As String
Dim strTransportadoraRazaoSocial As String
Dim strTransportadoraIE As String
Dim strTransportadoraUF As String
Dim strTransportadoraEmail As String
Dim strTransportadoraEmail2 As String
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
Dim strNFeTagInfRespTec As String
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
Dim strEmitenteCidade As String
Dim strEmitenteUf As String
Dim strNFeMsgRetornoSPSituacao As String
Dim strNFeMsgRetornoSPEmite As String
Dim strNFeMsgRetornoSPEmiteTamAjustadoBD As String
Dim strCodStatusInutilizacao As String
Dim strListaSugeridaMunicipiosIBGE As String
Dim strTextoCubagem As String
Dim strZerarPisCst As String
Dim strEmailXML As String
Dim strNFeRef As String
Dim s_pedido_nota As String
Dim s_pedido_aux As String

Dim strZerarCofinsCst As String
Dim strInfoAdicIbpt As String
Dim strEnderecoCep As String
Dim strEnderecoLogradouro As String
Dim strEnderecoNumero As String
Dim strEnderecoComplemento As String
Dim strEnderecoBairro As String
Dim strEnderecoCidade As String
Dim strEnderecoUf As String
Dim strPresComprador As String
Dim strInfoAdicParc As String
Dim strPagtoAntecipadoStatus As String
Dim strPagtoAntecipadoQuitadoStatus As String
Dim s_Texto_DIFAL_UF As String

' FLAGS
Dim blnIsDestinatarioPJ As Boolean
Dim blnHaProdutoCstIcms60 As Boolean
Dim blnErro As Boolean
Dim edicao_manual_liberada_foi_usada As Boolean
Dim blnExibirTotalTributos As Boolean
Dim blnHaProdutoSemDadosIbpt As Boolean
Dim blnIgnorarAtualizacaoNFnoPedido
Dim blnNotadeCompromisso As Boolean
Dim blnRemessaEntregaFutura As Boolean
Dim blnImprimeDadosFatura As Boolean
Dim blnIgnorarDIFAL As Boolean

' CONTADORES
Dim i As Integer
Dim j As Integer
Dim n As Long
Dim ic As Integer
Dim intNumItem As Integer
Dim iQtdConfirmaDuvidaEmit As Integer

' QUANTIDADES
Dim total_volumes_bd As Long
Dim qtde_linhas_nf As Integer
Dim lngMax As Long
Dim qtde_linhas As Integer
Dim lngAffectedRecords As Long

' CÓDIGOS E NSU
Dim intNfeRetornoSPSituacao As Integer
Dim intNfeRetornoSPEmite As Integer
Dim lngNsuNFeEmissao As Long
Dim intEmitente As Integer
Dim lngNsuNFeImagem As Long
Dim lngNFeUltNumeroNfEmitido As Long
Dim lngNFeUltSerieEmitida As Long
Dim lngNFeSerieManual As Long
Dim lngNFeNumeroNfManual As Long
Dim intContribuinteICMS As Integer
Dim intAnoPartilha As Integer

' BANCO DE DADOS
Dim t_DESTINATARIO As ADODB.Recordset
Dim t_TRANSPORTADORA As ADODB.Recordset
Dim t_IBPT As ADODB.Recordset
Dim t_PRODUTO As ADODB.Recordset
'Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim t_T1_NFE_INUTILIZA As ADODB.Recordset
Dim t_USUARIO_X_LOJA As ADODB.Recordset
Dim t_PEDIDO As ADODB.Recordset
Dim t_NFe_UF_PARAMETRO As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPEmite As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeEmite As New ADODB.Command
Dim dbcNFe As ADODB.Connection

' MOEDA
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
Dim total_peso_bruto_bd As Single
Dim total_peso_liquido_bd As Single
Dim cubagem_aux As Single
Dim cubagem_bruto As Single
Dim aliquota_icms_interestadual As Single

' VETORES
Dim v_nf() As TIPO_LINHA_NFe_EMISSAO_MANUAL
Dim vListaNFeRef() As String
Dim v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim v_pedido_nota() As String

' DADOS DE IMAGEM DA NFE
Dim rNFeImg As TIPO_NFe_IMG
Dim vNFeImgItem() As TIPO_NFe_IMG_ITEM
Dim vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP
Dim vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA
Dim vNFeImgPag() As TIPO_NFe_IMG_PAG

    On Error GoTo NFE_EMITE_TRATA_ERRO

'   CONSISTÊNCIAS
'   ~~~~~~~~~~~~~
'   EMITENTE
    If (cb_emitente.ListIndex = -1) Or (Trim$(cb_emitente) = "") Then
        aviso_erro "Selecione o emitente da NFe!!"
        cb_emitente.SetFocus
        Exit Sub
        End If
        
    s = cb_emitente
    s_aux = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then Exit For
        s_aux = s_aux & c
        Next
    intEmitente = CInt(s_aux)
    
'   DESTINATÁRIO
    If Trim$(c_cnpj_cpf_dest) = "" Then
        aviso_erro "Informe o CNPJ/CPF do destinatário!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
    
    If Not cnpj_cpf_ok(c_cnpj_cpf_dest) Then
        aviso_erro "CNPJ/CPF do destinatário inválido!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
        
'   INDICADOR DE PRESENÇA
    If Trim$(cb_indpres) = "" Then
        aviso_erro "Informe o indicador de presença!!"
        cb_indpres.SetFocus
        Exit Sub
        End If


'   PRODUTOS
    qtde_linhas = 0
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_produto(i)) <> "" Then qtde_linhas = qtde_linhas + 1
        Next
    
    If qtde_linhas = 0 Then
        aviso_erro "Nenhum produto foi informado!!"
        c_fabricante(c_fabricante.LBound).SetFocus
        Exit Sub
        End If
        
'   HÁ PRODUTOS REPETIDOS?
    For i = c_produto.LBound To (c_produto.UBound - 1)
        For j = (i + 1) To c_produto.UBound
            If (Trim$(c_produto(i)) <> "") And (Trim$(c_produto(j)) <> "") Then
                If (Trim$(c_fabricante(i)) = Trim$(c_fabricante(j))) And _
                    (Trim$(c_produto(i)) = Trim$(c_produto(j))) And _
                    (Trim$(c_NCM(i)) = Trim$(c_NCM(j))) Then
                    aviso_erro "Produto repetido nas linhas " & (i + 1) & " e " & (j + 1) & "!!"
                    c_produto(j).SetFocus
                    Exit Sub
                    End If
                End If
            Next
        Next

'   CONSISTE QUANTIDADE E VL UNITÁRIO DOS PRODUTOS
'   (se não se tratar de complemento de ICMS - LHGX)
    strNFeCodFinalidade = left$(Trim$(cb_finalidade), 1)
    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_produto(i)) <> "" Then
            If Not IsNumeric(c_qtde(i)) Then
                aviso_erro "Preencha a quantidade para o produto " & Trim$(c_produto(i)) & "!!"
                c_qtde(i).SetFocus
                Exit Sub
            ElseIf CInt(c_qtde(i)) = 0 Then
                aviso_erro "Informe a quantidade para o produto " & Trim$(c_produto(i)) & "!!"
                c_qtde(i).SetFocus
                Exit Sub
            ElseIf CInt(c_qtde(i)) < 0 Then
                aviso_erro "A quantidade do produto " & Trim$(c_produto(i)) & " não pode ser negativa!!"
                c_qtde(i).SetFocus
                Exit Sub
                End If
            
            If Trim$(c_vl_unitario(i)) = "" Then
                aviso_erro "Informe o valor unitário para o produto " & Trim$(c_produto(i)) & "!!"
                c_vl_unitario(i).SetFocus
                Exit Sub
            ElseIf converte_para_currency(c_vl_unitario(i)) < 0 Then
                aviso_erro "O valor unitário do produto " & Trim$(c_produto(i)) & " não pode ser negativo!!"
                c_vl_unitario(i).SetFocus
                Exit Sub
            ElseIf converte_para_currency(c_vl_outras_despesas_acessorias(i)) < 0 Then
                aviso_erro "O valor das outras despesas acessórias do produto " & Trim$(c_produto(i)) & " não pode ser negativo!!"
                c_vl_outras_despesas_acessorias(i).SetFocus
                Exit Sub
                End If
            If Trim$(c_unidade(i)) = "" Then
                aviso_erro "Informe a unidade para o produto " & Trim$(c_produto(i)) & "!!"
                c_unidade(i).SetFocus
                Exit Sub
                End If
            End If
        Next
        End If
        
'>  PEDIDO DA NOTA
'   LHGX - parece ser um bom lugar para a consistência, pois vai considerar os produttos
    blnIgnorarAtualizacaoNFnoPedido = False
    s_pedido_nota = Trim$(c_pedido_nota)
    If s_pedido_nota <> "" Then
        ' a rotina pedido_eh_do_emitente_atual já verifica se o pedido existe e se é do emitente atual
        If Not pedido_eh_do_emitente_atual(s_pedido_nota) Then
            c_pedido_nota.SetFocus
            Exit Sub
            End If
        
        '   verificar se os dados do cliente batem
        If Not obtem_info_cliente_pedido(s_pedido_nota, strCampo, strCnpjCpfAux, strPagtoAntecipadoStatus, strPagtoAntecipadoQuitadoStatus, s_aux) Then
            s_aux = "Não foi possível obter os dados do cliente do pedido " & s_pedido_nota & ": " & vbCrLf & _
                    s_aux
            aviso s_aux
            End If
            
        If Not ((strCnpjCpfAux <> "") And _
                (strCnpjCpfAux = retorna_so_digitos(c_cnpj_cpf_dest)) And _
                (Trim$(strCampo) = Trim$(c_nome_dest))) Then
            s_aux = "As informações do cliente estão divergentes com o pedido." & vbCrLf & vbCrLf & _
                "Confirma a associação desta nota com o pedido " & s_pedido_nota & "?"
            If Not confirma(s_aux) Then
                c_pedido_nota.SetFocus
                Exit Sub
                End If
            End If
                
        '   verificar se já foi feita tentativa de emissão automática
            If Not tentativa_NF_anterior_OK(s_pedido_nota) Then
                s_aux = "Foi informado um número de pedido que não teve tentativa anterior de emissão. " & vbCrLf & vbCrLf & _
                    "Confirma a associação desta nota com o pedido " & s_pedido_nota & "?"
                If Not confirma(s_aux) Then
                    c_pedido_nota.SetFocus
                    Exit Sub
                    End If
                End If
                
        
        
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
                
        '   CASO O PEDIDO PAI SEJA PARA PAGAMENTO ANTECIPADO, VERIFICA SE O PEDIDO FILHO ESTÁ QUITADO
        '   (não permitir emissão se não for nota de compromisso)
            If (strPagtoAntecipadoStatus = "1") And (strPagtoAntecipadoQuitadoStatus <> "1") Then
                If Not blnNotadeCompromisso Then
                    s = "Pedido " & s_pedido_nota & " se refere a venda futura não quitada!"
                    aviso s
                    GoSub NFE_EMITE_FECHA_TABELAS
                    aguarde INFO_NORMAL, m_id
                    Exit Sub
                    End If
                End If
        
                
        '   verificar se a lista de produtos bate
            If Not lista_de_produtos_OK(s_pedido_nota, s_aux) Then
                s_aux = "Pedido " & s_pedido_nota & ": " & vbCrLf & _
                        s_aux & vbCrLf & _
                        "Deseja prosseguir?"
                If Not confirma(s_aux) Then
                    c_pedido_nota.SetFocus
                    Exit Sub
                    End If
                End If
        End If
        
        
    If DESENVOLVIMENTO Then
        NFE_AMBIENTE = NFE_AMBIENTE_HOMOLOGACAO
    Else
        NFE_AMBIENTE = NFE_AMBIENTE_PRODUCAO
        End If
        
    ReDim vNFeImgItem(0)
    ReDim vNFeImgTagDup(0)
    ReDim vNFeImgNFeRef(0)
    ReDim vNFeImgPag(0)
        
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
    strNFeTagInfRespTec = ""
    strNFeInfAdicQuadroProdutos = ""
    strNFeInfAdicQuadroInfAdic = ""
    strNFeTagFat = ""
    strNFeTagDup = ""

    blnImprimeDadosFatura = False
    
'   OBTÉM TIPO DO DOCUMENTO FISCAL
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
    
    strNFeCodFinalidadeAux = retorna_finalidade_nfe(strCfopCodigo)
    If (strNFeCodFinalidade <> strNFeCodFinalidadeAux) And _
        (strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then
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
    
  ' T_PRODUTO
    Set t_PRODUTO = New ADODB.Recordset
    With t_PRODUTO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'  ' T_FIN_BOLETO_CEDENTE
'    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
'    With t_FIN_BOLETO_CEDENTE
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With

  ' T_NFE_EMITENTE
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
    
    ' t_USUARIO_X_LOJA
    Set t_USUARIO_X_LOJA = New ADODB.Recordset
    With t_USUARIO_X_LOJA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
  
  ' t_NFe_UF_PARAMETRO
    Set t_NFe_UF_PARAMETRO = New ADODB.Recordset
    With t_NFe_UF_PARAMETRO
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
  
  
'   SE FOI INFORMADO UM Nº DE PEDIDO, VERIFICA SE A ATUALIZAÇÃO DO CAMPO "OBSERVAÇÕES II" OU "OBSERVAÇÕES III" COM O Nº DA NOTA FISCAL PODE SER FEITA AUTOMATICAMENTE
    If s_pedido_nota <> "" Then
        If blnAtualizaNFnoPedido Then
            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
            t_PEDIDO.CursorType = BD_CURSOR_SOMENTE_LEITURA
            s = "SELECT * FROM t_PEDIDO WHERE ( pedido = '" & s_pedido_nota & "')"
            t_PEDIDO.Open s, dbc, , , adCmdText
            If Not t_PEDIDO.EOF Then
                s_aux = ""
                If opVenda.Value = True Then
                    If Trim$("" & t_PEDIDO("obs_2")) <> "" Then
                        s_aux = "O pedido " & s_pedido_nota & " já possui o campo 'Nº Nota Fiscal' preenchido com: " & Trim$("" & t_PEDIDO("obs_2")) & _
                                Chr(13) & _
                                "Atualizar o pedido com o novo número?"
                        End If
                ElseIf opRemessa.Value = True Then
                    If Trim$("" & t_PEDIDO("obs_3")) <> "" Then
                        s_aux = "O pedido " & s_pedido_nota & " já possui o campo 'NF Simples Remessa' preenchido com: " & Trim$("" & t_PEDIDO("obs_3")) & _
                                Chr(13) & _
                                "Atualizar o pedido com o novo número?"
                        End If
                ElseIf opVendaFutura.Value = True Then
                    If Trim$("" & t_PEDIDO("obs_4")) <> "" Then
                        s_aux = "O pedido " & s_pedido_nota & " já possui o campo 'NF Venda Futura' preenchido com: " & Trim$("" & t_PEDIDO("obs_4")) & _
                                Chr(13) & _
                                "Atualizar o pedido com o novo número?"
                        End If
                    End If
                If s_aux <> "" Then
                    If Not confirma(s_aux) Then blnIgnorarAtualizacaoNFnoPedido = True
                    End If
                End If
            End If
        End If
  
'   INICIALIZAÇÃO
    s_erro = ""
    rNFeImg.ide__indPag = "2"  ' Forma de pagamento: outros
    
'   OBTÉM OS DADOS DO EMITENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~
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
            " (id = " & CStr(intEmitente) & ")"
    
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    If t_NFE_EMITENTE.EOF Then
        aviso_erro "Dados do emitente não foram localizados no BD (id=" & CStr(intEmitente) & ")!!"
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
   
    rNFeImg.id_nfe_emitente = intEmitente
    
    'OBTÉM O INDICADOR DE PRESENÇA DO COMPRADOR NO ESTABELECIMENTO COMERCIAL NO MOMENTO DA OPERAÇÃO
    'selecionar através da combobox cb_indpres
    strPresComprador = left$(Trim$(cb_indpres), 1)
        
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
        
'   CONSULTA SITUAÇÃO DA NFe
    Set cmdNFeSituacao.ActiveConnection = dbcNFe
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
  
    
'   OBTÉM OS DADOS DOS PRODUTOS (TELA)
    ReDim v_nf(0)
    limpa_item_TIPO_LINHA_NFe_EMISSAO_MANUAL v_nf(UBound(v_nf))
    
    qtde_linhas_nf = 0
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_produto(i)) <> "" Then
            qtde_linhas_nf = qtde_linhas_nf + 1
            If v_nf(UBound(v_nf)).produto <> "" Then
                ReDim Preserve v_nf(UBound(v_nf) + 1)
                limpa_item_TIPO_LINHA_NFe_EMISSAO_MANUAL v_nf(UBound(v_nf))
                End If
            
            With v_nf(UBound(v_nf))
                .fabricante = Trim$(c_fabricante(i))
                .produto = Trim$(c_produto(i))
                .qtde = CInt(Trim$(c_qtde(i)))
                .vl_unitario = converte_para_currency(Trim$(c_vl_unitario(i)))
                .vl_outras_despesas_acessorias = converte_para_currency(Trim$(c_vl_outras_despesas_acessorias(i)))
                .infAdProd = Trim$(c_produto_obs(i))
                .xPed = Trim$(c_xPed(i))
                .nItemPed = Trim$(c_nItemPed(i))
                .uCom = Trim$(c_unidade(i))
                .uTrib = Trim$(c_unidade(i))
                .fcp = Trim$(c_fcp(i))
                .descricao_tela = Trim$(c_descricao(i))
                .CST_tela = Trim$(c_CST(i))
                .NCM_tela = Trim$(c_NCM(i))
                If cb_CFOP(i).ListIndex <> -1 Then
                    If Trim$(cb_CFOP(i)) <> "" Then
                        s = Trim$(cb_CFOP(i))
                        For j = 1 To Len(s)
                            c = Mid$(s, j, 1)
                            If c = " " Then Exit For
                            .CFOP_tela_formatado = .CFOP_tela_formatado & c
                            Next
                        .CFOP_tela = retorna_so_digitos(.CFOP_tela_formatado)
                        End If
                    End If
                If Trim$(cb_ICMS_item(i)) <> "" Then
                    .ICMS_tela = Trim$(cb_ICMS_item(i))
                    End If
                End With
            End If
        Next
    
'   OBTÉM OS DEMAIS DADOS DOS PRODUTOS (BD)
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            s = "SELECT" & _
                    " descricao," & _
                    " ean," & _
                    " qtde_volumes," & _
                    " peso," & _
                    " cubagem," & _
                    " perc_MVA_ST," & _
                    " (" & _
                        "SELECT TOP 1 cst FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=t_PRODUTO.fabricante) AND (tEI.produto=t_PRODUTO.produto) ORDER BY id_estoque DESC" & _
                    ") AS cst," & _
                    " (" & _
                        "SELECT TOP 1 ncm FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=t_PRODUTO.fabricante) AND (tEI.produto=t_PRODUTO.produto) ORDER BY id_estoque DESC" & _
                    ") AS ncm" & _
                " FROM t_PRODUTO" & _
                " WHERE" & _
                    " (fabricante = '" & Trim$(v_nf(i).fabricante) & "')" & _
                    " AND (produto = '" & Trim$(v_nf(i).produto) & "')"
            If t_PRODUTO.State <> adStateClosed Then t_PRODUTO.Close
            
            If bln_assist_pedido_ok Then
                t_PRODUTO.Open s, dbcAssist, , , adCmdText
            Else
                t_PRODUTO.Open s, dbc, , , adCmdText
                End If
            
            
            If t_PRODUTO.EOF Then
                v_nf(i).descricao_bd = ""
                v_nf(i).EAN = ""
                v_nf(i).NCM_bd = ""
                v_nf(i).CST_bd = ""
                v_nf(i).perc_MVA_ST = 0
                v_nf(i).qtde_volumes_total = 0
                v_nf(i).peso_total = 0
                v_nf(i).cubagem_total = 0
                
                If Not edicao_manual_liberada Then
                    If (strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then
                        aviso_erro "Falha ao localizar o produto no banco de dados: fabricante=" & Trim$(v_nf(i).fabricante) & ", produto=" & Trim$(v_nf(i).produto)
                        GoSub NFE_EMITE_FECHA_TABELAS
                        aguarde INFO_NORMAL, m_id
                        Exit Sub
                        End If
                Else
                    edicao_manual_liberada_foi_usada = True
                    End If
            Else
                v_nf(i).descricao_bd = Trim$("" & t_PRODUTO("descricao"))
                v_nf(i).EAN = Trim$("" & t_PRODUTO("ean"))
                v_nf(i).NCM_bd = Trim$("" & t_PRODUTO("ncm"))
                v_nf(i).CST_bd = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PRODUTO("cst")))
                v_nf(i).perc_MVA_ST = t_PRODUTO("perc_MVA_ST")
                n = 0
                If IsNumeric(t_PRODUTO("qtde_volumes")) Then n = CLng(t_PRODUTO("qtde_volumes"))
                v_nf(i).qtde_volumes_total = v_nf(i).qtde * n
                peso_aux = 0
                If IsNumeric(t_PRODUTO("peso")) Then peso_aux = CSng(t_PRODUTO("peso"))
                v_nf(i).peso_total = v_nf(i).qtde * peso_aux
                cubagem_aux = 0
                If IsNumeric(t_PRODUTO("cubagem")) Then cubagem_aux = CSng(t_PRODUTO("cubagem"))
                v_nf(i).cubagem_total = v_nf(i).qtde * cubagem_aux
                End If
            End If
        Next
  
  
'   DESCRIÇÃO => VERIFICA SE HOUVE ALTERAÇÃO NA DESCRIÇÃO DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If (Trim$(v_nf(i).descricao_tela) <> "") And (Trim$(v_nf(i).descricao_bd) <> "") Then
                If Trim$(v_nf(i).descricao_bd) <> Trim$(v_nf(i).descricao_tela) Then
                    edicao_manual_liberada_foi_usada = True
                    If s_msg <> "" Then s_msg = s_msg & vbCrLf
                    s_msg = s_msg & "Produto " & v_nf(i).produto & ": a descrição foi alterada de """ & v_nf(i).descricao_bd & """ para """ & v_nf(i).descricao_tela & """"
                    End If
                End If
            End If
        Next
    
    If s_msg <> "" Then
        s_msg = "Houve alteração na descrição do(s) seguinte(s) produto(s):" & _
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
    
'   PREPARA O CAMPO QUE ARMAZENA A DESCRIÇÃO A SER USADA
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            v_nf(i).descricao = v_nf(i).descricao_bd
            If Trim$(v_nf(i).descricao_tela) <> "" Then v_nf(i).descricao = Trim$(v_nf(i).descricao_tela)
            End If
        Next


'   CST => VERIFICA SE HOUVE ALTERAÇÃO NO CST DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If (Trim$(v_nf(i).CST_tela) <> "") And (Trim$(v_nf(i).CST_bd) <> "") Then
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

'   NCM => VERIFICA SE HOUVE ALTERAÇÃO NO NCM DE ALGUM PRODUTO E SOLICITA CONFIRMAÇÃO
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If (Trim$(v_nf(i).NCM_tela) <> "") And (Trim$(v_nf(i).NCM_bd) <> "") Then
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
    
    
'   CONSISTE DADOS
    s_msg = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            If Trim$(v_nf(i).descricao) = "" Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " NÃO possui descrição!!"
            ElseIf Trim$(v_nf(i).ncm) = "" Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " NÃO possui o código NCM!!"
            ElseIf (Len(Trim$(v_nf(i).ncm)) <> 0) And (Len(Trim$(v_nf(i).ncm)) <> 8) And (Len(Trim$(v_nf(i).ncm)) <> 2) Then
                If s_msg <> "" Then s_msg = s_msg & vbCrLf
                s_msg = s_msg & "O produto " & v_nf(i).produto & " - " & v_nf(i).descricao & " possui o código NCM com tamanho inválido!!"
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


'   OBTÉM DADOS DA TRANSPORTADORA
    strTransportadoraId = ""
    If (cb_transportadora.ListIndex <> -1) And (Trim$(cb_transportadora) <> "") Then
        If InStr(cb_transportadora, " - ") > 0 Then strTransportadoraId = Trim$(Mid$(cb_transportadora, 1, InStr(cb_transportadora, " - ") - 1))
        End If
        
    strTransportadoraCnpj = ""
    strTransportadoraRazaoSocial = ""
    strTransportadoraIE = ""
    strTransportadoraUF = ""
    strTransportadoraEmail = ""
    strTransportadoraEmail2 = ""
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
            strTransportadoraEmail2 = Trim$("" & t_TRANSPORTADORA("email2"))
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
            
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (cnpj_cpf='" & retorna_so_digitos(c_cnpj_cpf_dest) & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        s = "Destinatário com o CNPJ/CPF " & cnpj_cpf_formata(c_cnpj_cpf_dest) & " não foi encontrado no cadastro de clientes!!"
        aviso_erro s
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
    If usar_endereco_editado Then
        strEnderecoCep = retorna_so_digitos(endereco_editado__cep)
        strEnderecoLogradouro = endereco_editado__logradouro
        strEnderecoNumero = endereco_editado__numero
        strEnderecoComplemento = endereco_editado__complemento
        strEnderecoBairro = endereco_editado__bairro
        strEnderecoCidade = endereco_editado__cidade
        strEnderecoUf = endereco_editado__uf
    Else
        strEnderecoCep = retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep")))
        strEnderecoLogradouro = Trim("" & t_DESTINATARIO("endereco"))
        strEnderecoNumero = Trim$("" & t_DESTINATARIO("endereco_numero"))
        strEnderecoComplemento = Trim$("" & t_DESTINATARIO("endereco_complemento"))
        strEnderecoBairro = Trim$("" & t_DESTINATARIO("bairro"))
        strEnderecoCidade = Trim$("" & t_DESTINATARIO("cidade"))
        strEnderecoUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
        End If
    
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
            s = "O destinatário é do estado de " & UCase$(strEnderecoUf) & " cuja alíquota de ICMS é de " & strIcms & "%" & _
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
    
'   VERIFICA SE O CFOP ESTÁ COERENTE COM O CST DO ICMS
    s_confirma = ""
    For i = LBound(v_nf) To UBound(v_nf)
        If Trim$(v_nf(i).produto) <> "" Then
            strNFeCst = Trim$(right$(v_nf(i).cst, 2))
            strCfopCodigoAux = Trim$(v_nf(i).cfop)
            strCfopCodigoFormatadoAux = Trim$(v_nf(i).CFOP_formatado)
            s = "O produto " & v_nf(i).produto & " possui CST = " & strNFeCst & ", mas o CFOP selecionado é " & strCfopCodigoFormatadoAux
            If strNFeCst = "00" Then
                If (strCfopCodigoAux = "5102") Or (strCfopCodigoAux = "6102") Then s = ""
            ElseIf strNFeCst = "60" Then
                If (strCfopCodigoAux = "5405") Or (strCfopCodigoAux = "6404") Then s = ""
            Else
                If (strCfopCodigoAux <> "5102") And (strCfopCodigoAux <> "6102") And _
                   (strCfopCodigoAux <> "5405") And (strCfopCodigoAux <> "6404") Then s = ""
                End If
            
            If s <> "" Then
                If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
                s_confirma = s_confirma & s
                End If
            End If
        Next
        
    If s_confirma <> "" Then
        s_confirma = "Foram encontradas possíveis incoerências entre CFOP e CST:" & _
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
        
'   SE HÁ PEDIDO ESPECIFICANDO PAGAMENTO VIA BOLETO BANCÁRIO, CALCULA QUANTIDADE DE PARCELAS, DATAS E VALORES
'   DOS BOLETOS. ESSES DADOS SERÃO IMPRESSOS NA NF E TAMBÉM SALVOS NO BD, POIS SERVIRÃO DE BASE PARA A GERAÇÃO
'   DOS BOLETOS NO ARQUIVO DE REMESSA.
    If (param_geracaoboletos.campo_texto = "Manual") And blnExisteParcelamentoBoleto Then
        ReDim v_parcela_pagto(UBound(v_parcela_manual_boleto))
        v_parcela_pagto = v_parcela_manual_boleto
    Else
        ReDim v_parcela_pagto(0)
        ReDim v_pedido_nota(0)
        v_pedido_nota(0) = s_pedido_nota
        If Not geraDadosParcelasPagto(v_pedido_nota(), v_parcela_pagto(), s_erro) Then
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
            
        If s <> "" Then
            s = "Serão emitidas na NFe as seguintes informações de pagamento:" & Chr(13) & Chr(13) & s
            If DESENVOLVIMENTO Then
                aviso s
                End If
            End If
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
                         vbCrLf & _
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
    
        
    aguarde INFO_EXECUTANDO, "preparando emissão da NFe"
    
'   TAG OPERACIONAL
'   ~~~~~~~~~~~~~~~
    strNFeTagOperacional = "operacional;" & vbCrLf

'   EMAIL DO DESTINATÁRIO DA NFe
    rNFeImg.operacional__email = Trim("" & t_DESTINATARIO("email"))
    'testar se a concatenação de e-mails não excederá o limite
    If (Len(rNFeImg.operacional__email & ";" & strTransportadoraEmail) <= MAX_TAM_EMAIL_OPERACIONAL) Then
        If (Trim$(rNFeImg.operacional__email) <> "") And (Trim$(strTransportadoraEmail) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
        rNFeImg.operacional__email = rNFeImg.operacional__email & strTransportadoraEmail
        End If
    'testar se a concatenação de e-mails não excederá o limite
    If (Len(rNFeImg.operacional__email & ";" & strTransportadoraEmail2) <= MAX_TAM_EMAIL_OPERACIONAL) Then
        If (Trim$(rNFeImg.operacional__email) <> "") And (Trim$(strTransportadoraEmail2) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
        rNFeImg.operacional__email = rNFeImg.operacional__email & strTransportadoraEmail2
        End If
    strEmailXML = Trim("" & t_DESTINATARIO("email_xml"))
    'testar se a concatenação de e-mails não excederá o limite
    If (Len(rNFeImg.operacional__email & ";" & strEmailXML) <= MAX_TAM_EMAIL_OPERACIONAL) Then
        If Trim$(strEmailXML) <> "" Then
            If (Trim$(rNFeImg.operacional__email) <> "") Then rNFeImg.operacional__email = rNFeImg.operacional__email & ";"
            rNFeImg.operacional__email = rNFeImg.operacional__email & strEmailXML
            End If
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
        s_erro = "CNPJ/CPF do destinatário não está preenchido no cadastro de clientes!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Not cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
        s_erro = "CNPJ/CPF do destinatário está cadastrado com informação inválida!!"
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
        s_erro = "O nome do destinatário não está preenchido no cadastro de clientes!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O nome do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xNome = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xNome", rNFeImg.dest__xNome)
    
'   LOGRADOURO
    strCampo = strEnderecoLogradouro
    If strCampo = "" Then
        s_erro = "O endereço do destinatário não está preenchido!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O endereço do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xLgr = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xLgr", rNFeImg.dest__xLgr)
    
'   ENDEREÇO: NÚMERO
    strCampo = strEnderecoNumero
    If strCampo = "" Then
        s_erro = "O endereço do destinatário no cadastro do cliente deve ser preenchido corretamente para poder emitir a NFe!!" & vbCrLf & _
                 "As informações de número e complemento do endereço devem ser preenchidas nos campos adequados!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O número do endereço do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__nro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("nro", rNFeImg.dest__nro)
        
'   ENDEREÇO: COMPLEMENTO
    strCampo = strEnderecoComplemento
    If Len(strCampo) > 60 Then
        s_erro = "O campo complemento do endereço do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xCpl = strCampo
    If Len(strCampo) > 0 Then strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xCpl", rNFeImg.dest__xCpl)
    
'   BAIRRO
    strCampo = strEnderecoBairro
    If strCampo = "" Then
        s_erro = "O campo bairro no endereço do cliente não está preenchido no cadastro!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
    ElseIf Len(strCampo) > 60 Then
        s_erro = "O campo bairro no endereço do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xBairro = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xBairro", rNFeImg.dest__xBairro)
    
'   MUNICIPIO
    strCampo = strEnderecoCidade
    s_aux = strEnderecoUf
    If (strCampo <> "") And (s_aux <> "") Then strCampo = strCampo & "/"
    strCampo = strCampo & s_aux
    rNFeImg.dest__cMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cMun", rNFeImg.dest__cMun)
    
    strCampo = strEnderecoCidade
    If Len(strCampo) > 60 Then
        s_erro = "O campo cidade no endereço do destinatário excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xMun", rNFeImg.dest__xMun)
    
'   UF
    strCampo = strEnderecoUf
    If strCampo = "" Then
        s_erro = "O campo UF no endereço do destinatário não está preenchido no cadastro!!"
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
    strCampo = retorna_so_digitos(strEnderecoCep)
    If strCampo = "" Then
        s_erro = "O campo CEP no endereço do destinatário não está preenchido no cadastro!!"
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
            End If
        End If
        
    If strCampo <> "" Then
        rNFeImg.dest__fone = strCampo
        strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("fone", rNFeImg.dest__fone)
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
    total_volumes_bd = 0
    total_peso_bruto_bd = 0
    total_peso_liquido_bd = 0
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
                vNFeImgItem(UBound(vNFeImgItem)).det__uCom = .uCom
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uCom", vNFeImgItem(UBound(vNFeImgItem)).det__uCom)
                
            '   QUANTIDADE
                vNFeImgItem(UBound(vNFeImgItem)).det__qCom = NFeFormataNumero4Dec(.qtde)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qCom", vNFeImgItem(UBound(vNFeImgItem)).det__qCom)
                
            '   VALOR UNITÁRIO
                If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom = NFeFormataNumero4Dec(.vl_unitario)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom = NFeFormataNumero4Dec(0)
                    End If
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vUnCom", vNFeImgItem(UBound(vNFeImgItem)).det__vUnCom)
                
            '   VALOR TOTAL
                If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__vProd = NFeFormataMoeda2Dec(.qtde * .vl_unitario)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).det__vProd = NFeFormataMoeda2Dec(0)
                    End If
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("vProd", vNFeImgItem(UBound(vNFeImgItem)).det__vProd)
                
            '   cEANTrib - GTIN (Global Trade Item Number) da unidade tributável
                vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = .EAN
                'NFE 4.0 - EM BRANCO, INFORMAR SEM GTIN
                If vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "" Then vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib = "SEM GTIN"
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("cEANTrib", vNFeImgItem(UBound(vNFeImgItem)).det__cEANTrib)
            
            '   UNIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__uTrib = .uTrib
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("uTrib", vNFeImgItem(UBound(vNFeImgItem)).det__uTrib)
                
            '   QUANTIDADE TRIBUTÁVEL
                vNFeImgItem(UBound(vNFeImgItem)).det__qTrib = NFeFormataNumero4Dec(.qtde)
                strNFeTagDet = strNFeTagDet & vbTab & NFeFormataCampo("qTrib", vNFeImgItem(UBound(vNFeImgItem)).det__qTrib)
                
            '   VALOR UNITÁRIO DE TRIBUTAÇÃO
                If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                    vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib = NFeFormataNumero4Dec(.vl_unitario)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).det__vUnTrib = NFeFormataNumero4Dec(0)
                    End If
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
                If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) And (strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then
                    perc_IBPT = ibpt_aliquota_aplicavel(.cst, .percAliqNac, .percAliqImp)
                    vl_estimado_tributos = arredonda_para_monetario((.qtde * .vl_unitario) * (perc_IBPT / 100))
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
                vl_BC_ICMS = .qtde * .vl_unitario
            
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
                    vl_ICMS = (.qtde * .vl_unitario) * (perc_ICMS / 100)
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
                    
                '   NFE 4.0 - ENVIO FCP - automaticamente preenchido?
                    'strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pFCP", "0.00")
                    'strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vFCP", "0.00")
                    
                '   VALOR DO ICMS DESONERADO (ZERO, ATÉ RESOLUÇÃO EM CONTRÁRIO)
                    If vl_ICMSDeson <> 0 Then
                        vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson = NFeFormataMoeda2Dec(vl_ICMSDeson)
                        strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("vICMSDeson", vNFeImgItem(UBound(vNFeImgItem)).ICMS__vICMSDeson)
                        End If
            
            '   ICMS (CST=10): TRIBUTADA E COM COBRANÇA DO ICMS POR SUBSTITUIÇÃO TRIBUTÁRIA
                ElseIf strNFeCst = "10" Then
                    vl_ICMS = (.qtde * .vl_unitario) * (perc_ICMS / 100)
                    vl_ICMS = CCur(Format$(vl_ICMS, FORMATO_MOEDA))
                
                    If Not obtem_aliquota_ICMS_ST(rNFeImg.dest__UF, perc_ICMS_ST_aux, s_erro_aux) Then
                        s_erro = "Falha ao tentar obter a alíquota do ICMS ST para a UF: '" & rNFeImg.dest__UF & "'"
                        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
                        End If
                    perc_ICMS_ST = perc_ICMS_ST_aux
                    
                    vl_BC_ICMS_ST = calcula_BC_ICMS_ST((.qtde * .vl_unitario), .perc_MVA_ST)
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
                    " (UF='" & Trim$(strEnderecoUf) & "')"
                If t_NFe_UF_PARAMETRO.State <> adStateClosed Then t_NFe_UF_PARAMETRO.Close
                t_NFe_UF_PARAMETRO.Open s, dbc, , , adCmdText
                If Not t_NFe_UF_PARAMETRO.EOF Then
                    blnIgnorarDIFAL = t_NFe_UF_PARAMETRO("st_ignorar_difal") = 1
                    s_Texto_DIFAL_UF = Trim$("" & t_NFe_UF_PARAMETRO("texto_adicional"))
                    End If


                    
            '   OS CÁLCULOS DE PARTILHA FORAM MOVIDOS PARA CÁ DEVIDO À EXCLUSÃO DE ICMS E DIFAL DAS BASES DE CÁLCULO
            '   DE PIS E COFINS, CONFORME DECISÃO DO STF
            
'                A partir de 03/04/2023 condição abaixo de ICMS maior que zero começou a dar problema
'                (Nota Técnica 2022.005 SEFAZ)
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    (rNFeImg.dest__indIEDest = "9") And _
'                    Not blnIgnorarDIFAL And _
'                    Not cfop_eh_de_remessa(strCfopCodigo) And _
'                    (vl_ICMS > 0) Then
                    
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not blnIgnorarDIFAL And _
                    Not cfop_eh_de_remessa(strCfopCodigo) Then
                    
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
            
                vl_IPI = (.qtde * .vl_unitario) * (perc_IPI / 100)
                vl_IPI = CCur(Format$(vl_IPI, FORMATO_MOEDA))
                
            '   TAG PIS
            '   ~~~~~~~
                vl_PIS = 0
                vl_BC_PIS = 0
                
                strZerarPisCst = Trim$(left$(cb_zerar_PIS, 2))
                
                If strZerarPisCst = "" Then
                    vl_BC_PIS = .qtde * .vl_unitario
                    
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
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC = NFeFormataMoeda2Dec(vl_BC_PIS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC = NFeFormataMoeda2Dec(0)
                        End If
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).PIS__vBC)
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS = NFeFormataPercentual2Dec(perc_PIS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS = NFeFormataPercentual2Dec(0)
                        End If
                    strNFeTagPis = strNFeTagPis & vbTab & NFeFormataCampo("pPIS", vNFeImgItem(UBound(vNFeImgItem)).PIS__pPIS)
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS = NFeFormataMoeda2Dec(vl_PIS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).PIS__vPIS = NFeFormataMoeda2Dec(0)
                        End If
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
                    vl_BC_COFINS = .qtde * .vl_unitario
                    
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
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC = NFeFormataMoeda2Dec(vl_BC_COFINS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC = NFeFormataMoeda2Dec(0)
                        End If
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vBC", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vBC)
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS = NFeFormataPercentual2Dec(perc_COFINS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS = NFeFormataPercentual2Dec(0)
                        End If
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("pCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__pCOFINS)
                    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS = NFeFormataMoeda2Dec(vl_COFINS)
                    Else
                        vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS = NFeFormataMoeda2Dec(0)
                        End If
                    strNFeTagCofins = strNFeTagCofins & vbTab & NFeFormataCampo("vCOFINS", vNFeImgItem(UBound(vNFeImgItem)).COFINS__vCOFINS)
                Else
                    vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST = strZerarCofinsCst
                    strNFeTagCofins = vbTab & NFeFormataCampo("CST", vNFeImgItem(UBound(vNFeImgItem)).COFINS__CST)
                    End If
                
            '   TAG ICMSUFDest
            '   ~~~~~~~~~~~~~~
                
'                A partir de 03/04/2023 condição abaixo de ICMS maior que zero começou a dar problema
'                (Nota Técnica 2022.005 SEFAZ)
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    (rNFeImg.dest__indIEDest = "9") And _
'                    Not blnIgnorarDIFAL And _
'                    Not cfop_eh_de_remessa(strCfopCodigo) And _
'                    (vl_ICMS > 0) Then
                                
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
                
                
'                A partir de 03/04/2023 condição abaixo de ICMS maior que zero começou a dar problema
'                (Nota Técnica 2022.005 SEFAZ)
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    (rNFeImg.dest__indIEDest = "9") And _
'                    Not cfop_eh_de_remessa(strCfopCodigo) And _
'                    Not blnIgnorarDIFAL And _
'                    (vl_ICMS > 0) Then
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    Not blnIgnorarDIFAL Then

                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & _
                                            "ICMSUFDest;" & vbCrLf & strNFeTagIcmsUFDest
                    End If
                                
                
            '   INFORMAÇÕES ADICIONAIS DO PRODUTO
                vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd = .infAdProd
                If Trim$(vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd) <> "" Then
                    strNFeTagBlocoProduto = strNFeTagBlocoProduto & vbTab & NFeFormataCampo("infAdProd", vNFeImgItem(UBound(vNFeImgItem)).det__infAdProd)
                    End If
                
            '   QTDE DE VOLUMES
                total_volumes_bd = total_volumes_bd + .qtde_volumes_total
                
            '   PESO BRUTO
                total_peso_bruto_bd = total_peso_bruto_bd + .peso_total
                    
            '   PESO LIQUIDO
                total_peso_liquido_bd = total_peso_liquido_bd + .peso_total
                
            '   CUBAGEM TOTAL
                cubagem_bruto = cubagem_bruto + .cubagem_total
                
            '   TOTALIZAÇÃO
                vl_total_ICMS = vl_total_ICMS + vl_ICMS
                vl_total_ICMSDeson = vl_total_ICMSDeson + vl_ICMSDeson
                vl_total_ICMS_ST = vl_total_ICMS_ST + vl_ICMS_ST
                vl_total_produtos = vl_total_produtos + (.qtde * .vl_unitario)
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
    If (Trim$(c_total_volumes) <> "") And (total_volumes_bd <> 0) Then
        If CLng(c_total_volumes) <> total_volumes_bd Then
            s = "A quantidade total de volumes foi editada de " & CStr(total_volumes_bd) & " para " & c_total_volumes & vbCrLf & _
                "Continua mesmo assim?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If

    
'   PESO BRUTO TOTAL
'   ~~~~~~~~~~~~~~~~
    If (Trim$(c_total_peso_bruto) <> "") And (total_peso_bruto_bd <> 0) Then
        If formata_numero_3dec(converte_para_single(c_total_peso_bruto)) <> (total_peso_bruto_bd) Then
            s = "O peso bruto total foi editado de " & formata_numero_3dec(total_peso_bruto_bd) & " para " & c_total_peso_bruto & vbCrLf & _
                "Continua mesmo assim?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If
        End If

    
'   PESO LÍQUIDO TOTAL
'   ~~~~~~~~~~~~~~~~~~
    If (Trim$(c_total_peso_liquido) <> "") And (total_peso_liquido_bd <> 0) Then
        If formata_numero_3dec(converte_para_single(c_total_peso_liquido)) <> formata_numero_3dec(total_peso_liquido_bd) Then
            s = "O peso líquido total foi editado de " & formata_numero_3dec(total_peso_liquido_bd) & " para " & c_total_peso_liquido & vbCrLf & _
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

    'novo campo vICMSDeson (layout 3.10)
'   VALOR TOTAL DO ICMS DESONERADO
    rNFeImg.total__vICMSDeson = NFeFormataMoeda2Dec(vl_total_ICMSDeson)
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vICMSDeson", rNFeImg.total__vICMSDeson)
                             
    
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
    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
        rNFeImg.total__vProd = NFeFormataMoeda2Dec(vl_total_produtos)
    Else
        rNFeImg.total__vProd = NFeFormataMoeda2Dec(0)
        End If
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
    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
        rNFeImg.total__vPIS = NFeFormataMoeda2Dec(vl_total_PIS)
    Else
        rNFeImg.total__vPIS = NFeFormataMoeda2Dec(0)
        End If
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vPIS", rNFeImg.total__vPIS)
    
'   VALOR DO COFINS
    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
        rNFeImg.total__vCOFINS = NFeFormataMoeda2Dec(vl_total_COFINS)
    Else
        rNFeImg.total__vCOFINS = NFeFormataMoeda2Dec(0)
        End If
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
    If strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR Then
        rNFeImg.total__vNF = NFeFormataMoeda2Dec(vl_total_NF)
    Else
        rNFeImg.total__vNF = NFeFormataMoeda2Dec(0)
        End If
    strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vNF", rNFeImg.total__vNF)
                             
'   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    strInfoAdicIbpt = ""
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) Then
        rNFeImg.total__vTotTrib = NFeFormataMoeda2Dec(vl_total_estimado_tributos)
        strNFeTagValoresTotais = strNFeTagValoresTotais & _
                             vbTab & NFeFormataCampo("vTotTrib", rNFeImg.total__vTotTrib)
        If vl_total_NF = 0 Then
            perc_aux = 0
        Else
            perc_aux = 100 * (vl_total_estimado_tributos / vl_total_NF)
            End If
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
        rNFeImg.vol__qVol = retorna_so_digitos(CStr(total_volumes_bd))
        End If
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("qVol", rNFeImg.vol__qVol)
    
'   ESPÉCIE DOS VOLUMES TRANSPORTADOS
    rNFeImg.vol__esp = "VOLUME"
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("esp", rNFeImg.vol__esp)
    
'   PESO LÍQUIDO
    If Trim$(c_total_peso_liquido) <> "" Then
        rNFeImg.vol__pesoL = NFeFormataNumero3Dec(converte_para_single(Trim$(c_total_peso_liquido)))
    Else
        rNFeImg.vol__pesoL = NFeFormataNumero3Dec(total_peso_liquido_bd)
        End If
    strNFeTagVol = strNFeTagVol & _
                   vbTab & NFeFormataCampo("pesoL", rNFeImg.vol__pesoL)
    
'   PESO BRUTO
    If Trim$(c_total_peso_bruto) <> "" Then
        rNFeImg.vol__pesoB = NFeFormataNumero3Dec(converte_para_single(Trim$(c_total_peso_bruto)))
    Else
        rNFeImg.vol__pesoB = NFeFormataNumero3Dec(total_peso_bruto_bd)
        End If
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
    'se foi selecionado "90 - Sem pagamento, informar vPag zerado
    If left(cb_meio_pagto, 2) = "90" Then
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = "90"
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(0)
    'Se a nota é de entrada ou ajuste/devolução - sem pagamento
    ElseIf rNFeImg.ide__tpNF = "0" Or _
        strNFeCodFinalidade = "3" Or _
        strNFeCodFinalidade = "4" Then
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = "90"
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = NFeFormataMoeda2Dec(0)
    'se já estiver valendo a exigência de um tPag diferente de 99 (Outros), impedir emissão
    ElseIf (param_nftipopag.campo_inteiro = 1) And (left(cb_meio_pagto, 2) = "99") Then
        s = "Não é possível prosseguir com a emissão, pois o meio de pagamento não pode ser Outros!!"
        aviso_erro s
        cb_meio_pagto.SetFocus
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
    Else
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = left(cb_meio_pagto, 2)
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = rNFeImg.total__vNF
        End If
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("indPag", vNFeImgPag(UBound(vNFeImgPag)).pag__indPag)
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tPag", vNFeImgPag(UBound(vNFeImgPag)).pag__tPag)
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vPag", vNFeImgPag(UBound(vNFeImgPag)).pag__vPag)
    'informações do intermediador
    'strNFeTagPag = strNFeTagPag & vbTab & "infIntermed;"
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("CNPJ", "00776574000660")
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("idCadIntTran", "'23.209.013/0003-32")

    'Segundo informado pelo Valter (Target) em e-mail de 27/07/2017, o grupo vcard não deve ser informado no arquivo texto,
    'ele é preenchido pelo sistema
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vTroco", "0.00")
                              
    'teste informações cartão
    'strNFeTagPag = strNFeTagPag & vbTab & "card;"
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tpIntegra", "1")
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tBand", "01")  'visa
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("cAut", "A1B2C3D4E5F6G7H8I9J0")
    'strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vTroco", "0.00")

    
'   TAG INFADIC
'   ~~~~~~~~~~~
'   TEXTO DIGITADO
    If Trim$(c_dados_adicionais) <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & Trim$(c_dados_adicionais)
        End If
    
    If blnHaProdutoCstIcms60 Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = TEXTO_LEI_CST_ICMS_60 & strNFeInfAdicQuadroProdutos
        End If
    
    If (strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then
    '   INFORMA O TOTAL ESTIMADO DE TRIBUTOS?
    '   (a partir de 01/12/2022, esta mensagem não será exibida em NF's de transferência de estoque entre filiais - CFOP 5152
        If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) And (strInfoAdicIbpt <> "") And (strCfopCodigo <> "5152") Then
            If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
            strNFeInfAdicQuadroProdutos = strInfoAdicIbpt & strNFeInfAdicQuadroProdutos
            End If
        
    '   CUBAGEM
        strTextoCubagem = ""
        If cubagem_bruto > 0 Then strTextoCubagem = "CUB: " & formata_numero_2dec(cubagem_bruto) & " m3"
        If strTextoCubagem <> "" Then
            If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
            strNFeInfAdicQuadroProdutos = strTextoCubagem & strNFeInfAdicQuadroProdutos
            End If
        
    '   AVISO
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
        strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & "Fabricante não cobre avarias de peças plásticas, portanto, é necessário avaliar o equipamento no ato da entrega."
        
    '   TEXTO FIXO SOBRE REGIME ESPECIAL
        If txtFixoEspecifico <> "" Then
            If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & vbCrLf
            strNFeInfAdicQuadroProdutos = strNFeInfAdicQuadroProdutos & txtFixoEspecifico
            End If
        
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
        End If 'If (strNFeCodFinalidade <> NFE_FINALIDADE_NFE_COMPLEMENTAR)
        
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
    
    
'   TAG INFRESTEC
'   ~~~~~~~~~~~~~
    If (param_nfinformaresptec.campo_inteiro = 1) And (resptec_emissor.CNPJ <> "") Then
        strNFeTagInfRespTec = "infRespTec;" & vbCrLf & _
                                vbTab & NFeFormataCampo("CNPJ", resptec_emissor.CNPJ) & _
                                vbTab & NFeFormataCampo("xContato", resptec_emissor.nome) & _
                                vbTab & NFeFormataCampo("email", resptec_emissor.EMAIL) & _
                                vbTab & NFeFormataCampo("fone", resptec_emissor.telefone)
        End If
    
    
'   Nº DA NFE: AUTOMÁTICO OU MANUAL?
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If FLAG_NUMERACAO_MANUAL Then
    '   OBTÉM O NÚMERO DA ÚLTIMA NFe EMITIDA
        If Not NFeObtemUltimoNumeroEmitido(rNFeImg.id_nfe_emitente, lngNFeUltSerieEmitida, lngNFeUltNumeroNfEmitido, s_erro_aux) Then
            s = "Falha ao tentar consultar o número da última NFe emitida!!"
            If s_erro_aux <> "" Then s = s & vbCrLf
            s = s & s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        
    '   EXIBE PAINEL P/ USUÁRIO INFORMAR O Nº DA NFE MANUALMENTE
        f_NFE_NUM_MANUAL.lngNFeUltSerieEmitida = lngNFeUltSerieEmitida
        f_NFE_NUM_MANUAL.lngNFeUltNumeroNfEmitido = lngNFeUltNumeroNfEmitido
        f_NFE_NUM_MANUAL.strDescricaoEmitente = strEmitenteNf
        f_NFE_NUM_MANUAL.Show vbModal, Me
        If Not f_NFE_NUM_MANUAL.blnResultadoFormOk Then
            s = "Operação cancelada!!"
            aviso s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
            
        lngNFeSerieManual = f_NFE_NUM_MANUAL.lngNFeSerieManual
        lngNFeNumeroNfManual = f_NFE_NUM_MANUAL.lngNFeNumeroNfManual
        
    '   VERIFICA SE O Nº INFORMADO MANUALMENTE É POSTERIOR AO Nº DA ÚLTIMA NFe EMITIDA
        If lngNFeSerieManual <> lngNFeUltSerieEmitida Then
            s = "Não é permitido informar manualmente um nº de série da NFe diferente da série atual!"
            aviso s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
            
        If lngNFeNumeroNfManual > lngNFeUltNumeroNfEmitido Then
            s = "Não é permitido informar manualmente um número de NFe maior que o último número emitido automaticamente!"
            aviso s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        
    '   VERIFICA SE O Nº INFORMADO MANUALMENTE JÁ FOI USADO ANTERIORMENTE, CASO SIM, EXIBE O STATUS,
    '   INFORMAÇÕES BÁSICAS DA NFE E SOLICITA CONFIRMAÇÃO ATRAVÉS DA DIGITAÇÃO DA SENHA
        s = "SELECT" & _
                " data_hora," & _
                " dest__CNPJ," & _
                " dest__CPF," & _
                " dest__xNome," & _
                " dest__xLgr," & _
                " dest__nro," & _
                " dest__xMun," & _
                " dest__UF," & _
                " total__vNF" & _
            " FROM t_NFe_IMAGEM" & _
            " WHERE" & _
                " (id_nfe_emitente = " & CStr(rNFeImg.id_nfe_emitente) & ")" & _
                " AND (NFe_serie_NF = " & CStr(lngNFeSerieManual) & ")" & _
                " AND (NFe_numero_NF = " & CStr(lngNFeNumeroNfManual) & ")" & _
            " ORDER BY" & _
                " st_anulado ASC," & _
                " data_hora DESC"
        If t_NFe_IMAGEM.State <> adStateClosed Then t_NFe_IMAGEM.Close
        t_NFe_IMAGEM.Open s, dbc, , , adCmdText
        If Not t_NFe_IMAGEM.EOF Then
        '   SITUAÇÃO NO SISTEMA DA TARGET ONE
            cmdNFeSituacao.Parameters("NFe") = NFeFormataNumeroNF(lngNFeNumeroNfManual)
            cmdNFeSituacao.Parameters("Serie") = NFeFormataSerieNF(lngNFeSerieManual)
            Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
            intNfeRetornoSPSituacao = rsNFeRetornoSPSituacao("Retorno")
            strNFeMsgRetornoSPSituacao = Trim$("" & rsNFeRetornoSPSituacao("Mensagem"))
            
        '   VERIFICA SITUAÇÃO DA EMISSÃO ANTERIOR
            blnErro = False
            If (intNfeRetornoSPSituacao = 0) And (UCase$(Trim$(strNFeMsgRetornoSPSituacao)) = UCase$(Trim$("Aguardando processamento."))) Then blnErro = True
            If (intNfeRetornoSPSituacao = 1) And (UCase$(Trim$(strNFeMsgRetornoSPSituacao)) = UCase$(Trim$("Autorizada"))) Then blnErro = True
            
            If blnErro Then
                s = "Não é possível prosseguir com a emissão, pois já existe uma NFe com o mesmo número na seguinte situação:" & vbCrLf & _
                    CStr(intNfeRetornoSPSituacao) & " - " & strNFeMsgRetornoSPSituacao
                aviso_erro s
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
                
            strCnpjCpfAux = Trim$("" & t_NFe_IMAGEM("dest__CNPJ"))
            If strCnpjCpfAux = "" Then strCnpjCpfAux = Trim$("" & t_NFe_IMAGEM("dest__CPF"))
        '   JÁ HOUVE TENTATIVA ANTERIOR DE EMITIR ESTA NFe, ENTÃO CONFIRMA OPERAÇÃO VIA SOLICITAÇÃO DA SENHA DO USUÁRIO
            s = "ATENÇÃO:" & vbCrLf & _
                "Já houve uma tentativa anterior de emitir a NFe Nº " & NFeFormataNumeroNF(lngNFeNumeroNfManual) & " (Série: " & NFeFormataSerieNF(lngNFeSerieManual) & ")" & vbCrLf & _
                vbCrLf & _
                "Situação: " & intNfeRetornoSPSituacao & " - " & strNFeMsgRetornoSPSituacao & vbCrLf & _
                vbCrLf & _
                "Informações da tentativa anterior:" & vbCrLf & _
                "Data: " & Format$(t_NFe_IMAGEM("data_hora"), FORMATO_DATA_HORA) & vbCrLf & _
                "Valor: " & Format$(converte_para_currency(Trim$("" & t_NFe_IMAGEM("total__vNF"))), FORMATO_MOEDA) & vbCrLf & _
                "Cliente: " & cnpj_cpf_formata(strCnpjCpfAux) & " - " & Trim$("" & t_NFe_IMAGEM("dest__xNome")) & vbCrLf & _
                "Endereço: " & Trim$("" & t_NFe_IMAGEM("dest__xLgr")) & ", " & Trim$("" & t_NFe_IMAGEM("dest__nro")) & " - " & Trim$("" & t_NFe_IMAGEM("dest__xMun")) & " - " & Trim$("" & t_NFe_IMAGEM("dest__UF"))
            f_CONFIRMACAO_VIA_SENHA.strMensagemInformativa = s
            f_CONFIRMACAO_VIA_SENHA.strSenhaCorreta = usuario.senha
            f_CONFIRMACAO_VIA_SENHA.Show vbModal, Me
            If Not f_CONFIRMACAO_VIA_SENHA.blnResultadoFormOk Then
                s = "Operação cancelada!!"
                aviso s
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
            End If  'If Not t_NFe_IMAGEM.EOF
        End If  'If FLAG_NUMERACAO_MANUAL
    
    '   VERIFICAR DIVERGÊNCIA DE LOCAL DE DESTINO DA OPERAÇÃO
    s_confirma = ""
    'primeira situação: UFs diferentes e Local de Destino  <> Interestadual
    If rNFeImg.ide__tpNF <> "0" Then
        If (Trim$(rNFeImg.ide__idDest) <> "2") And (strEmitenteUf <> strEnderecoUf) Then
            If s_confirma <> "" Then s_confirma = s_confirma & vbCrLf
            s_confirma = s_confirma & "UF de origem e destino da Nota são diferentes, porém local de operação selecionado é " & vbCrLf & vbCrLf
            s_confirma = s_confirma & cb_loc_dest
            End If
        
        If (Trim$(rNFeImg.ide__idDest) <> "1") And (strEmitenteUf = strEnderecoUf) Then
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
    
'   SE FOR NOTA COMPLEMENTAR E NÃO HOUVER NOTA REFERENCIADA, INFORMAR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If (strNFeCodFinalidade = NFE_FINALIDADE_NFE_COMPLEMENTAR) And (Trim$(c_chave_nfe_ref) = "") Then
        s = "Nota fiscal complementar sendo impressa sem informação da nota referenciada. Operação pode gerar erro." & vbCrLf & _
            "Deseja prosseguir?"
        If Not confirma(s) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
'   SE FOR NOTA COMPLEMENTAR E NÃO HOUVER NOTA DADOS ADICIONAIS, INFORMAR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If (strNFeCodFinalidade = NFE_FINALIDADE_NFE_COMPLEMENTAR) And (Trim$(c_dados_adicionais) = "") Then
        s = "Nota fiscal complementar sem dados adicionais sobre a nota referenciada. Operação pode gerar erro." & vbCrLf & _
            "Deseja prosseguir?"
        If Not confirma(s) Then
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
'   CONFIRMAÇÃO FINAL
'   ~~~~~~~~~~~~~~~~~
    s = "Confirma a emissão da NFe?"
    If FLAG_NUMERACAO_MANUAL Then
        s = s & vbCrLf & vbCrLf & "Número da NFe informado manualmente:" & vbCrLf & _
            "Série NFe:  " & NFeFormataSerieNF(lngNFeSerieManual) & vbCrLf & _
            "Nº NFe:  " & NFeFormataNumeroNF(lngNFeNumeroNfManual)
        End If
    
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
    If FLAG_NUMERACAO_MANUAL Then
        strSerieNf = CStr(lngNFeSerieManual)
        strNumeroNf = CStr(lngNFeNumeroNfManual)
    Else
        aguarde INFO_EXECUTANDO, "obtendo próximo número de NF"
        If Not NFeObtemProximoNumero(rNFeImg.id_nfe_emitente, strSerieNf, strNumeroNf, s_erro_aux) Then
            s = "Falha ao tentar gerar o número para a NFe!!"
            If s_erro_aux <> "" Then s = s & vbCrLf
            s = s & s_erro_aux
            aviso_erro s
            GoSub NFE_EMITE_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
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
    '=== Campo indIntermed: para emissão manual, considerar sem intermediador
    If (param_nfintermediador.campo_inteiro = 1) And ((strPresComprador = "2") Or (strPresComprador = "3")) Then
        strNFeTagIdentificacao = strNFeTagIdentificacao & vbTab & NFeFormataCampo("indIntermed", "0") '0-Sem intermediador 1-Operação em site ou plataforma de terceiros
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
                   strNFeTagInfAdicionais & _
                   strNFeTagInfRespTec
    
    
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
    If s_pedido_nota <> "" Then
        t_NFe_EMISSAO("pedido") = s_pedido_nota
        End If
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
    t_NFe_EMISSAO("volumes_qtde_total_sistema") = total_volumes_bd
    t_NFe_EMISSAO("volumes_qtde_total_tela") = c_total_volumes
    t_NFe_EMISSAO("total_peso_bruto_sistema") = formata_numero_3dec(total_peso_bruto_bd)
    t_NFe_EMISSAO("total_peso_bruto_tela") = c_total_peso_bruto
    t_NFe_EMISSAO("total_peso_liquido_sistema") = formata_numero_3dec(total_peso_liquido_bd)
    t_NFe_EMISSAO("total_peso_liquido_tela") = c_total_peso_liquido
    If edicao_manual_liberada Or edicao_manual_liberada_foi_usada Then
        t_NFe_EMISSAO("edicao_manual_liberada") = 1
        End If
    
    s = RTrim$(c_dados_adicionais)
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
    
    
'   CANCELA DADOS DE HISTÓRICO ANTERIOR?
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If FLAG_NUMERACAO_MANUAL Then
        If intNfeRetornoSPEmite = 1 Then
        '   NFe FOI ACEITA, CANCELA DADOS DE EMISSÃO ANTERIOR
            s = "UPDATE t_NFe_EMISSAO SET" & _
                    " st_anulado = 1," & _
                    " dt_anulado = " & sqlMontaGetdateSomenteData() & "," & _
                    " dt_hr_anulado = getdate()," & _
                    " usuario_anulado = '" & usuario.id & "'" & _
                " WHERE" & _
                    " (id_nfe_emitente = " & CStr(rNFeImg.id_nfe_emitente) & ")" & _
                    " AND (NFe_serie_NF = " & CStr(lngNFeSerieManual) & ")" & _
                    " AND (NFe_numero_NF = " & CStr(lngNFeNumeroNfManual) & ")" & _
                    " AND (st_anulado = 0)" & _
                    " AND (id <> " & CStr(lngNsuNFeEmissao) & ")"
            dbc.Execute s, lngAffectedRecords
            
            s = "UPDATE t_NFe_IMAGEM SET" & _
                    " st_anulado = 1," & _
                    " dt_anulado = " & sqlMontaGetdateSomenteData() & "," & _
                    " dt_hr_anulado = getdate()," & _
                    " usuario_anulado = '" & usuario.id & "'" & _
                " WHERE" & _
                    " (id_nfe_emitente = " & CStr(rNFeImg.id_nfe_emitente) & ")" & _
                    " AND (NFe_serie_NF = " & CStr(lngNFeSerieManual) & ")" & _
                    " AND (NFe_numero_NF = " & CStr(lngNFeNumeroNfManual) & ")" & _
                    " AND (st_anulado = 0)" & _
                    " AND (id <> " & CStr(lngNsuNFeImagem) & ")"
            dbc.Execute s, lngAffectedRecords
        Else
        '   NFe FOI RECUSADA, CANCELA DADOS DESTA EMISSÃO
            s = "UPDATE t_NFe_EMISSAO SET" & _
                    " st_anulado = 1," & _
                    " dt_anulado = " & sqlMontaGetdateSomenteData() & "," & _
                    " dt_hr_anulado = getdate()," & _
                    " usuario_anulado = '" & usuario.id & "'" & _
                " WHERE" & _
                    " (id = " & CStr(lngNsuNFeEmissao) & ")"
            dbc.Execute s, lngAffectedRecords
            
            s = "UPDATE t_NFe_IMAGEM SET" & _
                    " st_anulado = 1," & _
                    " dt_anulado = " & sqlMontaGetdateSomenteData() & "," & _
                    " dt_hr_anulado = getdate()," & _
                    " usuario_anulado = '" & usuario.id & "'" & _
                " WHERE" & _
                    " (id = " & CStr(lngNsuNFeImagem) & ")"
            dbc.Execute s, lngAffectedRecords
            End If
        End If
    
    
'   GRAVA O LOG
'   ~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "gravando log"
    strLogComplemento = "Retorno SP=" & CStr(intNfeRetornoSPEmite) & " (" & IIf(intNfeRetornoSPEmite = 1, "Sucesso", "Falha") & ")" & _
                        "; Msg SP=" & strNFeMsgRetornoSPEmite & _
                        "; Série NFe=" & strSerieNf & _
                        "; Nº NFe=" & strNumeroNf & _
                        "; tela emissão=Painel Manual" & _
                        "; emitente=" & CStr(rNFeImg.id_nfe_emitente) & _
                        "; destinatário=" & cnpj_cpf_formata(c_cnpj_cpf_dest) & _
                        "; tipo=" & cb_tipo_NF & _
                        "; natureza operação=" & cb_natureza & _
                        "; ICMS=" & cb_icms & _
                        "; IPI=" & c_ipi & _
                        "; frete=" & cb_frete & _
                        "; zerar PIS=(" & Trim$(cb_zerar_PIS) & ")" & _
                        "; zerar COFINS=(" & Trim$(cb_zerar_COFINS) & ")" & _
                        "; transportadora=" & strTransportadoraId & _
                        "; finalidade=" & Trim$(cb_finalidade) & _
                        "; chave NFe referenciada=" & Trim$(c_chave_nfe_ref) & _
                        "; dados adicionais=" & Trim$(c_dados_adicionais)
    s_pedido_aux = ""
    If s_pedido_nota <> "" Then s_pedido_aux = s_pedido_nota
    Call grava_log(usuario.id, "", s_pedido_aux, "", OP_LOG_NFE_EMISSAO_MANUAL, strLogComplemento)
        
'   SUCESSO NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "processamento complementar"
    If intNfeRetornoSPEmite = 1 Then
    '   SE FOI INFORMADO UM Nº DE PEDIDO, ATUALIZA O CAMPO "OBSERVAÇÕES II" OU "OBSERVAÇÕES III" COM O Nº DA NOTA FISCAL
    '   (dependendo da opção selecionada)
        If blnAtualizaNFnoPedido And (Not blnIgnorarAtualizacaoNFnoPedido) And (s_pedido_nota <> "") Then
        ' T_PEDIDO
            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
            t_PEDIDO.CursorType = BD_CURSOR_EDICAO
            s = "SELECT * FROM t_PEDIDO WHERE ( pedido = '" & s_pedido_nota & "')"
            t_PEDIDO.Open s, dbc, , , adCmdText
            If Not t_PEDIDO.EOF Then
                If opVendaFutura.Value = True Then
                    t_PEDIDO("obs_4") = strNumeroNf
                    t_PEDIDO.Update
                ElseIf opRemessa.Value = True Then
                    t_PEDIDO("obs_3") = strNumeroNf
                    t_PEDIDO.Update
                Else
                    t_PEDIDO("obs_2") = strNumeroNf
                    t_PEDIDO.Update
                    End If
                End If
            End If
    
'   FALHA NA CHAMADA DA STORED PROCEDURE!!
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        aviso_erro "Falha na emissão da NFe:" & vbCrLf & strNFeMsgRetornoSPEmite
        End If
        
        
  ' LIMPA FORMULÁRIO
    c_num_serie_danfe = strSerieNf
    c_num_nfe_danfe = strNumeroNf
    For i = 0 To cb_emitente_danfe.ListCount - 1
        s = cb_emitente_danfe.List(i)
        s_aux = ""
        For j = 1 To Len(s)
            c = Mid$(s, j, 1)
            If Not IsNumeric(c) Then Exit For
            s_aux = s_aux & c
            Next
        If CStr(rNFeImg.id_nfe_emitente) = s_aux Then
            cb_emitente_danfe.ListIndex = i
            Exit For
            End If
        Next
        
    formulario_limpa
        
  ' EXIBE DADOS DA ÚLTIMA NFe EMITIDA
    l_serie_NF = strSerieNfNormalizado
    l_num_NF = strNumeroNfNormalizado
    l_emitente_NF = strEmitenteNf
        
    f_MAIN.l_serie_NF = strSerieNfNormalizado
    f_MAIN.l_num_NF = strNumeroNfNormalizado
    f_MAIN.l_emitente_NF = strEmitenteNf

    GoSub NFE_EMITE_FECHA_TABELAS
    
    c_cnpj_cpf_dest.SetFocus
    
    aguarde INFO_NORMAL, m_id
    
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
    bd_desaloca_recordset t_DESTINATARIO, True
    bd_desaloca_recordset t_TRANSPORTADORA, True
    bd_desaloca_recordset t_IBPT, True
    bd_desaloca_recordset t_PRODUTO, True
    'bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset t_NFe_IMAGEM, True
    bd_desaloca_recordset t_T1_NFE_INUTILIZA, True
    bd_desaloca_recordset t_USUARIO_X_LOJA, True
    bd_desaloca_recordset t_PEDIDO, True
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
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_TRANSPORTADORA As ADODB.Recordset
Dim vAliquotas() As String

    On Error GoTo FI_TRATA_ERRO
    
  ' t_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
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

'   EMITENTE
'   ~~~~~~~~
    s = "SELECT" & _
            " e.id," & _
            " e.razao_social" & _
        " FROM t_NFE_EMITENTE e" & _
        " INNER JOIN t_USUARIO_X_NFe_EMITENTE ue ON e.id = ue.id_nfe_emitente" & _
        " WHERE" & _
             " (e.st_ativo = 1)" & _
        " AND" & _
             " (ue.usuario = '" & usuario.id & "')" & _
        " ORDER BY" & _
            " id"
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    Do While Not t_NFE_EMITENTE.EOF
        s = Trim("" & t_NFE_EMITENTE("id")) & " - " & Trim("" & t_NFE_EMITENTE("razao_social"))
        cb_emitente.AddItem s
        cb_emitente_danfe.AddItem s
        t_NFE_EMITENTE.MoveNext
        Loop
    
    'If cb_emitente.ListCount = 1 Then cb_emitente.ListIndex = 0
    'If cb_emitente_danfe.ListCount = 1 Then cb_emitente_danfe.ListIndex = 0
    s = usuario.emit_id
    For i = 0 To cb_emitente.ListCount - 1
        If left(cb_emitente.List(i), Len(s)) = s Then
            cb_emitente.ListIndex = i
            cb_emitente_danfe.ListIndex = i
            Exit For
            End If
        Next
    
'   TRANSPORTADORA
'   ~~~~~~~~~~~~~~
    s = "SELECT" & _
            " id," & _
            " nome," & _
            " razao_social" & _
        " FROM t_TRANSPORTADORA" & _
        " ORDER BY" & _
            " id"
    t_TRANSPORTADORA.Open s, dbc, , , adCmdText
    cb_transportadora.AddItem ""
    Do While Not t_TRANSPORTADORA.EOF
        s = Trim("" & t_TRANSPORTADORA("nome"))
        If s = "" Then s = Trim("" & t_TRANSPORTADORA("razao_social"))
        s = Trim("" & t_TRANSPORTADORA("id")) & " - " & UCase$(s)
        cb_transportadora.AddItem s
        t_TRANSPORTADORA.MoveNext
        Loop
            
'   FINALIDADE DE EMISSÃO
'   ~~~~~~~~~~~~~~~~~~~~~
    cb_finalidade.Clear
    cb_finalidade.AddItem "1 - NFe Normal"
    cb_finalidade.AddItem "2 - NFe Complementar"
    cb_finalidade.AddItem "2 - NFe Complementar (Complemento de ICMS)"
    'LHGX - VER DEPOIS SE DÁ PARA MUDAR A COR
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
    
'   LOCAL DE DESTINO DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cb_loc_dest.Clear
    cb_loc_dest.AddItem "1 - INTERNA"
    cb_loc_dest.AddItem "2 - INTERESTADUAL"
    cb_loc_dest.AddItem "3 - EXTERIOR"
    
'   INDICADOR DE PRESENÇA DO COMPRADOR
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cb_indpres.Clear
    cb_indpres.AddItem "0 - NÃO SE APLICA"
    cb_indpres.AddItem "1 - PRESENCIAL"
    cb_indpres.AddItem "2 - INTERNET"
    cb_indpres.AddItem "3 - TELEATENDIMENTO"
    
'   NATUREZA DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~
    cb_natureza.Clear
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
        vAliquotas = Split(s_aux, vbCrLf)
        For i = LBound(vAliquotas) To UBound(vAliquotas)
            cb_icms.AddItem vAliquotas(i)
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
        End If
    
    For i = cb_ICMS_item.LBound To cb_ICMS_item.UBound
        cb_ICMS_item(i).Clear
        cb_ICMS_item(i).AddItem ""
        For j = 0 To (cb_icms.ListCount - 1)
            If Trim$(cb_icms.List(j)) <> "" Then cb_ICMS_item(i).AddItem cb_icms.List(j)
            Next
        Next
    
    'ajusta o ICMS de acordo com a UF do depósito
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
    
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    With c_dados_adicionais
        .FontName = FONTNAME_IMPRESSAO
        .FontSize = FONTSIZE_IMPRESSAO
        .FontBold = FONTBOLD_IMPRESSAO
        .FontItalic = FONTITALIC_IMPRESSAO
        End With
    
'   FRAME C/ CAMPOS P/ EDIÇÃO DO ENDEREÇO
    pn_endereco.Visible = False
    
'   HORÁRIO DE VERÃO
'   ~~~~~~~~~~~~~~~
    If ((Date >= InicioHorarioVerao(Year(Date))) And (Date <= TerminoHorarioVerao(Year(Date) + 1))) Or ((Date >= _
        InicioHorarioVerao(Year(Date) - 1)) And (Date <= TerminoHorarioVerao(Year(Date)))) Then
        
        End If
        
        
'   FORMA DE PAGAMENTO
'   ~~~~~~~~~~~~~~~~~~
    cb_forma_pagto.Clear
    cb_forma_pagto.AddItem "0 - À vista"
    cb_forma_pagto.AddItem "1 - À prazo"
    
    
'   MEIO DE PAGAMENTO
'   ~~~~~~~~~~~~~~~~~
    cb_meio_pagto.Clear
    cb_meio_pagto.AddItem "01 - Dinheiro"
    cb_meio_pagto.AddItem "02 - Cheque"
    cb_meio_pagto.AddItem "03 - Cartão de Crédito"
    cb_meio_pagto.AddItem "04 - Cartão de Débito"
    cb_meio_pagto.AddItem "05 - Crédito Loja"
    cb_meio_pagto.AddItem "10 - Vale Alimentação"
    cb_meio_pagto.AddItem "11 - Vale Refeição"
    cb_meio_pagto.AddItem "12 - Vale Presente"
    cb_meio_pagto.AddItem "13 - Vale Combustível"
    cb_meio_pagto.AddItem "15 - Boleto Bancário"
    If (param_nftipopag.campo_inteiro = 1) Then
        cb_meio_pagto.AddItem "16 - Depósito Bancário"
        cb_meio_pagto.AddItem "17 - PIX"
        cb_meio_pagto.AddItem "18 - Transf Bancária, Carteira Digital"
        cb_meio_pagto.AddItem "19 - Programa Fidelidade, Cashback, Crédito Virtual"
        End If
    cb_meio_pagto.AddItem "90 - Sem pagamento"
    cb_meio_pagto.AddItem "99 - Outros"
    
    
    GoSub FI_FECHA_TABELAS
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FI_TRATA_ERRO:
'=============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub FI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FI_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_TRANSPORTADORA, True
    Return
    
End Sub

Sub recalcula_totais()
Dim i As Integer
Dim vl_total_geral As Currency
Dim vl_total_outras_despesas_acessorias As Currency
Dim qtde_aux As Integer
Dim vl_aux As Currency
Dim vl_total As Currency
Dim blnTemDadosParaVlTotal As Boolean
Dim blnTemDadosParaVlOutrasDespesasAcessorias As Boolean

    For i = c_fabricante.LBound To c_fabricante.UBound
        If (Trim$(c_qtde(i)) = "") Or (Trim$(c_vl_unitario(i)) = "") Then
            c_vl_total(i) = ""
        Else
            blnTemDadosParaVlTotal = True
            qtde_aux = CInt(Trim$(c_qtde(i)))
            vl_aux = converte_para_currency(c_vl_unitario(i))
            vl_total = qtde_aux * vl_aux
            c_vl_total(i) = formata_moeda(vl_total)
            vl_total_geral = vl_total_geral + vl_total
            End If
        
        If Trim$(c_vl_outras_despesas_acessorias(i)) <> "" Then
            blnTemDadosParaVlOutrasDespesasAcessorias = True
            vl_total_outras_despesas_acessorias = vl_total_outras_despesas_acessorias + converte_para_currency(c_vl_outras_despesas_acessorias(i))
            End If
        Next
        
    If blnTemDadosParaVlTotal Then
        c_vl_total_geral = formata_moeda(vl_total_geral)
    Else
        c_vl_total_geral = ""
        End If
    
    If blnTemDadosParaVlOutrasDespesasAcessorias Then
        c_vl_total_outras_despesas_acessorias = formata_moeda(vl_total_outras_despesas_acessorias)
    Else
        c_vl_total_outras_despesas_acessorias = ""
        End If
        
End Sub

Sub fechar_modo_emissao_manual()
Dim s As String

    If ha_dados_preenchidos Then
        s = "Os dados preenchidos serão perdidos se o painel for alternado para o modo de emissão automática!!" & _
            vbCrLf & _
            "Continua assim mesmo?"
        If Not confirma(s) Then Exit Sub
        End If
    
    Unload f_EMISSAO_MANUAL

End Sub

Sub recalcula_totalizacao_itens()
Dim i As Integer
Dim qtde_aux As Integer
Dim blnTemDados As Boolean
Dim lngTotalVolumes As Long
Dim dblPesoBruto As Single
Dim dblPesoLiq As Single

    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_qtde(i)) <> "" Then
            blnTemDados = True
            qtde_aux = CInt(Trim$(c_qtde(i)))
            lngTotalVolumes = lngTotalVolumes + (qtde_aux * v_totalizacao_itens(i).qtde_volumes)
            dblPesoBruto = dblPesoBruto + (qtde_aux * v_totalizacao_itens(i).peso_bruto)
            dblPesoLiq = dblPesoLiq + (qtde_aux * v_totalizacao_itens(i).peso_liquido)
            End If
        Next
        
    If blnTemDados Then
        c_total_volumes = CStr(lngTotalVolumes)
        c_total_peso_bruto = formata_numero_3dec(dblPesoBruto)
        c_total_peso_liquido = formata_numero_3dec(dblPesoLiq)
    Else
        c_total_volumes = ""
        c_total_peso_bruto = ""
        c_total_peso_liquido = ""
        End If

End Sub

Sub tab_stop_configura()
Dim i As Integer

    b_dummy.TabIndex = 0
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
    b_end_editado_exclui.TabIndex = 0
    b_danfe.TabIndex = 0
    c_num_nfe_danfe.TabIndex = 0
    c_num_serie_danfe.TabIndex = 0
    cb_emitente_danfe.TabIndex = 0
    b_editar_endereco.TabIndex = 0
    b_fechar.TabIndex = 0
    b_edicao.TabIndex = 0
    b_emissao_automatica.TabIndex = 0
    b_destinatario.TabIndex = 0
    b_emite_numeracao_manual.TabIndex = 0
    b_imprime.TabIndex = 0
    c_dados_adicionais.TabIndex = 0
    c_pedido_nota.TabIndex = 0
    c_vl_total_outras_despesas_acessorias.TabIndex = 0
    c_vl_total_geral.TabIndex = 0
    c_total_peso_liquido.TabIndex = 0
    c_total_peso_bruto.TabIndex = 0
    c_total_volumes.TabIndex = 0
    
    For i = c_produto.UBound To c_produto.LBound Step -1
        c_fcp(i).TabIndex = 0
        c_unidade(i).TabIndex = 0
        c_nItemPed(i).TabIndex = 0
        c_xPed(i).TabIndex = 0
        c_nItemPed(i).TabIndex = 0
        cb_ICMS_item(i).TabIndex = 0
        c_NCM(i).TabIndex = 0
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
    
    c_chave_nfe_ref.TabIndex = 0
    cb_finalidade.TabIndex = 0
    cb_zerar_COFINS.TabIndex = 0
    cb_zerar_PIS.TabIndex = 0
    cb_transportadora.TabIndex = 0
    cb_loc_dest.TabIndex = 0
    cb_indpres.TabIndex = 0
    cb_frete.TabIndex = 0
    c_ipi.TabIndex = 0
    cb_icms.TabIndex = 0
    cb_natureza.TabIndex = 0
    cb_tipo_NF.TabIndex = 0
    c_nome_dest.TabIndex = 0
    c_cnpj_cpf_dest.TabIndex = 0
    cb_emitente.TabIndex = 0
    
End Sub

Sub trata_botao_endereco_edicao_cancela()

    pn_endereco.Visible = False

End Sub

Sub trata_botao_endereco_edicao_limpa()

    limpa_campos_endereco_edicao
    c_end_edicao_cep.SetFocus

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
    l_end_editado_cep = endereco_editado__cep
    
    endereco_editado__logradouro = Trim$(c_end_edicao_logradouro)
    l_end_editado_logradouro = endereco_editado__logradouro
    
    endereco_editado__numero = Trim$(c_end_edicao_numero)
    l_end_editado_numero = endereco_editado__numero
    
    endereco_editado__complemento = Trim$(c_end_edicao_complemento)
    l_end_editado_complemento = endereco_editado__complemento
    
    endereco_editado__bairro = Trim$(c_end_edicao_bairro)
    l_end_editado_bairro = endereco_editado__bairro
    
    endereco_editado__cidade = Trim$(c_end_edicao_cidade)
    l_end_editado_cidade = endereco_editado__cidade
    
    endereco_editado__uf = Trim$(c_end_edicao_uf)
    l_end_editado_uf = endereco_editado__uf
    
    usar_endereco_editado = True
    picEndereco.Visible = True
    picEndereco.ToolTipText = "Endereço: " & endereco_editado__logradouro & _
                            IIf(endereco_editado__numero = "", "", " Nº " & endereco_editado__numero) & _
                            IIf(endereco_editado__complemento = "", "", " Complemento: " & endereco_editado__complemento) & _
                            IIf(endereco_editado__bairro = "", "", " Bairro: " & endereco_editado__bairro) & _
                            IIf(endereco_editado__cidade = "", "", " Cidade: " & endereco_editado__cidade) & _
                            IIf(endereco_editado__uf = "", "", " UF: " & endereco_editado__uf) & _
                            IIf(endereco_editado__cep = "", "", " CEP: " & endereco_editado__cep)
    
End Sub

Sub trata_botao_endereco_editado_exclui()
    
    limpa_dados_endereco_editado
    
End Sub

Sub mostra_endereco_editado()

    aviso "Endereço: " & endereco_editado__logradouro & vbCrLf & _
            IIf(endereco_editado__numero = "", "", "Nº " & endereco_editado__numero & vbCrLf) & _
            IIf(endereco_editado__complemento = "", "", "Complemento: " & endereco_editado__complemento & vbCrLf) & _
            IIf(endereco_editado__bairro = "", "", "Bairro: " & endereco_editado__bairro & vbCrLf) & _
            IIf(endereco_editado__cidade = "", "", "Cidade: " & endereco_editado__cidade & vbCrLf) & _
            IIf(endereco_editado__uf = "", "", "UF: " & endereco_editado__uf & vbCrLf) & _
            IIf(endereco_editado__cep = "", "", "CEP: " & endereco_editado__cep)
    
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

Sub trata_botao_assistencia_tecnica()
    Dim conexao_ok As Boolean
    Dim s As String
    
    conexao_ok = Not (dbcAssist Is Nothing)
        
    If Not conexao_ok Then
        aguarde INFO_EXECUTANDO, "conectando ao banco dados de Assistência Técnica"
        If Not BD_Assist_inicia() Then
            s = "Falha ao conectar com o Banco de Dados de Assistência Técnica!!" & _
                vbCrLf & "As informações deverão ser digitadas !!"
            aviso_erro s
        Else
            conexao_ok = True
            End If
        aguarde INFO_NORMAL, m_id
        End If
        
    If conexao_ok Then
        f_PEDIDO_ASSISTENCIA.Show vbModal, Me
        End If
    
    If bln_assist_pedido_ok Then
        formulario_limpa_campos_itens_pedido
        limpa_dados_endereco_cadastro
        limpa_dados_endereco_editado
        limpa_campos_endereco_edicao
        obtem_info_itens_pedido_assistencia (s_assist_Pedido)
        c_cnpj_cpf_dest = retorna_so_digitos(s_assist_Cliente_CNPJ_CPF)
        c_cnpj_cpf_dest_LostFocus
        'bln_assist_pedido_ok = False
        End If
        
End Sub

Sub atualiza_valor_total_icms()
    
    Const NomeDestaRotina = "atualiza_valor_total_icms()"
    Dim s As String
    Dim vl_icms_total As Currency
    Dim vl_icms_calc As Currency
    Dim i As Integer
    Dim perc_ICMS_geral As Single
    Dim perc_ICMS_calc As Single
    Dim temItemPreenchido As Boolean
    Dim s_cst As String
    
    On Error GoTo AVTI_TRATA_ERRO
    
    ''se é uma operação de complemento de ICMS, sair
    'If (left$(Trim(cb_finalidade), 1) <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then Exit Sub
    
    'se não houver item preenchido, não realizar o cálculo
    temItemPreenchido = False
    For i = c_fabricante.LBound To c_fabricante.UBound
        If Trim(c_fabricante(i)) <> "" Then temItemPreenchido = True
        Next
    If Not temItemPreenchido Then Exit Sub
    
    If IsNumeric(cb_icms) Then
        perc_ICMS_geral = CSng(cb_icms)
    Else
        perc_ICMS_geral = 0
        End If
    
    vl_icms_total = 0
    For i = c_fabricante.LBound To c_fabricante.UBound
        
        s_cst = Trim$(right$(c_CST(i), 2))
        vl_icms_calc = converte_para_currency(c_vl_total(i))
        If IsNumeric(cb_ICMS_item(i)) Then
            perc_ICMS_calc = CSng(cb_ICMS_item(i))
        Else
            perc_ICMS_calc = perc_ICMS_geral
            End If
        
        If s_cst = "00" Then
            vl_icms_calc = vl_icms_calc * perc_ICMS_calc / 100
        ElseIf s_cst = "10" Then
            vl_icms_calc = vl_icms_calc * perc_ICMS_calc / 100
        ElseIf (s_cst = "40") Or (s_cst = "41") Or (s_cst = "50") Then
            vl_icms_calc = 0
        ElseIf s_cst = "60" Then
            vl_icms_calc = 0
        Else
            vl_icms_calc = vl_icms_calc * perc_ICMS_calc / 100
            End If
            
        'vl_icms_calc = vl_icms_calc * perc_ICMS_calc / 100
        vl_icms_total = vl_icms_total + vl_icms_calc
            
        Next
        
    c_vl_total_icms = formata_moeda(vl_icms_total)
    
    Exit Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
AVTI_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    

End Sub

Sub trata_tela_complemento_ICMS()
Dim s As String

    s = "ATENÇÃO!!!" & vbCrLf & vbCrLf & _
    "Será emitida um NOTA FISCAL DE COMPLEMENTO DE ICMS!" & vbCrLf & vbCrLf & _
    "Preencher o valor da nota fiscal e a alíquota complementar do ICMS" & vbCrLf & _
    "(O VALOR DA NOTA FISCAL NÃO SERÁ IMPRESSO, APENAS UTILIZADO NOS CÁLCULOS)!" & vbCrLf & vbCrLf & _
    "Devem ser informados o mesmo CFOP e o mesmo CST da nota fiscal original!" & vbCrLf & vbCrLf & _
    "Incluir o número da nota fiscal original no campo 'Dados Adicionais'!" & vbCrLf & vbCrLf & _
    "Deseja prosseguir?"
    If Not confirma(s) Then Exit Sub

    formulario_limpa_campos_itens_pedido
    c_fabricante(0) = "000"
    c_produto(0) = "000045"
    c_descricao(0) = "COMPLEMENTO DE ICMS"
    c_qtde(0) = 1
    c_vl_unitario(0) = "0,00"
    c_NCM(0) = "95069100"
    c_fcp(0) = ""
    c_unidade(0) = "UN"
    c_CST(0) = "000"
    c_vl_total(0) = "0,00"
    c_produto_obs(0) = ""
    
    'INDICADOR DE PRESENÇA: NÃO SE APLICA
    cb_indpres.ListIndex = 0

    c_dados_adicionais = c_dados_adicionais & "Nota Fiscal Complementar - Complemento de ICMS"

End Sub


Private Sub b_assistencia_tecnica_Click()
    
    trata_botao_assistencia_tecnica
    
End Sub

Private Sub b_cep_pesquisar_Click()

    trata_botao_pesquisa_cep
    
End Sub

Private Sub b_danfe_Click()

Const NomeDestaRotina = "b_danfe_Click()"
Dim i As Integer
Dim s As String
Dim s_aux As String
Dim c As String
Dim intEmitente As Integer

    On Error GoTo B_DANFE_CLICK_TRATA_ERRO
    
    If (cb_emitente_danfe.ListIndex = -1) Or (Trim$(cb_emitente_danfe) = "") Then
        aviso_erro "Selecione o emitente da NFe da qual deseja consultar a DANFE!!"
        cb_emitente_danfe.SetFocus
        Exit Sub
        End If
        
    If Trim$(c_num_serie_danfe) = "" Then
        aviso_erro "Informe o nº de série da NFe da qual deseja consultar a DANFE!!"
        c_num_serie_danfe.SetFocus
        Exit Sub
        End If
        
    If Trim$(c_num_nfe_danfe) = "" Then
        aviso_erro "Informe o nº da NFe da qual deseja consultar a DANFE!!"
        c_num_nfe_danfe.SetFocus
        Exit Sub
        End If
    
    s = cb_emitente_danfe
    s_aux = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then Exit For
        s_aux = s_aux & c
        Next
    intEmitente = CInt(s_aux)
    
    DANFE_CONSULTA_parametro_emitente intEmitente, CInt(c_num_serie_danfe), CLng(c_num_nfe_danfe)
    
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

Private Sub b_destinatario_Click()
' CONSTANTES
Const NomeDestaRotina = "b_destinatario_Click()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_endereco As String
Dim s_bairro As String
Dim s_cep As String
Dim s_cidade As String
Dim s_uf As String
Dim s_nome As String
Dim s_ie_rg As String
Dim s_info As String
Dim s_end_linha_1 As String
Dim s_end_linha_2 As String

' BANCO DE DADOS
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo BCLI_TRATA_ERRO
    
    If Trim$(c_cnpj_cpf_dest) = "" Then
        aviso_erro "Informe o CNPJ/CPF do destinatário!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
        
    If Not cnpj_cpf_ok(Trim$(c_cnpj_cpf_dest)) Then
        aviso_erro "CNPJ/CPF do destinatário é inválido!!"
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If
        
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
  ' T_DESTINATARIO
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    
    s_endereco = ""
    s_bairro = ""
    s_cep = ""
    s_cidade = ""
    s_uf = ""
    s_nome = ""
    s_ie_rg = ""
        
'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (cnpj_cpf='" & retorna_so_digitos(c_cnpj_cpf_dest) & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        s = "Cliente com CNPJ/CPF " & cnpj_cpf_formata(c_cnpj_cpf_dest) & " não foi encontrado!!"
        aviso_erro s
        GoSub BCLI_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If

    s_endereco = UCase$(Trim$("" & t_DESTINATARIO("endereco")))
    s_aux = UCase$(Trim$("" & t_DESTINATARIO("endereco_numero")))
    If s_aux <> "" Then s_endereco = s_endereco & ", " & s_aux
    s_aux = UCase$(Trim$("" & t_DESTINATARIO("endereco_complemento")))
    If s_aux <> "" Then s_endereco = s_endereco & " " & s_aux

'   BAIRRO
    s_bairro = UCase$(Trim$("" & t_DESTINATARIO("bairro")))

'   CEP
    s_cep = Trim$("" & t_DESTINATARIO("cep"))

'   CIDADE
    s_cidade = UCase$(Trim$("" & t_DESTINATARIO("cidade")))

'   UF
    s_uf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    
    s_end_linha_1 = s_endereco
    If (s_end_linha_1 <> "") And (s_bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & s_bairro
    
    s_end_linha_2 = s_cidade
    If (s_end_linha_2 <> "") And (s_uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & s_uf
    If (s_end_linha_2 <> "") And (s_cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & cep_formata(s_cep)
        
    If (s_end_linha_1 <> "") And (s_end_linha_2 <> "") Then s_end_linha_1 = s_end_linha_1 & vbCrLf
    
    s_info = s_nome & vbCrLf
    
    s_info = s_info & "CNPJ/CPF: " & cnpj_cpf_formata(c_cnpj_cpf_dest) & vbCrLf
    If s_ie_rg <> "" Then s_info = s_info & "IE/RG: " & s_ie_rg & vbCrLf
            
    s_info = s_info & _
             s_end_linha_1 & s_end_linha_2
    
    aviso s_info

    GoSub BCLI_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BCLI_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub BCLI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BCLI_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_DESTINATARIO, True
    Return
    
End Sub

Private Sub b_edicao_Click()
Dim i As Integer
Dim s As String

    If Not edicao_manual_liberada Then
    '   SOLICITA CONFIRMAÇÃO VIA SENHA
        s = "A liberação da edição manual desativa todas as regras de consistência dos produtos, permitindo o preenchimento com dados arbitrários!!" & vbCrLf & _
            "Continua assim mesmo?"
        f_CONFIRMACAO_VIA_SENHA.strMensagemInformativa = s
        f_CONFIRMACAO_VIA_SENHA.strSenhaCorreta = usuario.senha
        f_CONFIRMACAO_VIA_SENHA.Show vbModal, Me
        If Not f_CONFIRMACAO_VIA_SENHA.blnResultadoFormOk Then
            s = "Operação cancelada!!"
            aviso s
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
        End If
    
    edicao_manual_liberada = Not edicao_manual_liberada
    
    If edicao_manual_liberada Then
        b_edicao.Caption = "&Bloquear Edição"
    Else
        b_edicao.Caption = "Li&berar Edição"
        End If
    
    
    For i = c_fabricante.LBound To c_fabricante.UBound
        If edicao_manual_liberada Then
            c_descricao(i).Locked = False
        Else
            c_descricao(i).Locked = True
            End If
        Next

End Sub

Private Sub b_editar_endereco_Click()

    trata_botao_editar_endereco
    
End Sub

Private Sub b_emissao_automatica_Click()

    fechar_modo_emissao_manual
    
End Sub

Private Sub b_emite_numeracao_manual_Click()

    NFe_emite True
    
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

Private Sub b_end_editado_exclui_Click()

    trata_botao_endereco_editado_exclui
    
End Sub

Private Sub b_fechar_Click()

    fechar_programa
    
End Sub


Private Sub b_imprime_Click()

    If blnExisteParcelamentoBoleto Then
        aviso "Emissão de notas com parcelamento em manutenção, favor utilizar o painel automático"
        formulario_limpa
        blnExisteParcelamentoBoleto = False
        Exit Sub
        End If
    
    NFe_emite False
    
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

Private Sub c_chave_nfe_ref_GotFocus()

    c_chave_nfe_ref.Height = c_chave_nfe_ref.Height * 3
    
    With c_chave_nfe_ref
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_chave_nfe_ref_KeyPress(KeyAscii As Integer)
   
Dim executa_tab As Boolean
Dim s As String
Dim c As String

    If KeyAscii = 13 Then
    '   COMO O CAMPO ACEITA MÚLTIPLAS LINHAS, SÓ VAI P/ O PRÓXIMO CAMPO APÓS 2 "ENTER's" CONSECUTIVOS
        executa_tab = True
    '   CURSOR ESTÁ NO FINAL DO TEXTO (IGNORA "ENTER's" SUBSEQUENTES NO TEXTO) ?
        s = Mid$(c_chave_nfe_ref.Text, c_chave_nfe_ref.SelStart + 1)
        s = Replace$(s, vbCr, "")
        s = Replace$(s, vbLf, "")
        s = Trim$(s)
        If s <> "" Then executa_tab = False
    '   CARACTER ANTERIOR É "ENTER" ?
        If c_chave_nfe_ref.SelStart > 0 Then
            c = Mid$(c_chave_nfe_ref.Text, c_chave_nfe_ref.SelStart, 1)
            If (c <> Chr$(13)) And (c <> Chr$(10)) Then executa_tab = False
            End If
        
        If executa_tab Then
            KeyAscii = 0
            c_produto_obs(0).SetFocus
            End If
        
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    
    If KeyAscii = 0 Then Exit Sub

End Sub

Private Sub c_chave_nfe_ref_LostFocus()

Dim lista() As String
Dim s As String
Dim i As Integer
Dim j As Integer
    
    c_chave_nfe_ref.Height = c_chave_nfe_ref.Height / 3
    

    c_chave_nfe_ref = Trim$(c_chave_nfe_ref)
    If c_chave_nfe_ref = "" Then Exit Sub
    
    lista = Split(c_chave_nfe_ref, vbCrLf)
    For i = LBound(lista) To UBound(lista)
        s = Trim$(lista(i))
        If s <> "" Then
            If Len(s) <> 44 Then
                aviso_erro "Tamanho inválido para a chave de acesso da NFe referenciada!!" & vbCrLf & _
                           "(" & s & ")"
                c_chave_nfe_ref.SetFocus
                Exit Sub
                End If
            
            If Not nfe_chave_acesso_ok(s) Then
                aviso_erro "Número inválido para a chave de acesso da NFe referenciada!!" & vbCrLf & _
                           "(" & s & ")"
                c_chave_nfe_ref.SetFocus
                Exit Sub
                End If
                
            For j = i + 1 To UBound(lista)
                If s = Trim$(lista(j)) Then
                    aviso_erro "NFe referenciada repetida!!" & vbCrLf & _
                           "(" & s & ")"
                c_chave_nfe_ref.SetFocus
                Exit Sub
                    End If
                Next
                
            End If
        Next
    
End Sub

Private Sub c_cnpj_cpf_dest_GotFocus()

    With c_cnpj_cpf_dest
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_CST_GotFocus(Index As Integer)

    With c_CST(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_CST_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_CFOP(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub

Private Sub c_CST_LostFocus(Index As Integer)

    c_CST(Index) = retorna_so_digitos(Trim$(c_CST(Index)))
    
    If Trim$(c_CST(Index)) = "" Then Exit Sub
    
    If Len(Trim$(c_CST(Index))) <> 3 Then
        aviso_erro "Código de CST inválido!!" & vbCrLf & "Informe o código de CST com 3 dígitos!!"
        c_CST(Index).SetFocus
        Exit Sub
        End If
        
    atualiza_valor_total_icms
    
End Sub

Private Sub c_dados_adicionais_GotFocus()

    If Trim$(c_dados_adicionais) <> "" Then
        With c_dados_adicionais
            .SelStart = 0
            .SelLength = Len(.Text)
            End With
        Exit Sub
        End If
        
End Sub



Private Sub c_dados_adicionais_KeyPress(KeyAscii As Integer)

'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0
    
End Sub


Private Sub c_dados_adicionais_LostFocus()

    c_dados_adicionais = RTrimCrLf(c_dados_adicionais)
    
'   Filtra caracter separador definido pela Target One
    c_dados_adicionais = Replace(c_dados_adicionais, "|", "/")
    
End Sub

Private Sub c_descricao_GotFocus(Index As Integer)

    With c_descricao(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_descricao_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_produto_obs(Index).SetFocus
        Exit Sub
        End If

'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub


Private Sub c_descricao_LostFocus(Index As Integer)

    c_descricao(Index) = Trim$(c_descricao(Index))
    
'   Filtra caracter separador definido pela Target One
    c_descricao(Index) = Replace(c_descricao(Index), "|", "/")

End Sub


Private Sub c_end_edicao_bairro_GotFocus()

    c_end_edicao_bairro.SelStart = 0
    c_end_edicao_bairro.SelLength = Len(c_end_edicao_bairro)

End Sub


Private Sub c_end_edicao_bairro_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_end_edicao_cidade.SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_end_edicao_cep_GotFocus()

    c_end_edicao_cep.SelStart = 0
    c_end_edicao_cep.SelLength = Len(c_end_edicao_cep)
    
End Sub


Private Sub c_end_edicao_cep_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Trim$(c_end_edicao_cep) <> "" Then
            c_end_edicao_logradouro.SetFocus
            trata_botao_pesquisa_cep
            End If
        Exit Sub
        End If

    KeyAscii = filtra_numerico(KeyAscii)

End Sub


Private Sub c_end_edicao_cep_LostFocus()
    
    If Trim(c_end_edicao_cep) <> "" Then
        c_end_edicao_cep = cep_formata(c_end_edicao_cep)
        End If
    
End Sub


Private Sub c_end_edicao_cidade_GotFocus()

    c_end_edicao_cidade.SelStart = 0
    c_end_edicao_cidade.SelLength = Len(c_end_edicao_cidade)

End Sub


Private Sub c_end_edicao_cidade_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_end_edicao_uf.SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_end_edicao_complemento_GotFocus()

    c_end_edicao_complemento.SelStart = 0
    c_end_edicao_complemento.SelLength = Len(c_end_edicao_complemento)

End Sub


Private Sub c_end_edicao_complemento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_end_edicao_bairro.SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_end_edicao_logradouro_GotFocus()

    c_end_edicao_logradouro.SelStart = 0
    c_end_edicao_logradouro.SelLength = Len(c_end_edicao_logradouro)

End Sub


Private Sub c_end_edicao_logradouro_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_end_edicao_numero.SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_end_edicao_numero_GotFocus()

    c_end_edicao_numero.SelStart = 0
    c_end_edicao_numero.SelLength = Len(c_end_edicao_numero)

End Sub


Private Sub c_end_edicao_numero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_end_edicao_complemento.SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_end_edicao_uf_GotFocus()

    c_end_edicao_uf.SelStart = 0
    c_end_edicao_uf.SelLength = Len(c_end_edicao_uf)
    
End Sub


Private Sub c_end_edicao_uf_KeyPress(KeyAscii As Integer)

    KeyAscii = filtra_letra(KeyAscii)
    If KeyAscii <> 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub c_end_edicao_uf_LostFocus()

    c_end_edicao_uf = UCase$(Trim$(c_end_edicao_uf))
    
    If c_end_edicao_uf <> "" Then
        If Not UF_ok(c_end_edicao_uf) Then
            aviso_erro "UF inválida!!"
            c_end_edicao_uf.SetFocus
            Exit Sub
            End If
        End If
        
End Sub


Private Sub c_fabricante_Change(Index As Integer)

    c_fabricante(Index).ForeColor = vbBlack
    
End Sub

Private Sub c_fabricante_GotFocus(Index As Integer)

    With c_fabricante(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_fabricante_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_produto(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
End Sub


Private Sub c_fabricante_LostFocus(Index As Integer)
Const NomeDestaRotina = "c_fabricante_LostFocus()"
Dim s As String
Dim t As ADODB.Recordset
    
    On Error GoTo C_FABRICANTE_LF_TRATA_ERRO

    If Trim$(c_fabricante(Index)) = "" Then Exit Sub
    
    If Not edicao_manual_liberada Then
        c_fabricante(Index) = normaliza_codigo(Trim$(c_fabricante(Index)), TAM_MIN_FABRICANTE)
        End If
    
'   VERIFICA SE FABRICANTE ESTÁ CADASTRADO
    aguarde INFO_EXECUTANDO, "pesquisando fabricante no banco de dados"
    
'   T_FABRICANTE
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT" & _
            " fabricante," & _
            " nome," & _
            " razao_social" & _
        " FROM t_FABRICANTE" & _
        " WHERE" & _
            " (fabricante = '" & Trim$(c_fabricante(Index)) & "')"
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        c_fabricante(Index).ForeColor = vbBlack
    Else
        c_fabricante(Index).ForeColor = vbRed
        aguarde INFO_NORMAL, m_id
        If Not edicao_manual_liberada And (left$(Trim(cb_finalidade), 1) <> NFE_FINALIDADE_NFE_COMPLEMENTAR) Then
            aviso_erro "Fabricante não cadastrado!!"
            c_fabricante(Index).SetFocus
            End If
        GoSub C_FABRICANTE_LF_FECHA_TABELAS
        Exit Sub
        End If
    
    GoSub C_FABRICANTE_LF_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_FABRICANTE_LF_FECHA_TABELAS:
'=============================
    bd_desaloca_recordset t, True
    Return
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_FABRICANTE_LF_TRATA_ERRO:
'==========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub C_FABRICANTE_LF_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub


Private Sub c_fcp_GotFocus(Index As Integer)
    
    With c_fcp(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_fcp_KeyPress(Index As Integer, KeyAscii As Integer)

       If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index < 11) Then c_fabricante(Index + 1).SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_perc(c_fcp(Index), KeyAscii)

End Sub

Private Sub c_fcp_LostFocus(Index As Integer)
    Dim i As Single

    If IsNumeric(c_fcp(Index)) Then
        i = CSng(c_fcp(Index))
        If (i < 0) Or (i > 2) Then
            aviso_erro "Alíquota de Fundo de Combate à Pobreza inválida (máximo: 2%)!!"
            c_fcp(Index).SetFocus
            Exit Sub
            End If
        
        c_fcp(Index) = Format$(i, FORMATO_PERCENTUAL)
        End If

End Sub


Private Sub c_nItemPed_GotFocus(Index As Integer)
   
   With c_nItemPed(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
 
End Sub

Private Sub c_nItemPed_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        c_unidade(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub

Private Sub c_nItemPed_LostFocus(Index As Integer)
    
    c_nItemPed(Index) = Trim$(c_nItemPed(Index))
        
End Sub

Private Sub c_ipi_GotFocus()

    With c_ipi
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
        
End Sub

Private Sub c_ipi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_frete.SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_perc(c_ipi, KeyAscii)
    
End Sub




Private Sub c_ipi_LostFocus()
Dim i As Single

    If IsNumeric(c_ipi) Then
        i = CSng(c_ipi)
        If (i < 0) Or (i > 100) Then
            aviso_erro "Alíquota de IPI inválida !!"
            c_ipi.SetFocus
            Exit Sub
            End If
        
        c_ipi = Format$(i, FORMATO_PERCENTUAL)
        End If
        
End Sub



Private Sub c_NCM_GotFocus(Index As Integer)

    With c_NCM(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_NCM_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_ICMS_item(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub


Private Sub c_NCM_LostFocus(Index As Integer)

    c_NCM(Index) = Trim$(c_NCM(Index))
    
    If Trim$(c_NCM(Index)) = "" Then Exit Sub
    
    If (Len(Trim$(c_NCM(Index))) <> 8) And (Len(Trim$(c_NCM(Index))) <> 2) Then
        aviso_erro "Código de NCM inválido!!" & vbCrLf & "Informe o código de NCM com 8 dígitos ou 2 dígitos (gênero)!!"
        c_NCM(Index).SetFocus
        Exit Sub
        End If

End Sub


Private Sub c_num_nfe_danfe_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim s As String
Dim s_aux As String
Dim c As String
Dim intEmitente As Integer

    If KeyAscii = 13 Then
        KeyAscii = 0
        b_danfe.SetFocus
        If (cb_emitente_danfe.ListIndex <> -1) And (Trim$(cb_emitente_danfe) <> "") And (Trim$(c_num_serie_danfe) <> "") And (Trim$(c_num_nfe_danfe) <> "") Then
            s = cb_emitente_danfe
            s_aux = ""
            For i = 1 To Len(s)
                c = Mid$(s, i, 1)
                If Not IsNumeric(c) Then Exit For
                s_aux = s_aux & c
                Next
            intEmitente = CInt(s_aux)
            
            DANFE_CONSULTA_parametro_emitente intEmitente, CInt(c_num_serie_danfe), CLng(c_num_nfe_danfe)
            
            End If
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
End Sub


Private Sub c_num_nfe_danfe_LostFocus()

    c_num_nfe_danfe = Trim$(c_num_nfe_danfe)

End Sub


Private Sub c_cnpj_cpf_dest_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_tipo_NF.SetFocus
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_cnpj_cpf(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
End Sub


Private Sub c_cnpj_cpf_dest_LostFocus()
Const NomeDestaRotina = "c_cnpj_cpf_dest_LostFocus()"
Dim s As String
Dim t As ADODB.Recordset
Dim strDDD As String
Dim strTelRes As String
Dim strTelCel As String
Dim strTelCom As String
Dim strTelCom2 As String
Dim strRamal As String
Dim strSufixoRes As String
Dim strSufixoCom As String
Dim strInfoIE As String

    On Error GoTo C_CNPJ_CPF_DEST_LF_TRATA_ERRO
    
    c_nome_dest = ""
    l_IE = ""
    
    If Trim$(c_cnpj_cpf_dest) = "" Then Exit Sub
    
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
    
    'COM A FUNCIONALIDADE DE EMITIR NOTAS REFERENTES À ASSISTÊNCIA TÉCNICA, PODEMOS TER QUE SELECIONAR
    'OS DADOS DO CLIENTE VIA CONEXÃO COM O BANCO AssistenciaTecnica (dbcAssist)
    'PARA ISTO, VERIFICAREMOS SE O CNPJ/CPF DIGITADO É IGUAL AO ÚLTIMO CNPJ/CPF OBTIDO DA ASSISTÊNCIA
    'SE SIM, A PROCURA SERÁ VIA dbcAssist, SENÃO, SERÁ VIA dbc
    If bln_assist_pedido_ok And _
       (retorna_so_digitos(c_cnpj_cpf_dest) = retorna_so_digitos(s_assist_Cliente_CNPJ_CPF)) Then
        t.Open s, dbcAssist, , , adCmdText
    Else
        t.Open s, dbc, , , adCmdText
        bln_assist_pedido_ok = False
        End If
    
    If t.EOF Then
        aguarde INFO_NORMAL, m_id
        aviso_erro "CNPJ/CPF não cadastrado!!"
        c_cnpj_cpf_dest.SetFocus
        GoSub C_CNPJ_CPF_DEST_LF_FECHA_TABELAS
        Exit Sub
        End If
    
    c_nome_dest = Trim$("" & t("nome"))
    
    limpa_dados_endereco_cadastro
    limpa_dados_endereco_editado
    limpa_campos_endereco_edicao
    atualiza_dados_endereco_cadastro
    c_dados_adicionais = ""
    
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
    If s <> "" Then c_dados_adicionais = s
    
'   INFORMAÇÃO SE É CONTRIBUINTE DE ICMS
    If UCase$(Trim$("" & t("tipo"))) = ID_PJ Then
        Select Case t("contribuinte_icms_status")
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO: strInfoIE = "NC"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM: strInfoIE = "C"
            Case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO: strInfoIE = "I"
            Case Else: strInfoIE = ""
            End Select
    Else
        Select Case t("produtor_rural_status")
            Case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM: strInfoIE = "PR"
            Case Else: strInfoIE = ""
            End Select
        End If
    l_IE.Caption = strInfoIE

    
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

Private Sub c_num_serie_danfe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_num_nfe_danfe.SetFocus
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_num_serie_danfe_LostFocus()

    c_num_serie_danfe = Trim$(c_num_serie_danfe)
    
End Sub


Private Sub c_pedido_nota_GotFocus()

    With c_pedido_nota
        .SelStart = 0
        .SelLength = Len(.Text)
        End With


End Sub

Private Sub c_pedido_nota_KeyPress(KeyAscii As Integer)

Dim executa_tab As Boolean
Dim s As String
Dim c As String

    If KeyAscii = 13 Then
    '   COMO O CAMPO ACEITA MÚLTIPLAS LINHAS, SÓ VAI P/ O PRÓXIMO CAMPO APÓS 2 "ENTER's" CONSECUTIVOS
        executa_tab = True
    '   CURSOR ESTÁ NO FINAL DO TEXTO (IGNORA "ENTER's" SUBSEQUENTES NO TEXTO) ?
        s = Mid$(c_pedido_nota.Text, c_pedido_nota.SelStart + 1)
        s = Replace$(s, vbCr, "")
        s = Replace$(s, vbLf, "")
        s = Trim$(s)
        If s <> "" Then executa_tab = False
    '   CARACTER ANTERIOR É "ENTER" ?
        If c_pedido_nota.SelStart > 0 Then
            c = Mid$(c_pedido_nota.Text, c_pedido_nota.SelStart, 1)
            If (c <> Chr$(13)) And (c <> Chr$(10)) Then executa_tab = False
            End If
        
        If Not c_pedido_nota.MultiLine Then
            c_pedido_nota = normaliza_num_pedido(c_pedido_nota)
            If Len(c_pedido_nota) > 0 Then
                c_pedido_nota.SelStart = Len(c_pedido_nota)
            Else
                End If
            executa_tab = True
            End If
        
        If executa_tab Then
            KeyAscii = 0
            c_dados_adicionais.SetFocus
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

Private Sub c_pedido_nota_KeyUp(KeyCode As Integer, Shift As Integer)
Dim s_erro As String

    If Len(c_pedido_nota) > 0 Then
        If blnAtualizaNFnoPedido Then
            pn_pedido_nota.Visible = True
            pn_aviso_pedido_nota.Visible = False
        Else
            pn_pedido_nota.Visible = False
            pn_aviso_pedido_nota.Visible = True
            End If
    Else
        pn_pedido_nota.Visible = False
        pn_aviso_pedido_nota.Visible = False
        End If

End Sub

Private Sub c_pedido_nota_LostFocus()

Dim s_erro As String

    c_pedido_nota = normaliza_num_pedido(Trim$(c_pedido_nota))

    If Len(c_pedido_nota) <= 0 Then
        pn_pedido_nota.Visible = False
        pn_aviso_pedido_nota.Visible = False
        End If
    
    pnParcelasEmBoletos.Visible = False
    
    If Len(c_pedido_nota) > 0 Then
        'verificar se existe informação de parcelas em boleto
        If (param_geracaoboletos.campo_texto = "Manual") Then
            If c_pedido_nota <> "" Then
                ReDim v_pedido_manual_boleto(0)
                v_pedido_manual_boleto(UBound(v_pedido_manual_boleto)) = c_pedido_nota
                blnExisteParcelamentoBoleto = False
                'pnParcelasEmBoletos.Visible = False
            If ExisteDadosParcelasPagto(c_pedido_nota, s_erro) And _
                consultaDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
                    AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
                    If blnExisteParcelamentoBoleto Then
                        pnParcelasEmBoletos.Visible = True
                        pnParcelasEmBoletos.Enabled = False
                        c_dataparc.Enabled = False
                        pn_pedido_nota.Visible = True
                        pn_aviso_pedido_nota.Visible = True
                        opVenda.Value = True
                        End If
                ElseIf geraDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
                    AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
                    If blnExisteParcelamentoBoleto Then
                        pnParcelasEmBoletos.Visible = True
                        pnParcelasEmBoletos.Enabled = True
                        c_dataparc.Enabled = True
                        pn_pedido_nota.Visible = True
                        pn_aviso_pedido_nota.Visible = True
                        opVendaFutura.Value = True
                        End If
                Else
                    aviso "Erro: " & s_erro
                    End If
                                    
                End If
            End If
        End If

End Sub

Private Sub c_produto_Change(Index As Integer)

    c_produto(Index).ForeColor = vbBlack

End Sub

Private Sub c_produto_GotFocus(Index As Integer)

    With c_produto(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_produto_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Trim$(c_produto(Index)) = "" Then
            c_pedido_nota.SetFocus
        Else
            If edicao_manual_liberada Then
                c_descricao(Index).SetFocus
            Else
                c_produto_obs(Index).SetFocus
                End If
            End If
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_produto_LostFocus(Index As Integer)
Const NomeDestaRotina = "c_produto_LostFocus()"
Dim s As String
Dim t As ADODB.Recordset
    
    On Error GoTo C_PRODUTO_LF_TRATA_ERRO
    
    If Trim$(c_produto(Index)) = "" Then Exit Sub
    
    If Not edicao_manual_liberada Then
        c_produto(Index) = normaliza_codigo(Trim$(c_produto(Index)), TAM_MIN_PRODUTO)
        End If

'   BUSCA PELA DESCRIÇÃO DO PRODUTO
    aguarde INFO_EXECUTANDO, "pesquisando produto no banco de dados"
    
  ' T_PRODUTO
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = "SELECT" & _
            " fabricante," & _
            " descricao," & _
            " qtde_volumes," & _
            " peso," & _
            " (" & _
                "SELECT TOP 1 cst FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=t_PRODUTO.fabricante) AND (tEI.produto=t_PRODUTO.produto) ORDER BY id_estoque DESC" & _
            ") AS cst," & _
            " (" & _
                "SELECT TOP 1 ncm FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=t_PRODUTO.fabricante) AND (tEI.produto=t_PRODUTO.produto) ORDER BY id_estoque DESC" & _
            ") AS ncm" & _
        " FROM t_PRODUTO" & _
        " WHERE" & _
            " (produto = '" & Trim$(c_produto(Index)) & "')"
    If Trim$(c_fabricante(Index)) <> "" Then
        s = s & " AND (fabricante = '" & Trim$(c_fabricante(Index)) & "')"
        End If
        
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        c_produto(Index).ForeColor = vbBlack
    Else
        c_produto(Index).ForeColor = vbRed
        aguarde INFO_NORMAL, m_id
        If Not edicao_manual_liberada Then
            aviso_erro "Produto não cadastrado!!"
            c_produto(Index).SetFocus
            End If
        GoSub C_PRODUTO_LF_FECHA_TABELAS
        Exit Sub
        End If
        
    If t.RecordCount > 1 Then
        aguarde INFO_NORMAL, m_id
        If Not edicao_manual_liberada Then
            aviso_erro "Há mais de 1 produto com o código " & c_produto(Index) & "!!"
            c_produto(Index).SetFocus
            End If
        GoSub C_PRODUTO_LF_FECHA_TABELAS
        Exit Sub
        End If
    
    v_totalizacao_itens(Index).qtde_volumes = 0
    v_totalizacao_itens(Index).peso_bruto = 0
    v_totalizacao_itens(Index).peso_liquido = 0
    
    If Not t.EOF Then
        c_descricao(Index) = Trim$("" & t("descricao"))
        If Trim$(c_fabricante(Index)) = "" Then c_fabricante(Index) = Trim$("" & t("fabricante"))
        c_CST(Index) = cst_converte_codigo_entrada_para_saida(Trim$("" & t("cst")))
        c_NCM(Index) = Trim$("" & t("ncm"))
        
        If IsNumeric(t("qtde_volumes")) Then
            v_totalizacao_itens(Index).qtde_volumes = t("qtde_volumes")
            End If
        
        If IsNumeric(t("peso")) Then
            v_totalizacao_itens(Index).peso_bruto = t("peso")
            v_totalizacao_itens(Index).peso_liquido = t("peso")
            End If
            
        c_unidade(Index) = "PC"
        End If
    
    recalcula_totalizacao_itens
    
    atualiza_valor_total_icms
    
    GoSub C_PRODUTO_LF_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_PRODUTO_LF_FECHA_TABELAS:
'==========================
    bd_desaloca_recordset t, True
    Return
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_PRODUTO_LF_TRATA_ERRO:
'=======================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub C_PRODUTO_LF_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Private Sub c_produto_obs_GotFocus(Index As Integer)

    With c_produto_obs(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_produto_obs_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_qtde(Index).SetFocus
        Exit Sub
        End If

'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub


Private Sub c_produto_obs_LostFocus(Index As Integer)

    c_produto_obs(Index) = Trim$(c_produto_obs(Index))
    
'   Filtra caracter separador definido pela Target One
    c_produto_obs(Index) = Replace(c_produto_obs(Index), "|", "/")

End Sub


Private Sub c_qtde_GotFocus(Index As Integer)

    With c_qtde(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_qtde_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_vl_unitario(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_qtde_LostFocus(Index As Integer)

    c_qtde(Index) = Trim$(c_qtde(Index))
    
    recalcula_totais
    recalcula_totalizacao_itens
    
End Sub


Private Sub c_total_peso_bruto_GotFocus()

    With c_total_peso_bruto
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_total_peso_bruto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        c_total_peso_liquido.SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
        
    KeyAscii = filtra_peso(KeyAscii)

End Sub


Private Sub c_total_peso_bruto_LostFocus()
Dim s As String
Dim n As Single

    On Error GoTo C_TOTAL_PESO_BRUTO_LOSTFOCUS_TRATA_ERRO

    c_total_peso_bruto = formata_numero_3dec(converte_para_single(Trim$(c_total_peso_bruto)))
    If c_total_peso_bruto <> "" Then
        n = converte_para_single(c_total_peso_bruto)
        If n < 0 Then
            aviso_erro "O peso não pode ser negativo!!"
            c_total_peso_bruto.SetFocus
            Exit Sub
            End If
        End If
        
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_TOTAL_PESO_BRUTO_LOSTFOCUS_TRATA_ERRO:
'=======================================
    s = "O peso informado é inválido!!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    c_total_peso_bruto.SetFocus
    Exit Sub

End Sub


Private Sub c_total_peso_liquido_GotFocus()

    With c_total_peso_liquido
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_total_peso_liquido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_pedido_nota.SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
        
    KeyAscii = filtra_peso(KeyAscii)

End Sub


Private Sub c_total_peso_liquido_LostFocus()
Dim s As String
Dim n As Single

    On Error GoTo C_TOTAL_PESO_LIQUIDO_LOSTFOCUS_TRATA_ERRO

    c_total_peso_liquido = formata_numero_3dec(converte_para_single(Trim$(c_total_peso_liquido)))
    If c_total_peso_liquido <> "" Then
        n = converte_para_single(c_total_peso_liquido)
        If n < 0 Then
            aviso_erro "O peso não pode ser negativo!!"
            c_total_peso_liquido.SetFocus
            Exit Sub
            End If
        End If
        
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_TOTAL_PESO_LIQUIDO_LOSTFOCUS_TRATA_ERRO:
'=========================================
    s = "O peso informado é inválido!!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    c_total_peso_liquido.SetFocus
    Exit Sub

End Sub



Private Sub c_total_volumes_GotFocus()

    With c_total_volumes
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_total_volumes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_total_peso_bruto.SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_numerico(KeyAscii)

End Sub


Private Sub c_total_volumes_LostFocus()
Dim s As String
Dim i As Long

    On Error GoTo C_TOTAL_VOLUMES_LOSTFOCUS_TRATA_ERRO

    c_total_volumes = Trim$(c_total_volumes)
    If c_total_volumes <> "" Then
        i = CLng(c_total_volumes)
        If i < 0 Then
            aviso_erro "A quantidade não pode ser negativa!!"
            c_total_volumes.SetFocus
            Exit Sub
        ElseIf CStr(i) <> c_total_volumes Then
        '   LEMBRANDO QUE:
        '       CLng("1.5") = 15
        '       CLng("1,5") = 2
            aviso_erro "Número informado possui formato inválido para este campo!!"
            c_total_volumes.SetFocus
            Exit Sub
            End If
        End If
        
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
C_TOTAL_VOLUMES_LOSTFOCUS_TRATA_ERRO:
'====================================
    s = "A quantidade informada é inválida!!" & vbCrLf & CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    c_total_volumes.SetFocus
    Exit Sub

End Sub


Private Sub c_unidade_GotFocus(Index As Integer)

    With c_unidade(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_unidade_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = c_fabricante.UBound Then
            c_pedido_nota.SetFocus
        Else
            c_fabricante(Index + 1).SetFocus
            End If
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub

Private Sub c_unidade_LostFocus(Index As Integer)

    c_unidade(Index) = Trim$(c_unidade(Index))

End Sub

Private Sub c_vl_outras_despesas_acessorias_GotFocus(Index As Integer)

    With c_vl_outras_despesas_acessorias(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub c_vl_outras_despesas_acessorias_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_CST(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_moeda(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_vl_outras_despesas_acessorias_LostFocus(Index As Integer)

    If Trim$(c_vl_outras_despesas_acessorias(Index)) = "" Then
        recalcula_totais
        Exit Sub
        End If
    
    c_vl_outras_despesas_acessorias(Index) = formata_moeda(c_vl_outras_despesas_acessorias(Index))
    
    recalcula_totais

End Sub


Private Sub c_vl_total_Change(Index As Integer)

    atualiza_valor_total_icms
    
End Sub

Private Sub c_vl_unitario_GotFocus(Index As Integer)

    With c_vl_unitario(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_vl_unitario_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_vl_outras_despesas_acessorias(Index).SetFocus
        Exit Sub
        End If
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_moeda(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_vl_unitario_LostFocus(Index As Integer)

    If Trim$(c_vl_unitario(Index)) = "" Then
        recalcula_totais
        Exit Sub
        End If
    
    c_vl_unitario(Index) = formata_moeda(c_vl_unitario(Index))
    
    recalcula_totais
    
End Sub


Private Sub c_xPed_GotFocus(Index As Integer)

    With c_xPed(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
    
End Sub


Private Sub c_xPed_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_nItemPed(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0

End Sub

Private Sub c_xPed_LostFocus(Index As Integer)

    c_xPed(Index) = Trim$(c_xPed(Index))

End Sub


Private Sub cb_CFOP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        cb_CFOP(Index).ListIndex = -1
        End If
        
End Sub


Private Sub cb_CFOP_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_NCM(Index).SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_emitente_Click()
Dim i As Integer
Dim sAliquotaEmit As String
Dim blnAchouEmit As Boolean

    'iguala o emitente da DANFE
    i = cb_emitente.ListIndex
    cb_emitente_danfe.ListIndex = i

    'localizar as informações do emitente selecionado
    blnAchouEmit = False
    For i = LBound(vEmitsUsuario) To UBound(vEmitsUsuario)
        If vEmitsUsuario(i).c2 = left(cb_emitente, Len(vEmitsUsuario(i).c2)) Then
            usuario.emit = Mid$(vEmitsUsuario(i).c1, 1, Len(vEmitsUsuario(i).c1) - 5)
            usuario.emit_uf = Mid$(vEmitsUsuario(i).c1, Len(vEmitsUsuario(i).c1) - 2, 2)
            usuario.emit_id = vEmitsUsuario(i).c2
            txtFixoEspecifico = vEmitsUsuario(i).c3
            blnAchouEmit = True
            Exit For
            End If
        Next
    If Not blnAchouEmit Then
        aviso_erro "Problemas na localização das informações do emitente!"
        cb_emitente.SetFocus
        Exit Sub
        End If

    'ajusta o ICMS de acordo com a UF do depósito
    Select Case usuario.emit_uf
        Case "ES": sAliquotaEmit = "12"
        Case "MG": sAliquotaEmit = "18"
        Case "MS": sAliquotaEmit = "17"
        Case "RJ": sAliquotaEmit = "20"
        Case "SP": sAliquotaEmit = "18"
        Case "TO": sAliquotaEmit = "18"
        Case Else: sAliquotaEmit = "18"
        End Select
    
    For i = 0 To cb_icms.ListCount - 1
        If cb_icms.List(i) = sAliquotaEmit Then
            cb_icms.ListIndex = i
            Exit For
            End If
        Next
        
'   EXIBIR UF DO EMITENTE SELECIONADO NO LABEL EM DESTAQUE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    l_emitente_uf.Caption = usuario.emit_uf
        
    
End Sub

Private Sub cb_emitente_danfe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_num_serie_danfe.SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_emitente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_cnpj_cpf_dest.SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_finalidade_Click()
Const INDICE_COMPLEMENTO_ICMS = 2

    If cb_finalidade.ListIndex = INDICE_COMPLEMENTO_ICMS Then
        trata_tela_complemento_ICMS
        End If

End Sub

Private Sub cb_frete_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_transportadora.SetFocus
        Exit Sub
        End If
        
End Sub


Private Sub cb_icms_GotFocus()

    With cb_icms
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub



Private Sub cb_ICMS_item_GotFocus(Index As Integer)

    With cb_ICMS_item(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub


Private Sub cb_ICMS_item_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_xPed(Index).SetFocus
        Exit Sub
        End If

    KeyAscii = filtra_numerico(KeyAscii)

End Sub


Private Sub cb_ICMS_item_LostFocus(Index As Integer)

Dim i As Single

    If cb_ICMS_item(Index) <> "" Then
        If IsNumeric(cb_ICMS_item(Index)) Then
            i = CSng(cb_ICMS_item(Index))
            If (i < 0) Or (i > 100) Then
                aviso_erro "Alíquota de ICMS inválida!!"
                cb_ICMS_item(Index).SetFocus
                Exit Sub
                End If
            
            cb_ICMS_item(Index) = CStr(i)
            End If
        End If

    atualiza_valor_total_icms
    
End Sub


Private Sub cb_icms_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_ipi.SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_numerico(KeyAscii)
    
End Sub


Private Sub cb_icms_LostFocus()

Dim i As Single

    If cb_icms <> "" Then
        If IsNumeric(cb_icms) Then
            i = CSng(cb_icms)
            If (i < 0) Or (i > 100) Then
                aviso_erro "Alíquota de ICMS inválida !!"
                cb_icms.SetFocus
                Exit Sub
                End If
            
            cb_icms = CStr(i)
            End If
        End If
        
    atualiza_valor_total_icms
        
End Sub

Private Sub cb_indpres_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_natureza.SetFocus
        Exit Sub
        End If
        
End Sub

Private Sub cb_loc_dest_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        c_fabricante(0).SetFocus
        Exit Sub
        End If

End Sub

Private Sub cb_natureza_Click()
    ' - Se o código de natureza da operação inicia com 1 ou 5, trata-se de uma operação interna;
    '   se o código de natureza da operação inicia com 2 ou 6, trata-se de uma operação interestadual
    ' - Verificar também se o CFOP é de remessa ou se a descrição contém a palavra "remessa", para decidir
    '   onde incluir o número da nf em t_PEDIDO, se for o caso
    Dim digito As String
    Dim s_cfop As String
    
    digito = left(Trim(cb_natureza.Text), 1)
    If (digito = "1") Or (digito = "5") Then cb_loc_dest.ListIndex = 0
    If (digito = "2") Or (digito = "6") Then cb_loc_dest.ListIndex = 1

    s_cfop = left(Trim(cb_natureza.Text), 5)
    
    If cfop_eh_de_remessa(retorna_so_digitos(left(cb_natureza.Text, 5))) Or _
        (InStr(UCase(cb_natureza), "REMESSA") > 0) Then
        opRemessa.Value = True
    ElseIf ((s_cfop = "5.922") Or (s_cfop = "6.922")) Then
        opVendaFutura.Value = True
    Else
        opVenda.Value = True
        End If
    
    If (s_cfop = "5.915") Or (s_cfop = "6.152") Or (s_cfop = "5.949") Or (s_cfop = "6.949") Or _
       (s_cfop = "6.117") Or (s_cfop = "6.923") Or (s_cfop = "6.910") Then
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
        cb_icms.SetFocus
        Exit Sub
        End If
        
End Sub

Private Sub cb_tipo_NF_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_indpres.SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_transportadora_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_loc_dest.SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_zerar_COFINS_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        cb_zerar_COFINS.ListIndex = 0
        End If

End Sub


Private Sub cb_zerar_COFINS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_fabricante(0).SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_zerar_PIS_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        cb_zerar_PIS.ListIndex = 0
        End If

End Sub


Private Sub cb_zerar_PIS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cb_zerar_COFINS.SetFocus
        Exit Sub
        End If

End Sub


Private Sub Form_Activate()
Const NomeDestaRotina = "Form_Activate()"
Dim s As String

    On Error GoTo FORMACTIVATE_TRATA_ERRO

    If Not modulo_inicializacao_ok Then
        
      ' OK !!
        modulo_inicializacao_ok = True
        
        edicao_manual_liberada = False
        
        tab_stop_configura
        
        ReDim v_totalizacao_itens(c_produto.LBound To c_produto.UBound)
        
        relogio_Timer
        
        aguarde INFO_EXECUTANDO, "iniciando aplicativo"
                   
    '   PREPARA CAMPOS/CARREGA DADOS INICIAIS
        formulario_inicia
        
    '   LIMPA CAMPOS/POSICIONA DEFAULTS
        formulario_limpa
    
        c_num_serie_danfe = ""
        c_num_nfe_danfe = ""
        
    '   DADOS DA ÚLTIMA NFe EMITIDA
        l_serie_NF = f_MAIN.l_serie_NF
        l_num_NF = f_MAIN.l_num_NF
        l_emitente_NF = f_MAIN.l_emitente_NF
        
        Caption = Caption & " v" & m_id_versao
        
        If DESENVOLVIMENTO Then
            Caption = Caption & "  (Versão Exclusiva de Desenvolvimento/Homologação)"
            End If
        
        If cor_fundo_padrao <> "" Then
            Me.BackColor = cor_fundo_padrao
            End If
        
    '   EXIBIR UF DO EMITENTE SELECIONADO NO LABEL EM DESTAQUE
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        l_emitente_uf.Caption = usuario.emit_uf
        
        
    '   VARIÁVEL PARA VERIFICAR SE ATUALIZA NF NO PEDIDO
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        blnAtualizaNFnoPedido = param_atualizanfnopedido.campo_inteiro = 1
        pn_pedido_nota.Visible = False
        pn_aviso_pedido_nota.Visible = False
        
        aguarde INFO_NORMAL, m_id
        End If
    
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FORMACTIVATE_TRATA_ERRO:
'=======================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    Err.Clear
    aviso_erro s
    Exit Sub


End Sub

Private Sub Form_Click()

    b_dummy.SetFocus
    
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

'   EM EXECUÇÃO ?
    If em_execucao Then
        Cancel = True
        Exit Sub
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

Private Sub mnu_download_pdf_danfe_Click()
        executa_download_pdf_danfe_parametro_emitente Me

End Sub

Private Sub mnu_download_pdf_danfe_periodo_Click()
    
    executa_download_pdf_danfe_periodo_parametro_emitente Me
    
End Sub


Private Sub mnu_emissao_automatica_Click()

    fechar_modo_emissao_manual
    
End Sub

Private Sub mnu_FECHAR_Click()

    fechar_programa
    
End Sub

Private Sub picEndereco_Click()
    
    If usar_endereco_editado Then mostra_endereco_editado
    
End Sub

Private Sub relogio_Timer()
Dim s As String

    s = left$(Time$, 5)
    If Val(right$(Time$, 1)) Mod 2 Then Mid$(s, 3, 1) = " "
    agora = s

    hoje = Format$(Date, "dd/mm/yyyy")
    
End Sub


