VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form f_EMISSAO_NFE_COMPLEMENTAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal"
   ClientHeight    =   10380
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   15270
   Icon            =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_dummy 
      Appearance      =   0  'Flat
      Caption         =   "b_dummy"
      Height          =   345
      Left            =   7995
      TabIndex        =   138
      Top             =   -525
      Width           =   1350
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
      ItemData        =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0442
      Left            =   7830
      List            =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   131
      Top             =   1065
      Width           =   7290
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
      TabIndex        =   130
      Top             =   1695
      Width           =   975
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
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   129
      Top             =   1695
      Width           =   1020
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
      Left            =   3570
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   128
      Top             =   1695
      Width           =   3840
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
      TabIndex        =   127
      Top             =   1065
      Width           =   7290
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
      ItemData        =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0446
      Left            =   7830
      List            =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0448
      Style           =   2  'Dropdown List
      TabIndex        =   126
      Top             =   1695
      Width           =   7290
   End
   Begin TabDlg.SSTab tabNFe 
      Height          =   4425
      Left            =   120
      TabIndex        =   32
      Top             =   2220
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   7805
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Destinatário"
      TabPicture(0)   =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":044A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Itens"
      TabPicture(1)   =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0466
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "l_tit_fabricante"
      Tab(1).Control(1)=   "l_tit_produto"
      Tab(1).Control(2)=   "l_tit_descricao"
      Tab(1).Control(3)=   "l_tit_qtde"
      Tab(1).Control(4)=   "l_tit_vl_unitario"
      Tab(1).Control(5)=   "l_tit_vl_total"
      Tab(1).Control(6)=   "l_tit_vl_total_geral"
      Tab(1).Control(7)=   "l_tit_produto_obs"
      Tab(1).Control(8)=   "c_fabricante(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "c_produto(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "c_descricao(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "c_qtde(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "c_vl_unitario(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "c_vl_total(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "c_fabricante(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "c_produto(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "c_descricao(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "c_qtde(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "c_vl_unitario(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "c_vl_total(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "c_fabricante(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "c_produto(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "c_descricao(2)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "c_qtde(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "c_vl_unitario(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "c_vl_total(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "c_fabricante(3)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "c_produto(3)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "c_descricao(3)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "c_qtde(3)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "c_vl_unitario(3)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "c_vl_total(3)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "c_fabricante(4)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "c_produto(4)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "c_descricao(4)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "c_qtde(4)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "c_vl_unitario(4)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "c_vl_total(4)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "c_fabricante(5)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "c_produto(5)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "c_descricao(5)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "c_qtde(5)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "c_vl_unitario(5)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "c_vl_total(5)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "c_fabricante(6)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "c_produto(6)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "c_descricao(6)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "c_qtde(6)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "c_vl_unitario(6)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "c_vl_total(6)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "c_fabricante(7)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "c_produto(7)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "c_descricao(7)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "c_qtde(7)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "c_vl_unitario(7)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "c_vl_total(7)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "c_fabricante(8)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "c_produto(8)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "c_descricao(8)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "c_qtde(8)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "c_vl_unitario(8)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "c_vl_total(8)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "c_fabricante(9)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "c_produto(9)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "c_descricao(9)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "c_qtde(9)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "c_vl_unitario(9)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "c_vl_total(9)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "c_vl_total_geral"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "c_produto_obs(0)"
      Tab(1).Control(70)=   "c_produto_obs(1)"
      Tab(1).Control(71)=   "c_produto_obs(2)"
      Tab(1).Control(72)=   "c_produto_obs(3)"
      Tab(1).Control(73)=   "c_produto_obs(4)"
      Tab(1).Control(74)=   "c_produto_obs(5)"
      Tab(1).Control(75)=   "c_produto_obs(6)"
      Tab(1).Control(76)=   "c_produto_obs(7)"
      Tab(1).Control(77)=   "c_produto_obs(8)"
      Tab(1).Control(78)=   "c_produto_obs(9)"
      Tab(1).Control(79)=   "c_produto_obs(10)"
      Tab(1).Control(80)=   "c_vl_total(10)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "c_vl_unitario(10)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "c_qtde(10)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "c_descricao(10)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "c_produto(10)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "c_fabricante(10)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "c_produto_obs(11)"
      Tab(1).Control(87)=   "c_vl_total(11)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "c_vl_unitario(11)"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "c_qtde(11)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "c_descricao(11)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "c_produto(11)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "c_fabricante(11)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).ControlCount=   93
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":0482
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":049E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   3720
         Width           =   525
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3720
         Width           =   885
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3720
         Width           =   4485
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   11
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   111
         Top             =   3720
         Width           =   5505
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3435
         Width           =   525
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3435
         Width           =   885
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3435
         Width           =   4485
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   3435
         Width           =   615
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   3435
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   3435
         Width           =   1305
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   10
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   104
         Top             =   3435
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   9
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   103
         Top             =   3150
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   8
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   102
         Top             =   2865
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   7
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   101
         Top             =   2580
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   6
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   100
         Top             =   2295
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   5
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   99
         Top             =   2010
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   4
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   98
         Top             =   1725
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   3
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   97
         Top             =   1440
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   2
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   96
         Top             =   1155
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   1
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   95
         Top             =   870
         Width           =   5505
      End
      Begin VB.TextBox c_produto_obs 
         Height          =   285
         Index           =   0
         Left            =   -68895
         MaxLength       =   500
         TabIndex        =   94
         Top             =   585
         Width           =   5505
      End
      Begin VB.TextBox c_vl_total_geral 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -61470
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   4005
         Width           =   1305
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3150
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   3150
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   3150
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3150
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2865
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2865
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2865
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2865
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   2865
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2865
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2580
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2580
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2580
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2295
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2295
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2295
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2295
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2010
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2010
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2010
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2010
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2010
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2010
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1725
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1725
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1725
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1725
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1440
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1155
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1155
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1155
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1155
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   870
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   870
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   870
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   870
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   870
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   870
         Width           =   525
      End
      Begin VB.TextBox c_vl_total 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -61470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   585
         Width           =   1305
      End
      Begin VB.TextBox c_vl_unitario 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -62775
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   585
         Width           =   1305
      End
      Begin VB.TextBox c_qtde 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -63390
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox c_descricao 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -73380
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   585
         Width           =   4485
      End
      Begin VB.TextBox c_produto 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -74265
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   585
         Width           =   885
      End
      Begin VB.TextBox c_fabricante 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -74790
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   585
         Width           =   525
      End
      Begin VB.Label l_tit_produto_obs 
         AutoSize        =   -1  'True
         Caption         =   "Informações Adicionais"
         Height          =   195
         Left            =   -68880
         TabIndex        =   125
         Top             =   375
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
         Left            =   -62070
         TabIndex        =   124
         Top             =   4050
         Width           =   450
      End
      Begin VB.Label l_tit_vl_total 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total"
         Height          =   195
         Left            =   -60945
         TabIndex        =   123
         Top             =   375
         Width           =   765
      End
      Begin VB.Label l_tit_vl_unitario 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário"
         Height          =   195
         Left            =   -62430
         TabIndex        =   122
         Top             =   375
         Width           =   945
      End
      Begin VB.Label l_tit_qtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   -63255
         TabIndex        =   121
         Top             =   375
         Width           =   345
      End
      Begin VB.Label l_tit_descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -73365
         TabIndex        =   120
         Top             =   375
         Width           =   720
      End
      Begin VB.Label l_tit_produto 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   195
         Left            =   -74250
         TabIndex        =   119
         Top             =   375
         Width           =   555
      End
      Begin VB.Label l_tit_fabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabric"
         Height          =   195
         Left            =   -74775
         TabIndex        =   118
         Top             =   375
         Width           =   435
      End
   End
   Begin VB.Frame pn_nfe_referenciada 
      Caption         =   "NFe Referenciada"
      Height          =   690
      Left            =   120
      TabIndex        =   24
      Top             =   60
      Width           =   15060
      Begin VB.ComboBox cb_emitente_nfe_ref 
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
         ItemData        =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":04BA
         Left            =   855
         List            =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":04BC
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   225
         Width           =   6990
      End
      Begin VB.TextBox c_num_nfe_ref 
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
         Left            =   11265
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   225
         Width           =   1650
      End
      Begin VB.CommandButton b_nfe_ref 
         Caption         =   "NFe &Referenciada"
         Height          =   450
         Left            =   13410
         TabIndex        =   26
         Top             =   165
         Width           =   1500
      End
      Begin VB.TextBox c_num_serie_nfe_ref 
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
         Left            =   9000
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label l_tit_emitente_nfe_ref 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   300
         Width           =   615
      End
      Begin VB.Label l_tit_num_nfe_ref 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   10650
         TabIndex        =   30
         Top             =   300
         Width           =   525
      End
      Begin VB.Label l_tit_num_serie_nfe_ref 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   8325
         TabIndex        =   29
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Timer relogio 
      Interval        =   1000
      Left            =   14550
      Top             =   7740
   End
   Begin VB.Frame pnNumeroNFe 
      Caption         =   "Última NFe emitida"
      Height          =   690
      Left            =   120
      TabIndex        =   13
      Top             =   8895
      Width           =   15060
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
         Height          =   360
         Left            =   11265
         TabIndex        =   19
         Top             =   225
         Width           =   1650
      End
      Begin VB.Label l_tit_num_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   10650
         TabIndex        =   18
         Top             =   300
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
         Height          =   360
         Left            =   9000
         TabIndex        =   17
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label l_tit_serie_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   8325
         TabIndex        =   16
         Top             =   300
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
         Height          =   360
         Left            =   855
         TabIndex        =   15
         Top             =   225
         Width           =   6990
      End
      Begin VB.Label l_tit_emitente_NF 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   615
      End
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
      Left            =   7770
      TabIndex        =   12
      Top             =   6945
      Width           =   2115
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
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6945
      Width           =   5103
   End
   Begin VB.CommandButton b_fechar 
      Caption         =   "&Fechar"
      Height          =   450
      Left            =   7770
      TabIndex        =   10
      Top             =   8400
      Width           =   2115
   End
   Begin VB.Frame pnDanfe 
      Caption         =   "DANFE"
      Height          =   690
      Left            =   120
      TabIndex        =   2
      Top             =   9630
      Width           =   15060
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
         Left            =   9000
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   225
         Width           =   1170
      End
      Begin VB.CommandButton b_danfe 
         Caption         =   "D&ANFE"
         Height          =   450
         Left            =   13410
         TabIndex        =   5
         Top             =   165
         Width           =   1500
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
         Left            =   11265
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   225
         Width           =   1650
      End
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
         ItemData        =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":04BE
         Left            =   855
         List            =   "f_EMISSAO_NFE_COMPLEMENTAR.frx":04C0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   6990
      End
      Begin VB.Label l_tit_num_serie_Danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   8325
         TabIndex        =   9
         Top             =   300
         Width           =   585
      End
      Begin VB.Label l_tit_num_nfe_danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   10650
         TabIndex        =   8
         Top             =   300
         Width           =   525
      End
      Begin VB.Label l_tit_emitente_danfe 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton b_emissao_automatica 
      Caption         =   "Painel Emissão &Automática"
      Height          =   450
      Left            =   10515
      TabIndex        =   1
      Top             =   6945
      Width           =   2115
   End
   Begin VB.CommandButton b_emite_numeracao_manual 
      Caption         =   "Emitir NFe (Nº &Manual)"
      Height          =   450
      Left            =   7770
      TabIndex        =   0
      Top             =   7665
      Width           =   2115
   End
   Begin VB.Label l_tit_natureza 
      AutoSize        =   -1  'True
      Caption         =   "Natureza da Operação"
      Height          =   195
      Left            =   7845
      TabIndex        =   137
      Top             =   855
      Width           =   1620
   End
   Begin VB.Label l_tit_aliquota_icms 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota ICMS"
      Height          =   195
      Left            =   135
      TabIndex        =   136
      Top             =   1485
      Width           =   1035
   End
   Begin VB.Label l_tit_aliquota_IPI 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota IPI"
      Height          =   195
      Left            =   1875
      TabIndex        =   135
      Top             =   1485
      Width           =   840
   End
   Begin VB.Label l_tit_frete 
      AutoSize        =   -1  'True
      Caption         =   "Frete por Conta"
      Height          =   195
      Left            =   3585
      TabIndex        =   134
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Label l_tit_tipo_NF 
      AutoSize        =   -1  'True
      Caption         =   "Tipo do Documento Fiscal"
      Height          =   195
      Left            =   135
      TabIndex        =   133
      Top             =   855
      Width           =   1860
   End
   Begin VB.Label l_tit_transportadora 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora"
      Height          =   195
      Left            =   7845
      TabIndex        =   132
      Top             =   1485
      Width           =   1080
   End
   Begin VB.Label l_tit_dados_adicionais 
      AutoSize        =   -1  'True
      Caption         =   "Dados Adicionais"
      Height          =   195
      Left            =   135
      TabIndex        =   23
      Top             =   6735
      Width           =   1230
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
      Left            =   5430
      TabIndex        =   22
      Top             =   6945
      Width           =   1980
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
      Left            =   5430
      TabIndex        =   21
      Top             =   8145
      Width           =   1980
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
      Left            =   5430
      TabIndex        =   20
      Top             =   8550
      Width           =   1980
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
Attribute VB_Name = "f_EMISSAO_NFE_COMPLEMENTAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modulo_inicializacao_ok As Boolean

Dim m_blnDadosNFeReferenciadaOk As Boolean
Dim m_lngSerieNFeReferenciada As Long
Dim m_lngNumeroNFeReferenciada As Long
Dim m_strChaveAcessoNFeReferenciada As String
Dim m_rNFeImg As TIPO_NFe_IMG
Dim m_vNFeImgItem() As TIPO_NFe_IMG_ITEM
Dim m_vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP
Dim m_vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA
Dim m_vNFeImgPag() As TIPO_NFe_IMG_PAG

Private Const FONTNAME_IMPRESSAO = "Tahoma"
Private Const FONTSIZE_IMPRESSAO = 8
Private Const FONTBOLD_IMPRESSAO = True
Private Const FONTITALIC_IMPRESSAO = False

Sub formulario_limpa()
Const NomeDestaRotina = "formulario_limpa()"
Dim s As String
Dim i As Integer

    On Error GoTo FL_TRATA_ERRO

'   TRANSPORTADORA
'   ~~~~~~~~~~~~~~
    cb_transportadora.ListIndex = -1

'   ITENS
'   ~~~~~
    c_vl_total_geral = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_fabricante(i) = ""
        c_produto(i) = ""
        c_descricao(i) = ""
        c_qtde(i) = ""
        c_vl_unitario(i) = ""
        c_vl_total(i) = ""
        c_produto_obs(i) = ""
        Next

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
    
'   NATUREZA DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~
'   DEFAULT
    s = "5.102"
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
    
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    c_dados_adicionais = ""
           
'   Nº NFe REFERENCIADA
'   ~~~~~~~~~~~~~~~~~~~
    c_num_nfe_ref = ""
    
'   FOCO INICIAL
'   ~~~~~~~~~~~~
    If cb_emitente_nfe_ref.ListIndex = -1 Then
        cb_emitente_nfe_ref.SetFocus
    ElseIf Trim$(c_num_serie_nfe_ref) = "" Then
        c_num_serie_nfe_ref.SetFocus
    Else
        c_num_nfe_ref.SetFocus
        End If
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FL_TRATA_ERRO:
'=============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Sub formulario_inicia()

' CONSTANTES
Const NomeDestaRotina = "formulario_inicia()"

Dim s As String
Dim s_aux As String
Dim msg_erro As String
Dim v_CFOP() As TIPO_LISTA_CFOP
Dim i As Integer
Dim i_qtde As Integer
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_TRANSPORTADORA As ADODB.Recordset

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
            " id," & _
            " razao_social" & _
        " FROM t_NFE_EMITENTE" & _
        " WHERE" & _
            " (id = " & usuario.emit_id & ")" & _
        " ORDER BY" & _
            " id"
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    Do While Not t_NFE_EMITENTE.EOF
        s = Trim("" & t_NFE_EMITENTE("id")) & " - " & Trim("" & t_NFE_EMITENTE("razao_social"))
        cb_emitente_nfe_ref.AddItem s
        cb_emitente_danfe.AddItem s
        t_NFE_EMITENTE.MoveNext
        Loop
            
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
            
'   TIPO DO DOCUMENTO FISCAL
'   ~~~~~~~~~~~~~~~~~~~~~~~~
    cb_tipo_NF.Clear
    cb_tipo_NF.AddItem "0 - ENTRADA"
    cb_tipo_NF.AddItem "1 - SAÍDA"
    
'   NATUREZA DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~
    cb_natureza.Clear
    
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
                s = .codigo & String$(1, " ") & .descricao
                cb_natureza.AddItem s
                End If
            End With
        Next
       
'   ALÍQUOTA ICMS
'   ~~~~~~~~~~~~~
    cb_icms.Clear
    cb_icms.AddItem "0"
    cb_icms.AddItem "4"
    cb_icms.AddItem "7"
    cb_icms.AddItem "12"
    cb_icms.AddItem "17"
    cb_icms.AddItem "18"
    cb_icms.AddItem "20"
    
'   FRETE POR CONTA
'   ~~~~~~~~~~~~~~~
    cb_frete.Clear
    cb_frete.AddItem "0 - EMITENTE"
    cb_frete.AddItem "1 - DESTINATÁRIO"
    
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    With c_dados_adicionais
        .FontName = FONTNAME_IMPRESSAO
        .FontSize = FONTSIZE_IMPRESSAO
        .FontBold = FONTBOLD_IMPRESSAO
        .FontItalic = FONTITALIC_IMPRESSAO
        End With
    
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

Sub limpa_variaveis_NFe_referenciada()
Const NomeDestaRotina = "limpa_variaveis_NFe_referenciada()"
Dim s As String

    On Error GoTo LVNR_TRATA_ERRO
    
    m_blnDadosNFeReferenciadaOk = False
    m_lngSerieNFeReferenciada = 0
    m_lngNumeroNFeReferenciada = 0
    m_strChaveAcessoNFeReferenciada = ""
    
'   TIPO_NFe_IMG
    limpa_TIPO_NFe_IMG m_rNFeImg
        
'   TIPO_NFe_IMG_ITEM
    ReDim m_vNFeImgItem(0)
    limpa_TIPO_NFe_IMG_ITEM m_vNFeImgItem()
    
'   TIPO_NFe_IMG_TAG_DUP
    ReDim m_vNFeImgTagDup(0)
    limpa_TIPO_NFe_IMG_TAG_DUP m_vNFeImgTagDup()
        
'   TIPO_NFe_IMG_NFe_REFERENCIADA
    ReDim m_vNFeImgNFeRef(0)
    limpa_TIPO_NFe_IMG_NFe_REFERENCIADA m_vNFeImgNFeRef()
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LVNR_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Sub NFe_emite(ByVal FLAG_NUMERACAO_MANUAL As Boolean)
' ____________________________________________________________________________________________________
'|
'|  EMITE A NOTA FISCAL ELETRÔNICA (NFe) COMPLEMENTAR COM BASE NOS DADOS PREENCHIDOS MANUALMENTE.
'|


End Sub

Function ha_dados_preenchidos() As Boolean
Const NomeDestaRotina = "ha_dados_preenchidos()"
Dim i As Integer
Dim s As String

    On Error GoTo HDP_TRATA_ERRO

    ha_dados_preenchidos = True
    
    For i = c_produto.LBound To c_produto.UBound
        If Trim$(c_fabricante(i)) <> "" Then Exit Function
        If Trim$(c_produto(i)) <> "" Then Exit Function
        If Trim$(c_qtde(i)) <> "" Then Exit Function
        If converte_para_currency(c_vl_unitario(i)) <> 0 Then Exit Function
        Next
    
    If Trim$(c_dados_adicionais) <> "" Then Exit Function
    
'   TODO
    
    ha_dados_preenchidos = False
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
HDP_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function

End Function

Sub fechar_programa()

    On Error Resume Next
    
'   FECHA BANCO DE DADOS
    BD_Fecha
    BD_CEP_Fecha
    BD_Assist_Fecha
    
'   ENCERRA PROGRAMA
    End

End Sub

Sub fechar_modo_emissao_nfe_complementar()
Const NomeDestaRotina = "fechar_modo_emissao_nfe_complementar()"
Dim s As String

    On Error GoTo FMENC_TRATA_ERRO
    
    If ha_dados_preenchidos Then
        s = "Os dados preenchidos serão perdidos se o painel for alternado para o modo de emissão automática!!" & _
            vbCrLf & _
            "Continua assim mesmo?"
        If Not confirma(s) Then Exit Sub
        End If
        
    Unload f_EMISSAO_NFE_COMPLEMENTAR

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FMENC_TRATA_ERRO:
'================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

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
    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
    
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
'=========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub DANFE_CONSULTA_PARAM_EMITENTE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Sub NFe_referenciada_consulta(ByVal intNFeRefEmitente As Integer, ByVal strNFeRefSerie As String, ByVal strNFeRefNumero As String)
Const NomeDestaRotina = "NFe_referenciada_consulta()"
Dim s As String
Dim s_aux As String
Dim msg_erro As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim id_nfe_imagem As Long
Dim t_T1_NFE As ADODB.Recordset
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim dbcNFe As ADODB.Connection

    On Error GoTo NFE_REF_CONSULTA_TRATA_ERRO
    
    limpa_variaveis_NFe_referenciada
    
'   INICIALIZAÇÃO
'   ~~~~~~~~~~~~~
    m_lngSerieNFeReferenciada = CLng(strNFeRefSerie)
    m_lngNumeroNFeReferenciada = CLng(strNFeRefNumero)
    strNFeRefSerie = NFeFormataSerieNF(strNFeRefSerie)
    strNFeRefNumero = NFeFormataNumeroNF(strNFeRefNumero)
    
    aguarde INFO_EXECUTANDO, "consultando NFe referenciada no banco de dados"
    
'   t_T1_NFE
    Set t_T1_NFE = New ADODB.Recordset
    With t_T1_NFE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFe_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
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
            " (id = " & CStr(intNFeRefEmitente) & ")"
    t_NFE_EMITENTE.Open s, dbc, , , adCmdText
    If t_NFE_EMITENTE.EOF Then
        aviso_erro "Dados do emitente não foram localizados no BD (id=" & CStr(intNFeRefEmitente) & ")!!"
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
    Else
        strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
        strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
        strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
        strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
        End If
    
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
    
    s = "SELECT" & _
            " ChaveAcesso" & _
        " FROM NFE" & _
        " WHERE" & _
            " (Serie = '" & strNFeRefSerie & "')" & _
            " AND (Nfe = '" & strNFeRefNumero & "')"
    t_T1_NFE.Open s, dbcNFe, , , adCmdText
    If t_T1_NFE.EOF Then
        aviso_erro "NFe referenciada informada não está cadastrada no banco de dados!!"
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    If Trim$("" & t_T1_NFE("ChaveAcesso")) = "" Then
        aviso_erro "NFe referenciada informada não possui a informação da chave de acesso!!"
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
    ElseIf Len(retorna_so_digitos(Trim$("" & t_T1_NFE("ChaveAcesso")))) <> 44 Then
        aviso_erro "NFe referenciada informada possui chave de acesso com tamanho inválido (" & Trim$("" & t_T1_NFE("ChaveAcesso")) & ")!!"
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
        
    m_strChaveAcessoNFeReferenciada = Trim$("" & t_T1_NFE("ChaveAcesso"))
    m_vNFeImgNFeRef(0).refNFe = Trim$("" & t_T1_NFE("ChaveAcesso"))
    m_vNFeImgNFeRef(0).NFe_serie_NF_referenciada = m_lngSerieNFeReferenciada
    m_vNFeImgNFeRef(0).NFe_numero_NF_referenciada = m_lngNumeroNFeReferenciada
    
'   Lê dados completos da tabela de imagem da NFe referenciada
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM" & _
        " WHERE" & _
            " (id_nfe_emitente = " & CStr(intNFeRefEmitente) & ")" & _
            " AND (NFe_serie_NF = " & CStr(m_lngSerieNFeReferenciada) & ")" & _
            " AND (NFe_numero_NF = " & CStr(m_lngNumeroNFeReferenciada) & ")" & _
            " AND (st_anulado = 0)" & _
        " ORDER BY" & _
            " id DESC"
    t_NFe_IMAGEM.Open s, dbc, , , adCmdText
    If t_NFe_IMAGEM.EOF Then
        aviso_erro "Não foram localizados os dados da NF referenciada na tabela com os dados de imagem!!"
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    id_nfe_imagem = t_NFe_IMAGEM("id")
    
    aguarde INFO_EXECUTANDO, "recuperando dados da NFe referenciada"
    If Not le_NFe_imagem(id_nfe_imagem, m_rNFeImg, m_vNFeImgItem(), m_vNFeImgTagDup(), m_vNFeImgNFeRef(), m_vNFeImgPag(), msg_erro) Then
        aviso_erro "Falha ao recuperar os dados da NFe referenciada!!" & vbCrLf & msg_erro
        GoSub NFE_REF_CONSULTA_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If
    
    m_blnDadosNFeReferenciadaOk = True
    
    GoSub NFE_REF_CONSULTA_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_REF_CONSULTA_FECHA_TABELAS:
'==============================
  ' RECORDSETS
    bd_desaloca_recordset t_T1_NFE, True
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_IMAGEM, True
    
  ' CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_REF_CONSULTA_TRATA_ERRO:
'===========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub NFE_REF_CONSULTA_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
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

Private Sub b_emissao_automatica_Click()

    fechar_modo_emissao_nfe_complementar

End Sub

Private Sub b_emite_numeracao_manual_Click()

    NFe_emite True

End Sub

Private Sub b_fechar_Click()

    fechar_programa

End Sub

Private Sub b_imprime_Click()

    NFe_emite False

End Sub

Private Sub b_nfe_ref_Click()
Const NomeDestaRotina = "b_nfe_ref_Click()"
Dim i As Integer
Dim c As String
Dim s As String
Dim s_aux As String
Dim intNFeRefEmitente As Integer

    On Error GoTo B_NFE_REF_TRATA_ERRO
    
'   CONSISTÊNCIAS
'   ~~~~~~~~~~~~~
'   EMITENTE
    If (cb_emitente_nfe_ref.ListIndex = -1) Or (Trim$(cb_emitente_nfe_ref) = "") Then
        aviso_erro "Selecione o emitente da NFe referenciada!!"
        cb_emitente_nfe_ref.SetFocus
        Exit Sub
        End If
        
    s = cb_emitente_nfe_ref
    s_aux = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then Exit For
        s_aux = s_aux & c
        Next
    intNFeRefEmitente = CInt(s_aux)
    
'   Série da NFe referenciada
    If Trim$(c_num_serie_nfe_ref) = "" Then
        aviso_erro "Informe a série da NFe referenciada!!"
        c_num_serie_nfe_ref.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(c_num_serie_nfe_ref) Then
        aviso_erro "Número da série da NFe referenciada é inválido!!"
        c_num_serie_nfe_ref.SetFocus
        Exit Sub
        End If
    
'   Nº da NFe referenciada
    If Trim$(c_num_nfe_ref) = "" Then
        aviso_erro "Informe o número da NFe referenciada!!"
        c_num_nfe_ref.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(c_num_nfe_ref) Then
        aviso_erro "Número da NFe referenciada é inválido!!"
        c_num_nfe_ref.SetFocus
        Exit Sub
        End If
        
    NFe_referenciada_consulta intNFeRefEmitente, c_num_serie_nfe_ref, c_num_nfe_ref
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
B_NFE_REF_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
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


Private Sub c_num_nfe_ref_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim s As String
Dim s_aux As String
Dim c As String
Dim intEmitente As Integer

    If KeyAscii = 13 Then
        KeyAscii = 0
        b_nfe_ref.SetFocus
        If (cb_emitente_nfe_ref.ListIndex <> -1) And (Trim$(cb_emitente_nfe_ref) <> "") And (Trim$(c_num_serie_nfe_ref) <> "") And (Trim$(c_num_nfe_ref) <> "") Then
            s = cb_emitente_nfe_ref
            s_aux = ""
            For i = 1 To Len(s)
                c = Mid$(s, i, 1)
                If Not IsNumeric(c) Then Exit For
                s_aux = s_aux & c
                Next
            intEmitente = CInt(s_aux)
            
            NFe_referenciada_consulta intEmitente, c_num_serie_nfe_ref, c_num_nfe_ref
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


Private Sub c_num_nfe_ref_LostFocus()

    c_num_nfe_ref = Trim$(c_num_nfe_ref)

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


Private Sub c_num_serie_nfe_ref_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_num_nfe_ref.SetFocus
        Exit Sub
        End If
    
    
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_numerico(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub


Private Sub c_num_serie_nfe_ref_LostFocus()

    c_num_serie_nfe_ref = Trim$(c_num_serie_nfe_ref)

End Sub


Private Sub cb_emitente_danfe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_num_serie_danfe.SetFocus
        Exit Sub
        End If

End Sub


Private Sub cb_emitente_nfe_ref_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_num_serie_nfe_ref.SetFocus
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

End Sub

Private Sub Form_Unload(Cancel As Integer)

'   EM EXECUÇÃO ?
    If em_execucao Then
        Cancel = True
        Exit Sub
        End If

End Sub


Private Sub mnu_emissao_automatica_Click()

    fechar_modo_emissao_nfe_complementar

End Sub

Private Sub mnu_FECHAR_Click()

    fechar_programa

End Sub

Private Sub relogio_Timer()
Dim s As String

    s = left$(Time$, 5)
    If Val(right$(Time$, 1)) Mod 2 Then Mid$(s, 3, 1) = " "
    agora = s

    hoje = Format$(Date, "dd/mm/yyyy")

End Sub


