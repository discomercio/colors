VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form f_MAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal"
   ClientHeight    =   11610
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   20490
   Icon            =   "f_MAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11610
   ScaleWidth      =   20490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame pnItens 
      BackColor       =   &H00FFC0C0&
      Height          =   4750
      Left            =   120
      TabIndex        =   72
      Top             =   2160
      Width           =   20250
      Begin VB.TextBox c_vl_total_icms 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15660
         Locked          =   -1  'True
         TabIndex        =   655
         TabStop         =   0   'False
         Top             =   4365
         Width           =   1425
      End
      Begin VB.TextBox c_total_peso_liquido 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8185
         MaxLength       =   15
         TabIndex        =   654
         Top             =   4365
         Width           =   1095
      End
      Begin VB.TextBox c_total_peso_bruto 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5260
         MaxLength       =   15
         TabIndex        =   653
         Top             =   4365
         Width           =   1095
      End
      Begin VB.TextBox c_vl_total_outras_despesas_acessorias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11840
         Locked          =   -1  'True
         TabIndex        =   652
         TabStop         =   0   'False
         Top             =   4365
         Width           =   1305
      End
      Begin VB.TextBox c_total_volumes 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   651
         Top             =   4365
         Width           =   735
      End
      Begin VB.TextBox c_vl_total_geral 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10535
         Locked          =   -1  'True
         TabIndex        =   650
         TabStop         =   0   'False
         Top             =   4365
         Width           =   1305
      End
      Begin TabDlg.SSTab sstItens 
         Height          =   3855
         Left            =   50
         TabIndex        =   73
         Top             =   360
         Width           =   20100
         _ExtentX        =   35454
         _ExtentY        =   6800
         _Version        =   393216
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
         TabCaption(0)   =   "Itens 1 a 12"
         TabPicture(0)   =   "f_MAIN.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "c_NCM(11)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cb_ICMS_item(11)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "c_xPed(11)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "c_unidade(11)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "c_nItemPed(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "c_fcp(11)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "c_NCM(10)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cb_ICMS_item(10)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "c_xPed(10)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "c_unidade(10)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "c_nItemPed(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "c_fcp(10)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "c_fcp(9)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "c_nItemPed(9)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "c_unidade(9)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "c_xPed(9)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "cb_ICMS_item(9)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "c_NCM(9)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "c_fcp(8)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "c_nItemPed(8)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "c_unidade(8)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "c_xPed(8)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "cb_ICMS_item(8)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "c_NCM(8)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "c_fcp(7)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "c_nItemPed(7)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "c_unidade(7)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "c_xPed(7)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "cb_ICMS_item(7)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "c_NCM(7)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "c_fcp(6)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "c_nItemPed(6)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "c_unidade(6)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "c_xPed(6)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "cb_ICMS_item(6)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "c_NCM(6)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "c_fcp(5)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "c_nItemPed(5)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "c_unidade(5)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "c_xPed(5)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "cb_ICMS_item(5)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "c_NCM(5)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "c_fcp(4)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "c_nItemPed(4)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "c_unidade(4)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "c_xPed(4)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "cb_ICMS_item(4)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "c_NCM(4)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "c_fcp(3)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "c_nItemPed(3)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "c_unidade(3)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "c_xPed(3)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "cb_ICMS_item(3)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "c_NCM(3)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "c_fcp(2)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "c_nItemPed(2)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "c_unidade(2)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "c_xPed(2)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "cb_ICMS_item(2)"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "c_NCM(2)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "c_fcp(1)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "c_nItemPed(1)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "c_unidade(1)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "c_xPed(1)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "cb_ICMS_item(1)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "c_NCM(1)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "c_fcp(0)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "c_nItemPed(0)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "c_unidade(0)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "c_xPed(0)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "cb_ICMS_item(0)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "c_NCM(0)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "cb_CFOP(11)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "cb_CFOP(10)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "cb_CFOP(9)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "cb_CFOP(8)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "cb_CFOP(7)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "cb_CFOP(6)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "cb_CFOP(5)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).Control(79)=   "cb_CFOP(4)"
         Tab(0).Control(79).Enabled=   0   'False
         Tab(0).Control(80)=   "cb_CFOP(3)"
         Tab(0).Control(80).Enabled=   0   'False
         Tab(0).Control(81)=   "cb_CFOP(2)"
         Tab(0).Control(81).Enabled=   0   'False
         Tab(0).Control(82)=   "cb_CFOP(1)"
         Tab(0).Control(82).Enabled=   0   'False
         Tab(0).Control(83)=   "cb_CFOP(0)"
         Tab(0).Control(83).Enabled=   0   'False
         Tab(0).Control(84)=   "c_CST(0)"
         Tab(0).Control(84).Enabled=   0   'False
         Tab(0).Control(85)=   "c_CST(1)"
         Tab(0).Control(85).Enabled=   0   'False
         Tab(0).Control(86)=   "c_CST(2)"
         Tab(0).Control(86).Enabled=   0   'False
         Tab(0).Control(87)=   "c_CST(3)"
         Tab(0).Control(87).Enabled=   0   'False
         Tab(0).Control(88)=   "c_CST(4)"
         Tab(0).Control(88).Enabled=   0   'False
         Tab(0).Control(89)=   "c_CST(5)"
         Tab(0).Control(89).Enabled=   0   'False
         Tab(0).Control(90)=   "c_CST(6)"
         Tab(0).Control(90).Enabled=   0   'False
         Tab(0).Control(91)=   "c_CST(7)"
         Tab(0).Control(91).Enabled=   0   'False
         Tab(0).Control(92)=   "c_CST(8)"
         Tab(0).Control(92).Enabled=   0   'False
         Tab(0).Control(93)=   "c_CST(9)"
         Tab(0).Control(93).Enabled=   0   'False
         Tab(0).Control(94)=   "c_CST(10)"
         Tab(0).Control(94).Enabled=   0   'False
         Tab(0).Control(95)=   "c_CST(11)"
         Tab(0).Control(95).Enabled=   0   'False
         Tab(0).Control(96)=   "c_vl_outras_despesas_acessorias(0)"
         Tab(0).Control(96).Enabled=   0   'False
         Tab(0).Control(97)=   "c_vl_outras_despesas_acessorias(1)"
         Tab(0).Control(97).Enabled=   0   'False
         Tab(0).Control(98)=   "c_vl_outras_despesas_acessorias(2)"
         Tab(0).Control(98).Enabled=   0   'False
         Tab(0).Control(99)=   "c_vl_outras_despesas_acessorias(3)"
         Tab(0).Control(99).Enabled=   0   'False
         Tab(0).Control(100)=   "c_vl_outras_despesas_acessorias(4)"
         Tab(0).Control(100).Enabled=   0   'False
         Tab(0).Control(101)=   "c_vl_outras_despesas_acessorias(5)"
         Tab(0).Control(101).Enabled=   0   'False
         Tab(0).Control(102)=   "c_vl_outras_despesas_acessorias(6)"
         Tab(0).Control(102).Enabled=   0   'False
         Tab(0).Control(103)=   "c_vl_outras_despesas_acessorias(7)"
         Tab(0).Control(103).Enabled=   0   'False
         Tab(0).Control(104)=   "c_vl_outras_despesas_acessorias(8)"
         Tab(0).Control(104).Enabled=   0   'False
         Tab(0).Control(105)=   "c_vl_outras_despesas_acessorias(9)"
         Tab(0).Control(105).Enabled=   0   'False
         Tab(0).Control(106)=   "c_vl_outras_despesas_acessorias(10)"
         Tab(0).Control(106).Enabled=   0   'False
         Tab(0).Control(107)=   "c_vl_outras_despesas_acessorias(11)"
         Tab(0).Control(107).Enabled=   0   'False
         Tab(0).Control(108)=   "c_produto_obs(11)"
         Tab(0).Control(108).Enabled=   0   'False
         Tab(0).Control(109)=   "c_vl_total(11)"
         Tab(0).Control(109).Enabled=   0   'False
         Tab(0).Control(110)=   "c_vl_unitario(11)"
         Tab(0).Control(110).Enabled=   0   'False
         Tab(0).Control(111)=   "c_qtde(11)"
         Tab(0).Control(111).Enabled=   0   'False
         Tab(0).Control(112)=   "c_descricao(11)"
         Tab(0).Control(112).Enabled=   0   'False
         Tab(0).Control(113)=   "c_produto(11)"
         Tab(0).Control(113).Enabled=   0   'False
         Tab(0).Control(114)=   "c_fabricante(11)"
         Tab(0).Control(114).Enabled=   0   'False
         Tab(0).Control(115)=   "c_fabricante(10)"
         Tab(0).Control(115).Enabled=   0   'False
         Tab(0).Control(116)=   "c_produto(10)"
         Tab(0).Control(116).Enabled=   0   'False
         Tab(0).Control(117)=   "c_descricao(10)"
         Tab(0).Control(117).Enabled=   0   'False
         Tab(0).Control(118)=   "c_qtde(10)"
         Tab(0).Control(118).Enabled=   0   'False
         Tab(0).Control(119)=   "c_vl_unitario(10)"
         Tab(0).Control(119).Enabled=   0   'False
         Tab(0).Control(120)=   "c_vl_total(10)"
         Tab(0).Control(120).Enabled=   0   'False
         Tab(0).Control(121)=   "c_produto_obs(10)"
         Tab(0).Control(121).Enabled=   0   'False
         Tab(0).Control(122)=   "c_produto_obs(9)"
         Tab(0).Control(122).Enabled=   0   'False
         Tab(0).Control(123)=   "c_produto_obs(8)"
         Tab(0).Control(123).Enabled=   0   'False
         Tab(0).Control(124)=   "c_produto_obs(7)"
         Tab(0).Control(124).Enabled=   0   'False
         Tab(0).Control(125)=   "c_produto_obs(6)"
         Tab(0).Control(125).Enabled=   0   'False
         Tab(0).Control(126)=   "c_produto_obs(5)"
         Tab(0).Control(126).Enabled=   0   'False
         Tab(0).Control(127)=   "c_produto_obs(4)"
         Tab(0).Control(127).Enabled=   0   'False
         Tab(0).Control(128)=   "c_produto_obs(3)"
         Tab(0).Control(128).Enabled=   0   'False
         Tab(0).Control(129)=   "c_produto_obs(2)"
         Tab(0).Control(129).Enabled=   0   'False
         Tab(0).Control(130)=   "c_produto_obs(1)"
         Tab(0).Control(130).Enabled=   0   'False
         Tab(0).Control(131)=   "c_produto_obs(0)"
         Tab(0).Control(131).Enabled=   0   'False
         Tab(0).Control(132)=   "c_vl_total(9)"
         Tab(0).Control(132).Enabled=   0   'False
         Tab(0).Control(133)=   "c_vl_unitario(9)"
         Tab(0).Control(133).Enabled=   0   'False
         Tab(0).Control(134)=   "c_qtde(9)"
         Tab(0).Control(134).Enabled=   0   'False
         Tab(0).Control(135)=   "c_descricao(9)"
         Tab(0).Control(135).Enabled=   0   'False
         Tab(0).Control(136)=   "c_produto(9)"
         Tab(0).Control(136).Enabled=   0   'False
         Tab(0).Control(137)=   "c_fabricante(9)"
         Tab(0).Control(137).Enabled=   0   'False
         Tab(0).Control(138)=   "c_vl_total(8)"
         Tab(0).Control(138).Enabled=   0   'False
         Tab(0).Control(139)=   "c_vl_unitario(8)"
         Tab(0).Control(139).Enabled=   0   'False
         Tab(0).Control(140)=   "c_qtde(8)"
         Tab(0).Control(140).Enabled=   0   'False
         Tab(0).Control(141)=   "c_descricao(8)"
         Tab(0).Control(141).Enabled=   0   'False
         Tab(0).Control(142)=   "c_produto(8)"
         Tab(0).Control(142).Enabled=   0   'False
         Tab(0).Control(143)=   "c_fabricante(8)"
         Tab(0).Control(143).Enabled=   0   'False
         Tab(0).Control(144)=   "c_vl_total(7)"
         Tab(0).Control(144).Enabled=   0   'False
         Tab(0).Control(145)=   "c_vl_unitario(7)"
         Tab(0).Control(145).Enabled=   0   'False
         Tab(0).Control(146)=   "c_qtde(7)"
         Tab(0).Control(146).Enabled=   0   'False
         Tab(0).Control(147)=   "c_descricao(7)"
         Tab(0).Control(147).Enabled=   0   'False
         Tab(0).Control(148)=   "c_produto(7)"
         Tab(0).Control(148).Enabled=   0   'False
         Tab(0).Control(149)=   "c_fabricante(7)"
         Tab(0).Control(149).Enabled=   0   'False
         Tab(0).Control(150)=   "c_vl_total(6)"
         Tab(0).Control(150).Enabled=   0   'False
         Tab(0).Control(151)=   "c_vl_unitario(6)"
         Tab(0).Control(151).Enabled=   0   'False
         Tab(0).Control(152)=   "c_qtde(6)"
         Tab(0).Control(152).Enabled=   0   'False
         Tab(0).Control(153)=   "c_descricao(6)"
         Tab(0).Control(153).Enabled=   0   'False
         Tab(0).Control(154)=   "c_produto(6)"
         Tab(0).Control(154).Enabled=   0   'False
         Tab(0).Control(155)=   "c_fabricante(6)"
         Tab(0).Control(155).Enabled=   0   'False
         Tab(0).Control(156)=   "c_vl_total(5)"
         Tab(0).Control(156).Enabled=   0   'False
         Tab(0).Control(157)=   "c_vl_unitario(5)"
         Tab(0).Control(157).Enabled=   0   'False
         Tab(0).Control(158)=   "c_qtde(5)"
         Tab(0).Control(158).Enabled=   0   'False
         Tab(0).Control(159)=   "c_descricao(5)"
         Tab(0).Control(159).Enabled=   0   'False
         Tab(0).Control(160)=   "c_produto(5)"
         Tab(0).Control(160).Enabled=   0   'False
         Tab(0).Control(161)=   "c_fabricante(5)"
         Tab(0).Control(161).Enabled=   0   'False
         Tab(0).Control(162)=   "c_vl_total(4)"
         Tab(0).Control(162).Enabled=   0   'False
         Tab(0).Control(163)=   "c_vl_unitario(4)"
         Tab(0).Control(163).Enabled=   0   'False
         Tab(0).Control(164)=   "c_qtde(4)"
         Tab(0).Control(164).Enabled=   0   'False
         Tab(0).Control(165)=   "c_descricao(4)"
         Tab(0).Control(165).Enabled=   0   'False
         Tab(0).Control(166)=   "c_produto(4)"
         Tab(0).Control(166).Enabled=   0   'False
         Tab(0).Control(167)=   "c_fabricante(4)"
         Tab(0).Control(167).Enabled=   0   'False
         Tab(0).Control(168)=   "c_vl_total(3)"
         Tab(0).Control(168).Enabled=   0   'False
         Tab(0).Control(169)=   "c_vl_unitario(3)"
         Tab(0).Control(169).Enabled=   0   'False
         Tab(0).Control(170)=   "c_qtde(3)"
         Tab(0).Control(170).Enabled=   0   'False
         Tab(0).Control(171)=   "c_descricao(3)"
         Tab(0).Control(171).Enabled=   0   'False
         Tab(0).Control(172)=   "c_produto(3)"
         Tab(0).Control(172).Enabled=   0   'False
         Tab(0).Control(173)=   "c_fabricante(3)"
         Tab(0).Control(173).Enabled=   0   'False
         Tab(0).Control(174)=   "c_vl_total(2)"
         Tab(0).Control(174).Enabled=   0   'False
         Tab(0).Control(175)=   "c_vl_unitario(2)"
         Tab(0).Control(175).Enabled=   0   'False
         Tab(0).Control(176)=   "c_qtde(2)"
         Tab(0).Control(176).Enabled=   0   'False
         Tab(0).Control(177)=   "c_descricao(2)"
         Tab(0).Control(177).Enabled=   0   'False
         Tab(0).Control(178)=   "c_produto(2)"
         Tab(0).Control(178).Enabled=   0   'False
         Tab(0).Control(179)=   "c_fabricante(2)"
         Tab(0).Control(179).Enabled=   0   'False
         Tab(0).Control(180)=   "c_vl_total(1)"
         Tab(0).Control(180).Enabled=   0   'False
         Tab(0).Control(181)=   "c_vl_unitario(1)"
         Tab(0).Control(181).Enabled=   0   'False
         Tab(0).Control(182)=   "c_qtde(1)"
         Tab(0).Control(182).Enabled=   0   'False
         Tab(0).Control(183)=   "c_descricao(1)"
         Tab(0).Control(183).Enabled=   0   'False
         Tab(0).Control(184)=   "c_produto(1)"
         Tab(0).Control(184).Enabled=   0   'False
         Tab(0).Control(185)=   "c_fabricante(1)"
         Tab(0).Control(185).Enabled=   0   'False
         Tab(0).Control(186)=   "c_vl_total(0)"
         Tab(0).Control(186).Enabled=   0   'False
         Tab(0).Control(187)=   "c_vl_unitario(0)"
         Tab(0).Control(187).Enabled=   0   'False
         Tab(0).Control(188)=   "c_qtde(0)"
         Tab(0).Control(188).Enabled=   0   'False
         Tab(0).Control(189)=   "c_descricao(0)"
         Tab(0).Control(189).Enabled=   0   'False
         Tab(0).Control(190)=   "c_produto(0)"
         Tab(0).Control(190).Enabled=   0   'False
         Tab(0).Control(191)=   "c_fabricante(0)"
         Tab(0).Control(191).Enabled=   0   'False
         Tab(0).ControlCount=   192
         TabCaption(1)   =   "Itens 13 a 24"
         TabPicture(1)   =   "f_MAIN.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "c_fcp(12)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "c_nItemPed(12)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "c_unidade(12)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "c_xPed(12)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cb_ICMS_item(12)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "c_NCM(12)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cb_CFOP(12)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "c_CST(12)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "c_vl_outras_despesas_acessorias(12)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "c_produto_obs(12)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "c_vl_total(12)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "c_vl_unitario(12)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "c_qtde(12)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "c_descricao(12)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "c_produto(12)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "c_fabricante(12)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "c_fcp(13)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "c_nItemPed(13)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "c_unidade(13)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "c_xPed(13)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "cb_ICMS_item(13)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "c_NCM(13)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "cb_CFOP(13)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "c_CST(13)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "c_vl_outras_despesas_acessorias(13)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "c_produto_obs(13)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "c_vl_total(13)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "c_vl_unitario(13)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "c_qtde(13)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "c_descricao(13)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "c_produto(13)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "c_fabricante(13)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "c_fcp(14)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "c_nItemPed(14)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "c_unidade(14)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "c_xPed(14)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "cb_ICMS_item(14)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "c_NCM(14)"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "cb_CFOP(14)"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "c_CST(14)"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "c_vl_outras_despesas_acessorias(14)"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "c_produto_obs(14)"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "c_vl_total(14)"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "c_vl_unitario(14)"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "c_qtde(14)"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "c_descricao(14)"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "c_produto(14)"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "c_fabricante(14)"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "c_fcp(15)"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "c_nItemPed(15)"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "c_unidade(15)"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "c_xPed(15)"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "cb_ICMS_item(15)"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "c_NCM(15)"
         Tab(1).Control(53).Enabled=   0   'False
         Tab(1).Control(54)=   "cb_CFOP(15)"
         Tab(1).Control(54).Enabled=   0   'False
         Tab(1).Control(55)=   "c_CST(15)"
         Tab(1).Control(55).Enabled=   0   'False
         Tab(1).Control(56)=   "c_vl_outras_despesas_acessorias(15)"
         Tab(1).Control(56).Enabled=   0   'False
         Tab(1).Control(57)=   "c_produto_obs(15)"
         Tab(1).Control(57).Enabled=   0   'False
         Tab(1).Control(58)=   "c_vl_total(15)"
         Tab(1).Control(58).Enabled=   0   'False
         Tab(1).Control(59)=   "c_vl_unitario(15)"
         Tab(1).Control(59).Enabled=   0   'False
         Tab(1).Control(60)=   "c_qtde(15)"
         Tab(1).Control(60).Enabled=   0   'False
         Tab(1).Control(61)=   "c_descricao(15)"
         Tab(1).Control(61).Enabled=   0   'False
         Tab(1).Control(62)=   "c_produto(15)"
         Tab(1).Control(62).Enabled=   0   'False
         Tab(1).Control(63)=   "c_fabricante(15)"
         Tab(1).Control(63).Enabled=   0   'False
         Tab(1).Control(64)=   "c_fcp(16)"
         Tab(1).Control(64).Enabled=   0   'False
         Tab(1).Control(65)=   "c_nItemPed(16)"
         Tab(1).Control(65).Enabled=   0   'False
         Tab(1).Control(66)=   "c_unidade(16)"
         Tab(1).Control(66).Enabled=   0   'False
         Tab(1).Control(67)=   "c_xPed(16)"
         Tab(1).Control(67).Enabled=   0   'False
         Tab(1).Control(68)=   "cb_ICMS_item(16)"
         Tab(1).Control(68).Enabled=   0   'False
         Tab(1).Control(69)=   "c_NCM(16)"
         Tab(1).Control(69).Enabled=   0   'False
         Tab(1).Control(70)=   "cb_CFOP(16)"
         Tab(1).Control(70).Enabled=   0   'False
         Tab(1).Control(71)=   "c_CST(16)"
         Tab(1).Control(71).Enabled=   0   'False
         Tab(1).Control(72)=   "c_vl_outras_despesas_acessorias(16)"
         Tab(1).Control(72).Enabled=   0   'False
         Tab(1).Control(73)=   "c_produto_obs(16)"
         Tab(1).Control(73).Enabled=   0   'False
         Tab(1).Control(74)=   "c_vl_total(16)"
         Tab(1).Control(74).Enabled=   0   'False
         Tab(1).Control(75)=   "c_vl_unitario(16)"
         Tab(1).Control(75).Enabled=   0   'False
         Tab(1).Control(76)=   "c_qtde(16)"
         Tab(1).Control(76).Enabled=   0   'False
         Tab(1).Control(77)=   "c_descricao(16)"
         Tab(1).Control(77).Enabled=   0   'False
         Tab(1).Control(78)=   "c_produto(16)"
         Tab(1).Control(78).Enabled=   0   'False
         Tab(1).Control(79)=   "c_fabricante(16)"
         Tab(1).Control(79).Enabled=   0   'False
         Tab(1).Control(80)=   "c_fcp(17)"
         Tab(1).Control(80).Enabled=   0   'False
         Tab(1).Control(81)=   "c_nItemPed(17)"
         Tab(1).Control(81).Enabled=   0   'False
         Tab(1).Control(82)=   "c_unidade(17)"
         Tab(1).Control(82).Enabled=   0   'False
         Tab(1).Control(83)=   "c_xPed(17)"
         Tab(1).Control(83).Enabled=   0   'False
         Tab(1).Control(84)=   "cb_ICMS_item(17)"
         Tab(1).Control(84).Enabled=   0   'False
         Tab(1).Control(85)=   "c_NCM(17)"
         Tab(1).Control(85).Enabled=   0   'False
         Tab(1).Control(86)=   "cb_CFOP(17)"
         Tab(1).Control(86).Enabled=   0   'False
         Tab(1).Control(87)=   "c_CST(17)"
         Tab(1).Control(87).Enabled=   0   'False
         Tab(1).Control(88)=   "c_vl_outras_despesas_acessorias(17)"
         Tab(1).Control(88).Enabled=   0   'False
         Tab(1).Control(89)=   "c_produto_obs(17)"
         Tab(1).Control(89).Enabled=   0   'False
         Tab(1).Control(90)=   "c_vl_total(17)"
         Tab(1).Control(90).Enabled=   0   'False
         Tab(1).Control(91)=   "c_vl_unitario(17)"
         Tab(1).Control(91).Enabled=   0   'False
         Tab(1).Control(92)=   "c_qtde(17)"
         Tab(1).Control(92).Enabled=   0   'False
         Tab(1).Control(93)=   "c_descricao(17)"
         Tab(1).Control(93).Enabled=   0   'False
         Tab(1).Control(94)=   "c_produto(17)"
         Tab(1).Control(94).Enabled=   0   'False
         Tab(1).Control(95)=   "c_fabricante(17)"
         Tab(1).Control(95).Enabled=   0   'False
         Tab(1).Control(96)=   "c_fcp(18)"
         Tab(1).Control(96).Enabled=   0   'False
         Tab(1).Control(97)=   "c_nItemPed(18)"
         Tab(1).Control(97).Enabled=   0   'False
         Tab(1).Control(98)=   "c_unidade(18)"
         Tab(1).Control(98).Enabled=   0   'False
         Tab(1).Control(99)=   "c_xPed(18)"
         Tab(1).Control(99).Enabled=   0   'False
         Tab(1).Control(100)=   "cb_ICMS_item(18)"
         Tab(1).Control(100).Enabled=   0   'False
         Tab(1).Control(101)=   "c_NCM(18)"
         Tab(1).Control(101).Enabled=   0   'False
         Tab(1).Control(102)=   "cb_CFOP(18)"
         Tab(1).Control(102).Enabled=   0   'False
         Tab(1).Control(103)=   "c_CST(18)"
         Tab(1).Control(103).Enabled=   0   'False
         Tab(1).Control(104)=   "c_vl_outras_despesas_acessorias(18)"
         Tab(1).Control(104).Enabled=   0   'False
         Tab(1).Control(105)=   "c_produto_obs(18)"
         Tab(1).Control(105).Enabled=   0   'False
         Tab(1).Control(106)=   "c_vl_total(18)"
         Tab(1).Control(106).Enabled=   0   'False
         Tab(1).Control(107)=   "c_vl_unitario(18)"
         Tab(1).Control(107).Enabled=   0   'False
         Tab(1).Control(108)=   "c_qtde(18)"
         Tab(1).Control(108).Enabled=   0   'False
         Tab(1).Control(109)=   "c_descricao(18)"
         Tab(1).Control(109).Enabled=   0   'False
         Tab(1).Control(110)=   "c_produto(18)"
         Tab(1).Control(110).Enabled=   0   'False
         Tab(1).Control(111)=   "c_fabricante(18)"
         Tab(1).Control(111).Enabled=   0   'False
         Tab(1).Control(112)=   "c_fcp(19)"
         Tab(1).Control(112).Enabled=   0   'False
         Tab(1).Control(113)=   "c_nItemPed(19)"
         Tab(1).Control(113).Enabled=   0   'False
         Tab(1).Control(114)=   "c_unidade(19)"
         Tab(1).Control(114).Enabled=   0   'False
         Tab(1).Control(115)=   "c_xPed(19)"
         Tab(1).Control(115).Enabled=   0   'False
         Tab(1).Control(116)=   "cb_ICMS_item(19)"
         Tab(1).Control(116).Enabled=   0   'False
         Tab(1).Control(117)=   "c_NCM(19)"
         Tab(1).Control(117).Enabled=   0   'False
         Tab(1).Control(118)=   "cb_CFOP(19)"
         Tab(1).Control(118).Enabled=   0   'False
         Tab(1).Control(119)=   "c_CST(19)"
         Tab(1).Control(119).Enabled=   0   'False
         Tab(1).Control(120)=   "c_vl_outras_despesas_acessorias(19)"
         Tab(1).Control(120).Enabled=   0   'False
         Tab(1).Control(121)=   "c_produto_obs(19)"
         Tab(1).Control(121).Enabled=   0   'False
         Tab(1).Control(122)=   "c_vl_total(19)"
         Tab(1).Control(122).Enabled=   0   'False
         Tab(1).Control(123)=   "c_vl_unitario(19)"
         Tab(1).Control(123).Enabled=   0   'False
         Tab(1).Control(124)=   "c_qtde(19)"
         Tab(1).Control(124).Enabled=   0   'False
         Tab(1).Control(125)=   "c_descricao(19)"
         Tab(1).Control(125).Enabled=   0   'False
         Tab(1).Control(126)=   "c_produto(19)"
         Tab(1).Control(126).Enabled=   0   'False
         Tab(1).Control(127)=   "c_fabricante(19)"
         Tab(1).Control(127).Enabled=   0   'False
         Tab(1).Control(128)=   "c_fcp(20)"
         Tab(1).Control(128).Enabled=   0   'False
         Tab(1).Control(129)=   "c_nItemPed(20)"
         Tab(1).Control(129).Enabled=   0   'False
         Tab(1).Control(130)=   "c_unidade(20)"
         Tab(1).Control(130).Enabled=   0   'False
         Tab(1).Control(131)=   "c_xPed(20)"
         Tab(1).Control(131).Enabled=   0   'False
         Tab(1).Control(132)=   "cb_ICMS_item(20)"
         Tab(1).Control(132).Enabled=   0   'False
         Tab(1).Control(133)=   "c_NCM(20)"
         Tab(1).Control(133).Enabled=   0   'False
         Tab(1).Control(134)=   "cb_CFOP(20)"
         Tab(1).Control(134).Enabled=   0   'False
         Tab(1).Control(135)=   "c_CST(20)"
         Tab(1).Control(135).Enabled=   0   'False
         Tab(1).Control(136)=   "c_vl_outras_despesas_acessorias(20)"
         Tab(1).Control(136).Enabled=   0   'False
         Tab(1).Control(137)=   "c_produto_obs(20)"
         Tab(1).Control(137).Enabled=   0   'False
         Tab(1).Control(138)=   "c_vl_total(20)"
         Tab(1).Control(138).Enabled=   0   'False
         Tab(1).Control(139)=   "c_vl_unitario(20)"
         Tab(1).Control(139).Enabled=   0   'False
         Tab(1).Control(140)=   "c_qtde(20)"
         Tab(1).Control(140).Enabled=   0   'False
         Tab(1).Control(141)=   "c_descricao(20)"
         Tab(1).Control(141).Enabled=   0   'False
         Tab(1).Control(142)=   "c_produto(20)"
         Tab(1).Control(142).Enabled=   0   'False
         Tab(1).Control(143)=   "c_fabricante(20)"
         Tab(1).Control(143).Enabled=   0   'False
         Tab(1).Control(144)=   "c_fcp(21)"
         Tab(1).Control(144).Enabled=   0   'False
         Tab(1).Control(145)=   "c_nItemPed(21)"
         Tab(1).Control(145).Enabled=   0   'False
         Tab(1).Control(146)=   "c_unidade(21)"
         Tab(1).Control(146).Enabled=   0   'False
         Tab(1).Control(147)=   "c_xPed(21)"
         Tab(1).Control(147).Enabled=   0   'False
         Tab(1).Control(148)=   "cb_ICMS_item(21)"
         Tab(1).Control(148).Enabled=   0   'False
         Tab(1).Control(149)=   "c_NCM(21)"
         Tab(1).Control(149).Enabled=   0   'False
         Tab(1).Control(150)=   "cb_CFOP(21)"
         Tab(1).Control(150).Enabled=   0   'False
         Tab(1).Control(151)=   "c_CST(21)"
         Tab(1).Control(151).Enabled=   0   'False
         Tab(1).Control(152)=   "c_vl_outras_despesas_acessorias(21)"
         Tab(1).Control(152).Enabled=   0   'False
         Tab(1).Control(153)=   "c_produto_obs(21)"
         Tab(1).Control(153).Enabled=   0   'False
         Tab(1).Control(154)=   "c_vl_total(21)"
         Tab(1).Control(154).Enabled=   0   'False
         Tab(1).Control(155)=   "c_vl_unitario(21)"
         Tab(1).Control(155).Enabled=   0   'False
         Tab(1).Control(156)=   "c_qtde(21)"
         Tab(1).Control(156).Enabled=   0   'False
         Tab(1).Control(157)=   "c_descricao(21)"
         Tab(1).Control(157).Enabled=   0   'False
         Tab(1).Control(158)=   "c_produto(21)"
         Tab(1).Control(158).Enabled=   0   'False
         Tab(1).Control(159)=   "c_fabricante(21)"
         Tab(1).Control(159).Enabled=   0   'False
         Tab(1).Control(160)=   "c_fcp(22)"
         Tab(1).Control(160).Enabled=   0   'False
         Tab(1).Control(161)=   "c_nItemPed(22)"
         Tab(1).Control(161).Enabled=   0   'False
         Tab(1).Control(162)=   "c_unidade(22)"
         Tab(1).Control(162).Enabled=   0   'False
         Tab(1).Control(163)=   "c_xPed(22)"
         Tab(1).Control(163).Enabled=   0   'False
         Tab(1).Control(164)=   "cb_ICMS_item(22)"
         Tab(1).Control(164).Enabled=   0   'False
         Tab(1).Control(165)=   "c_NCM(22)"
         Tab(1).Control(165).Enabled=   0   'False
         Tab(1).Control(166)=   "cb_CFOP(22)"
         Tab(1).Control(166).Enabled=   0   'False
         Tab(1).Control(167)=   "c_CST(22)"
         Tab(1).Control(167).Enabled=   0   'False
         Tab(1).Control(168)=   "c_vl_outras_despesas_acessorias(22)"
         Tab(1).Control(168).Enabled=   0   'False
         Tab(1).Control(169)=   "c_produto_obs(22)"
         Tab(1).Control(169).Enabled=   0   'False
         Tab(1).Control(170)=   "c_vl_total(22)"
         Tab(1).Control(170).Enabled=   0   'False
         Tab(1).Control(171)=   "c_vl_unitario(22)"
         Tab(1).Control(171).Enabled=   0   'False
         Tab(1).Control(172)=   "c_qtde(22)"
         Tab(1).Control(172).Enabled=   0   'False
         Tab(1).Control(173)=   "c_descricao(22)"
         Tab(1).Control(173).Enabled=   0   'False
         Tab(1).Control(174)=   "c_produto(22)"
         Tab(1).Control(174).Enabled=   0   'False
         Tab(1).Control(175)=   "c_fabricante(22)"
         Tab(1).Control(175).Enabled=   0   'False
         Tab(1).Control(176)=   "c_fcp(23)"
         Tab(1).Control(176).Enabled=   0   'False
         Tab(1).Control(177)=   "c_nItemPed(23)"
         Tab(1).Control(177).Enabled=   0   'False
         Tab(1).Control(178)=   "c_unidade(23)"
         Tab(1).Control(178).Enabled=   0   'False
         Tab(1).Control(179)=   "c_xPed(23)"
         Tab(1).Control(179).Enabled=   0   'False
         Tab(1).Control(180)=   "cb_ICMS_item(23)"
         Tab(1).Control(180).Enabled=   0   'False
         Tab(1).Control(181)=   "c_NCM(23)"
         Tab(1).Control(181).Enabled=   0   'False
         Tab(1).Control(182)=   "cb_CFOP(23)"
         Tab(1).Control(182).Enabled=   0   'False
         Tab(1).Control(183)=   "c_CST(23)"
         Tab(1).Control(183).Enabled=   0   'False
         Tab(1).Control(184)=   "c_vl_outras_despesas_acessorias(23)"
         Tab(1).Control(184).Enabled=   0   'False
         Tab(1).Control(185)=   "c_produto_obs(23)"
         Tab(1).Control(185).Enabled=   0   'False
         Tab(1).Control(186)=   "c_vl_total(23)"
         Tab(1).Control(186).Enabled=   0   'False
         Tab(1).Control(187)=   "c_vl_unitario(23)"
         Tab(1).Control(187).Enabled=   0   'False
         Tab(1).Control(188)=   "c_qtde(23)"
         Tab(1).Control(188).Enabled=   0   'False
         Tab(1).Control(189)=   "c_descricao(23)"
         Tab(1).Control(189).Enabled=   0   'False
         Tab(1).Control(190)=   "c_produto(23)"
         Tab(1).Control(190).Enabled=   0   'False
         Tab(1).Control(191)=   "c_fabricante(23)"
         Tab(1).Control(191).Enabled=   0   'False
         Tab(1).ControlCount=   192
         TabCaption(2)   =   "Itens 25 a 36"
         TabPicture(2)   =   "f_MAIN.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "c_fabricante(24)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "c_produto(24)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "c_descricao(24)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "c_qtde(24)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "c_vl_unitario(24)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "c_vl_total(24)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "c_produto_obs(24)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "c_vl_outras_despesas_acessorias(24)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "c_CST(24)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "cb_CFOP(24)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "c_NCM(24)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "cb_ICMS_item(24)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "c_xPed(24)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "c_unidade(24)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "c_nItemPed(24)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "c_fcp(24)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "c_fabricante(25)"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "c_produto(25)"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "c_descricao(25)"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "c_qtde(25)"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "c_vl_unitario(25)"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "c_vl_total(25)"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "c_produto_obs(25)"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "c_vl_outras_despesas_acessorias(25)"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).Control(24)=   "c_CST(25)"
         Tab(2).Control(24).Enabled=   0   'False
         Tab(2).Control(25)=   "cb_CFOP(25)"
         Tab(2).Control(25).Enabled=   0   'False
         Tab(2).Control(26)=   "c_NCM(25)"
         Tab(2).Control(26).Enabled=   0   'False
         Tab(2).Control(27)=   "cb_ICMS_item(25)"
         Tab(2).Control(27).Enabled=   0   'False
         Tab(2).Control(28)=   "c_xPed(25)"
         Tab(2).Control(28).Enabled=   0   'False
         Tab(2).Control(29)=   "c_unidade(25)"
         Tab(2).Control(29).Enabled=   0   'False
         Tab(2).Control(30)=   "c_nItemPed(25)"
         Tab(2).Control(30).Enabled=   0   'False
         Tab(2).Control(31)=   "c_fcp(25)"
         Tab(2).Control(31).Enabled=   0   'False
         Tab(2).Control(32)=   "c_fabricante(26)"
         Tab(2).Control(32).Enabled=   0   'False
         Tab(2).Control(33)=   "c_produto(26)"
         Tab(2).Control(33).Enabled=   0   'False
         Tab(2).Control(34)=   "c_descricao(26)"
         Tab(2).Control(34).Enabled=   0   'False
         Tab(2).Control(35)=   "c_qtde(26)"
         Tab(2).Control(35).Enabled=   0   'False
         Tab(2).Control(36)=   "c_vl_unitario(26)"
         Tab(2).Control(36).Enabled=   0   'False
         Tab(2).Control(37)=   "c_vl_total(26)"
         Tab(2).Control(37).Enabled=   0   'False
         Tab(2).Control(38)=   "c_produto_obs(26)"
         Tab(2).Control(38).Enabled=   0   'False
         Tab(2).Control(39)=   "c_vl_outras_despesas_acessorias(26)"
         Tab(2).Control(39).Enabled=   0   'False
         Tab(2).Control(40)=   "c_CST(26)"
         Tab(2).Control(40).Enabled=   0   'False
         Tab(2).Control(41)=   "cb_CFOP(26)"
         Tab(2).Control(41).Enabled=   0   'False
         Tab(2).Control(42)=   "c_NCM(26)"
         Tab(2).Control(42).Enabled=   0   'False
         Tab(2).Control(43)=   "cb_ICMS_item(26)"
         Tab(2).Control(43).Enabled=   0   'False
         Tab(2).Control(44)=   "c_xPed(26)"
         Tab(2).Control(44).Enabled=   0   'False
         Tab(2).Control(45)=   "c_unidade(26)"
         Tab(2).Control(45).Enabled=   0   'False
         Tab(2).Control(46)=   "c_nItemPed(26)"
         Tab(2).Control(46).Enabled=   0   'False
         Tab(2).Control(47)=   "c_fcp(26)"
         Tab(2).Control(47).Enabled=   0   'False
         Tab(2).Control(48)=   "c_fabricante(27)"
         Tab(2).Control(48).Enabled=   0   'False
         Tab(2).Control(49)=   "c_produto(27)"
         Tab(2).Control(49).Enabled=   0   'False
         Tab(2).Control(50)=   "c_descricao(27)"
         Tab(2).Control(50).Enabled=   0   'False
         Tab(2).Control(51)=   "c_qtde(27)"
         Tab(2).Control(51).Enabled=   0   'False
         Tab(2).Control(52)=   "c_vl_unitario(27)"
         Tab(2).Control(52).Enabled=   0   'False
         Tab(2).Control(53)=   "c_vl_total(27)"
         Tab(2).Control(53).Enabled=   0   'False
         Tab(2).Control(54)=   "c_produto_obs(27)"
         Tab(2).Control(54).Enabled=   0   'False
         Tab(2).Control(55)=   "c_vl_outras_despesas_acessorias(27)"
         Tab(2).Control(55).Enabled=   0   'False
         Tab(2).Control(56)=   "c_CST(27)"
         Tab(2).Control(56).Enabled=   0   'False
         Tab(2).Control(57)=   "cb_CFOP(27)"
         Tab(2).Control(57).Enabled=   0   'False
         Tab(2).Control(58)=   "c_NCM(27)"
         Tab(2).Control(58).Enabled=   0   'False
         Tab(2).Control(59)=   "cb_ICMS_item(27)"
         Tab(2).Control(59).Enabled=   0   'False
         Tab(2).Control(60)=   "c_xPed(27)"
         Tab(2).Control(60).Enabled=   0   'False
         Tab(2).Control(61)=   "c_unidade(27)"
         Tab(2).Control(61).Enabled=   0   'False
         Tab(2).Control(62)=   "c_nItemPed(27)"
         Tab(2).Control(62).Enabled=   0   'False
         Tab(2).Control(63)=   "c_fcp(27)"
         Tab(2).Control(63).Enabled=   0   'False
         Tab(2).Control(64)=   "c_fabricante(28)"
         Tab(2).Control(64).Enabled=   0   'False
         Tab(2).Control(65)=   "c_produto(28)"
         Tab(2).Control(65).Enabled=   0   'False
         Tab(2).Control(66)=   "c_descricao(28)"
         Tab(2).Control(66).Enabled=   0   'False
         Tab(2).Control(67)=   "c_qtde(28)"
         Tab(2).Control(67).Enabled=   0   'False
         Tab(2).Control(68)=   "c_vl_unitario(28)"
         Tab(2).Control(68).Enabled=   0   'False
         Tab(2).Control(69)=   "c_vl_total(28)"
         Tab(2).Control(69).Enabled=   0   'False
         Tab(2).Control(70)=   "c_produto_obs(28)"
         Tab(2).Control(70).Enabled=   0   'False
         Tab(2).Control(71)=   "c_vl_outras_despesas_acessorias(28)"
         Tab(2).Control(71).Enabled=   0   'False
         Tab(2).Control(72)=   "c_CST(28)"
         Tab(2).Control(72).Enabled=   0   'False
         Tab(2).Control(73)=   "cb_CFOP(28)"
         Tab(2).Control(73).Enabled=   0   'False
         Tab(2).Control(74)=   "c_NCM(28)"
         Tab(2).Control(74).Enabled=   0   'False
         Tab(2).Control(75)=   "cb_ICMS_item(28)"
         Tab(2).Control(75).Enabled=   0   'False
         Tab(2).Control(76)=   "c_xPed(28)"
         Tab(2).Control(76).Enabled=   0   'False
         Tab(2).Control(77)=   "c_unidade(28)"
         Tab(2).Control(77).Enabled=   0   'False
         Tab(2).Control(78)=   "c_nItemPed(28)"
         Tab(2).Control(78).Enabled=   0   'False
         Tab(2).Control(79)=   "c_fcp(28)"
         Tab(2).Control(79).Enabled=   0   'False
         Tab(2).Control(80)=   "c_fabricante(29)"
         Tab(2).Control(80).Enabled=   0   'False
         Tab(2).Control(81)=   "c_produto(29)"
         Tab(2).Control(81).Enabled=   0   'False
         Tab(2).Control(82)=   "c_descricao(29)"
         Tab(2).Control(82).Enabled=   0   'False
         Tab(2).Control(83)=   "c_qtde(29)"
         Tab(2).Control(83).Enabled=   0   'False
         Tab(2).Control(84)=   "c_vl_unitario(29)"
         Tab(2).Control(84).Enabled=   0   'False
         Tab(2).Control(85)=   "c_vl_total(29)"
         Tab(2).Control(85).Enabled=   0   'False
         Tab(2).Control(86)=   "c_produto_obs(29)"
         Tab(2).Control(86).Enabled=   0   'False
         Tab(2).Control(87)=   "c_vl_outras_despesas_acessorias(29)"
         Tab(2).Control(87).Enabled=   0   'False
         Tab(2).Control(88)=   "c_CST(29)"
         Tab(2).Control(88).Enabled=   0   'False
         Tab(2).Control(89)=   "cb_CFOP(29)"
         Tab(2).Control(89).Enabled=   0   'False
         Tab(2).Control(90)=   "c_NCM(29)"
         Tab(2).Control(90).Enabled=   0   'False
         Tab(2).Control(91)=   "cb_ICMS_item(29)"
         Tab(2).Control(91).Enabled=   0   'False
         Tab(2).Control(92)=   "c_xPed(29)"
         Tab(2).Control(92).Enabled=   0   'False
         Tab(2).Control(93)=   "c_unidade(29)"
         Tab(2).Control(93).Enabled=   0   'False
         Tab(2).Control(94)=   "c_nItemPed(29)"
         Tab(2).Control(94).Enabled=   0   'False
         Tab(2).Control(95)=   "c_fcp(29)"
         Tab(2).Control(95).Enabled=   0   'False
         Tab(2).Control(96)=   "c_fabricante(30)"
         Tab(2).Control(96).Enabled=   0   'False
         Tab(2).Control(97)=   "c_produto(30)"
         Tab(2).Control(97).Enabled=   0   'False
         Tab(2).Control(98)=   "c_descricao(30)"
         Tab(2).Control(98).Enabled=   0   'False
         Tab(2).Control(99)=   "c_qtde(30)"
         Tab(2).Control(99).Enabled=   0   'False
         Tab(2).Control(100)=   "c_vl_unitario(30)"
         Tab(2).Control(100).Enabled=   0   'False
         Tab(2).Control(101)=   "c_vl_total(30)"
         Tab(2).Control(101).Enabled=   0   'False
         Tab(2).Control(102)=   "c_produto_obs(30)"
         Tab(2).Control(102).Enabled=   0   'False
         Tab(2).Control(103)=   "c_vl_outras_despesas_acessorias(30)"
         Tab(2).Control(103).Enabled=   0   'False
         Tab(2).Control(104)=   "c_CST(30)"
         Tab(2).Control(104).Enabled=   0   'False
         Tab(2).Control(105)=   "cb_CFOP(30)"
         Tab(2).Control(105).Enabled=   0   'False
         Tab(2).Control(106)=   "c_NCM(30)"
         Tab(2).Control(106).Enabled=   0   'False
         Tab(2).Control(107)=   "cb_ICMS_item(30)"
         Tab(2).Control(107).Enabled=   0   'False
         Tab(2).Control(108)=   "c_xPed(30)"
         Tab(2).Control(108).Enabled=   0   'False
         Tab(2).Control(109)=   "c_unidade(30)"
         Tab(2).Control(109).Enabled=   0   'False
         Tab(2).Control(110)=   "c_nItemPed(30)"
         Tab(2).Control(110).Enabled=   0   'False
         Tab(2).Control(111)=   "c_fcp(30)"
         Tab(2).Control(111).Enabled=   0   'False
         Tab(2).Control(112)=   "c_fabricante(31)"
         Tab(2).Control(112).Enabled=   0   'False
         Tab(2).Control(113)=   "c_produto(31)"
         Tab(2).Control(113).Enabled=   0   'False
         Tab(2).Control(114)=   "c_descricao(31)"
         Tab(2).Control(114).Enabled=   0   'False
         Tab(2).Control(115)=   "c_qtde(31)"
         Tab(2).Control(115).Enabled=   0   'False
         Tab(2).Control(116)=   "c_vl_unitario(31)"
         Tab(2).Control(116).Enabled=   0   'False
         Tab(2).Control(117)=   "c_vl_total(31)"
         Tab(2).Control(117).Enabled=   0   'False
         Tab(2).Control(118)=   "c_produto_obs(31)"
         Tab(2).Control(118).Enabled=   0   'False
         Tab(2).Control(119)=   "c_vl_outras_despesas_acessorias(31)"
         Tab(2).Control(119).Enabled=   0   'False
         Tab(2).Control(120)=   "c_CST(31)"
         Tab(2).Control(120).Enabled=   0   'False
         Tab(2).Control(121)=   "cb_CFOP(31)"
         Tab(2).Control(121).Enabled=   0   'False
         Tab(2).Control(122)=   "c_NCM(31)"
         Tab(2).Control(122).Enabled=   0   'False
         Tab(2).Control(123)=   "cb_ICMS_item(31)"
         Tab(2).Control(123).Enabled=   0   'False
         Tab(2).Control(124)=   "c_xPed(31)"
         Tab(2).Control(124).Enabled=   0   'False
         Tab(2).Control(125)=   "c_unidade(31)"
         Tab(2).Control(125).Enabled=   0   'False
         Tab(2).Control(126)=   "c_nItemPed(31)"
         Tab(2).Control(126).Enabled=   0   'False
         Tab(2).Control(127)=   "c_fcp(31)"
         Tab(2).Control(127).Enabled=   0   'False
         Tab(2).Control(128)=   "c_fabricante(32)"
         Tab(2).Control(128).Enabled=   0   'False
         Tab(2).Control(129)=   "c_produto(32)"
         Tab(2).Control(129).Enabled=   0   'False
         Tab(2).Control(130)=   "c_descricao(32)"
         Tab(2).Control(130).Enabled=   0   'False
         Tab(2).Control(131)=   "c_qtde(32)"
         Tab(2).Control(131).Enabled=   0   'False
         Tab(2).Control(132)=   "c_vl_unitario(32)"
         Tab(2).Control(132).Enabled=   0   'False
         Tab(2).Control(133)=   "c_vl_total(32)"
         Tab(2).Control(133).Enabled=   0   'False
         Tab(2).Control(134)=   "c_produto_obs(32)"
         Tab(2).Control(134).Enabled=   0   'False
         Tab(2).Control(135)=   "c_vl_outras_despesas_acessorias(32)"
         Tab(2).Control(135).Enabled=   0   'False
         Tab(2).Control(136)=   "c_CST(32)"
         Tab(2).Control(136).Enabled=   0   'False
         Tab(2).Control(137)=   "cb_CFOP(32)"
         Tab(2).Control(137).Enabled=   0   'False
         Tab(2).Control(138)=   "c_NCM(32)"
         Tab(2).Control(138).Enabled=   0   'False
         Tab(2).Control(139)=   "cb_ICMS_item(32)"
         Tab(2).Control(139).Enabled=   0   'False
         Tab(2).Control(140)=   "c_xPed(32)"
         Tab(2).Control(140).Enabled=   0   'False
         Tab(2).Control(141)=   "c_unidade(32)"
         Tab(2).Control(141).Enabled=   0   'False
         Tab(2).Control(142)=   "c_nItemPed(32)"
         Tab(2).Control(142).Enabled=   0   'False
         Tab(2).Control(143)=   "c_fcp(32)"
         Tab(2).Control(143).Enabled=   0   'False
         Tab(2).Control(144)=   "c_fabricante(33)"
         Tab(2).Control(144).Enabled=   0   'False
         Tab(2).Control(145)=   "c_produto(33)"
         Tab(2).Control(145).Enabled=   0   'False
         Tab(2).Control(146)=   "c_descricao(33)"
         Tab(2).Control(146).Enabled=   0   'False
         Tab(2).Control(147)=   "c_qtde(33)"
         Tab(2).Control(147).Enabled=   0   'False
         Tab(2).Control(148)=   "c_vl_unitario(33)"
         Tab(2).Control(148).Enabled=   0   'False
         Tab(2).Control(149)=   "c_vl_total(33)"
         Tab(2).Control(149).Enabled=   0   'False
         Tab(2).Control(150)=   "c_produto_obs(33)"
         Tab(2).Control(150).Enabled=   0   'False
         Tab(2).Control(151)=   "c_vl_outras_despesas_acessorias(33)"
         Tab(2).Control(151).Enabled=   0   'False
         Tab(2).Control(152)=   "c_CST(33)"
         Tab(2).Control(152).Enabled=   0   'False
         Tab(2).Control(153)=   "cb_CFOP(33)"
         Tab(2).Control(153).Enabled=   0   'False
         Tab(2).Control(154)=   "c_NCM(33)"
         Tab(2).Control(154).Enabled=   0   'False
         Tab(2).Control(155)=   "cb_ICMS_item(33)"
         Tab(2).Control(155).Enabled=   0   'False
         Tab(2).Control(156)=   "c_xPed(33)"
         Tab(2).Control(156).Enabled=   0   'False
         Tab(2).Control(157)=   "c_unidade(33)"
         Tab(2).Control(157).Enabled=   0   'False
         Tab(2).Control(158)=   "c_nItemPed(33)"
         Tab(2).Control(158).Enabled=   0   'False
         Tab(2).Control(159)=   "c_fcp(33)"
         Tab(2).Control(159).Enabled=   0   'False
         Tab(2).Control(160)=   "c_fabricante(34)"
         Tab(2).Control(160).Enabled=   0   'False
         Tab(2).Control(161)=   "c_produto(34)"
         Tab(2).Control(161).Enabled=   0   'False
         Tab(2).Control(162)=   "c_descricao(34)"
         Tab(2).Control(162).Enabled=   0   'False
         Tab(2).Control(163)=   "c_qtde(34)"
         Tab(2).Control(163).Enabled=   0   'False
         Tab(2).Control(164)=   "c_vl_unitario(34)"
         Tab(2).Control(164).Enabled=   0   'False
         Tab(2).Control(165)=   "c_vl_total(34)"
         Tab(2).Control(165).Enabled=   0   'False
         Tab(2).Control(166)=   "c_produto_obs(34)"
         Tab(2).Control(166).Enabled=   0   'False
         Tab(2).Control(167)=   "c_vl_outras_despesas_acessorias(34)"
         Tab(2).Control(167).Enabled=   0   'False
         Tab(2).Control(168)=   "c_CST(34)"
         Tab(2).Control(168).Enabled=   0   'False
         Tab(2).Control(169)=   "cb_CFOP(34)"
         Tab(2).Control(169).Enabled=   0   'False
         Tab(2).Control(170)=   "c_NCM(34)"
         Tab(2).Control(170).Enabled=   0   'False
         Tab(2).Control(171)=   "cb_ICMS_item(34)"
         Tab(2).Control(171).Enabled=   0   'False
         Tab(2).Control(172)=   "c_xPed(34)"
         Tab(2).Control(172).Enabled=   0   'False
         Tab(2).Control(173)=   "c_unidade(34)"
         Tab(2).Control(173).Enabled=   0   'False
         Tab(2).Control(174)=   "c_nItemPed(34)"
         Tab(2).Control(174).Enabled=   0   'False
         Tab(2).Control(175)=   "c_fcp(34)"
         Tab(2).Control(175).Enabled=   0   'False
         Tab(2).Control(176)=   "c_fabricante(35)"
         Tab(2).Control(176).Enabled=   0   'False
         Tab(2).Control(177)=   "c_produto(35)"
         Tab(2).Control(177).Enabled=   0   'False
         Tab(2).Control(178)=   "c_descricao(35)"
         Tab(2).Control(178).Enabled=   0   'False
         Tab(2).Control(179)=   "c_qtde(35)"
         Tab(2).Control(179).Enabled=   0   'False
         Tab(2).Control(180)=   "c_vl_unitario(35)"
         Tab(2).Control(180).Enabled=   0   'False
         Tab(2).Control(181)=   "c_vl_total(35)"
         Tab(2).Control(181).Enabled=   0   'False
         Tab(2).Control(182)=   "c_produto_obs(35)"
         Tab(2).Control(182).Enabled=   0   'False
         Tab(2).Control(183)=   "c_vl_outras_despesas_acessorias(35)"
         Tab(2).Control(183).Enabled=   0   'False
         Tab(2).Control(184)=   "c_CST(35)"
         Tab(2).Control(184).Enabled=   0   'False
         Tab(2).Control(185)=   "cb_CFOP(35)"
         Tab(2).Control(185).Enabled=   0   'False
         Tab(2).Control(186)=   "c_NCM(35)"
         Tab(2).Control(186).Enabled=   0   'False
         Tab(2).Control(187)=   "cb_ICMS_item(35)"
         Tab(2).Control(187).Enabled=   0   'False
         Tab(2).Control(188)=   "c_xPed(35)"
         Tab(2).Control(188).Enabled=   0   'False
         Tab(2).Control(189)=   "c_unidade(35)"
         Tab(2).Control(189).Enabled=   0   'False
         Tab(2).Control(190)=   "c_nItemPed(35)"
         Tab(2).Control(190).Enabled=   0   'False
         Tab(2).Control(191)=   "c_fcp(35)"
         Tab(2).Control(191).Enabled=   0   'False
         Tab(2).ControlCount=   192
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   649
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   645
            MaxLength       =   8
            TabIndex        =   648
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   647
            Top             =   360
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   646
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   645
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   644
            TabStop         =   0   'False
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   4
            TabIndex        =   643
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   645
            MaxLength       =   8
            TabIndex        =   642
            Top             =   645
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   641
            Top             =   645
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   640
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   639
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   638
            TabStop         =   0   'False
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   4
            TabIndex        =   637
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   645
            MaxLength       =   8
            TabIndex        =   636
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   635
            Top             =   930
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   634
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   633
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   632
            TabStop         =   0   'False
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   120
            MaxLength       =   4
            TabIndex        =   631
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   645
            MaxLength       =   8
            TabIndex        =   630
            Top             =   1215
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   629
            Top             =   1215
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   628
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   627
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   626
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   120
            MaxLength       =   4
            TabIndex        =   625
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   645
            MaxLength       =   8
            TabIndex        =   624
            Top             =   1500
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   623
            Top             =   1500
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   622
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   621
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   620
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   120
            MaxLength       =   4
            TabIndex        =   619
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   645
            MaxLength       =   8
            TabIndex        =   618
            Top             =   1785
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   617
            Top             =   1785
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   616
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   615
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   614
            TabStop         =   0   'False
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   120
            MaxLength       =   4
            TabIndex        =   613
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   645
            MaxLength       =   8
            TabIndex        =   612
            Top             =   2070
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   611
            Top             =   2070
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   610
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   609
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   608
            TabStop         =   0   'False
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   120
            MaxLength       =   4
            TabIndex        =   607
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   645
            MaxLength       =   8
            TabIndex        =   606
            Top             =   2355
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   605
            Top             =   2355
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   604
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   603
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   602
            TabStop         =   0   'False
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   120
            MaxLength       =   4
            TabIndex        =   601
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   645
            MaxLength       =   8
            TabIndex        =   600
            Top             =   2640
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   599
            Top             =   2640
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   598
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   597
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   596
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   120
            MaxLength       =   4
            TabIndex        =   595
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   645
            MaxLength       =   8
            TabIndex        =   594
            Top             =   2925
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   593
            Top             =   2925
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   592
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   591
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   590
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   0
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   589
            Top             =   360
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   1
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   588
            Top             =   645
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   2
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   587
            Top             =   930
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   3
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   586
            Top             =   1215
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   4
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   585
            Top             =   1500
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   5
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   584
            Top             =   1785
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   6
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   583
            Top             =   2070
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   7
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   582
            Top             =   2355
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   8
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   581
            Top             =   2640
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   9
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   580
            Top             =   2925
            Width           =   1705
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   10
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   579
            Top             =   3210
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   578
            TabStop         =   0   'False
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   577
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   576
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   575
            Top             =   3210
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   645
            MaxLength       =   8
            TabIndex        =   574
            Top             =   3210
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   120
            MaxLength       =   4
            TabIndex        =   573
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   120
            MaxLength       =   4
            TabIndex        =   572
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   645
            MaxLength       =   8
            TabIndex        =   571
            Top             =   3495
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   570
            Top             =   3495
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   8550
            MaxLength       =   5
            TabIndex        =   569
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   9165
            MaxLength       =   18
            TabIndex        =   568
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   10470
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   567
            TabStop         =   0   'False
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   11
            Left            =   6855
            MaxLength       =   500
            TabIndex        =   566
            Top             =   3495
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   565
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   564
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   563
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   562
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   561
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   560
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   559
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   558
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   557
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   556
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   555
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   11775
            MaxLength       =   18
            TabIndex        =   554
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   553
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   552
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   551
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   550
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   549
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   548
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   547
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   546
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   545
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   544
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   543
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   13080
            MaxLength       =   4
            TabIndex        =   542
            Top             =   360
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   0
            ItemData        =   "f_MAIN.frx":0496
            Left            =   13605
            List            =   "f_MAIN.frx":0498
            Style           =   2  'Dropdown List
            TabIndex        =   541
            Top             =   360
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   1
            ItemData        =   "f_MAIN.frx":049A
            Left            =   13605
            List            =   "f_MAIN.frx":049C
            Style           =   2  'Dropdown List
            TabIndex        =   540
            Top             =   645
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   2
            ItemData        =   "f_MAIN.frx":049E
            Left            =   13605
            List            =   "f_MAIN.frx":04A0
            Style           =   2  'Dropdown List
            TabIndex        =   539
            Top             =   930
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   3
            ItemData        =   "f_MAIN.frx":04A2
            Left            =   13605
            List            =   "f_MAIN.frx":04A4
            Style           =   2  'Dropdown List
            TabIndex        =   538
            Top             =   1215
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   4
            ItemData        =   "f_MAIN.frx":04A6
            Left            =   13605
            List            =   "f_MAIN.frx":04A8
            Style           =   2  'Dropdown List
            TabIndex        =   537
            Top             =   1500
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   5
            ItemData        =   "f_MAIN.frx":04AA
            Left            =   13605
            List            =   "f_MAIN.frx":04AC
            Style           =   2  'Dropdown List
            TabIndex        =   536
            Top             =   1785
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   6
            ItemData        =   "f_MAIN.frx":04AE
            Left            =   13605
            List            =   "f_MAIN.frx":04B0
            Style           =   2  'Dropdown List
            TabIndex        =   535
            Top             =   2070
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   7
            ItemData        =   "f_MAIN.frx":04B2
            Left            =   13605
            List            =   "f_MAIN.frx":04B4
            Style           =   2  'Dropdown List
            TabIndex        =   534
            Top             =   2355
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   8
            ItemData        =   "f_MAIN.frx":04B6
            Left            =   13605
            List            =   "f_MAIN.frx":04B8
            Style           =   2  'Dropdown List
            TabIndex        =   533
            Top             =   2640
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   9
            ItemData        =   "f_MAIN.frx":04BA
            Left            =   13605
            List            =   "f_MAIN.frx":04BC
            Style           =   2  'Dropdown List
            TabIndex        =   532
            Top             =   2925
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   10
            ItemData        =   "f_MAIN.frx":04BE
            Left            =   13605
            List            =   "f_MAIN.frx":04C0
            Style           =   2  'Dropdown List
            TabIndex        =   531
            Top             =   3210
            Width           =   1985
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   11
            ItemData        =   "f_MAIN.frx":04C2
            Left            =   13605
            List            =   "f_MAIN.frx":04C4
            Style           =   2  'Dropdown List
            TabIndex        =   530
            Top             =   3495
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   529
            Top             =   360
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
            Left            =   16485
            TabIndex        =   528
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   527
            Top             =   375
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   0
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   526
            Top             =   375
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   0
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   525
            Top             =   375
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
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   524
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   523
            Top             =   645
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
            Index           =   1
            Left            =   16485
            TabIndex        =   522
            Top             =   645
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   521
            Top             =   645
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   1
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   520
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   1
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   519
            Top             =   645
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
            Index           =   1
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   518
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   517
            Top             =   930
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
            Index           =   2
            Left            =   16485
            TabIndex        =   516
            Top             =   930
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   515
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   2
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   514
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   2
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   513
            Top             =   930
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
            Index           =   2
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   512
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   511
            Top             =   1215
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
            Index           =   3
            Left            =   16485
            TabIndex        =   510
            Top             =   1215
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   509
            Top             =   1215
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   3
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   508
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   3
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   507
            Top             =   1215
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
            Index           =   3
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   506
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   505
            Top             =   1500
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
            Index           =   4
            Left            =   16485
            TabIndex        =   504
            Top             =   1500
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   503
            Top             =   1500
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   4
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   502
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   4
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   501
            Top             =   1500
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
            Index           =   4
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   500
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   499
            Top             =   1785
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
            Index           =   5
            Left            =   16485
            TabIndex        =   498
            Top             =   1785
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   497
            Top             =   1785
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   5
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   496
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   5
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   495
            Top             =   1785
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
            Index           =   5
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   494
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   493
            Top             =   2070
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
            Index           =   6
            Left            =   16485
            TabIndex        =   492
            Top             =   2070
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   491
            Top             =   2070
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   6
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   490
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   6
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   489
            Top             =   2070
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
            Index           =   6
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   488
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   487
            Top             =   2355
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
            Index           =   7
            Left            =   16485
            TabIndex        =   486
            Top             =   2355
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   485
            Top             =   2355
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   7
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   484
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   7
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   483
            Top             =   2355
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
            Index           =   7
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   482
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   481
            Top             =   2640
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
            Index           =   8
            Left            =   16485
            TabIndex        =   480
            Top             =   2640
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   479
            Top             =   2640
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   8
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   478
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   8
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   477
            Top             =   2640
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
            Index           =   8
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   476
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   475
            Top             =   2925
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
            Index           =   9
            Left            =   16485
            TabIndex        =   474
            Top             =   2925
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   473
            Top             =   2925
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   9
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   472
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   9
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   471
            Top             =   2925
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
            Index           =   9
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   470
            Top             =   2925
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
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   469
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   10
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   468
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   10
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   467
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   466
            Top             =   3210
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
            Index           =   10
            Left            =   16485
            TabIndex        =   465
            Top             =   3210
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   464
            Top             =   3210
            Width           =   885
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
            Left            =   19470
            MaxLength       =   6
            TabIndex        =   463
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   11
            Left            =   18000
            MaxLength       =   6
            TabIndex        =   462
            Top             =   3495
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   11
            Left            =   18795
            MaxLength       =   6
            TabIndex        =   461
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   17070
            MaxLength       =   15
            TabIndex        =   460
            Top             =   3495
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
            Left            =   16485
            TabIndex        =   459
            Top             =   3495
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   15600
            MaxLength       =   8
            TabIndex        =   458
            Top             =   3495
            Width           =   885
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
            Index           =   12
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   457
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   12
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   456
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   12
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   455
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   454
            Top             =   360
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
            Index           =   12
            Left            =   -58515
            TabIndex        =   453
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   452
            Top             =   360
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   12
            ItemData        =   "f_MAIN.frx":04C6
            Left            =   -61395
            List            =   "f_MAIN.frx":04C8
            Style           =   2  'Dropdown List
            TabIndex        =   451
            Top             =   360
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   450
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   449
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   12
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   448
            Top             =   360
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   447
            TabStop         =   0   'False
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   446
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   445
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   444
            Top             =   360
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   443
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   442
            Top             =   360
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
            Index           =   13
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   441
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   13
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   440
            Top             =   645
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   13
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   439
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   438
            Top             =   645
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
            Index           =   13
            Left            =   -58515
            TabIndex        =   437
            Top             =   645
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   436
            Top             =   645
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   13
            ItemData        =   "f_MAIN.frx":04CA
            Left            =   -61395
            List            =   "f_MAIN.frx":04CC
            Style           =   2  'Dropdown List
            TabIndex        =   435
            Top             =   645
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   434
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   433
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   13
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   432
            Top             =   645
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   431
            TabStop         =   0   'False
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   430
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   429
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   428
            Top             =   645
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   427
            Top             =   645
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   426
            Top             =   645
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
            Index           =   14
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   425
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   14
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   424
            Top             =   930
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   14
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   423
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   422
            Top             =   930
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
            Index           =   14
            Left            =   -58515
            TabIndex        =   421
            Top             =   930
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   420
            Top             =   930
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   14
            ItemData        =   "f_MAIN.frx":04CE
            Left            =   -61395
            List            =   "f_MAIN.frx":04D0
            Style           =   2  'Dropdown List
            TabIndex        =   419
            Top             =   930
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   418
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   417
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   14
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   416
            Top             =   930
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   415
            TabStop         =   0   'False
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   414
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   413
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   412
            Top             =   930
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   411
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   410
            Top             =   930
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
            Index           =   15
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   409
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   15
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   408
            Top             =   1215
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   15
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   407
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   406
            Top             =   1215
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
            Index           =   15
            Left            =   -58515
            TabIndex        =   405
            Top             =   1215
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   404
            Top             =   1215
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   15
            ItemData        =   "f_MAIN.frx":04D2
            Left            =   -61395
            List            =   "f_MAIN.frx":04D4
            Style           =   2  'Dropdown List
            TabIndex        =   403
            Top             =   1215
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   402
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   401
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   15
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   400
            Top             =   1215
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   399
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   398
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   397
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   396
            Top             =   1215
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   395
            Top             =   1215
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   394
            Top             =   1215
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
            Index           =   16
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   393
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   16
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   392
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   16
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   391
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   390
            Top             =   1500
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
            Index           =   16
            Left            =   -58515
            TabIndex        =   389
            Top             =   1500
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   388
            Top             =   1500
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   16
            ItemData        =   "f_MAIN.frx":04D6
            Left            =   -61395
            List            =   "f_MAIN.frx":04D8
            Style           =   2  'Dropdown List
            TabIndex        =   387
            Top             =   1500
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   386
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   385
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   16
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   384
            Top             =   1500
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   383
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   382
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   381
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   380
            Top             =   1500
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   379
            Top             =   1500
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   378
            Top             =   1500
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
            Index           =   17
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   377
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   17
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   376
            Top             =   1785
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   17
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   375
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   374
            Top             =   1785
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
            Index           =   17
            Left            =   -58515
            TabIndex        =   373
            Top             =   1785
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   372
            Top             =   1785
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   17
            ItemData        =   "f_MAIN.frx":04DA
            Left            =   -61395
            List            =   "f_MAIN.frx":04DC
            Style           =   2  'Dropdown List
            TabIndex        =   371
            Top             =   1785
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   370
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   369
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   17
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   368
            Top             =   1785
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   367
            TabStop         =   0   'False
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   366
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   365
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   364
            Top             =   1785
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   363
            Top             =   1785
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   362
            Top             =   1785
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
            Index           =   18
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   361
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   18
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   360
            Top             =   2070
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   18
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   359
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   358
            Top             =   2070
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
            Index           =   18
            Left            =   -58515
            TabIndex        =   357
            Top             =   2070
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   356
            Top             =   2070
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   18
            ItemData        =   "f_MAIN.frx":04DE
            Left            =   -61395
            List            =   "f_MAIN.frx":04E0
            Style           =   2  'Dropdown List
            TabIndex        =   355
            Top             =   2070
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   354
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   353
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   18
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   352
            Top             =   2070
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   351
            TabStop         =   0   'False
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   350
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   349
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   348
            Top             =   2070
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   347
            Top             =   2070
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   346
            Top             =   2070
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
            Index           =   19
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   345
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   19
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   344
            Top             =   2355
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   19
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   343
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   342
            Top             =   2355
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
            Index           =   19
            Left            =   -58515
            TabIndex        =   341
            Top             =   2355
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   340
            Top             =   2355
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   19
            ItemData        =   "f_MAIN.frx":04E2
            Left            =   -61395
            List            =   "f_MAIN.frx":04E4
            Style           =   2  'Dropdown List
            TabIndex        =   339
            Top             =   2355
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   338
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   337
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   19
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   336
            Top             =   2355
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   335
            TabStop         =   0   'False
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   334
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   333
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   332
            Top             =   2355
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   331
            Top             =   2355
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   330
            Top             =   2355
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
            Index           =   20
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   329
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   20
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   328
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   20
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   327
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   326
            Top             =   2640
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
            Index           =   20
            Left            =   -58515
            TabIndex        =   325
            Top             =   2640
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   324
            Top             =   2640
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   20
            ItemData        =   "f_MAIN.frx":04E6
            Left            =   -61395
            List            =   "f_MAIN.frx":04E8
            Style           =   2  'Dropdown List
            TabIndex        =   323
            Top             =   2640
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   322
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   321
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   20
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   320
            Top             =   2640
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   319
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   318
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   317
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   316
            Top             =   2640
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   315
            Top             =   2640
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   314
            Top             =   2640
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
            Index           =   21
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   313
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   21
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   312
            Top             =   2925
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   21
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   311
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   310
            Top             =   2925
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
            Index           =   21
            Left            =   -58515
            TabIndex        =   309
            Top             =   2925
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   308
            Top             =   2925
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   21
            ItemData        =   "f_MAIN.frx":04EA
            Left            =   -61395
            List            =   "f_MAIN.frx":04EC
            Style           =   2  'Dropdown List
            TabIndex        =   307
            Top             =   2925
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   306
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   305
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   21
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   304
            Top             =   2925
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   303
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   302
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   301
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   300
            Top             =   2925
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   299
            Top             =   2925
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   298
            Top             =   2925
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
            Index           =   22
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   297
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   22
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   296
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   22
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   295
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   294
            Top             =   3210
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
            Index           =   22
            Left            =   -58515
            TabIndex        =   293
            Top             =   3210
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   292
            Top             =   3210
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   22
            ItemData        =   "f_MAIN.frx":04EE
            Left            =   -61395
            List            =   "f_MAIN.frx":04F0
            Style           =   2  'Dropdown List
            TabIndex        =   291
            Top             =   3210
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   290
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   289
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   22
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   288
            Top             =   3210
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   287
            TabStop         =   0   'False
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   286
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   285
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   284
            Top             =   3210
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   283
            Top             =   3210
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   282
            Top             =   3210
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
            Index           =   23
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   281
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   23
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   280
            Top             =   3495
            Width           =   735
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   23
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   279
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   278
            Top             =   3495
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
            Index           =   23
            Left            =   -58515
            TabIndex        =   277
            Top             =   3495
            Width           =   585
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   276
            Top             =   3495
            Width           =   885
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   23
            ItemData        =   "f_MAIN.frx":04F2
            Left            =   -61395
            List            =   "f_MAIN.frx":04F4
            Style           =   2  'Dropdown List
            TabIndex        =   275
            Top             =   3495
            Width           =   1985
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   274
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   273
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   23
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   272
            Top             =   3495
            Width           =   1705
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   271
            TabStop         =   0   'False
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   270
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   269
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   268
            Top             =   3495
            Width           =   5320
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   267
            Top             =   3495
            Width           =   885
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   266
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   265
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   264
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   263
            Top             =   360
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   262
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   261
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   260
            TabStop         =   0   'False
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   24
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   259
            Top             =   360
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   258
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   257
            Top             =   360
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   24
            ItemData        =   "f_MAIN.frx":04F6
            Left            =   -61395
            List            =   "f_MAIN.frx":04F8
            Style           =   2  'Dropdown List
            TabIndex        =   256
            Top             =   360
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   255
            Top             =   360
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
            Index           =   24
            Left            =   -58515
            TabIndex        =   254
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   253
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   24
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   252
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   24
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   251
            Top             =   360
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
            Index           =   24
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   250
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   249
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   248
            Top             =   645
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   247
            Top             =   645
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   246
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   245
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   244
            TabStop         =   0   'False
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   25
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   243
            Top             =   645
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   242
            Top             =   645
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   241
            Top             =   645
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   25
            ItemData        =   "f_MAIN.frx":04FA
            Left            =   -61395
            List            =   "f_MAIN.frx":04FC
            Style           =   2  'Dropdown List
            TabIndex        =   240
            Top             =   645
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   239
            Top             =   645
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
            Index           =   25
            Left            =   -58515
            TabIndex        =   238
            Top             =   645
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   237
            Top             =   645
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   25
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   236
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   25
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   235
            Top             =   645
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
            Index           =   25
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   234
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   233
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   232
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   231
            Top             =   930
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   230
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   229
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   228
            TabStop         =   0   'False
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   26
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   227
            Top             =   930
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   226
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   225
            Top             =   930
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   26
            ItemData        =   "f_MAIN.frx":04FE
            Left            =   -61395
            List            =   "f_MAIN.frx":0500
            Style           =   2  'Dropdown List
            TabIndex        =   224
            Top             =   930
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   223
            Top             =   930
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
            Index           =   26
            Left            =   -58515
            TabIndex        =   222
            Top             =   930
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   221
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   26
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   220
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   26
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   219
            Top             =   930
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
            Index           =   26
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   218
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   217
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   216
            Top             =   1215
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   215
            Top             =   1215
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   214
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   213
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   27
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   211
            Top             =   1215
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   210
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   209
            Top             =   1215
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   27
            ItemData        =   "f_MAIN.frx":0502
            Left            =   -61395
            List            =   "f_MAIN.frx":0504
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   1215
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   207
            Top             =   1215
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
            Index           =   27
            Left            =   -58515
            TabIndex        =   206
            Top             =   1215
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   205
            Top             =   1215
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   27
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   204
            Top             =   1215
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   27
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   203
            Top             =   1215
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
            Index           =   27
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   202
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   201
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   200
            Top             =   1500
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   199
            Top             =   1500
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   198
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   197
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   28
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   195
            Top             =   1500
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   194
            Top             =   1500
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   193
            Top             =   1500
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   28
            ItemData        =   "f_MAIN.frx":0506
            Left            =   -61395
            List            =   "f_MAIN.frx":0508
            Style           =   2  'Dropdown List
            TabIndex        =   192
            Top             =   1500
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   191
            Top             =   1500
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
            Index           =   28
            Left            =   -58515
            TabIndex        =   190
            Top             =   1500
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   189
            Top             =   1500
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   28
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   188
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   28
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   187
            Top             =   1500
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
            Index           =   28
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   186
            Top             =   1500
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   185
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   184
            Top             =   1785
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   183
            Top             =   1785
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   182
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   181
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   29
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   179
            Top             =   1785
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   178
            Top             =   1785
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   177
            Top             =   1785
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   29
            ItemData        =   "f_MAIN.frx":050A
            Left            =   -61395
            List            =   "f_MAIN.frx":050C
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   1785
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   175
            Top             =   1785
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
            Index           =   29
            Left            =   -58515
            TabIndex        =   174
            Top             =   1785
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   173
            Top             =   1785
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   29
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   172
            Top             =   1785
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   29
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   171
            Top             =   1785
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
            Index           =   29
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   170
            Top             =   1785
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   169
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   168
            Top             =   2070
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   167
            Top             =   2070
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   166
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   165
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   30
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   163
            Top             =   2070
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   162
            Top             =   2070
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   161
            Top             =   2070
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   30
            ItemData        =   "f_MAIN.frx":050E
            Left            =   -61395
            List            =   "f_MAIN.frx":0510
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   2070
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   159
            Top             =   2070
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
            Index           =   30
            Left            =   -58515
            TabIndex        =   158
            Top             =   2070
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   157
            Top             =   2070
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   30
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   156
            Top             =   2070
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   30
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   155
            Top             =   2070
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
            Index           =   30
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   154
            Top             =   2070
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   153
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   152
            Top             =   2355
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   151
            Top             =   2355
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   150
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   149
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   31
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   147
            Top             =   2355
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   146
            Top             =   2355
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   145
            Top             =   2355
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   31
            ItemData        =   "f_MAIN.frx":0512
            Left            =   -61395
            List            =   "f_MAIN.frx":0514
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   2355
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   143
            Top             =   2355
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
            Index           =   31
            Left            =   -58515
            TabIndex        =   142
            Top             =   2355
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   141
            Top             =   2355
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   31
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   140
            Top             =   2355
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   31
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   139
            Top             =   2355
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
            Index           =   31
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   138
            Top             =   2355
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   137
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   136
            Top             =   2640
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   135
            Top             =   2640
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   134
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   133
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   32
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   131
            Top             =   2640
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   130
            Top             =   2640
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   129
            Top             =   2640
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   32
            ItemData        =   "f_MAIN.frx":0516
            Left            =   -61395
            List            =   "f_MAIN.frx":0518
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   2640
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   127
            Top             =   2640
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
            Index           =   32
            Left            =   -58515
            TabIndex        =   126
            Top             =   2640
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   125
            Top             =   2640
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   32
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   124
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   32
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   123
            Top             =   2640
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
            Index           =   32
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   122
            Top             =   2640
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   121
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   120
            Top             =   2925
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   119
            Top             =   2925
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   118
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   117
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   33
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   115
            Top             =   2925
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   114
            Top             =   2925
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   113
            Top             =   2925
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   33
            ItemData        =   "f_MAIN.frx":051A
            Left            =   -61395
            List            =   "f_MAIN.frx":051C
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   2925
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   111
            Top             =   2925
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
            Index           =   33
            Left            =   -58515
            TabIndex        =   110
            Top             =   2925
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   109
            Top             =   2925
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   33
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   108
            Top             =   2925
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   33
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   107
            Top             =   2925
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
            Index           =   33
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   106
            Top             =   2925
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   105
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   104
            Top             =   3210
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   103
            Top             =   3210
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   102
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   101
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   34
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   99
            Top             =   3210
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   98
            Top             =   3210
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   97
            Top             =   3210
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   34
            ItemData        =   "f_MAIN.frx":051E
            Left            =   -61395
            List            =   "f_MAIN.frx":0520
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   3210
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   95
            Top             =   3210
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
            Index           =   34
            Left            =   -58515
            TabIndex        =   94
            Top             =   3210
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   93
            Top             =   3210
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   34
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   92
            Top             =   3210
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   34
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   91
            Top             =   3210
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
            Index           =   34
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   90
            Top             =   3210
            Width           =   525
         End
         Begin VB.TextBox c_fabricante 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -74880
            MaxLength       =   4
            TabIndex        =   89
            Top             =   3495
            Width           =   525
         End
         Begin VB.TextBox c_produto 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -74355
            MaxLength       =   8
            TabIndex        =   88
            Top             =   3495
            Width           =   885
         End
         Begin VB.TextBox c_descricao 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -73470
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   87
            Top             =   3495
            Width           =   5320
         End
         Begin VB.TextBox c_qtde 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -66450
            MaxLength       =   5
            TabIndex        =   86
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_vl_unitario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -65835
            MaxLength       =   18
            TabIndex        =   85
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_vl_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -64530
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_produto_obs 
            Height          =   285
            Index           =   35
            Left            =   -68145
            MaxLength       =   500
            TabIndex        =   83
            Top             =   3495
            Width           =   1705
         End
         Begin VB.TextBox c_vl_outras_despesas_acessorias 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -63225
            MaxLength       =   18
            TabIndex        =   82
            Top             =   3495
            Width           =   1305
         End
         Begin VB.TextBox c_CST 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -61920
            MaxLength       =   4
            TabIndex        =   81
            Top             =   3495
            Width           =   525
         End
         Begin VB.ComboBox cb_CFOP 
            Height          =   315
            Index           =   35
            ItemData        =   "f_MAIN.frx":0522
            Left            =   -61395
            List            =   "f_MAIN.frx":0524
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   3495
            Width           =   1985
         End
         Begin VB.TextBox c_NCM 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -59400
            MaxLength       =   8
            TabIndex        =   79
            Top             =   3495
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
            Index           =   35
            Left            =   -58515
            TabIndex        =   78
            Top             =   3495
            Width           =   585
         End
         Begin VB.TextBox c_xPed 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   -57930
            MaxLength       =   15
            TabIndex        =   77
            Top             =   3495
            Width           =   885
         End
         Begin VB.TextBox c_unidade 
            Height          =   285
            Index           =   35
            Left            =   -56205
            MaxLength       =   6
            TabIndex        =   76
            Top             =   3495
            Width           =   615
         End
         Begin VB.TextBox c_nItemPed 
            Height          =   285
            Index           =   35
            Left            =   -57000
            MaxLength       =   6
            TabIndex        =   75
            Top             =   3495
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
            Index           =   35
            Left            =   -55530
            MaxLength       =   6
            TabIndex        =   74
            Top             =   3495
            Width           =   525
         End
      End
      Begin VB.Label l_tit_vl_total_icms 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   676
         Top             =   4410
         Width           =   960
      End
      Begin VB.Label l_tit_FCP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "%FCP"
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
         Left            =   19530
         TabIndex        =   675
         Top             =   135
         Width           =   495
      End
      Begin VB.Label l_tit_nItemPed 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "nItemPed"
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
         Left            =   17945
         TabIndex        =   674
         Top             =   135
         Width           =   810
      End
      Begin VB.Label l_tit_unidade 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unid"
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
         Left            =   18955
         TabIndex        =   673
         Top             =   135
         Width           =   405
      End
      Begin VB.Label l_tit_xPed 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "xPed"
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
         Left            =   17245
         TabIndex        =   672
         Top             =   135
         Width           =   435
      End
      Begin VB.Label l_tit_ICMS_item 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "ICMS"
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
         Left            =   16560
         TabIndex        =   671
         Top             =   135
         Width           =   465
      End
      Begin VB.Label l_tit_total_peso_liquido 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   6915
         TabIndex        =   670
         Top             =   4410
         Width           =   1170
      End
      Begin VB.Label l_tit_total_peso_bruto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   3825
         TabIndex        =   669
         Top             =   4410
         Width           =   1335
      End
      Begin VB.Label l_tit_NCM 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "NCM"
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
         Left            =   15675
         TabIndex        =   668
         Top             =   135
         Width           =   420
      End
      Begin VB.Label l_tit_CFOP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "CFOP"
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
         Left            =   14080
         TabIndex        =   667
         Top             =   135
         Width           =   495
      End
      Begin VB.Label l_tit_CST 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "CST"
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
         Left            =   13245
         TabIndex        =   666
         Top             =   135
         Width           =   375
      End
      Begin VB.Label l_tit_vl_outras_despesas_acessorias 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dsp Acessórias"
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
         Left            =   11805
         TabIndex        =   665
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label l_tit_total_volumes 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   664
         Top             =   4410
         Width           =   720
      End
      Begin VB.Label l_tit_produto_obs 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inf Adicionais"
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
         Left            =   7200
         TabIndex        =   663
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label l_tit_vl_total_geral 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   9975
         TabIndex        =   662
         Top             =   4410
         Width           =   450
      End
      Begin VB.Label l_tit_vl_total 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Valor Total"
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
         Left            =   10705
         TabIndex        =   661
         Top             =   135
         Width           =   945
      End
      Begin VB.Label l_tit_vl_unitario 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Valor Unitário"
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
         Left            =   9370
         TabIndex        =   660
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label l_tit_qtde 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qtde"
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
         Left            =   8745
         TabIndex        =   659
         Top             =   135
         Width           =   420
      End
      Begin VB.Label l_tit_descricao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descrição"
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
         Left            =   3000
         TabIndex        =   658
         Top             =   135
         Width           =   870
      End
      Begin VB.Label l_tit_produto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Produto"
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
         Left            =   800
         TabIndex        =   657
         Top             =   135
         Width           =   675
      End
      Begin VB.Label l_tit_fabricante 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fabric"
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
         Left            =   150
         TabIndex        =   656
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.CommandButton b_emissao_nfe_triangular 
      Caption         =   "NFe &Triangular"
      Enabled         =   0   'False
      Height          =   450
      Left            =   15555
      TabIndex        =   64
      Top             =   7980
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame pnParcelasEmBoletos 
      Caption         =   "Parcelas em Boletos"
      Height          =   4575
      Left            =   5400
      TabIndex        =   53
      Top             =   6960
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chk_InfoAdicParc 
         Caption         =   "Incluir parcelas no campo de  Informações Adicionais"
         Height          =   360
         Left            =   120
         TabIndex        =   71
         Top             =   4080
         Width           =   5175
      End
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
         TabIndex        =   56
         Top             =   3120
         Width           =   945
      End
      Begin VB.CommandButton b_parc_edicao_ok 
         Height          =   390
         Left            =   360
         Picture         =   "f_MAIN.frx":0526
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3650
         Width           =   690
      End
      Begin VB.CommandButton b_parc_edicao_cancela 
         Height          =   390
         Left            =   1560
         Picture         =   "f_MAIN.frx":0778
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3650
         Width           =   690
      End
      Begin VB.CommandButton b_recalculaparc 
         Caption         =   "&Reagendar Parcelas Seguintes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         TabIndex        =   61
         Top             =   3600
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
         TabIndex        =   59
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
         TabIndex        =   57
         Top             =   3120
         Width           =   1260
      End
      Begin MSComctlLib.ListView lvParcBoletos 
         Height          =   2415
         Left            =   120
         TabIndex        =   54
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
         TabIndex        =   60
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label l_tit_dataparc 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   1560
         TabIndex        =   58
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label l_tit_numparc 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   360
         TabIndex        =   55
         Top             =   2880
         Width           =   540
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
      Left            =   11175
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "f_MAIN.frx":0BEB
      Top             =   1560
      Width           =   9015
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
      ItemData        =   "f_MAIN.frx":0C18
      Left            =   300
      List            =   "f_MAIN.frx":0C1A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   8301
   End
   Begin VB.CheckBox chk_HorVerao 
      Caption         =   "Horário de Verão"
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
      Left            =   11175
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame pnZerarAliquotas 
      Height          =   1365
      Left            =   14400
      TabIndex        =   45
      Top             =   0
      Width           =   5895
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
         Left            =   690
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
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
         Left            =   690
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   270
         Width           =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Zerar"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   840
         Width           =   375
      End
      Begin VB.Label l_zerar_1 
         AutoSize        =   -1  'True
         Caption         =   "Zerar"
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   240
         Width           =   375
      End
      Begin VB.Label l_tit_zerar_COFINS 
         AutoSize        =   -1  'True
         Caption         =   "COFINS"
         Height          =   195
         Left            =   0
         TabIndex        =   47
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label l_tit_zerar_PIS 
         AutoSize        =   -1  'True
         Caption         =   " PIS"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.Frame pnInfoFilaPedido 
      Caption         =   "Fila de Solicitações de Emissão de NFe"
      Height          =   570
      Left            =   240
      TabIndex        =   44
      Top             =   10815
      Width           =   4980
      Begin VB.CommandButton b_fila_remove 
         Height          =   390
         Left            =   3360
         Picture         =   "f_MAIN.frx":0C1C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   150
         Width           =   465
      End
      Begin VB.CommandButton b_fila_pause 
         Height          =   390
         Left            =   3855
         Picture         =   "f_MAIN.frx":10D9
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   150
         Width           =   465
      End
      Begin VB.CommandButton b_fila_play 
         Height          =   390
         Left            =   4350
         Picture         =   "f_MAIN.frx":13A5
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   150
         Width           =   465
      End
      Begin VB.Image imgFilasEmits 
         Appearance      =   0  'Flat
         Height          =   135
         Left            =   2520
         Picture         =   "f_MAIN.frx":167C
         Stretch         =   -1  'True
         ToolTipText     =   "Existem pendências de outros CDs!"
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblQtdeFilaSolicitacoesEmissaoNFe 
         AutoSize        =   -1  'True
         Caption         =   "00 solicitações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   270
         Width           =   1290
      End
   End
   Begin VB.Frame pnPedidoInfo 
      Caption         =   "Informações do Pedido"
      Height          =   2145
      Left            =   11160
      TabIndex        =   43
      Top             =   9240
      Width           =   9210
      Begin VB.TextBox c_info_pedido 
         Height          =   1740
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   240
         Width           =   8940
      End
   End
   Begin VB.CommandButton b_emissao_nfe_complementar 
      Caption         =   "NFe Com&plementar"
      Enabled         =   0   'False
      Height          =   450
      Left            =   15555
      TabIndex        =   18
      Top             =   8715
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton b_emite_numeracao_manual 
      Caption         =   "Emitir NFe (Nº &Manual)"
      Height          =   450
      Left            =   13320
      TabIndex        =   15
      Top             =   7980
      Width           =   2115
   End
   Begin VB.CommandButton b_emissao_manual 
      Caption         =   "Painel Emissão M&anual"
      Height          =   450
      Left            =   15555
      TabIndex        =   17
      Top             =   7260
      Width           =   2115
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
      Left            =   2235
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   2340
   End
   Begin VB.Frame pnDanfe 
      Caption         =   "DANFE"
      Height          =   2010
      Left            =   17880
      TabIndex        =   40
      Top             =   7200
      Width           =   2430
      Begin VB.CommandButton b_danfe 
         Caption         =   "D&ANFE"
         Height          =   390
         Left            =   390
         TabIndex        =   20
         Top             =   1470
         Width           =   1650
      End
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
         Left            =   390
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   525
         Width           =   1650
      End
      Begin VB.Label l_tit_pedido_Danfe 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido(s)"
         Height          =   195
         Left            =   390
         TabIndex        =   41
         Top             =   315
         Width           =   885
      End
   End
   Begin VB.CommandButton b_fechar 
      Caption         =   "&Fechar"
      Height          =   450
      Left            =   13320
      TabIndex        =   16
      Top             =   8715
      Width           =   2115
   End
   Begin VB.Timer relogio 
      Interval        =   1000
      Left            =   12480
      Top             =   7200
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
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   7260
      Width           =   4995
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
      Left            =   9795
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   300
      Width           =   3780
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
      Left            =   5955
      MaxLength       =   6
      TabIndex        =   3
      Top             =   300
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
      Left            =   4740
      TabIndex        =   2
      Top             =   300
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
      ItemData        =   "f_MAIN.frx":8C3E
      Left            =   300
      List            =   "f_MAIN.frx":8C40
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   930
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
      Left            =   300
      MaxLength       =   9
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   1650
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
      Left            =   13320
      TabIndex        =   14
      Top             =   7260
      Width           =   2115
   End
   Begin VB.CommandButton b_dummy 
      Appearance      =   0  'Flat
      Caption         =   "b_dummy"
      Height          =   345
      Left            =   5565
      TabIndex        =   29
      Top             =   -525
      Width           =   1350
   End
   Begin VB.Frame pnNumeroNFe 
      Caption         =   "Última NFe emitida"
      Height          =   1485
      Left            =   240
      TabIndex        =   36
      Top             =   9240
      Width           =   4980
      Begin VB.Label l_tit_emitente_NF 
         AutoSize        =   -1  'True
         Caption         =   "Emitente"
         Height          =   195
         Left            =   195
         TabIndex        =   39
         Top             =   825
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
         Left            =   180
         TabIndex        =   23
         Top             =   1035
         Width           =   4500
      End
      Begin VB.Label l_tit_serie_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série"
         Height          =   195
         Left            =   195
         TabIndex        =   38
         Top             =   240
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
         Left            =   180
         TabIndex        =   21
         Top             =   450
         Width           =   1710
      End
      Begin VB.Label l_tit_num_NF 
         AutoSize        =   -1  'True
         Caption         =   "Nº NFe"
         Height          =   195
         Left            =   2985
         TabIndex        =   37
         Top             =   225
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
         Left            =   2970
         TabIndex        =   22
         Top             =   435
         Width           =   1710
      End
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
      Left            =   7155
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   2460
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
      Left            =   13680
      TabIndex        =   68
      Top             =   300
      Width           =   585
   End
   Begin VB.Label l_tit_IE 
      AutoSize        =   -1  'True
      Caption         =   "IE"
      Height          =   195
      Left            =   13920
      TabIndex        =   67
      Top             =   90
      Width           =   150
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
      Left            =   9480
      TabIndex        =   66
      Top             =   1270
      Width           =   825
   End
   Begin VB.Label l_tit_emitente_uf 
      AutoSize        =   -1  'True
      Caption         =   "UF do Emitente"
      Height          =   195
      Left            =   9375
      TabIndex        =   65
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label l_tit_chave_nfe_ref 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso NFe Referenciada"
      Height          =   195
      Left            =   11190
      TabIndex        =   52
      Top             =   1350
      Width           =   2610
   End
   Begin VB.Label l_tit_finalidade 
      AutoSize        =   -1  'True
      Caption         =   "Finalidade"
      Height          =   195
      Left            =   315
      TabIndex        =   51
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label l_tit_loc_dest 
      AutoSize        =   -1  'True
      Caption         =   "Local de Destino da Operação"
      Height          =   195
      Left            =   7170
      TabIndex        =   50
      Top             =   90
      Width           =   2175
   End
   Begin VB.Label l_tit_tipo_NF 
      AutoSize        =   -1  'True
      Caption         =   "Tipo do Documento Fiscal"
      Height          =   195
      Left            =   2250
      TabIndex        =   42
      Top             =   90
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
      Left            =   11190
      TabIndex        =   13
      Top             =   8865
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
      Left            =   11190
      TabIndex        =   12
      Top             =   8460
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
      Left            =   11190
      TabIndex        =   11
      Top             =   7260
      Width           =   1980
      WordWrap        =   -1  'True
   End
   Begin VB.Label l_tit_dados_adicionais 
      AutoSize        =   -1  'True
      Caption         =   "Dados Adicionais"
      Height          =   195
      Left            =   255
      TabIndex        =   35
      Top             =   7050
      Width           =   1230
   End
   Begin VB.Label l_tit_frete 
      AutoSize        =   -1  'True
      Caption         =   "Frete por Conta"
      Height          =   195
      Left            =   9810
      TabIndex        =   34
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label l_tit_aliquota_IPI 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota IPI"
      Height          =   195
      Left            =   5970
      TabIndex        =   33
      Top             =   90
      Width           =   840
   End
   Begin VB.Label l_tit_aliquota_icms 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota ICMS"
      Height          =   195
      Left            =   4755
      TabIndex        =   32
      Top             =   90
      Width           =   1035
   End
   Begin VB.Label l_tit_natureza 
      AutoSize        =   -1  'True
      Caption         =   "Natureza da Operação"
      Height          =   195
      Left            =   315
      TabIndex        =   31
      Top             =   720
      Width           =   1620
   End
   Begin VB.Label l_tit_pedido 
      AutoSize        =   -1  'True
      Caption         =   "Nº Pedido"
      Height          =   195
      Left            =   315
      TabIndex        =   30
      Top             =   90
      Width           =   720
   End
   Begin VB.Menu mnu_ARQUIVO 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_emissao_manual 
         Caption         =   "&Modo de Emissão Manual"
      End
      Begin VB.Menu mnu_emissao_nfe_complementar 
         Caption         =   "Modo de Emissão NFe &Complementar"
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
Attribute VB_Name = "f_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modulo_inicializacao_ok As Boolean
Dim pedido_anterior As String
Dim dt_hr_ult_atualizacao_qtde_fila_solicitacoes_emissao_NFe As Date
Dim blnFilaSolicitacoesEmissaoNFeEmTratamento As Boolean
Dim inumparcela As Integer
Dim v_pedido_manual_boleto() As String
Dim v_parcela_manual_boleto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
Dim blnExisteParcelamentoBoleto As Boolean

Private Const FONTNAME_IMPRESSAO = "Tahoma"
Private Const FONTSIZE_IMPRESSAO = 8
Private Const FONTBOLD_IMPRESSAO = True
Private Const FONTITALIC_IMPRESSAO = False
Private Const FORMATO_PERCENTUAL = "##0.00"

Private Sub b_emissao_nfe_triangular_Click()
    
    If blnNotaTriangularAtiva Then
        sPedidoTriangular = ""
        sPedidoDANFETelaAnterior = ""
        sNFAnteriorSerie = ""
        sNFAnteriorNumero = ""
        sNFAnteriorEmitente = ""
        exibe_form_emissao_nfe_triangular
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


Private Sub c_dataparc_LostFocus()

    c_dataparc = Trim$(c_dataparc)
    If c_dataparc = "" Then Exit Sub
    
    data_ok c_dataparc
    
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



Sub atualiza_tela_qtde_fila_solicitacoes_emissao_NFe()
Dim s As String
Dim s_erro As String
Dim iTotal As Integer
Dim iEmit As Integer

    lblQtdeFilaSolicitacoesEmissaoNFe = ""
    imgFilasEmits.Visible = False
    If obtem_qtde_fila_solicitacoes_emissao_NFe(c_pedido, iTotal, iEmit, s_erro) Then
        If iEmit = 0 Then
            s = "Nenhuma solicitação"
        ElseIf iEmit = 1 Then
            s = CStr(iEmit) & " solicitação"
        Else
            s = CStr(iEmit) & " solicitações"
            End If
        lblQtdeFilaSolicitacoesEmissaoNFe = s
        If iTotal > iEmit Then imgFilasEmits.Visible = True
    ElseIf s_erro <> "" Then
        aviso_erro s_erro
        End If
        
    dt_hr_ult_atualizacao_qtde_fila_solicitacoes_emissao_NFe = Now

End Sub


Private Sub exibe_form_emissao_manual()

    Hide
    f_EMISSAO_MANUAL.Show vbModal, Me
    Me.Visible = True
    
End Sub



Private Sub exibe_form_emissao_nfe_complementar()

    Hide
    f_EMISSAO_NFE_COMPLEMENTAR.Show vbModal, Me
    Me.Visible = True
    
End Sub

Private Sub exibe_form_emissao_nfe_triangular()

    Hide
    f_EMISSAO_NFE_TRIANGULAR.Show vbModal, Me
    Me.Visible = True
    
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

Sub formulario_exibe_itens_pedido(ByVal pedido_selecionado As String)
Const NomeDestaRotina = "formulario_exibe_itens_pedido()"
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
Dim s_NFe_xPed As String


    On Error GoTo FEIP_TRATA_ERRO
    
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
            " pedido," & _
            " NFe_xPed" & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & pedido_selecionado & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        aviso_erro "O pedido " & pedido_selecionado & " NÃO está cadastrado!!"
        GoSub FEIP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        c_pedido.SetFocus
        Exit Sub
        End If
    
    s_NFe_xPed = Trim$("" & t_PEDIDO("NFe_xPed"))
    
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
            GoSub FEIP_FECHA_TABELAS
            aguarde INFO_NORMAL, m_id
            Exit Sub
            End If
            
        c_fabricante(intIndice) = Trim$("" & t_PEDIDO_ITEM("fabricante"))
        c_produto(intIndice) = Trim$("" & t_PEDIDO_ITEM("produto"))
        c_descricao(intIndice) = Trim$("" & t_PEDIDO_ITEM("descricao"))
        
        c_CST(intIndice) = cst_converte_codigo_entrada_para_saida(Trim$("" & t_PEDIDO_ITEM("cst")))
        c_NCM(intIndice) = Trim$("" & t_PEDIDO_ITEM("ncm"))
        
        If s_NFe_xPed <> "" Then
            c_xPed(intIndice) = s_NFe_xPed
            End If
            
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
        
    atualiza_valor_total_icms
    
    
    'posicionar na tab onde está o último item
    If intIndice > 24 Then
        sstItens.Tab = 2
    ElseIf intIndice > 12 Then
        sstItens.Tab = 1
    Else
        sstItens.Tab = 0
        End If
    
    GoSub FEIP_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FEIP_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub FEIP_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FEIP_FECHA_TABELAS:
'==================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    Return
    
End Sub

Sub formulario_limpa()

Dim s As String
Dim s_aux As String
Dim i As Integer
Dim vAliquotas() As String

'   Nº PEDIDO
'   ~~~~~~~~~
    c_pedido = ""
    pedido_anterior = ""
    
'   ITENS
'   ~~~~~
    formulario_limpa_campos_itens_pedido
        
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
    'No caso de PBJ, fixar 5.102, senão 6.108
    If (usuario.emit_id = "4006") Then
        s = "5.102"
    Else
        s = "6.108"
        End If
    For i = 0 To cb_natureza.ListCount - 1
        If left$(cb_natureza.List(i), Len(s)) = s Then
            cb_natureza.ListIndex = i
            Exit For
            End If
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
    
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    c_dados_adicionais = ""
    
'   INFORMAÇÕES DO PEDIDO
'   ~~~~~~~~~~~~~~~~~~~~~
    c_info_pedido = ""
    
'   PARCELAS EM BOLETOS
'   ~~~~~~~~~~~~~~~~~~~
    pnParcelasEmBoletos.Visible = False
    
    
'   INFO CONTRIBUINTE
'   ~~~~~~~~~~~~~~~~~
    l_IE.Caption = ""
    
'   FOCO INICIAL
'   ~~~~~~~~~~~~
    c_pedido.SetFocus
    
End Sub


Sub DANFE_consulta(ByVal relacaoPedidos As String)

' CONSTANTES
Const NomeDestaRotina = "DANFE_consulta()"
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
Dim intIdBoletoCedente As Integer
Dim lngNFeSerieNF As Long
Dim lngNFeNumeroNF As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

' VETORES
Dim v() As String
Dim v_pedido() As String
Dim v_danfe() As String

' BANCO DE DADOS
Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim dbcNFe As ADODB.Connection

    On Error GoTo DANFE_CONSULTA_TRATA_ERRO
    
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
    
  ' T_FIN_BOLETO_CEDENTE
    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
    With t_FIN_BOLETO_CEDENTE
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
    
    
'   PARA CADA PEDIDO DA LISTA, OBTÉM E EXIBE A DANFE
    s_alerta_erro = ""
    For ic = LBound(v_pedido) To UBound(v_pedido)
        strPedido = Trim$(v_pedido(ic))
        If strPedido <> "" Then
            aguarde INFO_EXECUTANDO, "consultando situação da NFe"
            
            s = "SELECT" & _
                    " id_boleto_cedente," & _
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
                
            intIdBoletoCedente = t_NFe_EMISSAO("id_boleto_cedente")
            
            s = "SELECT" & _
                    " nome_empresa," & _
                    " NFe_T1_servidor_BD," & _
                    " NFe_T1_nome_BD," & _
                    " NFe_T1_usuario_BD," & _
                    " NFe_T1_senha_BD" & _
                " FROM t_FIN_BOLETO_CEDENTE" & _
                " WHERE" & _
                    " (id = " & CStr(intIdBoletoCedente) & ")"
            If t_FIN_BOLETO_CEDENTE.State <> adStateClosed Then t_FIN_BOLETO_CEDENTE.Close
            t_FIN_BOLETO_CEDENTE.Open s, dbc, , , adCmdText
            If t_FIN_BOLETO_CEDENTE.EOF Then
                If s_alerta_erro <> "" Then s_alerta_erro = s_alerta_erro & vbCrLf
                s_alerta_erro = s_alerta_erro & "Pedido " & strPedido & ": falha ao localizar o registro em t_FIN_BOLETO_CEDENTE (id=" & CStr(intIdBoletoCedente) & ")!!"
                GoTo PROXIMO_PEDIDO
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
        
        
    GoSub DANFE_CONSULTA_FECHA_TABELAS
    
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
DANFE_CONSULTA_FECHA_TABELAS:
'============================
  ' RECORDSETS
    bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
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

Sub DANFE_CONSULTA_parametro_emitente_original(ByVal relacaoPedidos As String)
'OBS: ESTA ROTINA TEVE O NOME ALTERADO MAS PERMANECE NO SISTEMA PARA EVENTUAL REUTILIZAÇÃO
'A NOVA ROTINA (ABAIXO DESTA) PREVÊ A POSSIBILIDADE DE CONSULTAR DANFE'S GERADAS EM OPERAÇÕES TRIANGULARES

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

' VETORES
Dim v() As String
Dim v_pedido() As String
Dim v_danfe() As String

' BANCO DE DADOS
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
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
Dim strPastaEmitente As String
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
                        " apelido," & _
                        " cnpj," & _
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
                        strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
                        
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
                        strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
                        
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
                        " cnpj," & _
                        " apelido," & _
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
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
                    
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

Sub formulario_limpa_campos_itens_pedido()
Dim i As Integer
    
    c_vl_total_outras_despesas_acessorias = ""
    c_vl_total_geral = ""
    c_vl_total_icms = ""
    c_total_volumes = ""
    For i = c_fabricante.LBound To c_fabricante.UBound
        c_fcp(i) = ""
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
        
    sstItens.Tab = 0
        
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

Function obtem_info_pedido(ByVal pedido As String, ByRef strResposta As String, _
                            ByRef strEndEntregaFormatado As String, _
                            ByRef strEndEntregaUf As String, _
                            ByRef strEndClienteUf As String, _
                            ByRef strNFeTextoConstar As String, _
                            ByRef strInfoIE As String, _
                            ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_pedido()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_endereco As String
Dim s_bairro As String
Dim s_cep As String
Dim s_cidade As String
Dim s_uf As String
Dim s_nome As String
Dim s_cnpj_cpf As String
Dim s_ie_rg As String
Dim s_obs_1 As String
Dim s_info As String
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

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo OIP_TRATA_ERRO
    
    obtem_info_pedido = False
    strMsgErro = ""
    strResposta = ""
    strEndEntregaFormatado = ""
    strEndEntregaUf = ""
    strEndClienteUf = ""
    strInfoIE = ""
    
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
        
    s_endereco = ""
    s_bairro = ""
    s_cep = ""
    s_cidade = ""
    s_uf = ""
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
            " pedido, st_entrega, id_cliente, obs_1, st_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep, NFe_texto_Constar" & _
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
        
    '   ENDEREÇO DE ENTREGA
        If (s_end_entrega = "") And (CLng(t_PEDIDO("st_end_entrega")) <> 0) Then
            s_end_entrega = formata_endereco(Trim("" & t_PEDIDO("EndEtg_endereco")), Trim("" & t_PEDIDO("EndEtg_endereco_numero")), Trim("" & t_PEDIDO("EndEtg_endereco_complemento")), Trim("" & t_PEDIDO("EndEtg_bairro")), Trim("" & t_PEDIDO("EndEtg_cidade")), Trim("" & t_PEDIDO("EndEtg_uf")), Trim("" & t_PEDIDO("EndEtg_cep")))
            s_end_entrega = UCase$(s_end_entrega)
            strEndEntregaFormatado = s_end_entrega
            strEndEntregaUf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
            If s_end_entrega <> "" Then s_end_entrega = vbCrLf & "ENTREGA: " & s_end_entrega
            End If
    
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        s_id_cliente = Trim$("" & t_PEDIDO("id_cliente"))
        
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
        GoSub OIP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        

'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (id='" & s_id_cliente & "')"
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        strMsgErro = "Cliente com nº registro " & s_id_cliente & " não foi encontrado!!"
        GoSub OIP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
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
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   CNPJ/CPF
    s_cnpj_cpf = Trim$("" & t_DESTINATARIO("cnpj_cpf"))


'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    
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
    If (s_end_linha_1 <> "") And (s_bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & s_bairro
    
    s_end_linha_2 = s_cidade
    If (s_end_linha_2 <> "") And (s_uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & s_uf
    If (s_end_linha_2 <> "") And (s_cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & cep_formata(s_cep)
        
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
    
    GoSub OIP_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

    strResposta = s_info
    obtem_info_pedido = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIP_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIP_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIP_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_DESTINATARIO, True
    Return
    
End Function

Function obtem_info_pedido_memorizada(ByVal pedido As String, ByRef strResposta As String, _
                                    ByRef strEndEntregaFormatado As String, _
                                    ByRef strEndEntregaUf As String, _
                                    ByRef strEndClienteUf As String, _
                                    ByRef strNFeTextoConstar As String, _
                                    ByRef strInfoIE As String, _
                                    ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_pedido_memorizada()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_endereco As String
Dim s_bairro As String
Dim s_cep As String
Dim s_cidade As String
Dim s_uf As String
Dim s_nome As String
Dim s_cnpj_cpf As String
Dim s_ie_rg As String
Dim s_obs_1 As String
Dim s_info As String
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

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo OIPM_TRATA_ERRO
    
    obtem_info_pedido_memorizada = False
    strMsgErro = ""
    strResposta = ""
    strEndEntregaFormatado = ""
    strEndEntregaUf = ""
    strEndClienteUf = ""
    strInfoIE = ""
    
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
        
    s_endereco = ""
    s_bairro = ""
    s_cep = ""
    s_cidade = ""
    s_uf = ""
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
            " pedido, st_entrega, id_cliente, obs_1, st_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep, NFe_texto_Constar" & _
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
        
    '   ENDEREÇO DE ENTREGA
        If (s_end_entrega = "") And (CLng(t_PEDIDO("st_end_entrega")) <> 0) Then
            s_end_entrega = formata_endereco(Trim("" & t_PEDIDO("EndEtg_endereco")), Trim("" & t_PEDIDO("EndEtg_endereco_numero")), Trim("" & t_PEDIDO("EndEtg_endereco_complemento")), Trim("" & t_PEDIDO("EndEtg_bairro")), Trim("" & t_PEDIDO("EndEtg_cidade")), Trim("" & t_PEDIDO("EndEtg_uf")), Trim("" & t_PEDIDO("EndEtg_cep")))
            s_end_entrega = UCase$(s_end_entrega)
            strEndEntregaFormatado = s_end_entrega
            strEndEntregaUf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
            If s_end_entrega <> "" Then s_end_entrega = vbCrLf & "ENTREGA: " & s_end_entrega
            End If
            
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        s_id_cliente = Trim$("" & t_PEDIDO("id_cliente"))
        
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
        GoSub OIPM_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        

'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
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
        " WHERE (pedido = '" & Trim$(pedido) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PJ & "')"
    If param_nfmemooendentrega.campo_inteiro = 1 Then
        s = s & " UNION" & _
            " SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_uf else EndEtg_uf end as uf, " & _
                " endereco_cnpj_cpf as cnpj_cpf, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_logradouro else EndEtg_endereco end as endereco, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_bairro else EndEtg_bairro end as bairro, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cidade else EndEtg_cidade end as cidade, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_cep else EndEtg_cep end as cep, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_numero else EndEtg_endereco_numero end as endereco_numero, " & _
                " case when ltrim(rtrim(EndEtg_endereco)) = '' or isnull(EndEtg_endereco, '') = '' then endereco_complemento else EndEtg_endereco_complemento end as endereco_complemento, " & _
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
            " WHERE (pedido = '" & Trim$(pedido) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PF & "')"
    Else
        s = s & " UNION" & _
            " SELECT" & _
                " pedido, id_cliente, st_memorizacao_completa_enderecos, " & _
                " endereco_uf as uf, " & _
                " endereco_cnpj_cpf as cnpj_cpf, " & _
                " endereco_logradouro as endereco, " & _
                " endereco_bairro as bairro, " & _
                " endereco_cidade as cidade, " & _
                " endereco_cep as cep, " & _
                " endereco_numero as endereco_numero, " & _
                " endereco_complemento as endereco_complemento, " & _
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
            " WHERE (pedido = '" & Trim$(pedido) & "')" & " AND (endereco_tipo_pessoa = '" & ID_PF & "')"
        End If
    t_DESTINATARIO.Open s, dbc, , , adCmdText
    If t_DESTINATARIO.EOF Then
        strMsgErro = "Problemas na localização do endereço memorizado no pedido " & Trim$(pedido) & "!!"
        GoSub OIPM_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If


    '   (SÓ INFORMAR O ENDEREÇO DE ENTREGA SE FOR DIFERENTE DO ENDEREÇO DA NOTA)
    If (UCase$(Trim$("" & t_DESTINATARIO("endereco"))) = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco")))) And _
        (UCase$(Trim$("" & t_DESTINATARIO("endereco_numero"))) = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_numero")))) And _
        (UCase$(Trim$("" & t_DESTINATARIO("endereco_complemento"))) = UCase$(Trim("" & t_PEDIDO("EndEtg_endereco_complemento")))) And _
        (UCase$(Trim$("" & t_DESTINATARIO("bairro"))) = UCase$(Trim("" & t_PEDIDO("EndEtg_bairro")))) And _
        (UCase$(Trim$("" & t_DESTINATARIO("cidade"))) = UCase$(Trim("" & t_PEDIDO("EndEtg_cidade")))) Then
        
        s_end_entrega = ""
        
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
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   CNPJ/CPF
    s_cnpj_cpf = Trim$("" & t_DESTINATARIO("cnpj_cpf"))


'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    
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
    If (s_end_linha_1 <> "") And (s_bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & s_bairro
    
    s_end_linha_2 = s_cidade
    If (s_end_linha_2 <> "") And (s_uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & s_uf
    If (s_end_linha_2 <> "") And (s_cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & cep_formata(s_cep)
        
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
    
    GoSub OIPM_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

    strResposta = s_info
    obtem_info_pedido_memorizada = True
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPM_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIPM_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPM_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    Return
    
End Function

Function obtem_qtde_fila_solicitacoes_emissao_NFe(ByVal pedido_em_tratamento As String, ByRef intQtdeFilasTodas As Integer, ByRef intQtdeFilaAtual As Integer, ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_qtde_fila_solicitacoes_emissao_NFe()"
' DECLARAÇÕES
Dim s As String
Dim s_campo_select_adicional As String
Dim s_join_adicional As String
' BANCO DE DADOS
Dim t As ADODB.Recordset

    On Error GoTo OQFSENFE_TRATA_ERRO
    
    obtem_qtde_fila_solicitacoes_emissao_NFe = False
    intQtdeFilasTodas = 0
    intQtdeFilaAtual = 0
    strMsgErro = ""
    
    pedido_em_tratamento = Trim$("" & pedido_em_tratamento)
    pedido_em_tratamento = normaliza_num_pedido(pedido_em_tratamento)

    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    s_campo_select_adicional = " ISNULL(SUM(CASE WHEN tP.id_nfe_emitente = " & usuario.emit_id & " THEN 1 ELSE 0 END), 0) AS qtde_emitente_atual"
    s_join_adicional = ""

    s = "SELECT" & _
            " COUNT(*) AS qtde," & _
            s_campo_select_adicional & _
        " FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA tPNES" & _
            " INNER JOIN t_PEDIDO tP ON (tP.pedido=tPNES.pedido)" & _
            s_join_adicional & _
        " WHERE" & _
            " (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
            " AND (Len(Coalesce(tP.transportadora_id,'')) > 0)" & _
            " AND (tP.st_entrega <> '" & Trim(CStr(ST_ENTREGA_CANCELADO)) & "')" & _
            " AND (" & _
                    "(ult_requisicao_fila_data_hora IS NULL)" & _
                    " OR " & _
                    "(DateDiff(ss, ult_requisicao_fila_data_hora, getdate()) >= " & MAX_TIMEOUT_REGISTRO_REQUISITADO_FILA_EM_SEG & ")" & _
                ")"

    
    If pedido_em_tratamento <> "" Then
        s = s & " AND (tPNES.pedido <> '" & pedido_em_tratamento & "')"
        End If
    
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        intQtdeFilasTodas = t("qtde")
        intQtdeFilaAtual = t("qtde_emitente_atual")
        End If
    
    GoSub OQFSENFE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
    obtem_qtde_fila_solicitacoes_emissao_NFe = True
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OQFSENFE_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OQFSENFE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    strMsgErro = s
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OQFSENFE_FECHA_TABELAS:
'======================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return

End Function

Function ObtemEnderecoEntrega(ByRef strEndereco As String, ByRef strEnderecoNumero As String, ByRef strEnderecoComplemento As String, ByRef strBairro As String, ByRef strCidade As String, ByRef strUF As String, ByRef strCep As String, ByRef strEnderecoCompletoFormatado As String, ByRef strMsgErro As String) As Boolean
Const NomeDestaRotina = "ObtemEnderecoEntrega()"
Dim i As Integer
Dim j As Integer
Dim qtde_pedidos As Integer
Dim qtde_clientes As Integer
Dim s As String
Dim s_aux As String
Dim s_pedido As String
Dim s_filtro_pedidos As String
Dim s_lista_pedidos As String
Dim v() As String
Dim v_pedido() As String
Dim t_PEDIDO As ADODB.Recordset

    On Error GoTo OEE_TRATA_ERRO

    ObtemEnderecoEntrega = False
    
    strEndereco = ""
    strEnderecoNumero = ""
    strEnderecoComplemento = ""
    strBairro = ""
    strCidade = ""
    strUF = ""
    strCep = ""
    strEnderecoCompletoFormatado = ""
    strMsgErro = ""
    
    s_filtro_pedidos = ""
    s_lista_pedidos = ""
    qtde_clientes = 0
    
    ReDim v_pedido(0)
    v_pedido(UBound(v_pedido)) = ""
    qtde_pedidos = 0
    v = Split(c_pedido, vbCrLf)
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
            s_aux = "|" & Trim$(v(i)) & "|"
            If InStr(s_lista_pedidos, s_aux) = 0 Then
                s_lista_pedidos = s_lista_pedidos & s_aux
                If v_pedido(UBound(v_pedido)) <> "" Then ReDim Preserve v_pedido(UBound(v_pedido) + 1)
                v_pedido(UBound(v_pedido)) = Trim$(v(i))
                s_pedido = Trim$(v(i))
                qtde_pedidos = qtde_pedidos + 1
                End If
            End If
        Next
    
    'Há algum pedido informado ?
    If qtde_pedidos = 0 Then
        strMsgErro = "Não é possível obter o endereço de entrega porque não foi informado nenhum pedido !!"
        Exit Function
        End If
            
    'T_PEDIDO
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            If s_filtro_pedidos <> "" Then s_filtro_pedidos = s_filtro_pedidos & ","
            s_filtro_pedidos = s_filtro_pedidos & "'" & Trim$(v_pedido(i)) & "'"
            End If
        Next
    
    'No caso de haver mais do que 1 pedido, verifica se todos são do mesmo cliente
    s = "SELECT Count(*) As qtde_clientes FROM (SELECT DISTINCT id_cliente FROM t_PEDIDO WHERE pedido IN (" & s_filtro_pedidos & ")) __t_AUX"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If Not t_PEDIDO.EOF Then
        If Not IsNull(t_PEDIDO("qtde_clientes")) Then qtde_clientes = t_PEDIDO("qtde_clientes")
        End If
        
    If qtde_clientes > 1 Then
        strMsgErro = "Não é possível obter o endereço de entrega porque os pedidos não são do mesmo cliente !!"
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        
    s = "SELECT TOP 1 * FROM t_PEDIDO WHERE pedido IN (" & s_filtro_pedidos & ") AND (st_end_entrega <> 0)"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbc, , , adCmdText
    If t_PEDIDO.EOF Then
        strMsgErro = "Não é possível obter o endereço de entrega porque não há endereço cadastrado !!"
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        
    If CLng(t_PEDIDO("st_end_entrega")) <> 0 Then
        strEndereco = Trim("" & t_PEDIDO("EndEtg_endereco"))
        strEnderecoNumero = Trim("" & t_PEDIDO("EndEtg_endereco_numero"))
        strEnderecoComplemento = Trim("" & t_PEDIDO("EndEtg_endereco_complemento"))
        strBairro = Trim("" & t_PEDIDO("EndEtg_bairro"))
        strCidade = Trim("" & t_PEDIDO("EndEtg_cidade"))
        strUF = Trim("" & t_PEDIDO("EndEtg_uf"))
        strCep = Trim("" & t_PEDIDO("EndEtg_cep"))
        strEnderecoCompletoFormatado = formata_endereco(strEndereco, strEnderecoNumero, strEnderecoComplemento, strBairro, strCidade, strUF, strCep)
        strEnderecoCompletoFormatado = UCase$(strEnderecoCompletoFormatado)
        If strEnderecoCompletoFormatado <> "" Then ObtemEnderecoEntrega = True
        End If
    
    GoSub OEE_FECHA_TABELAS
    
    aguarde INFO_NORMAL, m_id

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OEE_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OEE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OEE_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    Return
    
End Function

Sub NFe_emite(ByVal FLAG_NUMERACAO_MANUAL As Boolean)
' __________________________________________________________________________________________
'|
'|  EMITE A NOTA FISCAL ELETRÔNICA (NFe) COM BASE NO PEDIDO
'|  ESPECIFICADO E NOS DEMAIS PARÂMETROS PREENCHIDOS MANUALMENTE.
'|
'|  OS PRODUTOS (T_PEDIDO_ITEM) COM PRECO_NF = R$ 0,00 SÃO
'|  RELATIVOS A BRINDES E DEVEM SER TOTALMENTE IGNORADOS.
'|  OS BRINDES ACOMPANHAM OS OUTROS PRODUTOS DENTRO DA MESMA CAIXA.
'|

' CONSTANTES
Const NomeDestaRotina = "NFe_emite()"
Const MAX_LINHAS_NOTA_FISCAL_DEFAULT = 34
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
Dim strTransportadoraEmail2 As String
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
Dim strNFeTagInfRespTec As String
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
Dim strMarketplaceCodOrigemGrupo As String
Dim strMarketPlaceCNPJ As String
Dim strMarketPlaceCadIntTran As String
Dim strPagtoAntecipadoStatus As Integer
Dim strPagtoAntecipadoQuitadoStatus As Integer
Dim s_Texto_DIFAL_UF As String
Dim strCnpjIntermediadorPagto As String
Dim strMetodoPagto As String
Dim strMeioPagtoSefaz As String

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
Dim blnZerarDIFAL As Boolean
Dim blnEncontrouMeioPagtoSkyHub As Boolean

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
Dim intInformarIntermediadorPagto As Integer

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
Dim t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT As ADODB.Recordset
Dim t_CFG_MKTP_INTERMEDIADOR_PAGTO As ADODB.Recordset
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
Dim vl_pST As Currency
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
    strNFeTagFat = ""
    strNFeTagDup = ""
    strNFeTagInfRespTec = ""
    
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
        
  't_NFe_UF_PARAMETRO
    Set t_NFe_UF_PARAMETRO = New ADODB.Recordset
    With t_NFe_UF_PARAMETRO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  't_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT
    Set t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT = New ADODB.Recordset
    With t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
  
  't_CFG_MKTP_INTERMEDIADOR_PAGTO
    Set t_CFG_MKTP_INTERMEDIADOR_PAGTO = New ADODB.Recordset
    With t_CFG_MKTP_INTERMEDIADOR_PAGTO
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
    strMetodoPagto = ""
    intInformarIntermediadorPagto = 0
    strCnpjIntermediadorPagto = ""
    strMeioPagtoSefaz = ""
    strMarketplaceCodOrigemGrupo = ""
    blnEncontrouMeioPagtoSkyHub = False

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
                
                'If (strPagtoAntecipadoQuitadoStatus = "1") And (strPagtoAntecipadoQuitadoStatus = "1") Then
                '    If s_erro <> "" Then s_erro = s_erro & vbCrLf
                '    s_erro = s_erro & "Pedido " & Trim$(v_pedido(i)) & " está com pagamento antecipado quitado, emitir nota de venda no painel manual!!"
                '    End If
                
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
                
            'obter as informações de marketplace
            If (s_erro = "") And (param_nfintermediador.campo_inteiro = 1) And (strPedidoBSMarketplace <> "") And (strMarketplaceCodOrigem <> "") Then
                s = "SELECT o.codigo, o.descricao, og.parametro_campo_texto, og.parametro_2_campo_texto, og.parametro_3_campo_flag, o.codigo_pai  " & _
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
                    strMarketplaceCodOrigemGrupo = Trim$("" & t_CODIGO_DESCRICAO("codigo_pai"))
                    End If
                    
                'verificar tabela de configuração de marketplaces para obter parametros
                If strMarketPlaceCNPJ <> "" Then
                    s = "SELECT * FROM t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT WHERE pedido = '" & Trim$(v_pedido(i)) & "' "
                    If t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT.State <> adStateClosed Then t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT.Close
                    t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT.Open s, dbc, , , adCmdText
                    If Not t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT.EOF Then
                        strMetodoPagto = Trim$("" & t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT("method"))
                        If strMetodoPagto <> "" Then
                            s = "SELECT * FROM t_CFG_MKTP_INTERMEDIADOR_PAGTO " & _
                                "WHERE IdCodigoDescricaoCodigo = '" & strMarketplaceCodOrigemGrupo & "' " & _
                                "AND PaymentMethod = '" & strMetodoPagto & "' "
                            If t_CFG_MKTP_INTERMEDIADOR_PAGTO.State <> adStateClosed Then t_CFG_MKTP_INTERMEDIADOR_PAGTO.Close
                            t_CFG_MKTP_INTERMEDIADOR_PAGTO.Open s, dbc, , , adCmdText
                            If Not t_CFG_MKTP_INTERMEDIADOR_PAGTO.EOF Then
                                intInformarIntermediadorPagto = t_CFG_MKTP_INTERMEDIADOR_PAGTO("StInformarIntermediadorPagto")
                                strCnpjIntermediadorPagto = Trim$("" & t_CFG_MKTP_INTERMEDIADOR_PAGTO("CnpjIntermediadorPagto"))
                                strMeioPagtoSefaz = Trim$("" & t_CFG_MKTP_INTERMEDIADOR_PAGTO("CodigoMeioPagtoSefaz"))
                                blnEncontrouMeioPagtoSkyHub = True
                                End If
                            End If
                                                
                        'se não encontrar meio pagto, procurar OUTROS
                        If Not blnEncontrouMeioPagtoSkyHub Then
                            s = "SELECT * FROM t_CFG_MKTP_INTERMEDIADOR_PAGTO " & _
                                "WHERE IdCodigoDescricaoCodigo = '" & strMarketplaceCodOrigemGrupo & "' " & _
                                "AND PaymentMethod = '" & "*_OUTROS_*" & "' "
                            If t_CFG_MKTP_INTERMEDIADOR_PAGTO.State <> adStateClosed Then t_CFG_MKTP_INTERMEDIADOR_PAGTO.Close
                            t_CFG_MKTP_INTERMEDIADOR_PAGTO.Open s, dbc, , , adCmdText
                            If Not t_CFG_MKTP_INTERMEDIADOR_PAGTO.EOF Then
                                intInformarIntermediadorPagto = t_CFG_MKTP_INTERMEDIADOR_PAGTO("StInformarIntermediadorPagto")
                                strCnpjIntermediadorPagto = Trim$("" & t_CFG_MKTP_INTERMEDIADOR_PAGTO("CnpjIntermediadorPagto"))
                                strMeioPagtoSefaz = Trim$("" & t_CFG_MKTP_INTERMEDIADOR_PAGTO("CodigoMeioPagtoSefaz"))
                                End If
                            End If
                                                
                        End If
                    End If
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
        aviso_erro "Problemas na identificação do emitente!!"
        GoSub NFE_EMITE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
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
        If (param_nfmemooendentrega.campo_inteiro = 1) Then strEndEtgUf = UCase$(Trim$("" & t_DESTINATARIO("uf_end_nota")))
        If (param_nfmemooendentrega.campo_inteiro = 1) Then
            strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf_end_nota")))
        Else
            strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
            End If
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
        strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
        End If
        
    
'   CONFIRMA ALÍQUOTA DO ICMS
    If obtem_aliquota_ICMS(usuario.emit_uf, strEndClienteUf, aliquota_icms_interestadual) Then
        strIcms = Trim$(CStr(aliquota_icms_interestadual))
    Else
        strIcms = ""
        End If
    
    If (strIcms <> "") And (cb_icms <> "") Then
        If (CSng(strIcms) <> CSng(cb_icms)) Then
            s = "O destinatário é do estado de " & strEndClienteUf & " cuja alíquota de ICMS é de " & strIcms & "%" & _
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
            's = "Pedido " & Trim$(v_pedido(i)) & " se refere a venda futura não quitada!"
            s = "Pedido se refere a venda futura não quitada!"
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
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("endereco_end_nota"))
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
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("numero_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("endereco_numero"))
        End If
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
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("complemento_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("endereco_complemento"))
        End If
    If Len(strCampo) > 60 Then
        s_erro = "O campo complemento do endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xCpl = strCampo
    If Len(strCampo) > 0 Then strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xCpl", rNFeImg.dest__xCpl)
    
'   BAIRRO
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("bairro_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("bairro"))
        End If
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
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("cidade_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("cidade"))
        End If
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        s_aux = Trim$("" & t_DESTINATARIO("uf_end_nota"))
    Else
        s_aux = Trim$("" & t_DESTINATARIO("uf"))
        End If
    If (strCampo <> "") And (s_aux <> "") Then strCampo = strCampo & "/"
    strCampo = strCampo & s_aux
    rNFeImg.dest__cMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("cMun", rNFeImg.dest__cMun)
    
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("cidade_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("cidade"))
        End If
    If Len(strCampo) > 60 Then
        s_erro = "O campo cidade no endereço do cliente excede o tamanho máximo (60 caracteres)!!"
        GoTo NFE_EMITE_ENCERRA_POR_ERRO_CONSISTENCIA
        End If
    rNFeImg.dest__xMun = strCampo
    strNFeTagDestinatario = strNFeTagDestinatario & vbTab & NFeFormataCampo("xMun", rNFeImg.dest__xMun)
    
'   UF
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = Trim$("" & t_DESTINATARIO("uf_end_nota"))
    Else
        strCampo = Trim$("" & t_DESTINATARIO("uf"))
        End If
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
    If blnExisteMemorizacaoEndereco And (strEndClienteUf = strEndEtgUf) And (param_nfmemooendentrega.campo_inteiro = 1) Then
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep_end_nota")))
    Else
        strCampo = retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep")))
        End If
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
                vl_pST = 0
                
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
                    
                '   ALIQUOTA SUPORTADA PELO CONSUMIDOR FINAL
                    vl_pST = 0
                    strNFeTagIcms = strNFeTagIcms & vbTab & NFeFormataCampo("pST", NFeFormataMoeda2Dec(vl_pST))
                
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
                blnZerarDIFAL = False
                s_Texto_DIFAL_UF = ""
                
                s = "SELECT " & _
                    "st_ignorar_difal, " & _
                    "st_zerar_difal, " & _
                    "texto_adicional" & _
                    " FROM t_NFe_UF_PARAMETRO" & _
                    " WHERE" & _
                    " (UF='" & Trim$(strEndClienteUf) & "')"
                If t_NFe_UF_PARAMETRO.State <> adStateClosed Then t_NFe_UF_PARAMETRO.Close
                t_NFe_UF_PARAMETRO.Open s, dbc, , , adCmdText
                If Not t_NFe_UF_PARAMETRO.EOF Then
                    blnIgnorarDIFAL = t_NFe_UF_PARAMETRO("st_ignorar_difal") = 1
                    blnZerarDIFAL = t_NFe_UF_PARAMETRO("st_zerar_difal") = 1
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
                            
            '   OBS: EXISTE IPI APENAS NA EMISSÃO DE NFe PARA DEVOLUÇÃO AO FORNECEDOR
                If blnIgnorarDIFAL Then
                
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
'                A partir de 03/04/2023 condição abaixo de ICMS maior que zero começou a dar problema
'                (Nota Técnica 2022.005 SEFAZ)
'                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
'                    (rNFeImg.dest__indIEDest = "9") And _
'                    Not cfop_eh_de_remessa(strCfopCodigo) And _
'                    Not blnIgnorarDIFAL And _
'                    (vl_ICMS > 0) Then

                If blnZerarDIFAL Then
                    perc_fcp = 0
                    perc_ICMS_interna_UF_dest = 0
                    vl_ICMS_UF_dest = 0
                    vl_ICMS_UF_remet = 0
                    End If

                
                If PARTILHA_ICMS_ATIVA And (rNFeImg.ide__idDest = "2") And _
                    (rNFeImg.dest__indIEDest = "9") And _
                    Not cfop_eh_de_remessa(strCfopCodigo) And _
                    Not blnIgnorarDIFAL Then
                
                    strNFeTagIcmsUFDest = ""
                    
                '   VALOR DA BC DO ICMS NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest = NFeFormataMoeda2Dec(vl_BC_ICMS)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__vBCUFDest)
                '   PERCENTUAL DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest = NFeFormataPercentual2Dec(perc_fcp)
                    strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("pFCPUFDest", vNFeImgItem(UBound(vNFeImgItem)).ICMSUFDest__pFCPUFDest)
                
                    'VALOR DA BASE DE CÁLCULO DO ICMS RELATIVO AO FUNDO DE COMBATE À POBREZA NA UF DE DESTINO
                    '(lhgx) obs: manter esta linha comentada, pois podemos ter problema com o resultado no ambiente de produção
                    'strNFeTagIcmsUFDest = strNFeTagIcmsUFDest & vbTab & NFeFormataCampo("vBCFCPUFDest", NFeFormataMoeda2Dec(vl_BC_ICMS))
                
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
    'Se a operação envolve marketplace, substituir com os valores obtidos
    ElseIf (rNFeImg.ide__tpNF = "1") And _
        (param_nfintermediador.campo_inteiro = 1) And _
        (strMarketplaceCodOrigem <> "") And _
        (strMeioPagtoSefaz <> "") Then
        vNFeImgPag(UBound(vNFeImgPag)).pag__indPag = "0"
        vNFeImgPag(UBound(vNFeImgPag)).pag__tPag = strMeioPagtoSefaz
        s = strMetodoPagto
        vNFeImgPag(UBound(vNFeImgPag)).pag__vPag = rNFeImg.total__vNF
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
        (param_contingencia_meio_pagamento_cartao.campo_inteiro = 1) Or _
        (strMeioPagtoSefaz = "99") Then
        If s <> "" Then strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("xPag", s)
        End If
    strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("vPag", vNFeImgPag(UBound(vNFeImgPag)).pag__vPag)
    'Segundo informado pelo Valter (Target) em e-mail de 27/07/2017, o grupo vcard não deve ser informado no arquivo texto,
    'ele é preenchido pelo sistema
    'ATUALIZAÇÃO: a partir de 2022, após a nota técnica 2020.006 v 1.30 da SEFAZ, mudou-se o entendimento e o grupo card
    'passou a ser preenchido quando o meio de pagamento for cartão de crédito e o CNPJ do intermediador do pagamento existir
    If (param_nfintermediador.campo_inteiro = 1) And _
        (intInformarIntermediadorPagto = 1) And _
        (strCnpjIntermediadorPagto <> "") Then
        strNFeTagPag = strNFeTagPag & vbTab & "card;"
        strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("tpIntegra", "1")
        strNFeTagPag = strNFeTagPag & vbTab & NFeFormataCampo("CNPJ", strCnpjIntermediadorPagto)
        End If
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
    
'   ENDEREÇO DE ENTREGA (INF ADICIONAIS)
    blnTemEndEtg = ObtemEnderecoEntrega(strEndEtgEndereco, strEndEtgEnderecoNumero, strEndEtgEnderecoComplemento, strEndEtgBairro, strEndEtgCidade, strEndEtgUf, strEndEtgCep, strEndEtgEnderecoCompletoFormatado, s_erro_aux)
    If blnTemEndEtg Then
        strEndEtgEnderecoCompletoFormatado = "ENTREGA: " & strEndEtgEnderecoCompletoFormatado
    '   SÓ É PERMITIDO USAR UM ENDEREÇO DE ENTREGA DIFERENTE DENTRO DE UM MESMO ESTADO
        If (UCase$(strEndEtgUf) = UCase$(strEndClienteUf)) Then
            If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
            strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & strEndEtgEnderecoCompletoFormatado
            End If
        End If
    
'   TEXTO DIGITADO
    If Trim$(c_dados_adicionais) <> "" Then
        If strNFeInfAdicQuadroInfAdic <> "" Then strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & vbCrLf
        strNFeInfAdicQuadroInfAdic = strNFeInfAdicQuadroInfAdic & Trim$(c_dados_adicionais)
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
'   (a partir de 01/12/2022, esta mensagem não será exibida em NF's de transferência de estoque entre filiais - CFOP 5152
    If blnExibirTotalTributos And (Not blnHaProdutoSemDadosIbpt) And (strInfoAdicIbpt <> "") And (strCfopCodigo <> "5152") Then
        If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
        strNFeInfAdicQuadroProdutos = strInfoAdicIbpt & strNFeInfAdicQuadroProdutos
        End If
    
'   Nº PEDIDO (NA 1ª LINHA) + CUBAGEM
    strTextoCubagem = ""
    If cubagem_bruto > 0 Then strTextoCubagem = Space$(20) & "CUB: " & formata_numero_2dec(cubagem_bruto) & " m3"
    If strNFeInfAdicQuadroProdutos <> "" Then strNFeInfAdicQuadroProdutos = vbCrLf & strNFeInfAdicQuadroProdutos
    strNFeInfAdicQuadroProdutos = Join(v_pedido, ", ") & strTextoCubagem & strNFeInfAdicQuadroProdutos
    
'   INFORMAÇÕES SOBRE PARTILHA DO ICMS
    If PARTILHA_ICMS_ATIVA And _
        Not blnIgnorarDIFAL Then
        'DIFAL- suprimir texto em notas de entrada/devolução
        If (rNFeImg.ide__tpNF <> "0") And _
            (strNFeCodFinalidade <> "3") And _
            (strNFeCodFinalidade <> "4") And _
                Not tem_instricao_virtual(usuario.emit_id, rNFeImg.dest__UF) Then
            If (vl_total_ICMSUFDest > 0) Or blnZerarDIFAL Then
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
    

'   TAG ENTREGA
'   ~~~~~~~~~~~
    If blnTemEndEtg Then
    '   SÓ É PERMITIDO USAR UM ENDEREÇO DE ENTREGA DIFERENTE DENTRO DE UM MESMO ESTADO
        If (UCase$(strEndEtgUf) <> UCase$(strEndClienteUf)) Then
            s = "ATENÇÃO!!" & vbCrLf & _
                "O pedido informa um endereço de entrega localizado em outra UF, portanto, nenhuma informação sobre o endereço de entrega será adicionada automaticamente nesta NF!!" & vbCrLf & _
                vbCrLf & _
                "Endereço do Cliente:" & vbCrLf & _
                vbTab & UCase$(formata_endereco(Trim("" & t_DESTINATARIO("endereco")), Trim$("" & t_DESTINATARIO("endereco_numero")), Trim$("" & t_DESTINATARIO("endereco_complemento")), Trim$("" & t_DESTINATARIO("bairro")), Trim$("" & t_DESTINATARIO("cidade")), Trim$("" & t_DESTINATARIO("uf")), retorna_so_digitos(Trim$("" & t_DESTINATARIO("cep"))))) & vbCrLf & _
                vbCrLf & _
                "Endereço de Entrega:" & vbCrLf & _
                vbTab & UCase$(formata_endereco(strEndEtgEndereco, strEndEtgEnderecoNumero, strEndEtgEnderecoComplemento, strEndEtgBairro, strEndEtgCidade, strEndEtgUf, strEndEtgCep)) & _
                vbCrLf & vbCrLf & _
                "Continua mesmo assim?"
            If Not confirma(s) Then
                GoSub NFE_EMITE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
        Else
        '   NO MOMENTO, A SEFAZ ACEITA ENDEREÇO DE ENTREGA DIFERENTE DO ENDEREÇO DE CADASTRO SOMENTE P/ PJ
        '   (SÓ INFORMAR O ENDEREÇO DE ENTREGA SE FOR DIFERENTE DO ENDEREÇO DO CLIENTE)
            If (rNFeImg.dest__xLgr <> strEndEtgEndereco) And _
                (rNFeImg.dest__nro <> strEndEtgEnderecoNumero) And _
                (rNFeImg.dest__xCpl <> strEndEtgEnderecoComplemento) And _
                (rNFeImg.dest__xBairro <> strEndEtgBairro) And _
                (rNFeImg.dest__xMun <> strEndEtgCidade) Then

                If cnpj_cpf_ok(strDestinatarioCnpjCpf) Then
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
                End If
            End If
        End If
        
'   TAG INFRESTEC
'   ~~~~~~~~~~~~~
    If (param_nfinformaresptec.campo_inteiro = 1) And (resptec_emissor.CNPJ <> "") Then
        strNFeTagInfRespTec = "infRespTec;" & vbCrLf & _
                                vbTab & NFeFormataCampo("CNPJ", resptec_emissor.CNPJ) & _
                                vbTab & NFeFormataCampo("xContato", resptec_emissor.nome) & _
                                vbTab & NFeFormataCampo("email", resptec_emissor.EMAIL) & _
                                vbTab & NFeFormataCampo("fone", resptec_emissor.telefone)
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
    
'   SE HOUVER MAIS DE UMA CONFIRMAÇÃO DE EMISSÃO QUE PODEM GERAR NFe PARA UM EMITENTE INDEVIDO, CONFIRMAR NOVAMENTE
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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


'   SE O PEDIDO ESTIVER NA FILA DE SOLICITAÇÃO DE EMISSÃO DE NFE, SINALIZA QUE JÁ FOI TRATADO
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            If Not marca_status_atendido_fila_solicitacoes_emissao_NFe(Trim$(v_pedido(i)), rNFeImg.id_nfe_emitente, CLng(strSerieNf), CLng(strNumeroNf), s_erro_aux) Then
                s = "Não é possível prosseguir com a emissão, pois houve falha ao atualizar os dados da fila de solicitações de emissão de NFe!!" & vbCrLf & _
                    s_erro_aux
                aviso_erro s
                GoSub NFE_EMITE_FECHA_TABELAS
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
    cmdNFeEmite.Parameters.Append cmdNFeEmite.CreateParameter("Arquivo", adVarChar, adParamInput, Len(strNFeArquivo), strNFeArquivo)
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
'   OBS: ESTE PROCESSAMENTO É REALIZADO APENAS AO INFORMAR MANUALMENTE
'        O Nº DA NFe, POIS O OBJETIVO É EVITAR QUE EXISTAM 2 EMISSÕES
'        VÁLIDAS C/ O MESMO Nº DE NFe.
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
    If LBound(v_pedido) = UBound(v_pedido) Then
        strLogPedido = v_pedido(UBound(v_pedido))
    Else
        strLogPedido = ""
        End If
    strLogComplemento = "Retorno SP=" & CStr(intNfeRetornoSPEmite) & " (" & IIf(intNfeRetornoSPEmite = 1, "Sucesso", "Falha") & ")" & _
                        "; Msg SP=" & strNFeMsgRetornoSPEmite & _
                        "; Série NFe=" & strSerieNf & _
                        "; Nº NFe=" & strNumeroNf & _
                        "; tela emissão=Painel Automático" & _
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
                        "; dados adicionais=" & Trim$(c_dados_adicionais)
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
    formulario_limpa
        
  ' EXIBE DADOS DA ÚLTIMA NFe EMITIDA
    l_serie_NF = strSerieNfNormalizado
    l_num_NF = strNumeroNfNormalizado
    l_emitente_NF = strEmitenteNf
        
    GoSub NFE_EMITE_FECHA_TABELAS
    
    If blnFilaSolicitacoesEmissaoNFeEmTratamento Then
    '   AO PREENCHER C/ O PRÓXIMO PEDIDO DA FILA, A QTDE PENDENTE NA FILA É ATUALIZADA AUTOMATICAMENTE
        preenche_prox_pedido_fila_solicitacoes_emissao_NFe
    Else
    '   ATUALIZA A QTDE PENDENTE NA FILA, POIS O PEDIDO INFORMADO MANUALMENTE PODE TER SIDO UM QUE CONSTAVA NA FILA
        atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
        End If
    
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
    bd_desaloca_recordset t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT, True
    bd_desaloca_recordset t_CFG_MKTP_INTERMEDIADOR_PAGTO, True
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


'Function calculaDataPrimeiroBoleto(ByVal intPrazoEmissaoPrimeiroBoleto As Integer) As Date
'
'Dim dtResposta As Date
'
'    If intPrazoEmissaoPrimeiroBoleto <= 29 Then
'        dtResposta = Date + 30
'    Else
'        'dtResposta = Date + intPrazoEmissaoPrimeiroBoleto + 7
'        'REMOÇÃO DOS 07 DIAS ADICIONAIS, A PEDIDO DO CARLOS
'        dtResposta = Date + intPrazoEmissaoPrimeiroBoleto
'        End If
'
'    calculaDataPrimeiroBoleto = dtResposta
'
'End Function


Sub formulario_inicia()

Dim s As String
Dim s_aux As String
Dim msg_erro As String
Dim v_CFOP() As TIPO_LISTA_CFOP
Dim i As Integer
Dim j As Integer
Dim i_qtde As Integer
Dim vAliquotas() As String

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
    

'   LOCAL DE DESTINO DA OPERAÇÃO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cb_loc_dest.Clear
    cb_loc_dest.AddItem "1 - INTERNA"
    cb_loc_dest.AddItem "2 - INTERESTADUAL"
    cb_loc_dest.AddItem "3 - EXTERIOR"
    
    
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
    cb_icms.Clear
    cb_icms.AddItem "0"
    cb_icms.AddItem "4"
    cb_icms.AddItem "7"
    cb_icms.AddItem "12"
    cb_icms.AddItem "17"
    cb_icms.AddItem "18"
    cb_icms.AddItem "20"
    
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
    
'   DADOS ADICIONAIS
'   ~~~~~~~~~~~~~~~~
    With c_dados_adicionais
        .FontName = FONTNAME_IMPRESSAO
        .FontSize = FONTSIZE_IMPRESSAO
        .FontBold = FONTBOLD_IMPRESSAO
        .FontItalic = FONTITALIC_IMPRESSAO
        End With
        
'   BOTÃO NOTA TRIANGULAR
'   ~~~~~~~~~~~~~~~~~~~~~
    b_emissao_nfe_triangular.Visible = blnNotaTriangularAtiva
    b_emissao_nfe_triangular.Enabled = blnNotaTriangularAtiva
    
End Sub

'Private Function geraDadosParcelasPagto(v_pedido() As String, v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO, ByRef strMsgErro As String) As Boolean
'' __________________________________________________________________________________________
''|
''|  ANALISA O(S) PEDIDO(S) PARA VERIFICAR SE HÁ ALGUM QUE ESPECIFICA PAGAMENTO VIA BOLETO.
''|  EM CASO AFIRMATIVO, CALCULA A QUANTIDADE DE PARCELAS, DATAS E VALORES.
''|
'
'Dim s As String
'Dim s_where As String
'Dim i As Integer
'Dim j As Integer
'Dim intQtdeTotalPedidos As Integer
'Dim intQtdePedidosPagtoBoleto As Integer
'Dim intQtdeTotalParcelas As Integer
'Dim intQtdePlanoContas As Integer
'Dim vlTotalPedido As Currency
'Dim vlTotalFormaPagto As Currency
'Dim vlDiferencaArredondamento As Currency
'Dim vlDiferencaArredondamentoRestante As Currency
'Dim vlRateio As Currency
'Dim dtUltimoPagtoCalculado As Date
'Dim blnPagtoPorBoleto As Boolean
'Dim strTipoParcelamento As String
'Dim strListaPedidosPagtoBoleto As String
'Dim strListaPedidosPagtoNaoBoleto As String
'Dim vPedidoCalculoParcelas() As TIPO_PEDIDO_CALCULO_PARCELAS_BOLETO
'
'' BANCO DE DADOS
'Dim t_PEDIDO As ADODB.Recordset
'Dim t_PEDIDO_ITEM As ADODB.Recordset
'Dim tAux As ADODB.Recordset
'
'    On Error GoTo GDPP_TRATA_ERRO
'
'    geraDadosParcelasPagto = False
'
'    strMsgErro = ""
'    ReDim v_parcela_pagto(0)
'
'    ReDim vPedidoCalculoParcelas(0)
'    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pedido = ""
'
'  ' T_PEDIDO
'    Set t_PEDIDO = New ADODB.Recordset
'    With t_PEDIDO
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
'
'  ' T_PEDIDO_ITEM
'    Set t_PEDIDO_ITEM = New ADODB.Recordset
'    With t_PEDIDO_ITEM
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
'
'  ' tAux
'    Set tAux = New ADODB.Recordset
'    With tAux
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
'
'    For i = LBound(v_pedido) To UBound(v_pedido)
'        If Trim$(v_pedido(i)) <> "" Then
'            s = "SELECT" & _
'                    " t_PEDIDO__BASE.tipo_parcelamento," & _
'                    " t_PEDIDO__BASE.av_forma_pagto," & _
'                    " t_PEDIDO__BASE.pc_qtde_parcelas," & _
'                    " t_PEDIDO__BASE.pc_valor_parcela," & _
'                    " t_PEDIDO__BASE.pce_forma_pagto_entrada," & _
'                    " t_PEDIDO__BASE.pce_forma_pagto_prestacao," & _
'                    " t_PEDIDO__BASE.pce_entrada_valor," & _
'                    " t_PEDIDO__BASE.pce_prestacao_qtde," & _
'                    " t_PEDIDO__BASE.pce_prestacao_valor," & _
'                    " t_PEDIDO__BASE.pce_prestacao_periodo," & _
'                    " t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
'                    " t_PEDIDO__BASE.pse_forma_pagto_demais_prest," & _
'                    " t_PEDIDO__BASE.pse_prim_prest_valor," & _
'                    " t_PEDIDO__BASE.pse_prim_prest_apos," & _
'                    " t_PEDIDO__BASE.pse_demais_prest_qtde," & _
'                    " t_PEDIDO__BASE.pse_demais_prest_valor," & _
'                    " t_PEDIDO__BASE.pse_demais_prest_periodo," & _
'                    " t_PEDIDO__BASE.pu_forma_pagto," & _
'                    " t_PEDIDO__BASE.pu_valor," & _
'                    " t_PEDIDO__BASE.pu_vencto_apos" & _
'                " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
'                    " ON (SUBSTRING(t_PEDIDO.pedido,1," & CStr(TAM_MIN_ID_PEDIDO) & ")=t_PEDIDO__BASE.pedido)" & _
'                " WHERE" & _
'                    " (t_PEDIDO.pedido='" & Trim$(v_pedido(i)) & "')"
'            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
'            t_PEDIDO.Open s, dbc, , , adCmdText
'            If t_PEDIDO.EOF Then
'                If strMsgErro <> "" Then strMsgErro = strMsgErro & vbCrLf
'                strMsgErro = strMsgErro & "Pedido " & Trim$(v_pedido(i)) & " não está cadastrado!!"
'            Else
'                intQtdeTotalPedidos = intQtdeTotalPedidos + 1
'
'                strTipoParcelamento = Trim$("" & t_PEDIDO("tipo_parcelamento"))
'                blnPagtoPorBoleto = False
'                If strTipoParcelamento = CStr(COD_FORMA_PAGTO_A_VISTA) Then
'                    If Trim$("" & t_PEDIDO("av_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
'                    If Trim$("" & t_PEDIDO("pce_forma_pagto_entrada")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                    If Trim$("" & t_PEDIDO("pce_forma_pagto_prestacao")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
'                    If Trim$("" & t_PEDIDO("pse_forma_pagto_prim_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                    If Trim$("" & t_PEDIDO("pse_forma_pagto_demais_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
'                    If Trim$("" & t_PEDIDO("pu_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
'                    End If
'
'                If (Trim$(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pedido) <> "") Then
'                    ReDim Preserve vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas) + 1)
'                    End If
'
'                With vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas))
'                    .pedido = Trim$(v_pedido(i))
'                    .tipo_parcelamento = t_PEDIDO("tipo_parcelamento")
'                    .av_forma_pagto = t_PEDIDO("av_forma_pagto")
'                    .pu_forma_pagto = t_PEDIDO("pu_forma_pagto")
'                    .pu_valor = t_PEDIDO("pu_valor")
'                    .pu_vencto_apos = t_PEDIDO("pu_vencto_apos")
'                    .pc_qtde_parcelas = t_PEDIDO("pc_qtde_parcelas")
'                    .pc_valor_parcela = t_PEDIDO("pc_valor_parcela")
'                    .pce_forma_pagto_entrada = t_PEDIDO("pce_forma_pagto_entrada")
'                    .pce_forma_pagto_prestacao = t_PEDIDO("pce_forma_pagto_prestacao")
'                    .pce_entrada_valor = t_PEDIDO("pce_entrada_valor")
'                    .pce_prestacao_qtde = t_PEDIDO("pce_prestacao_qtde")
'                    .pce_prestacao_valor = t_PEDIDO("pce_prestacao_valor")
'                    .pce_prestacao_periodo = t_PEDIDO("pce_prestacao_periodo")
'                    .pse_forma_pagto_prim_prest = t_PEDIDO("pse_forma_pagto_prim_prest")
'                    .pse_forma_pagto_demais_prest = t_PEDIDO("pse_forma_pagto_demais_prest")
'                    .pse_prim_prest_valor = t_PEDIDO("pse_prim_prest_valor")
'                    .pse_prim_prest_apos = t_PEDIDO("pse_prim_prest_apos")
'                    .pse_demais_prest_qtde = t_PEDIDO("pse_demais_prest_qtde")
'                    .pse_demais_prest_valor = t_PEDIDO("pse_demais_prest_valor")
'                    .pse_demais_prest_periodo = t_PEDIDO("pse_demais_prest_periodo")
'                    End With
'
'            '   CALCULA O VALOR TOTAL DESTE PEDIDO
'                s = "SELECT" & _
'                        " p.pedido," & _
'                        " Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
'                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" & _
'                    " WHERE" & _
'                        " (p.pedido = '" & Trim$(v_pedido(i)) & "')" & _
'                    " GROUP BY" & _
'                        " p.pedido" & _
'                    " UNION " & _
'                    " SELECT" & _
'                        " p.pedido," & _
'                        " -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
'                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" & _
'                    " WHERE" & _
'                        " (p.pedido = '" & Trim$(v_pedido(i)) & "')" & _
'                    " GROUP BY" & _
'                        " p.pedido"
'
'                s = "SELECT" & _
'                        " pedido," & _
'                        " Sum(vl_total) AS vl_total" & _
'                    " FROM" & _
'                        "(" & _
'                            s & _
'                        ") t" & _
'                    " GROUP BY" & _
'                        " pedido"
'
'                If tAux.State <> adStateClosed Then tAux.Close
'                tAux.Open s, dbc, , , adCmdText
'                If tAux.EOF Then
'                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalDestePedido = 0
'                Else
'                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalDestePedido = tAux("vl_total")
'                    End If
'
'            '   CALCULA O VALOR TOTAL DA FAMÍLIA DE PEDIDOS
'                s = "SELECT" & _
'                        " Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
'                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" & _
'                    " WHERE" & _
'                        " (p.pedido LIKE '" & retorna_num_pedido_base(Trim$(v_pedido(i))) & BD_CURINGA_TODOS & "')" & _
'                        " AND (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
'                    " UNION " & _
'                    " SELECT" & _
'                        " -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
'                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" & _
'                    " WHERE" & _
'                        " (p.pedido LIKE '" & retorna_num_pedido_base(Trim$(v_pedido(i))) & BD_CURINGA_TODOS & "')"
'
'                s = "SELECT" & _
'                        " Sum(vl_total) AS vl_total" & _
'                    " FROM" & _
'                        "(" & _
'                            s & _
'                        ") t"
'
'                If tAux.State <> adStateClosed Then tAux.Close
'                tAux.Open s, dbc, , , adCmdText
'                If tAux.EOF Then
'                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalFamiliaPedidos = 0
'                Else
'                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalFamiliaPedidos = tAux("vl_total")
'                    End If
'
'            '   CALCULA A RAZÃO ENTRE OS VALORES DESTE PEDIDO E A FAMÍLIA DE PEDIDOS
'                With vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas))
'                    If .vlTotalFamiliaPedidos = 0 Then
'                        .razaoValorPedidoFilhote = 0
'                    Else
'                        .razaoValorPedidoFilhote = .vlTotalDestePedido / .vlTotalFamiliaPedidos
'                        End If
'                    End With
'
'                If blnPagtoPorBoleto Then
'                    intQtdePedidosPagtoBoleto = intQtdePedidosPagtoBoleto + 1
'                    If strListaPedidosPagtoBoleto <> "" Then strListaPedidosPagtoBoleto = strListaPedidosPagtoBoleto & ", "
'                    strListaPedidosPagtoBoleto = strListaPedidosPagtoBoleto & Trim$(v_pedido(i))
'                Else
'                    If strListaPedidosPagtoNaoBoleto <> "" Then strListaPedidosPagtoNaoBoleto = strListaPedidosPagtoNaoBoleto & ", "
'                    strListaPedidosPagtoNaoBoleto = strListaPedidosPagtoNaoBoleto & Trim$(v_pedido(i))
'                    End If
'                End If
'            End If
'        Next
'
'
'
''   SE HOUVER ALGUM PEDIDO QUE DEFINA PAGAMENTO POR BOLETO, OS DADOS DE PAGAMENTO SERÃO IMPRESSOS NA NF.
''   ENTRETANTO, QUANDO HÁ MAIS DE 2 PEDIDOS, A FORMA DE PAGAMENTO DEVE SER IDÊNTICA P/ QUE SE POSSA SOMAR
''   OS VALORES DE CADA PARCELA, CASO CONTRÁRIO SERÁ RETORNADA UMA MENSAGEM DE ERRO PARA EXIBIÇÃO.
'
''   NÃO HÁ PEDIDOS POR BOLETOS!
'    If intQtdePedidosPagtoBoleto = 0 Then
'        geraDadosParcelasPagto = True
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'
'  ' HÁ PEDIDOS QUE SÃO POR BOLETO E OUTROS QUE NÃO
'    If intQtdePedidosPagtoBoleto <> intQtdeTotalPedidos Then
'        strMsgErro = "Há pedido(s) que especifica(m) pagamento via boleto bancário e há pedido(s) que especifica(m) outro(s) meio(s) de pagamento:" & Chr(13) & _
'                     "Pagamento via boleto bancário: " & strListaPedidosPagtoBoleto & Chr(13) & _
'                     "Pagamento via outros meios: " & strListaPedidosPagtoNaoBoleto & Chr(13) & _
'                     Chr(13) & _
'                     "Não é possível gerar os dados de pagamento para impressão na NFe!!"
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'
'  ' HÁ MAIS DO QUE 1 PEDIDO P/ SER PAGO POR BOLETO
'    If intQtdePedidosPagtoBoleto > 1 Then
'      ' HÁ PEDIDOS QUE ESPECIFICAM DIFERENTES FORMAS DE PAGAMENTO
'        For i = LBound(vPedidoCalculoParcelas) To (UBound(vPedidoCalculoParcelas) - 1)
'            If vPedidoCalculoParcelas(i).tipo_parcelamento <> vPedidoCalculoParcelas(i + 1).tipo_parcelamento Then
'                If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                strMsgErro = strMsgErro & "Pedido " & vPedidoCalculoParcelas(i).pedido & "=" & descricao_tipo_parcelamento(vPedidoCalculoParcelas(i).tipo_parcelamento) & _
'                             " e pedido " & vPedidoCalculoParcelas(i + 1).pedido & "=" & descricao_tipo_parcelamento(vPedidoCalculoParcelas(i + 1).tipo_parcelamento)
'                End If
'            Next
'
'        If strMsgErro <> "" Then
'            strMsgErro = "Os pedidos especificam diferentes formas de pagamento!!" & _
'                        Chr(13) & _
'                        strMsgErro & _
'                        Chr(13) & _
'                        Chr(13) & _
'                        "Não é possível gerar os dados de pagamento para impressão na NFe!!"
'            GoSub GDPP_FECHA_TABELAS
'            Exit Function
'            End If
'
'      ' HÁ PEDIDOS QUE P/ UMA FORMA DE PAGAMENTO DEFINEM DIFERENTES PRAZOS DE PAGAMENTO
'        For i = LBound(vPedidoCalculoParcelas) To (UBound(vPedidoCalculoParcelas) - 1)
'        '   PARCELADO COM ENTRADA
'        '   ~~~~~~~~~~~~~~~~~~~~~
'            If CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
'                If vPedidoCalculoParcelas(i).pce_forma_pagto_entrada <> vPedidoCalculoParcelas(i + 1).pce_forma_pagto_entrada Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na forma de pagamento da entrada: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pce_forma_pagto_entrada) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pce_forma_pagto_entrada) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pce_forma_pagto_prestacao <> vPedidoCalculoParcelas(i + 1).pce_forma_pagto_prestacao Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na forma de pagamento das prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pce_forma_pagto_prestacao) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pce_forma_pagto_prestacao) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pce_prestacao_qtde <> vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na quantidade de prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pce_prestacao_qtde) & " " & IIf(vPedidoCalculoParcelas(i).pce_prestacao_qtde > 1, "prestações", "prestação") & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde) & " " & IIf(vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde > 1, "prestações", "prestação") & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pce_prestacao_periodo <> vPedidoCalculoParcelas(i + 1).pce_prestacao_periodo Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência no período de vencimento das prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pce_prestacao_periodo) & " dias) e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pce_prestacao_periodo) & " dias)"
'                    End If
'
'        '   PARCELADO SEM ENTRADA
'        '   ~~~~~~~~~~~~~~~~~~~~~
'            ElseIf CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
'                If vPedidoCalculoParcelas(i).pse_forma_pagto_prim_prest <> vPedidoCalculoParcelas(i + 1).pse_forma_pagto_prim_prest Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na forma de pagamento da 1ª prestação: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pse_forma_pagto_prim_prest) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pse_forma_pagto_prim_prest) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pse_forma_pagto_demais_prest <> vPedidoCalculoParcelas(i + 1).pse_forma_pagto_demais_prest Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na forma de pagamento das demais prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pse_forma_pagto_demais_prest) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pse_forma_pagto_demais_prest) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pse_prim_prest_apos <> vPedidoCalculoParcelas(i + 1).pse_prim_prest_apos Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência no prazo de pagamento da 1ª prestação: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_prim_prest_apos) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_prim_prest_apos) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pse_demais_prest_qtde <> vPedidoCalculoParcelas(i + 1).pse_demais_prest_qtde Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência na quantidade de prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_demais_prest_qtde) & ") e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_demais_prest_qtde) & ")"
'                    End If
'
'                If vPedidoCalculoParcelas(i).pse_demais_prest_periodo <> vPedidoCalculoParcelas(i + 1).pse_demais_prest_periodo Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência no período de vencimento das prestações: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_demais_prest_periodo) & " dias) e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_demais_prest_periodo) & " dias)"
'                    End If
'
'        '   PARCELA ÚNICA
'        '   ~~~~~~~~~~~~~
'            ElseIf CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
'                If vPedidoCalculoParcelas(i).pu_vencto_apos <> vPedidoCalculoParcelas(i + 1).pu_vencto_apos Then
'                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
'                    strMsgErro = strMsgErro & "Divergência no prazo de vencimento da parcela única: " & _
'                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pu_vencto_apos) & " dia(s)) e " & _
'                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pu_vencto_apos) & " dia(s))"
'                    End If
'                End If
'            Next
'
'        If strMsgErro <> "" Then
'            strMsgErro = "Os pedidos especificam diferentes prazos e/ou condições de pagamento para a mesma forma de pagamento: " & descricao_tipo_parcelamento(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) & "!!" & _
'                        Chr(13) & _
'                        Chr(13) & _
'                        strMsgErro & _
'                        Chr(13) & _
'                        Chr(13) & _
'                        "Não é possível gerar os dados de pagamento para impressão na NFe!!"
'            GoSub GDPP_FECHA_TABELAS
'            Exit Function
'            End If
'        End If
'
'
'  ' HÁ MAIS DO QUE 1 PEDIDO P/ SER PAGO POR BOLETO
'    If intQtdePedidosPagtoBoleto > 1 Then
'        s_where = ""
'        For i = LBound(v_pedido) To UBound(v_pedido)
'            If Trim$(v_pedido(i)) <> "" Then
'                If s_where <> "" Then s_where = s_where & " OR"
'                s_where = s_where & " (pedido='" & Trim$(v_pedido(i)) & "')"
'                End If
'            Next
'
'        s = "SELECT DISTINCT" & _
'                " id_plano_contas_empresa," & _
'                " id_plano_contas_grupo," & _
'                " id_plano_contas_conta," & _
'                " natureza" & _
'            " FROM t_PEDIDO tP" & _
'                " INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" & _
'            " WHERE" & _
'                s_where
'
'        If tAux.State <> adStateClosed Then tAux.Close
'        tAux.Open s, dbc, , , adCmdText
'        intQtdePlanoContas = 0
'        Do While Not tAux.EOF
'            intQtdePlanoContas = intQtdePlanoContas + 1
'            tAux.MoveNext
'            Loop
'
'        If intQtdePlanoContas > 1 Then
'            strMsgErro = "Os pedidos são de lojas que especificam diferentes planos de conta!!" & _
'                        Chr(13) & _
'                        Chr(13) & _
'                        "Não é possível gerar os dados de pagamento para impressão na NFe!!"
'            GoSub GDPP_FECHA_TABELAS
'            Exit Function
'            End If
'        End If
'
'
'  ' HOUVE ALGUM ERRO?
'    If strMsgErro <> "" Then
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'
'  ' OBTÉM O VALOR TOTAL
'  ' ~~~~~~~~~~~~~~~~~~~
'    For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'        With vPedidoCalculoParcelas(i)
'            If Trim$(.pedido) <> "" Then
'                vlTotalPedido = vlTotalPedido + .vlTotalDestePedido
'            '   DADOS DO RATEIO NO CASO DE PAGAMENTO À VISTA
'                If CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
'                    If Trim$("" & v_parcela_pagto(0).strDadosRateio) <> "" Then v_parcela_pagto(0).strDadosRateio = v_parcela_pagto(0).strDadosRateio & "|"
'                    v_parcela_pagto(0).strDadosRateio = v_parcela_pagto(0).strDadosRateio & .pedido & "=" & CStr(.vlTotalDestePedido)
'                    End If
'                End If
'            End With
'        Next
'
'
'  ' CONSISTE VALOR TOTAL C/ A SOMA DOS VALORES DEFINIDOS NA FORMA DE PAGTO
'  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'        With vPedidoCalculoParcelas(i)
'            If CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
'                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pu_valor * .razaoValorPedidoFilhote)
'            ElseIf CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
'                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pce_entrada_valor * .razaoValorPedidoFilhote)
'                vlTotalFormaPagto = vlTotalFormaPagto + CInt(.pce_prestacao_qtde) * arredonda_para_monetario(.pce_prestacao_valor * .razaoValorPedidoFilhote)
'            ElseIf CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
'                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pse_prim_prest_valor * .razaoValorPedidoFilhote)
'                vlTotalFormaPagto = vlTotalFormaPagto + CInt(.pse_demais_prest_qtde) * arredonda_para_monetario(.pse_demais_prest_valor * .razaoValorPedidoFilhote)
'                End If
'            End With
'        Next
'
'    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
'        vlTotalFormaPagto = vlTotalPedido
'        End If
'
'    vlDiferencaArredondamento = vlTotalPedido - vlTotalFormaPagto
'    vlDiferencaArredondamentoRestante = vlDiferencaArredondamento
'
'    If Abs(vlDiferencaArredondamento) > 1 Then
'        strMsgErro = "A soma dos valores definidos na forma de pagamento (" & Format$(vlTotalFormaPagto, FORMATO_MOEDA) & ") não coincide com o valor total do(s) pedido(s) (" & Format$(vlTotalPedido, FORMATO_MOEDA) & ")!!" & _
'                     Chr(13) & _
'                     "Não é possível gerar os dados de pagamento para impressão na NFe!!"
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'  ' CALCULA OS DADOS DAS PARCELAS DOS BOLETOS
'  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  ' LEMBRANDO QUE:
'  '     SE O PRAZO DEFINIDO PARA O 1º BOLETO FOR ATÉ 29 DIAS ENTÃO:
'  '         VENCIMENTO = DATA EM QUE A NF ESTÁ SENDO EMITIDA + 30 DIAS
'  '     SENÃO
'  '         VENCIMENTO = DATA EM QUE A NF ESTÁ SENDO EMITIDA + PRAZO DEFINIDO PELO CLIENTE + 7 DIAS
'
''   À VISTA
''   ~~~~~~~
'    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
'        With v_parcela_pagto(0)
'            .intNumDestaParcela = 1
'            .intNumTotalParcelas = 1
'            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).av_forma_pagto
'            .vlValor = vlTotalPedido
'            .dtVencto = Date + 30
'            End With
'        End If
'
'
''   PARCELA ÚNICA
''   ~~~~~~~~~~~~~
'    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
'        With v_parcela_pagto(0)
'            .intNumDestaParcela = 1
'            .intNumTotalParcelas = 1
'            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pu_forma_pagto
'            .dtVencto = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pu_vencto_apos)
'            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pu_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
'                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
'                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(arredonda_para_monetario(vPedidoCalculoParcelas(i).pu_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote))
'                Next
'            End With
'        End If
'
'
''   PARCELADO COM ENTRADA
''   ~~~~~~~~~~~~~~~~~~~~~
'    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
'      ' ENTRADA
'        With v_parcela_pagto(0)
'            .intNumDestaParcela = 1
'            intQtdeTotalParcelas = 1
'            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada
'            End With
'
'      ' ENTRADA É POR BOLETO?
'        If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'            dtUltimoPagtoCalculado = Date + 30
'        Else
'            dtUltimoPagtoCalculado = Date
'            End If
'
'        With v_parcela_pagto(0)
'            .dtVencto = dtUltimoPagtoCalculado
'            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pce_entrada_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
'                vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(i).pce_entrada_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
'                If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'                    If vlDiferencaArredondamentoRestante <> 0 Then
'                        .vlValor = .vlValor + vlDiferencaArredondamentoRestante
'                        vlRateio = vlRateio + vlDiferencaArredondamentoRestante
'                        vlDiferencaArredondamentoRestante = 0
'                        End If
'                    End If
'                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
'                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(vlRateio)
'                Next
'            End With
'
'      ' PRESTAÇÕES
'        For i = 1 To vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_qtde
'            intQtdeTotalParcelas = intQtdeTotalParcelas + 1
'            If v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela <> 0 Then
'                ReDim Preserve v_parcela_pagto(UBound(v_parcela_pagto) + 1)
'                End If
'
'            With v_parcela_pagto(UBound(v_parcela_pagto))
'                .intNumDestaParcela = intQtdeTotalParcelas
'                .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao
'                End With
'
'        '   PRESTAÇÕES SÃO POR BOLETO?
'            If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'            '   A ENTRADA NÃO FOI PAGA POR BOLETO!
'                If intQtdeTotalParcelas = 1 Then
'                '   ESTA PRESTAÇÃO SERÁ O 1º BOLETO DA SÉRIE
'                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
'                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                    ElseIf CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) <= 29 Then
'                        dtUltimoPagtoCalculado = DateAdd("d", 30, dtUltimoPagtoCalculado)
'                    Else
'                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
'                        End If
'                Else
'                  ' CALCULA A DATA DOS DEMAIS BOLETOS
'                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
'                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                    Else
'                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
'                        End If
'                    End If
'            Else
'            '   CÁLCULO P/ PRESTAÇÕES QUE NÃO SÃO POR BOLETO
'                If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
'                    dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                Else
'                    dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
'                    End If
'                End If
'
'            With v_parcela_pagto(UBound(v_parcela_pagto))
'                .dtVencto = dtUltimoPagtoCalculado
'
'                For j = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'                    .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(j).pce_prestacao_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
'                    vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(j).pce_prestacao_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
'                    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'                        If vlDiferencaArredondamentoRestante <> 0 Then
'                            .vlValor = .vlValor + vlDiferencaArredondamentoRestante
'                            vlRateio = vlRateio + vlDiferencaArredondamentoRestante
'                            vlDiferencaArredondamentoRestante = 0
'                            End If
'                        End If
'                    If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
'                    .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(j).pedido & "=" & CStr(vlRateio)
'                    Next
'                End With
'            Next
'
'        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'            v_parcela_pagto(i).intNumTotalParcelas = intQtdeTotalParcelas
'            Next
'        End If
'
'
''   PARCELADO SEM ENTRADA
''   ~~~~~~~~~~~~~~~~~~~~~
'    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
'    '   1ª PRESTAÇÃO
'        With v_parcela_pagto(0)
'            .intNumDestaParcela = 1
'            intQtdeTotalParcelas = 1
'            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest
'            End With
'
'    '   1ª PRESTAÇÃO É POR BOLETO?
'        If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'            With v_parcela_pagto(0)
'                dtUltimoPagtoCalculado = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos)
'                End With
'        Else
'            dtUltimoPagtoCalculado = DateAdd("d", vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos, Date)
'            End If
'
'        With v_parcela_pagto(0)
'            .dtVencto = dtUltimoPagtoCalculado
'            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pse_prim_prest_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
'                vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(i).pse_prim_prest_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
'                If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'                    If vlDiferencaArredondamentoRestante <> 0 Then
'                        .vlValor = .vlValor + vlDiferencaArredondamentoRestante
'                        vlRateio = vlRateio + vlDiferencaArredondamentoRestante
'                        vlDiferencaArredondamentoRestante = 0
'                        End If
'                    End If
'                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
'                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(vlRateio)
'                Next
'            End With
'
'    '   DEMAIS PRESTAÇÕES
'        For i = 1 To vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_qtde
'            intQtdeTotalParcelas = intQtdeTotalParcelas + 1
'            If v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela <> 0 Then
'                ReDim Preserve v_parcela_pagto(UBound(v_parcela_pagto) + 1)
'                End If
'
'            With v_parcela_pagto(UBound(v_parcela_pagto))
'                .intNumDestaParcela = intQtdeTotalParcelas
'                .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest
'                End With
'
'        '   DEMAIS PRESTAÇÕES SÃO POR BOLETO?
'            If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'            '   A 1ª PRESTAÇÃO NÃO FOI PAGA POR BOLETO!
'                If intQtdeTotalParcelas = 1 Then
'                '   ESTA PRESTAÇÃO SERÁ O 1º BOLETO DA SÉRIE
'                    If (CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos) + _
'                        CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo)) >= 30 Then
'
'                        If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
'                            dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                        Else
'                            dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
'                            End If
'                    Else
'                        dtUltimoPagtoCalculado = DateAdd("d", 30, Date)
'                        End If
'                Else
'                  ' CALCULA A DATA DOS DEMAIS BOLETOS
'                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
'                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                    Else
'                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
'                        End If
'                    End If
'            Else
'            '   CÁLCULO P/ PRESTAÇÕES QUE NÃO SÃO POR BOLETO
'                If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
'                    dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
'                Else
'                    dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
'                    End If
'                End If
'
'            With v_parcela_pagto(UBound(v_parcela_pagto))
'                .dtVencto = dtUltimoPagtoCalculado
'                For j = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
'                    .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(j).pse_demais_prest_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
'                    vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(j).pse_demais_prest_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
'                    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'                        If vlDiferencaArredondamentoRestante <> 0 Then
'                            .vlValor = .vlValor + vlDiferencaArredondamentoRestante
'                            vlRateio = vlRateio + vlDiferencaArredondamentoRestante
'                            vlDiferencaArredondamentoRestante = 0
'                            End If
'                        End If
'                    If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
'                    .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(j).pedido & "=" & CStr(vlRateio)
'                    Next
'                End With
'            Next
'
'        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'            v_parcela_pagto(i).intNumTotalParcelas = intQtdeTotalParcelas
'            Next
'        End If
'
'
'    geraDadosParcelasPagto = True
'
'    GoSub GDPP_FECHA_TABELAS
'
'Exit Function
'
'
'
'
'
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GDPP_TRATA_ERRO:
''===============
'    strMsgErro = CStr(Err) & ": " & Error$(Err)
'    GoSub GDPP_FECHA_TABELAS
'    Exit Function
'
'
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GDPP_FECHA_TABELAS:
''==================
'  ' RECORDSETS
'    bd_desaloca_recordset t_PEDIDO, True
'    bd_desaloca_recordset t_PEDIDO_ITEM, True
'    bd_desaloca_recordset tAux, True
'    Return
'
'
'End Function
'
'Private Function gravaDadosParcelaPagto(ByVal numNF As Long, v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO, ByRef strMsgErro As String) As Boolean
'' __________________________________________________________________________________________
''|
''|  GRAVA AS INFORMAÇÕES DOS BOLETOS NO BANCO DE DADOS
''|
'
'Dim s As String
'Dim s_where As String
'Dim s_pedido_aux As String
'Dim i As Integer
'Dim j As Integer
'Dim intNsuNfParcelaPagto As Long
'Dim intNsuNfParcelaPagtoItem As Long
'Dim intQtdeParcelas As Integer
'Dim intQtdeParcelasBoleto As Integer
'Dim intRecordsAffected As Long
'Dim strIdCliente As String
'Dim v_pedido() As String
'Dim v_pedido_aux() As String
'
'' BANCO DE DADOS
'Dim t As ADODB.Recordset
'
'    On Error GoTo GDPP_TRATA_ERRO
'
'    gravaDadosParcelaPagto = False
'
'    strMsgErro = ""
'
''   TEM DADOS P/ GRAVAR?
'    intQtdeParcelas = 0
'    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'        If v_parcela_pagto(i).intNumDestaParcela > 0 Then
'            intQtdeParcelas = intQtdeParcelas + 1
'            End If
'
'        If CStr(v_parcela_pagto(i).id_forma_pagto) = CStr(ID_FORMA_PAGTO_BOLETO) Then
'            intQtdeParcelasBoleto = intQtdeParcelasBoleto + 1
'            End If
'        Next
'
'    If (intQtdeParcelas = 0) Then
'        gravaDadosParcelaPagto = True
'        Exit Function
'        End If
'
''   RECORDSET
'    Set t = New ADODB.Recordset
'    With t
'        .CursorType = BD_CURSOR_SOMENTE_LEITURA
'        .LockType = BD_POLITICA_LOCKING
'        .CacheSize = BD_CACHE_CONSULTA
'        End With
'
''   OBTÉM IDENTIFICAÇÃO DO CLIENTE
''   LEMBRANDO QUE GARANTIDAMENTE TODOS OS PEDIDOS SÃO DO MESMO CLIENTE
'    v_pedido = Split(v_parcela_pagto(UBound(v_parcela_pagto)).strDadosRateio, "|")
'    v_pedido_aux = Split(v_pedido(LBound(v_pedido)), "=")
'    s_pedido_aux = Trim$(v_pedido_aux(LBound(v_pedido_aux)))
'
'    s = "SELECT" & _
'            " c.id" & _
'        " FROM t_PEDIDO p" & _
'            " INNER JOIN t_CLIENTE c" & _
'                " ON p.id_cliente=c.id" & _
'        " WHERE" & _
'            " p.pedido = '" & s_pedido_aux & "'"
'    If t.State <> adStateClosed Then t.Close
'    t.Open s, dbc, , , adCmdText
'    If Not t.EOF Then
'        strIdCliente = Trim$("" & t("id"))
'    Else
'        strMsgErro = "Falha ao tentar localizar a identificação do cliente!!"
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'
''   GRAVA REGISTRO PRINCIPAL
''   ~~~~~~~~~~~~~~~
'    dbc.BeginTrans
''   ~~~~~~~~~~~~~~~
''   SE HOUVER DADOS DE PARCELAS CADASTRADOS ANTERIORMENTE NO STATUS INICIAL P/ ESTE(S) PEDIDO(S),
''   CANCELA-OS ANTES DE CADASTRAR OS NOVOS DADOS
'    s_where = ""
'    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'        With v_parcela_pagto(i)
'            If .intNumDestaParcela <> 0 Then
'                v_pedido = Split(.strDadosRateio, "|")
'                For j = LBound(v_pedido) To UBound(v_pedido)
'                    If Trim$(v_pedido(j)) <> "" Then
'                        v_pedido_aux = Split(v_pedido(j), "=")
'                        s_pedido_aux = Trim$(v_pedido_aux(LBound(v_pedido_aux)))
'                        If s_pedido_aux <> "" Then
'                            If InStr(s_where, s_pedido_aux) = 0 Then
'                                If s_where <> "" Then s_where = s_where & " OR"
'                                s_where = s_where & " (pedido='" & Trim$(v_pedido_aux(LBound(v_pedido_aux))) & "')"
'                                End If
'                            End If
'                        End If
'                    Next
'                End If
'            End With
'        Next
'
'    If s_where <> "" Then
'        s = "SELECT DISTINCT" & _
'                " tpp.id" & _
'            " FROM t_FIN_NF_PARCELA_PAGTO tpp" & _
'                " INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tppi" & _
'                    " ON (tpp.id=tppi.id_nf_parcela_pagto)" & _
'                " INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tppir" & _
'                    " ON (tppi.id=tppir.id_nf_parcela_pagto_item)" & _
'            " WHERE" & _
'                " (tpp.status = " & NF_PARCELA_PAGTO__STATUS_INICIAL & ")" & _
'                " AND (" & s_where & ")"
'        If t.State <> adStateClosed Then t.Close
'        t.Open s, dbc, , , adCmdText
'        Do While Not t.EOF
'            s = "UPDATE" & _
'                    " t_FIN_NF_PARCELA_PAGTO" & _
'                " SET" & _
'                    " status = " & NF_PARCELA_PAGTO__STATUS_CANCELADO & _
'                " WHERE" & _
'                    " (id = " & t("id") & ")" & _
'                    " AND (status = " & NF_PARCELA_PAGTO__STATUS_INICIAL & ")"
'            Call dbc.Execute(s, intRecordsAffected)
'            If intRecordsAffected = 0 Then
'                strMsgErro = "Falha ao tentar cancelar registros anteriores dos dados de pagamento do(s) pedido(s) especificado(s)!!"
'            '   ~~~~~~~~~~~~~~~~~
'                dbc.RollbackTrans
'            '   ~~~~~~~~~~~~~~~~~
'                GoSub GDPP_FECHA_TABELAS
'                Exit Function
'                End If
'            t.MoveNext
'            Loop
'        End If
'
''   OBTÉM NSU
'    If Not geraNsu(NSU_T_FIN_NF_PARCELA_PAGTO, intNsuNfParcelaPagto, strMsgErro) Then
'        If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
'        strMsgErro = "Falha ao gravar os dados de pagamento!!" & strMsgErro
'    '   ~~~~~~~~~~~~~~~~~
'        dbc.RollbackTrans
'    '   ~~~~~~~~~~~~~~~~~
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
'    On Error GoTo GDPP_TRATA_ERRO_TRANSACAO
''   LEMBRANDO QUE DT_CADASTRO, DT_ULT_ATUALIZACAO E STATUS SÃO INSERIDOS VIA DEFAULT DAS COLUNAS
'    s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO (" & _
'            "id," & _
'            "id_cliente," & _
'            "numero_NF," & _
'            "qtde_parcelas," & _
'            "qtde_parcelas_boleto," & _
'            "usuario_cadastro," & _
'            "usuario_ult_atualizacao" & _
'        ") VALUES (" & _
'            CStr(intNsuNfParcelaPagto) & "," & _
'            "'" & strIdCliente & "'," & _
'            CStr(numNF) & "," & _
'            CStr(intQtdeParcelas) & "," & _
'            CStr(intQtdeParcelasBoleto) & "," & _
'            "'" & Trim$(usuario.id) & "'," & _
'            "'" & Trim$(usuario.id) & "'" & _
'        ")"
'    Call dbc.Execute(s, intRecordsAffected)
'    If intRecordsAffected = 0 Then
'        strMsgErro = "Falha ao tentar inserir registro principal dos dados de pagamento!!"
'    '   ~~~~~~~~~~~~~~~~~
'        dbc.RollbackTrans
'    '   ~~~~~~~~~~~~~~~~~
'        GoSub GDPP_FECHA_TABELAS
'        Exit Function
'        End If
'
''   GRAVA REGISTRO DAS PARCELAS
'    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
'    '   OBTÉM NSU
'        If Not geraNsu(NSU_T_FIN_NF_PARCELA_PAGTO_ITEM, intNsuNfParcelaPagtoItem, strMsgErro) Then
'            If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
'            strMsgErro = "Falha ao gravar os dados de pagamento!!" & strMsgErro
'        '   ~~~~~~~~~~~~~~~~~
'            dbc.RollbackTrans
'        '   ~~~~~~~~~~~~~~~~~
'            GoSub GDPP_FECHA_TABELAS
'            Exit Function
'            End If
'
'        With v_parcela_pagto(i)
'            If .intNumDestaParcela <> 0 Then
'                s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO_ITEM (" & _
'                        "id," & _
'                        "id_nf_parcela_pagto," & _
'                        "num_parcela," & _
'                        "forma_pagto," & _
'                        "dt_vencto," & _
'                        "valor" & _
'                    ") VALUES (" & _
'                        CStr(intNsuNfParcelaPagtoItem) & "," & _
'                        CStr(intNsuNfParcelaPagto) & "," & _
'                        CStr(.intNumDestaParcela) & "," & _
'                        CStr(.id_forma_pagto) & "," & _
'                        sqlMontaDateParaSqlDateTime(.dtVencto) & "," & _
'                        sqlFormataDecimal(.vlValor) & _
'                    ")"
'                Call dbc.Execute(s, intRecordsAffected)
'                If intRecordsAffected = 0 Then
'                    strMsgErro = "Falha ao tentar inserir registro da parcela " & .intNumDestaParcela & "!!"
'                '   ~~~~~~~~~~~~~~~~~
'                    dbc.RollbackTrans
'                '   ~~~~~~~~~~~~~~~~~
'                    GoSub GDPP_FECHA_TABELAS
'                    Exit Function
'                    End If
'
'                v_pedido = Split(.strDadosRateio, "|")
'                For j = LBound(v_pedido) To UBound(v_pedido)
'                    If Trim$(v_pedido(j)) <> "" Then
'                        v_pedido_aux = Split(v_pedido(j), "=")
'                        s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO (" & _
'                                "id_nf_parcela_pagto_item," & _
'                                "pedido," & _
'                                "id_nf_parcela_pagto," & _
'                                "valor" & _
'                            ") VALUES (" & _
'                                CStr(intNsuNfParcelaPagtoItem) & "," & _
'                                "'" & Trim$(v_pedido_aux(LBound(v_pedido_aux))) & "'," & _
'                                CStr(intNsuNfParcelaPagto) & "," & _
'                                sqlFormataDecimal(CCur(Trim$(v_pedido_aux(UBound(v_pedido_aux))))) & _
'                            ")"
'                        Call dbc.Execute(s, intRecordsAffected)
'                        If intRecordsAffected = 0 Then
'                            strMsgErro = "Falha ao tentar inserir registro do rateio da parcela " & .intNumDestaParcela & "!!"
'                        '   ~~~~~~~~~~~~~~~~~
'                            dbc.RollbackTrans
'                        '   ~~~~~~~~~~~~~~~~~
'                            GoSub GDPP_FECHA_TABELAS
'                            Exit Function
'                            End If
'                        End If
'                    Next
'                End If
'            End With
'        Next
'
''   ~~~~~~~~~~~~~~~
'    dbc.CommitTrans
''   ~~~~~~~~~~~~~~~
'    On Error GoTo GDPP_TRATA_ERRO
'
'    gravaDadosParcelaPagto = True
'
'    GoSub GDPP_FECHA_TABELAS
'
'Exit Function
'
'
'
'
'
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GDPP_TRATA_ERRO:
''===============
'    strMsgErro = CStr(Err) & ": " & Error$(Err)
'    GoSub GDPP_FECHA_TABELAS
'    Exit Function
'
'
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GDPP_TRATA_ERRO_TRANSACAO:
''=========================
'    strMsgErro = CStr(Err) & ": " & Error$(Err)
'    On Error Resume Next
'    dbc.RollbackTrans
'    GoSub GDPP_FECHA_TABELAS
'    Exit Function
'
'
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GDPP_FECHA_TABELAS:
''==================
'  ' RECORDSETS
'    bd_desaloca_recordset t, True
'    Return
'
'End Function

Sub pedido_preenche_dados_tela(ByVal pedido As String)
Dim s_resp As String
Dim s_end_entrega As String
Dim s_end_entrega_uf As String
Dim s_end_cliente_uf As String
Dim s_NFe_texto_constar As String
Dim strIE As String
Dim s_erro As String
Dim pos_inicio_cnpj_cpf As Integer
Dim s_possivel_cpf As String
Dim s As String

    c_pedido = pedido
    
'   MUDOU O NÚMERO DO PEDIDO?
    If pedido_anterior = Trim$(pedido) Then Exit Sub
    pedido_anterior = Trim$(pedido)
    
    If (Trim$(pedido) = "") And blnFilaSolicitacoesEmissaoNFeEmTratamento Then trata_botao_fila_pause
                
'   EXIBE OS ITENS DO PEDIDO NA TELA
    formulario_exibe_itens_pedido Trim$(pedido)
           
    c_info_pedido = ""
    strIE = ""
    If pedido <> "" Then
        'verificar se os dados do cliente devem vir da memorização no pedido
        If param_pedidomemorizacaoenderecos.campo_inteiro = 1 Then
            If obtem_info_pedido_memorizada(pedido, s_resp, s_end_entrega, s_end_entrega_uf, s_end_cliente_uf, s_NFe_texto_constar, strIE, s_erro) Then
                c_info_pedido = s_resp
                c_dados_adicionais = s_NFe_texto_constar
            ElseIf s_erro <> "" Then
                aviso_erro s_erro
                End If
        Else
            If obtem_info_pedido(pedido, s_resp, s_end_entrega, s_end_entrega_uf, s_end_cliente_uf, s_NFe_texto_constar, strIE, s_erro) Then
                c_info_pedido = s_resp
                c_dados_adicionais = s_NFe_texto_constar
            ElseIf s_erro <> "" Then
                aviso_erro s_erro
                End If
            End If
        End If
    l_IE.Caption = strIE
        
    atualiza_tela_qtde_fila_solicitacoes_emissao_NFe

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
        
    'se a UF de entrega for diferente da UF do cliente e for pessoa física, emitir aviso
    If (s_end_entrega_uf <> "") And (s_end_entrega_uf <> s_end_cliente_uf) Then
        pos_inicio_cnpj_cpf = InStr(s_resp, "CNPJ/CPF: ")
        If (pos_inicio_cnpj_cpf > 0) Then
            'Para evitar mexer nas rotinas obtem_info_pedido_memorizada e obtem_info_pedido,
            'vamos analisar se existe um CPF na variável de retorno s_resp.
            'Para isto, pegaremos a substring com as 14 posições posteriores à string 'CNPJ/CPF: '
            'Se esta substring estiver no formato XXX.XXX.XXX/XX e for um CPF válido, emitiremos
            'o alerta de endereço de entrega diferente, para que o usuário receba este aviso antes
            'mesmo de iniciar a emissão da NF
            s_possivel_cpf = Mid(s_resp, pos_inicio_cnpj_cpf + Len("CNPJ/CPF: "), 14)
            If Mid(s_possivel_cpf, 4, 1) = "." And _
                Mid(s_possivel_cpf, 8, 1) = "." And _
                Mid(s_possivel_cpf, 12, 1) = "/" And _
                cnpj_cpf_ok(retorna_so_digitos(s_possivel_cpf)) Then
                s = "ATENÇÃO!!" & vbCrLf & _
                    "Pedido de PESSOA FÍSICA com endereço de entrega localizado em outra UF!!" & vbCrLf & _
                    vbCrLf & _
                    s_resp
                aviso s
                End If
            End If
        End If

End Sub

Sub preenche_prox_pedido_fila_solicitacoes_emissao_NFe()
' CONSTANTES
Const NomeDestaRotina = "preenche_prox_pedido_fila_solicitacoes_emissao_NFe()"
' DECLARAÇÕES
Dim s As String
Dim strPedido As String
Dim lngRecordsAffected As Long
Dim int_st_end_entrega As Integer
Dim intQtdeTentativas As Integer
Dim blnBuscarNovoPedido As Boolean
Dim lngId As Long
Dim strUsuario As String
Dim s_erro As String
Dim s_cliente_uf As String
Dim s_entrega_uf As String
Dim s_tipo_pessoa As String
' BANCO DE DADOS
Dim t As ADODB.Recordset

    On Error GoTo PPPFSENFE_TRATA_ERRO
    
    'se houver nota triangular em emissão, interromper processo
    If blnNotaTriangularAtiva Then
        If NFeExisteNotaTriangularEmEmissao(lngId, strUsuario, s_erro) Then
            trata_botao_fila_pause
            aviso "Nota triangular em emissão pelo usuário " & strUsuario & ", aguarde!!!"
            Exit Sub
            End If
        End If
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    blnBuscarNovoPedido = True
    Do While blnBuscarNovoPedido
        blnBuscarNovoPedido = False
        Do While True
            strPedido = ""
            intQtdeTentativas = intQtdeTentativas + 1
            If param_pedidomemorizacaoenderecos.campo_inteiro = 1 Then
                s = "SELECT TOP 10" & _
                        " tPNES.id," & _
                        " tPNES.pedido," & _
                        " tP.st_end_entrega," & _
                        " tP.EndEtg_uf," & _
                        " tP.endereco_tipo_pessoa as tipo_pessoa," & _
                        " tP.endereco_uf AS cli_uf" & _
                    " FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA tPNES" & _
                        " INNER JOIN t_PEDIDO tP ON (tP.pedido=tPNES.pedido)" & _
                    " WHERE" & _
                        " (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
                        " AND (Len(Coalesce(tP.transportadora_id,'')) > 0)" & _
                        " AND (tP.st_entrega <> '" & Trim(CStr(ST_ENTREGA_CANCELADO)) & "')" & _
                        " AND (tP.id_nfe_emitente = " & usuario.emit_id & ")" & _
                        " AND (" & _
                            "(ult_requisicao_fila_data_hora IS NULL)" & _
                            " OR " & _
                            "(DateDiff(ss, ult_requisicao_fila_data_hora, getdate()) >= " & MAX_TIMEOUT_REGISTRO_REQUISITADO_FILA_EM_SEG & ")" & _
                            ")"
                            
                s = s & _
                    " ORDER BY" & _
                        " id"
            
            Else
                s = "SELECT TOP 10" & _
                        " tPNES.id," & _
                        " tPNES.pedido," & _
                        " tP.st_end_entrega," & _
                        " tP.EndEtg_uf," & _
                        " tC.tipo as tipo_pessoa," & _
                        " tC.uf AS cli_uf" & _
                    " FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA tPNES" & _
                        " INNER JOIN t_PEDIDO tP ON (tP.pedido=tPNES.pedido)" & _
                        " INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
                    " WHERE" & _
                        " (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
                        " AND (Len(Coalesce(tP.transportadora_id,'')) > 0)" & _
                        " AND (tP.st_entrega <> '" & Trim(CStr(ST_ENTREGA_CANCELADO)) & "')" & _
                        " AND (tP.id_nfe_emitente = " & usuario.emit_id & ")" & _
                        " AND (" & _
                            "(ult_requisicao_fila_data_hora IS NULL)" & _
                            " OR " & _
                            "(DateDiff(ss, ult_requisicao_fila_data_hora, getdate()) >= " & MAX_TIMEOUT_REGISTRO_REQUISITADO_FILA_EM_SEG & ")" & _
                            ")"
                            
                s = s & _
                    " ORDER BY" & _
                        " id"
                End If
            If t.State <> adStateClosed Then t.Close
            t.Open s, dbc, , , adCmdText
            If t.EOF Then Exit Do
                
            strPedido = Trim$("" & t("pedido"))
            int_st_end_entrega = CLng(t("st_end_entrega"))
            s_entrega_uf = Trim$("" & t("EndEtg_uf"))
            s_cliente_uf = Trim$("" & t("cli_uf"))
            s_tipo_pessoa = Trim$("" & t("tipo_pessoa"))
            s = "UPDATE t_PEDIDO_NFe_EMISSAO_SOLICITADA SET" & _
                    " ult_requisicao_fila_data_hora = getdate()," & _
                    " ult_requisicao_fila_usuario = '" & usuario.id & "'" & _
                " WHERE" & _
                    " (id = " & t("id") & ")" & _
                    " AND (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
                    " AND (" & _
                        "(ult_requisicao_fila_data_hora IS NULL)" & _
                        " OR " & _
                        "(DateDiff(ss, ult_requisicao_fila_data_hora, getdate()) >= " & MAX_TIMEOUT_REGISTRO_REQUISITADO_FILA_EM_SEG & ")" & _
                        ")"
            dbc.Execute s, lngRecordsAffected
            If lngRecordsAffected = 1 Then
                Exit Do
            Else
                strPedido = ""
                End If
            
            If intQtdeTentativas >= 5 Then Exit Do
            Loop
                    
        'para abrir a tela de operação triangular: lhgx
        '- o parâmetro deve estar habilitado
        '- deve haver endereço de entrega
        '- a UF de entrega deve ser diferente da UF do cliente
        If blnNotaTriangularAtiva And _
           (strPedido <> "") And _
           (s_tipo_pessoa = "PJ") And _
           (int_st_end_entrega <> 0) And _
           (s_entrega_uf <> "") And _
           (s_cliente_uf <> s_entrega_uf) Then
            sPedidoTriangular = strPedido
            If c_pedido_danfe <> "" Then
                sPedidoDANFETelaAnterior = c_pedido_danfe
                End If
            aguarde INFO_NORMAL, m_id
            f_EMISSAO_NFE_TRIANGULAR.Show vbModal, Me
            'o número do pedido será enviado à variável global sPedidoTriangular
            'se houver o tratamento do pedido no painel de emissão triangular, a variavél retornará vazia
            'se não houver, o tratamento deve ser feito no form principal
            If sPedidoTriangular = strPedido Then
                trata_botao_fila_pause
                pedido_preenche_dados_tela strPedido
            Else
                atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
                blnBuscarNovoPedido = True
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
        Else
            pedido_preenche_dados_tela strPedido
            End If
        Loop
    
    If strPedido = "" Then trata_botao_fila_pause
    
    GoSub PPPFSENFE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPPFSENFE_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    aviso_erro s
    GoSub PPPFSENFE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    Exit Sub
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPPFSENFE_FECHA_TABELAS:
'=======================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Sub

Sub recalcula_totais()
Dim i As Integer
Dim vl_total_outras_despesas_acessorias As Currency
Dim blnTemDados As Boolean

    For i = c_fabricante.LBound To c_fabricante.UBound
        If Trim$(c_produto(i)) <> "" Then
            If Trim$(c_vl_outras_despesas_acessorias(i)) <> "" Then
                blnTemDados = True
                vl_total_outras_despesas_acessorias = vl_total_outras_despesas_acessorias + converte_para_currency(c_vl_outras_despesas_acessorias(i))
                End If
            End If
        Next
    
    If blnTemDados Then
        c_vl_total_outras_despesas_acessorias = formata_moeda(vl_total_outras_despesas_acessorias)
    Else
        c_vl_total_outras_despesas_acessorias = ""
        End If
    
End Sub

Sub tab_stop_configura()
Dim i As Integer

    b_dummy.TabIndex = 0
    c_info_pedido.TabIndex = 0
    b_fila_play.TabIndex = 0
    b_fila_pause.TabIndex = 0
    b_fila_remove.TabIndex = 0
    b_danfe.TabIndex = 0
    c_pedido_danfe.TabIndex = 0
    b_emissao_nfe_complementar.TabIndex = 0
    b_emissao_manual.TabIndex = 0
    b_fechar.TabIndex = 0
    b_emite_numeracao_manual.TabIndex = 0
    b_imprime.TabIndex = 0
    c_dados_adicionais.TabIndex = 0
    c_vl_total_outras_despesas_acessorias.TabIndex = 0
    c_vl_total_geral.TabIndex = 0
    c_vl_total_icms.TabIndex = 0
    c_total_volumes.TabIndex = 0
    
    For i = c_produto.UBound To c_produto.LBound Step -1
        c_fcp(i).TabIndex = 0
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
    cb_frete.TabIndex = 0
    cb_natureza.TabIndex = 0
    c_ipi.TabIndex = 0
    cb_icms.TabIndex = 0
    cb_tipo_NF.TabIndex = 0
    c_pedido.TabIndex = 0
    
End Sub

Sub trata_botao_fila_pause()

    On Error Resume Next
    
    b_fila_play.Enabled = True
    b_fila_pause.Enabled = False
    b_fila_play.SetFocus
    
    blnFilaSolicitacoesEmissaoNFeEmTratamento = False

End Sub


Sub trata_botao_fila_play()

    b_fila_pause.Enabled = True
    b_fila_play.Enabled = False
    b_fila_pause.SetFocus
    
    blnFilaSolicitacoesEmissaoNFeEmTratamento = True
    preenche_prox_pedido_fila_solicitacoes_emissao_NFe

End Sub


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
    
    atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
    
    GoSub TBFR_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
    pedido_preenche_dados_tela ""
    aviso "Pedido " & c_pedido & " removido com sucesso da fila de solicitações de emissão de NFe!!"
    
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
            
        vl_icms_calc = vl_icms_calc * perc_ICMS_calc / 100
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

Sub carrega_CFOPs_sem_partilha()

Dim msg_erro As String

    If le_arquivo_REMESSA_CFOP(vCFOPsSemPartilha, msg_erro) Then
        If msg_erro <> "" Then aviso "Arquivo de CFOP's sem partilha não carregado, será utilizada lista pré-existente"
        End If
    
End Sub


Sub carrega_UFs_inscricao_virtual()

Dim msg_erro As String

    If Not le_UFs_INSCRICAO_VIRTUAL(vCUFsInscricaoVirtual, msg_erro) Then
        If msg_erro <> "" Then aviso "Lista de UF's com inscrição virtual não carregada!!!"
        End If
    
End Sub

Private Sub ajusta_visualizacoes_emitente()
Dim i As Integer
Dim sAliquotaEmit As String
Dim msg_erro As String
Dim aliquota_icms As String

    pnInfoFilaPedido.Caption = "Emitente - " & usuario.emit

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
Private Sub atualiza_fila_emitente()
Dim qtdEmits As Integer

    If Not imgFilasEmits.Visible Then Exit Sub
    
'   SELEÇÃO DO EMITENTE A SER UTILIZADO
    If obtem_emitentes_usuario(usuario.id, vEmitsUsuario, qtdEmits) Then
        If qtdEmits = 1 Then
            usuario.emit = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, 1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 5)
            usuario.emit_uf = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 2, 2)
            usuario.emit_id = vEmitsUsuario(UBound(vEmitsUsuario)).c2
            txtFixoEspecifico = vEmitsUsuario(UBound(vEmitsUsuario)).c3
        Else
            f_CD.Show vbModal, Me
            End If
    Else
        aviso_erro "Nenhum Emitente habilitado para o usuário!!"
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If
        
'   EXIBIR INFORMAÕES DO EMITENTE SELECIONADO
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ajusta_visualizacoes_emitente
                
    atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
    
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

Private Sub b_emissao_manual_Click()
Dim emit_anterior As String

    emit_anterior = usuario.emit
    
    exibe_form_emissao_manual
    
    If usuario.emit <> emit_anterior Then
        ajusta_visualizacoes_emitente
        atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
        End If
    
End Sub

Private Sub b_emissao_nfe_complementar_Click()

    exibe_form_emissao_nfe_complementar
    
End Sub


Private Sub b_emite_numeracao_manual_Click()

    NFe_emite True
    
End Sub

Private Sub b_fechar_Click()

'   ---- DESATIVADO HORÁRIO DE VERÃO
'    If Not configura_registry_usuario_horario_verao(intHorarioVeraoAtivoInicio, chk_HorVerao.Value) Then
'        aviso "Não foi possível gravar as configurações de horário de verão no sistema!"
'        End If
   
    If Not configura_registry_usuario_info_parcelas(intInfoAdicParcInicio, chk_InfoAdicParc.Value) Then
        aviso "Não foi possível gravar as configurações de informações adicionais de parcelas!"
        End If
   
   '~~~
    End
   '~~~
    
End Sub


Private Sub b_fila_pause_Click()

    trata_botao_fila_pause
    
End Sub

Private Sub b_fila_play_Click()

    trata_botao_fila_play
    
End Sub

Private Sub b_fila_remove_Click()

    trata_botao_fila_remove
    
End Sub

Private Sub b_imprime_Click()
Dim lngId As Long
Dim strUsuario As String
Dim msg_erro As String

    If blnNotaTriangularAtiva Then
        If NFeExisteNotaTriangularEmEmissao(lngId, strUsuario, msg_erro) Then
            aviso "Nota triangular sendo emitida pelo usuário " & strUsuario & ", aguarde!!!"
            Exit Sub
            End If
    Else
        If msg_erro <> "" Then aviso msg_erro
        End If

    NFe_emite False
    
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
Dim s_end_entrega As String
Dim strMsgErro As String

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

Private Sub c_fcp_GotFocus(Index As Integer)
    
    With c_fcp(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_fcp_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = 11) Then
            sstItens.Tab = 1
            c_fabricante(Index + 1).SetFocus
        ElseIf (Index = 23) Then
            sstItens.Tab = 2
            c_fabricante(Index + 1).SetFocus
        ElseIf (Index < 35) Then
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


Private Sub c_nItemPed_GotFocus(Index As Integer)
   
   With c_nItemPed(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
 
End Sub

Private Sub c_nItemPed_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = c_fabricante.UBound Then
            c_dados_adicionais.SetFocus
        Else
            If Trim$(c_produto(Index + 1)) = "" Then
                c_dados_adicionais.SetFocus
            Else
                c_produto_obs(Index + 1).SetFocus
                End If
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

Private Sub c_nItemPed_LostFocus(Index As Integer)
    
    c_nItemPed(Index) = Trim$(c_nItemPed(Index))
        
End Sub

Private Sub c_pedido_danfe_KeyPress(KeyAscii As Integer)

Dim executa_tab As Boolean
Dim s As String
Dim c As String

    If KeyAscii = 13 Then
    '   COMO O CAMPO ACEITA MÚLTIPLAS LINHAS, SÓ VAI P/ O PRÓXIMO CAMPO APÓS 2 "ENTER's" CONSECUTIVOS
        executa_tab = True
    '   CURSOR ESTÁ NO FINAL DO TEXTO (IGNORA "ENTER's" SUBSEQUENTES NO TEXTO) ?
        s = Mid$(c_pedido_danfe.Text, c_pedido_danfe.SelStart + 1)
        s = Replace$(s, vbCr, "")
        s = Replace$(s, vbLf, "")
        s = Trim$(s)
        If s <> "" Then executa_tab = False
    '   CARACTER ANTERIOR É "ENTER" ?
        If c_pedido_danfe.SelStart > 0 Then
            c = Mid$(c_pedido_danfe.Text, c_pedido_danfe.SelStart, 1)
            If (c <> Chr$(13)) And (c <> Chr$(10)) Then executa_tab = False
            End If
        
        If executa_tab Then
            KeyAscii = 0
            b_danfe.SetFocus
            End If
        
        If Not c_pedido_danfe.MultiLine Then
            c_pedido_danfe = normaliza_num_pedido(c_pedido_danfe)
            If Len(c_pedido_danfe) > 0 Then c_pedido_danfe.SelStart = Len(c_pedido_danfe)
            DANFE_CONSULTA_parametro_emitente Trim$(c_pedido_danfe)
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


Private Sub c_pedido_Danfe_LostFocus()

Dim s As String
Dim i As Integer
Dim j As Integer
Dim v() As String
Dim v_pedido() As String

    c_pedido_danfe = Trim$(c_pedido_danfe)
    
    s = normaliza_lista_pedidos(c_pedido_danfe)
    If s <> "" Then c_pedido_danfe = s
    
    'CONSISTÊNCIA
    ReDim v_pedido(0)
    v_pedido(UBound(v_pedido)) = ""
    v = Split(c_pedido_danfe, vbCrLf)
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
        '   REPETIDO ?
            For j = LBound(v_pedido) To UBound(v_pedido)
                If Trim$(v(i)) = v_pedido(j) Then
                    aviso_erro "Pedido " & Trim$(v(i)) & " está repetido na lista !!"
                    c_pedido_danfe.SetFocus
                    Exit Sub
                    End If
                Next
                
            If v_pedido(UBound(v_pedido)) <> "" Then ReDim Preserve v_pedido(UBound(v_pedido) + 1)
            v_pedido(UBound(v_pedido)) = Trim$(v(i))
            End If
        Next

End Sub


Private Sub c_pedido_GotFocus()

    With c_pedido
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

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
Dim s_erro As String

    c_pedido = Trim$(c_pedido)
    
    s = normaliza_num_pedido(c_pedido)
    If s <> "" Then
        c_pedido = s
        If Not pedido_eh_do_emitente_atual(c_pedido) Then Exit Sub
        End If
    
    pedido_preenche_dados_tela c_pedido
    
'    'verificar se existe informação de parcelas em boleto
'    If (param_geracaoboletos.campo_texto = "Manual") Then
'        If s <> "" Then
'            ReDim v_pedido_manual_boleto(0)
'            v_pedido_manual_boleto(UBound(v_pedido_manual_boleto)) = s
'            blnExisteParcelamentoBoleto = False
'            pnParcelasEmBoletos.Visible = False
'            If geraDadosParcelasPagto(v_pedido_manual_boleto(), v_parcela_manual_boleto(), s_erro) Then
'                AdicionaListaParcelasEmBoletos v_parcela_manual_boleto()
'                End If
'            End If
'        End If
        
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
        c_vl_outras_despesas_acessorias(Index).SetFocus
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


Private Sub c_total_volumes_GotFocus()

    With c_total_volumes
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
        
End Sub

Private Sub c_total_volumes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_dados_adicionais.SetFocus
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



Private Sub c_unidade_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
'        If Index = c_fabricante.UBound Then
'            c_pedido_nota.SetFocus
'        Else
'            c_fabricante(Index + 1).SetFocus
'            End If
        c_fcp(Index).SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
'   Filtra caracter separador definido pela Target One
    If Chr(KeyAscii) = "|" Then KeyAscii = 0


End Sub

Private Sub c_valorparc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        b_parc_edicao_ok.SetFocus
        Exit Sub
        End If
        
    If KeyAscii = 8 Then Exit Sub ' BACKSPACE
    If KeyAscii = 1 Then Exit Sub ' CTRL+A
    If KeyAscii = 3 Then Exit Sub ' CTRL+C
    If KeyAscii = 22 Then Exit Sub ' CTRL+V
    
    KeyAscii = filtra_moeda(KeyAscii)
    If KeyAscii = 0 Then Exit Sub

End Sub

Private Sub c_valorparc_LostFocus()

    If Trim(c_valorparc) = "" Then Exit Sub
    
    c_valorparc = formata_moeda(c_valorparc)

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


Private Sub cb_frete_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        c_dados_adicionais.SetFocus
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


Private Sub cb_natureza_Click()
    ' Se o código de natureza da operação inicia com 1 ou 5, trata-se de uma operação interna;
    ' se o código de natureza da operação inicia com 2 ou 6, trata-se de uma operação interestadual
    Dim digito As String
    Dim s_cfop As String
    
    digito = left(Trim(cb_natureza.Text), 1)
    If (digito = "1") Or (digito = "5") Then cb_loc_dest.ListIndex = 0
    If (digito = "2") Or (digito = "6") Then cb_loc_dest.ListIndex = 1
    
    s_cfop = left(Trim(cb_natureza.Text), 5)
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
        cb_natureza.SetFocus
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
        c_dados_adicionais.SetFocus
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


Private Sub chk_HorVerao_Click()

'   ---- DESATIVADO HORÁRIO DE VERÃO
'    If chk_HorVerao.Value = 1 Then
'        blnHorarioVerao = True
'    Else
'        blnHorarioVerao = False
'        End If

    blnHorarioVerao = False

End Sub

Private Sub chk_InfoAdicParc_Click()
    
    If chk_InfoAdicParc.Value = 1 Then
        blnInfoAdicParc = True
    Else
        blnInfoAdicParc = False
        End If
    
End Sub

Private Sub Form_Activate()
Const NomeDestaRotina = "Form_Activate()"
Dim s As String
Dim msg_erro As String
Dim qtdEmits As Integer
Dim i As Integer
Dim sAliquotaEmit As String
Dim cor_inicial As String

    On Error GoTo FORMACTIVATE_TRATA_ERRO

    If Not modulo_inicializacao_ok Then
        
      ' OK !!
        modulo_inicializacao_ok = True
        
        tab_stop_configura
        
        relogio_Timer
        
        aguarde INFO_EXECUTANDO, "iniciando aplicativo"
        
      ' CONFIGURAÇÃO REGIONAL ESTÁ OK ?
        If Not verifica_configuracao_regional() Then
            s = "Há parâmetros da configuração regional que NÃO estão de acordo com as necessidades deste programa !!" & _
                vbCrLf & "Deseja que esses parâmetros sejam corrigidos agora ?"
            If Not confirma(s) Then
                aviso_erro "O programa não pode prosseguir enquanto a configuração regional não for corrigida !!"
               '~~~
                End
               '~~~
                End If
          
            If verifica_configuracao_regional(True) Then
                s = "A configuração regional foi alterada com sucesso !!" & _
                    vbCrLf & "O programa será encerrado agora e deve ser executado novamente para que possa operar corretamente !!"
                alerta s
            Else
                s = "Não foi possível alterar a configuração regional automaticamente !!" & _
                    vbCrLf & "Execute este programa novamente para tentar outra vez ou então faça a configuração manualmente !!"
                alerta s
                End If
                  
           '~~~
            End
           '~~~
            End If
        
      ' CONFIGURA PARÂMETROS DO CLIENT DO SQL SERVER NO REGISTRY DO WINDOWS
        If Not configura_registry_client_sql_server(msg_erro) Then
            s = "Falha ao configurar acesso do cliente do banco de dados !!" & _
                vbCrLf & "Não é possível continuar !!"
            If msg_erro <> "" Then s = s & vbCrLf & vbCrLf & msg_erro
            aviso_erro s
           '~~~
            End
           '~~~
            End If
                    
      ' LÊ PARÂMETROS P/ CONEXÃO AO BD
        If Not le_arquivo_ini(msg_erro) Then
            s = "Falha ao ler arquivo de configuração !!" & _
                vbCrLf & "Não é possível continuar !!"
            If msg_erro <> "" Then s = s & vbCrLf & vbCrLf & msg_erro
            aviso_erro s
           '~~~
            End
           '~~~
            End If
            
        
    '   PREPARA CAMPOS/CARREGA DADOS INICIAIS
        formulario_inicia
    
'   REPOSICIONANDO A LIMPEZA DO FORMULÁRIO PARA DEPOIS DE CONECTAR COM O BD,
'   PARA POSSIBILITAR O CARREGAMENTO DAS ALÍQUOTAS DE ICMS
'    '   LIMPA CAMPOS/POSICIONA DEFAULTS
'        formulario_limpa
        
        b_fila_pause.Enabled = False
        
        c_pedido_danfe = ""
        
    '   TODO
        If Not DESENVOLVIMENTO Then mnu_emissao_nfe_complementar.Enabled = False
        
        Caption = Caption & " v" & m_id_versao
        
        If DESENVOLVIMENTO Then
            Caption = Caption & "  (Versão Exclusiva de Desenvolvimento/Homologação)"
            aviso "Versão apenas para testes de desenvolvimento/homologação!!"
            End If
        
      ' SELECIONA O BD
        aguarde INFO_NORMAL, m_id
        If Trim$(bd_selecionado.NOME_BD) = "" Then
            aviso_erro "Não há informações suficientes para conectar ao banco de dados!!"
           '~~~
            End
           '~~~
            End If
        
      ' INICIA BD
        aguarde INFO_EXECUTANDO, "conectando ao banco de dados"
        If Not BD_inicia() Then
            s = "Falha ao conectar com o Banco de Dados!!" & _
                vbCrLf & "Não é possível continuar!!"
            aviso_erro s
           '~~~
            End
           '~~~
            End If
        
        If Not BD_CEP_inicia() Then
            s = "Falha ao conectar com o banco de dados de CEP!!" & _
                vbCrLf & "Não é possível continuar!!"
            aviso_erro s
           '~~~
            End
           '~~~
            End If
            
    '   CARREGA AS UF's POR CNPJ QUE POSSUEM INSCRIÇÃO VIRTUAL
    '   (rotina está aqui e não no Form_Load devido à necessidade de consulta ao BD
        carrega_UFs_inscricao_virtual
            
    '   LIMPA CAMPOS/POSICIONA DEFAULTS
        formulario_limpa
        
        
    '   PARÂMETROS DA T_VERSAO
        obtem_parametros_t_versao cor_fundo_padrao, identificador_ambiente_padrao
        If cor_fundo_padrao <> "" Then
            cor_fundo_padrao = converte_cor_Web2VB(cor_fundo_padrao)
            Me.BackColor = cor_fundo_padrao
            ' SE A COR DE FUNDO DO BANCO DE DADOS É DIFERENTE, GRAVAR NO REGISTRY
            If cor_fundo_padrao <> cor_inicial Then
                If Not configura_registry_usuario_cor_fundo_padrao(converte_cor_VB2Web(cor_fundo_padrao)) Then
                    aviso "Não foi possível gravar as configurações de cor de fundo para futuros acessos!"
                    End If
                End If
        Else
            ' SE EXISTIR UMA COR DE FUNDO GRAVADA NO REGISTRY, UTILIZAR
            If le_registry_usuario_cor_fundo_padrao(cor_inicial) Then
                cor_inicial = converte_cor_Web2VB(cor_inicial)
                Me.BackColor = cor_inicial
                End If
            End If
    
    '   OBTER O PARÂMETRO DA OPERAÇÃO TRIANGULAR
        get_registro_t_parametro "NF_FlagOperacaoTriangular", param_notatriangular
        blnNotaTriangularAtiva = param_notatriangular.campo_inteiro = 1
        b_emissao_nfe_triangular.Visible = blnNotaTriangularAtiva
        b_emissao_nfe_triangular.Enabled = blnNotaTriangularAtiva
        
    '   LOGIN
        aguarde INFO_NORMAL, m_id
        f_LOGIN.Show vbModal, Me
        Set painel_ativo = Me
        
    '   NÍVEL DE ACESSO
        If Not usuario.perfil_acesso_ok Then
            aviso_erro "ACESSO NEGADO !!" & vbCrLf & "Você não possui o perfil de acesso necessário !!"
          ' ENCERRA O PROGRAMA
            BD_Fecha
            BD_CEP_Fecha
           '~~~
            End
           '~~~
            End If
            
    '   OBTER O PARÂMETRO DA GERAÇÃO DE BOLETOS
        get_registro_t_parametro "NF_GeracaoBoletos", param_geracaoboletos
            
    '   OBTER O PARÂMETRO DA ATUALIZAÇÃO DO NÚMERO DA NF NO PEDIDO PARA EMISSÕES MANUAIS
        get_registro_t_parametro "NF_FlagAtualizaNFnoPedido", param_atualizanfnopedido
    
    '   OBTER O PARÂMETRO DA MEMORIZAÇÃO DOS ENDEREÇOS NA T_PEDIDO
        get_registro_t_parametro "Flag_Pedido_MemorizacaoCompletaEnderecos", param_pedidomemorizacaoenderecos
        
    '   OBTER O PARÂMETRO DO ENDEREÇO DE ENTREGA NA NOTA FISCAL
        get_registro_t_parametro "NF_MemorizacaoUsarEnderecoEntrega", param_nfmemooendentrega
    
    '   OBTER O PARÂMETRO DO INTERMEDIADOR
        get_registro_t_parametro "NF_IntermediadorAtivo", param_nfintermediador
    
    '   OBTER O PARÂMETRO DO TIPO DE PAGAMENTO
        get_registro_t_parametro "NF_TipoPagamentoObrigatorio", param_nftipopag
    
    '   OBTER O PARÂMETRO DO INTERMEDIADOR
        get_registro_t_parametro "NF_BC_Pis_Cofins_Excluir_ICMS", param_bc_pis_cofins_icms
    
    '   OBTER O PARÂMETRO DO TIPO DE PAGAMENTO
        get_registro_t_parametro "NF_BC_Pis_Cofins_Excluir_DIFAL", param_bc_pis_cofins_difal
        
    '   OBTER O PARÂMETRO DA CONTINGÊNCIA PARA O MEIO DE PAGAMENTO CARTÃO DE CRÉDITO
        get_registro_t_parametro "NF_Contingencia_MeioPagamento_Cartao", param_contingencia_meio_pagamento_cartao
    
    '   OBTER O PARÂMETRO DA CONTINGÊNCIA PARA QUALQUER MEIO DE PAGAMENTO
        get_registro_t_parametro "NF_Contingencia_MeioPagamento_Geral", param_contingencia_meio_pagamento_geral
        
    '   OBTER O PARÂMETRO DO RESPONSÁVEL TÉCNICO DO SOFTWARE EMISSOR
        get_registro_t_parametro "NF_Informa_Resp_Tec", param_nfinformaresptec
    
    '   SELEÇÃO DO EMITENTE A SER UTILIZADO
        If obtem_emitentes_usuario(usuario.id, vEmitsUsuario, qtdEmits) Then
            If qtdEmits = 1 Then
                usuario.emit = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, 1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 5)
                usuario.emit_uf = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 2, 2)
                usuario.emit_id = vEmitsUsuario(UBound(vEmitsUsuario)).c2
                txtFixoEspecifico = vEmitsUsuario(UBound(vEmitsUsuario)).c3
            Else
                f_CD.Show vbModal, Me
                End If
        Else
            aviso_erro "Nenhum Emitente habilitado para o usuário!!"
          ' ENCERRA O PROGRAMA
            BD_Fecha
            BD_CEP_Fecha
           '~~~
            End
           '~~~
            End If
            
    '   EXIBIR EMITENTE SELECIONADO NA FILA DE IMPRESSÃO
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        pnInfoFilaPedido.Caption = "Emitente - " & usuario.emit
    
    '   EXIBIR UF DO EMITENTE SELECIONADO NO LABEL EM DESTAQUE
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        l_emitente_uf.Caption = usuario.emit_uf
                
        l_IE.Caption = ""
        
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

            
'   ---- DESATIVADO HORÁRIO DE VERÃO
        blnHorarioVerao = False
        intHorarioVeraoAtivo = 0
        chk_HorVerao.Value = 0
    
'    '   HORÁRIO DE VERÃO
'        blnHorarioVerao = False
'        'TENTA LER OS PARÂMETROS GRAVADOS NO REGISTRY
'        If Not le_registry_usuario_horario_verao(intHorarioVeraoAtivo, sHorarioVeraoData) Then
'            intHorarioVeraoAtivo = 0
'            sHorarioVeraoData = ""
'            End If
'        intHorarioVeraoAtivoInicio = intHorarioVeraoAtivo
'
'        If ((Date >= InicioHorarioVerao(Year(Date))) And (Date <= TerminoHorarioVerao(Year(Date)))) Or ((Date >= _
'            InicioHorarioVerao(Year(Date) - 1)) And (Date <= TerminoHorarioVerao(Year(Date) - 1))) Then
'
'            'DETECTADO HORÁRIO DE VERÃO
'            blnHorarioVerao = True
'
'            If intHorarioVeraoAtivo = 0 Then
'                If sHorarioVeraoData = "" Then
'                    '1 - SE A OPÇÃO NUNCA FOI GRAVADA NO REGISTRY, ATIVAR AUTOMATICAMENTE
'                    intHorarioVeraoAtivo = 1
'                    blnHorarioVerao = True
'                    aviso "Horário de verão detectado. Caso necessário, desmarcar a opção Horário de Verão!"
'                    Else
'                        '2 - SE A OPÇÃO DE HORÁRIO DE VERÃO ATIVO ESTIVER DESMARCADA HÁ MAIS DE 07 DIAS,
'                        '    PERGUNTAR SE DESEJA ATIVAR
'                        If Now - CDate(sHorarioVeraoData) > 7 Then
'                            If confirma("Horário de verão detectado. deseja ativar a opção Horário de Verão?") Then
'                                intHorarioVeraoAtivo = 1
'                                blnHorarioVerao = True
'                                End If
'                            End If
'                    End If
'                End If
'            Else
'            'NÃO FOI DETECTADO HORÁRIO DE VERÃO
'                If intHorarioVeraoAtivo = 1 Then
'                        'SE A OPÇÃO DE HORÁRIO DE VERÃO ATIVO ESTIVER MARCADA HÁ MAIS DE 07 DIAS,
'                        'PERGUNTAR SE DESEJA ATIVAR
'                        If Now - CDate(sHorarioVeraoData) > 7 Then
'                            If confirma("Horário de verão não detectado. deseja desativar a opção Horário de Verão?") Then
'                                intHorarioVeraoAtivo = 0
'                                blnHorarioVerao = False
'                                End If
'                            End If
'                    End If
'            End If
'
'        chk_HorVerao.Value = intHorarioVeraoAtivo
'
        'INFORMAÇÕES ADICIONAIS DE PARCELAS
        blnInfoAdicParc = False
        If Not le_registry_usuario_info_parcelas(intInfoAdicParc) Then
            intInfoAdicParc = 0
            End If
        intInfoAdicParcInicio = intInfoAdicParc
        
        chk_InfoAdicParc.Value = intInfoAdicParc
        blnInfoAdicParc = (intInfoAdicParc = 1)


        'AVISO PENDÊNCIAS COM OPERAÇÕES TRIANGULARES, SE HOUVER
        If blnNotaTriangularAtiva Then
            sAvisosAExibir = RetornaOperacoesTriangularesPendentes
            If sAvisosAExibir <> "" Then sAvisosAExibir = sAvisosAExibir & vbCrLf & vbCrLf
            sAvisosAExibir = sAvisosAExibir & RetornaNumeracaoRemessaPendente
            If sAvisosAExibir <> "" Then
                f_AVISOS.Show vbModal, Me
                End If
            End If

        atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
        
        aguarde INFO_NORMAL, m_id
        End If
        
    'formulario_limpa
            
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
    
    lblQtdeFilaSolicitacoesEmissaoNFe = ""
    
    ScaleMode = vbPixels
    
    CriaListaParcelasEmBoletos
    carrega_CFOPs_sem_partilha
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

'   EM EXECUÇÃO ?
    If em_execucao Then
        Cancel = True
        Exit Sub
        End If

'   FECHA BANCO DE DADOS
    BD_Fecha
    BD_CEP_Fecha
    BD_Assist_Fecha
    End
    
End Sub

Private Sub imgFilasEmits_Click()

    atualiza_fila_emitente
    
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


Private Sub mnu_emissao_manual_Click()

    exibe_form_emissao_manual
    
End Sub

Private Sub mnu_emissao_nfe_complementar_Click()
    
    exibe_form_emissao_nfe_complementar
    
End Sub


Private Sub mnu_FECHAR_Click()

    Unload Me
    
End Sub

Private Sub pnInfoFilaPedido_DblClick()

    atualiza_fila_emitente
    
End Sub

Private Sub relogio_Timer()
Dim s As String
Dim n As Long

    s = left$(Time$, 5)
    If Val(right$(Time$, 1)) Mod 2 Then Mid$(s, 3, 1) = " "
    agora = s

    hoje = Format$(Date, "dd/mm/yyyy")
    
    If dt_hr_ult_atualizacao_qtde_fila_solicitacoes_emissao_NFe > 0 Then
        n = DateDiff("s", dt_hr_ult_atualizacao_qtde_fila_solicitacoes_emissao_NFe, Now)
        If n >= (1 * 60) Then atualiza_tela_qtde_fila_solicitacoes_emissao_NFe
        End If
        
End Sub


