<%
'======================================================================================
'*******   SITE: ARTVEN3 (BONSHOP) OU ARTVEN (F�BRICA)   *******
'======================================================================================
	Const ID_PARAM_SITE = "ArtBS"
'	Const ID_PARAM_SITE = "AssTec"
'	Const ID_PARAM_SITE = "ArtFab"
	
	
	Const IdFormTargetArtvenBonshop = "fArtBS"
	Const IdFormTargetArtvenFabricante = "fArtFab"
	Const IdFormTargetAssistenciaTecnica = "fAssTec"



'======================================================================================
'*******   HOR�RIO DA MANUTEN��O   *******
'======================================================================================
'	HOR�RIO DE MANUTEN��O DO SISTEMA
	Const TRATAR_HORARIO_MANUTENCAO_SISTEMA = False
	Const HORARIO_INICIO_MANUTENCAO_SISTEMA = "20:45"
	Const HORARIO_TERMINO_MANUTENCAO_SISTEMA = "21:45"

'	HOR�RIO DE REBOOT DO SERVIDOR
	Const TRATAR_HORARIO_REBOOT_SERVIDOR = True
	Const HORARIO_INICIO_REBOOT_SERVIDOR = "04:30"
	Const HORARIO_TERMINO_REBOOT_SERVIDOR = "05:15"



'======================================================================================
'*******   PAR�METROS   *******
'======================================================================================
'	IDENTIFICA��O DO AMBIENTE: 1=OLD01, 2=Assist�ncia T�cnica, 3=DIS
	Const ID_AMBIENTE = 3

	Const ID_AMBIENTE__OLD01 = 1
	Const ID_AMBIENTE__AT = 2
	Const ID_AMBIENTE__DIS = 3

	Const COD_FC_AMBIENTE = "L"

	Const SERVIDOR_BD = "Win2019DB,29981"
	Const NOME_BD = "Artven2DIS"
	Const USUARIO_BD = "artven2"
	Const SCHEMA_BD = "dbo"
	Const SENHA_BD = "000000007bbddfed3b85e97d47b91b"

'	DADOS P/ CONEX�O USADOS PELO RELAT�RIO DE PEDIDOS INDICADORES AO GRAVAR OS LAN�AMENTOS NO FLUXO DE CAIXA
	Const SERVIDOR_BD_RPIFC = "Win2019DB,29981"
	Const NOME_BD_RPIFC = "Artven2DIS_ModOrcto1"
	Const USUARIO_BD_RPIFC = "orcto_lab1"
	Const SENHA_BD_RPIFC = "000000007bbddfed3b85e97d47b91b"

	Const SERVIDOR_BD_AT = "Win2019DB,29981"
	Const NOME_BD_AT = "AssistenciaTecnica"
	Const USUARIO_BD_AT = "assistencia_tecnica"
	Const SENHA_BD_AT = "000000007bbddfed3b85e97d47b91b"

	Const SERVIDOR_BD_BS = "Win2019DB,29981"
	Const NOME_BD_BS = "Artven2DIS"
	Const USUARIO_BD_BS = "artven2"
	Const SENHA_BD_BS = "000000007bbddfed3b85e97d47b91b"

'	Const SERVIDOR_BD_CEP = "Ora9iServer"
	Const SERVIDOR_BD_CEP = "Win2019DB,29981"
	Const NOME_BD_CEP = "CEP_Homologacao"
	Const USUARIO_BD_CEP = "CEP_user"
	Const SENHA_BD_CEP = "000000007bbddfed3b85e97d47b91b"



	Const FATOR_BD = "1209"
	
	Const VISANET_SHOPID = ""
	Const VISANET_CFGLOJA = "OLD01"
	Const VISANET_CFGLOJA_BS = "OLD01ClienteVisaBS"

	Const DIR_VISANET_COMPONENTES = "C:\Projetos\Sites\VisaNet\componentes_vbv"
	Const URL_VISANET_COMPONENTES = "http://182.168.2.110/componentes_vbv"

'	CIELO (NOVA PLATAFORMA QUE SUBSTITUI A VISANET E SE BASEIA EM WEB SERVICES)
'	CIELO__NUMERO_CIELO = N� DA AFILIA��O
	Const CIELO__NUMERO_CIELO = ""
	Const CIELO__CHAVE_CIELO = ""
	Const CIELO_WEB_SERVICE_ENDERECO = "https://qasecommerce.cielo.com.br/servicos/ecommwsec.do"
	
'	CADASTRAR EMPRESA + CD (WMS) NA ENTRADA DE MERCADORIAS NO ESTOQUE
	Const CADASTRAR_WMS_CD_ENTRADA_ESTOQUE = True

'	BRASPAG + CLEARSALE
	Const USAR_BRASPAG_CLEARSALE = True
	
'	BRASPAG + CYBERSOURCE
	Const BRASPAG_OWNER_OLD01 = 1
	Const BRASPAG_OWNER_OLD02 = 3
	Const BRASPAG_OWNER_DIS = 4
	
	Const CLEARSALE_DF_KEY = "kxbqf6bnvf"
	Const CLEARSALE_DF_APP = "98a91b6eec" ' "fe71a6a4d1" ' ARCLUBE=fe71a6a4d1  BONSHOP=98a91b6eec

'	Const BRASPAG_DIS_PAG_MERCHANT_ID = "BBE2B085-858B-46B4-999D-D1D675525231" 'USADO NA TRANSA��O DO PAGADOR
	Const BRASPAG_DIS_PAG_MERCHANT_ID = "7A2F5804-A541-421A-B73B-5C5384D41CC0" 'AMBIENTE SANDBOX
	Const BRASPAG_DIS_AF_MERCHANT_ID = "62D7C0A7-505B-E311-93FD-0026B939D54B" 'USADO NA TRANSA��O DO ANTIFRAUDE
	Const BRASPAG_DIS_DF_MERCHANT_ID = "braspag_bonshop" ' USADO NO DEVICE FINGERPRINT
	Const BRASPAG_DIS_ORG_ID = "1snn5n9w"

	Const BRASPAG_OLD01_PAG_MERCHANT_ID = "" 'USADO NA TRANSA��O DO PAGADOR
	Const BRASPAG_OLD01_AF_MERCHANT_ID = "" 'USADO NA TRANSA��O DO ANTIFRAUDE
	Const BRASPAG_OLD01_DF_MERCHANT_ID = "" ' USADO NO DEVICE FINGERPRINT
	Const BRASPAG_OLD01_ORG_ID = ""
	
	Const BRASPAG_OLD02_PAG_MERCHANT_ID = "" 'USADO NA TRANSA��O DO PAGADOR
	Const BRASPAG_OLD02_AF_MERCHANT_ID = "" 'USADO NA TRANSA��O DO ANTIFRAUDE
	Const BRASPAG_OLD02_DF_MERCHANT_ID = "" ' USADO NO DEVICE FINGERPRINT
	Const BRASPAG_OLD02_ORG_ID = ""
	
	Const BRASPAG_ANTIFRAUDE_VERSION = "1.1"
	Const BRASPAG_PAGADOR_VERSION = "1.0"
	
	Const BRASPAG_AMBIENTE_HOMOLOGACAO = True
	Const BRASPAG_WS_ENDERECO_AF_PAG_COMBO_TRANSACTION = "https://homologacao.braspag.com.br/AntiFraudeWS/antifraud.asmx"
	Const BRASPAG_WS_ENDERECO_PAGADOR_QUERY = "https://transactionsandbox.pagador.com.br/services/pagadorQuery.asmx"
	Const BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION = "https://transactionsandbox.pagador.com.br/webservice/pagadortransaction.asmx"
	Const BRASPAG_WS_ENDERECO_ANTIFRAUDE_QUERY = "https://homologacao.braspag.com.br/AntiFraudeWS/antifraud.asmx"
	Const BRASPAG_WS_ENDERECO_ANTIFRAUDE_TRANSACTION = "https://homologacao.braspag.com.br/AntiFraudeWS/antifraud.asmx"
	
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__VISA = True
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__MASTERCARD = True
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__AMEX = True
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__ELO = True
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__HIPERCARD = True
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__DINERS = False
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__DISCOVER = False
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__AURA = False
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__JCB = False
	Const BRASPAG_OLD01_BANDEIRA_HABILITADA__CELULAR = False
	
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__VISA = True
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__MASTERCARD = True
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__AMEX = True
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__ELO = True
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__HIPERCARD = True
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__DINERS = False
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__DISCOVER = False
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__AURA = False
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__JCB = False
	Const BRASPAG_OLD02_BANDEIRA_HABILITADA__CELULAR = False
	
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__VISA = True
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__MASTERCARD = True
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__AMEX = True
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__ELO = True
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__HIPERCARD = True
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__DINERS = False
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__DISCOVER = False
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__AURA = False
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__JCB = False
	Const BRASPAG_DIS_BANDEIRA_HABILITADA__CELULAR = False


'	CIELO (NOVA PLATAFORMA QUE SUBSTITUI A VISANET E SE BASEIA EM WEB SERVICES)
	Const CIELO_VERSAO_TRANSACAO = "1.2.1"
	Const CIELO_XML_ENCODING = "ISO-8859-1"
	
	Const FLAG_CIELO__EXIBIR_MSG_AVISO_TRANSACOES_NAO_AUTORIZADAS = False

	Const CIELO_BANDEIRA_HABILITADA__VISA = True
	Const CIELO_BANDEIRA_HABILITADA__MASTERCARD = True
	Const CIELO_BANDEIRA_HABILITADA__AMEX = False
	Const CIELO_BANDEIRA_HABILITADA__ELO = True
	Const CIELO_BANDEIRA_HABILITADA__DINERS = True
	Const CIELO_BANDEIRA_HABILITADA__DISCOVER = True
	Const CIELO_BANDEIRA_HABILITADA__AURA = False
	Const CIELO_BANDEIRA_HABILITADA__JCB = False
	Const CIELO_BANDEIRA_HABILITADA__CELULAR = True
	
	Const CIELO_LOGOTIPO_LOJA = "EMPRESA_med.gif"

	Const BRASPAG_LOGOTIPO_LOJA = "logo_DIS_med.png"

	Const SITE_PARCEIRO_HEADER__LOGOTIPO = "centraldepedidos.gif"
	Const SITE_PARCEIRO_HEADER__ALT_IMG_TEXT = "Bonshop"
	
	Const SITE_PARCEIRO_HEADER_OLD03__LOGOTIPO = "logo_OLD03.png"
	Const SITE_PARCEIRO_HEADER_OLD03__ALT_IMG_TEXT = "OLD03"
	
	Const SITE_CLIENTE_HEADER__LOGOTIPO = "centraldepedidos.gif"
	Const SITE_CLIENTE_HEADER__ALT_IMG_TEXT = "Bonshop"
	
	Const SITE_CLIENTE_TITULO_JANELA = "Pagamento"
	Const SITE_CLIENTE_URL_BASE = "https://web.bonshop.com.br/pagamento/ClienteCartao/"
	Const SITE_CLIENTE_EXIBIR_LOGO_SSL = True
	Const SITE_CLIENTE_USAR_PROTOCOLO_HTTPS = True
	Const URL_FILE__E_LOGO_TOP_BS_CSS = "../Global/eLogoTopBS.css?v=003"
	Const URL_FILE__E_LOGO_TOP_OLD03_CSS = "../Global/eLogoTopOLD03.css?v=002"
	
'	URL PARA ACESSO NA P�GINA DE LOGIN DO PARCEIRO (P�GINA C/ LOGIN DUPLO)
	Const URL_BASE_RELATIVA_SITE_ARTVEN3 = "/art3HotFix5"
	Const URL_BASE_RELATIVA_SITE_ASSISTENCIA_TECNICA = "/assistenciatecnica"
	
'	Const URL_SITE_CLIENTE_PAGTO_CIELO = "http://web.bonshop.com.br/pagamento/clientecartao/default.asp"
	Const URL_SITE_CLIENTE_PAGTO_CIELO = "http://localhost/art3/clientecartao/default.asp"

	Const NUMERO_LOJA_ECOMMERCE_AR_CLUBE = "201"
	Const NUMERO_LOJA_OLD03 = "300"
	Const NUMERO_LOJA_OLD03_BONIFICACAO = "301"
	Const NUMERO_LOJA_OLD03_ASSISTENCIA = "302"
	Const NUMERO_LOJA_MARCELO_ARTVEN = "305"
	Const NUMERO_LOJA_TRANSFERENCIA = "01"
	Const NUMERO_LOJA_KITS = "02"
	Const NUMERO_LOJA_INDICACAO_OBRA = "299"

	Const FLAG_PEDIDO_CHAMADO_EMAIL_USAR_REPLY_TO = True

	Const PREFIXO_PEDIDO_MAGENTO_AR_CLUBE = "1"
	Const PREFIXO_PEDIDO_MAGENTO_BONSHOP = "2"
	Const TAMANHO_PEDIDO_MAGENTO = 9

	Const OP_PESQ_PEDIDO_MAGENTO_BONSHOP = "BSMAGE"
	Const OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE = "ARCLUBE"
	Const OP_PESQ_PEDIDO_MARKETPLACE_AR_CLUBE = "MARKETPLACE"


'	HTTPS_OBRIGATORIO_HABILITADO
'		False = permite o usu�rio usar o protocolo de sua prefer�ncia (HTTP ou HTTPS)
'		True = o usu�rio � obrigado a usar o protocolo HTTPS, portanto, caso o acesso tenha sido via HTTP, ser� realizado um redirecionamento autom�tico para HTTPS
'
'	HTTPS_OBRIGATORIO_URL_REDIRECT
'		Este par�metro � usado quando HTTPS_OBRIGATORIO_HABILITADO = True
'		Quando se define uma URL espec�fica para redirecionamento, se houver a necessidade de usar uma porta personalizada, o n�mero da porta deve ser especificada diretamente no par�metro HTTPS_OBRIGATORIO_URL_REDIRECT
'		O valor deste par�metro corresponde � URL base, se houver conte�do adicional do endere�o informado pelo usu�rio e obtido atrav�s de ServerVariables("URL"), esse conte�do ser� adicionado automaticamente
'			pela rotina de tratamento do redirecionamento.
'		Poss�veis valores:
'			String vazia = o redirecionamento ser� feito para a mesma URL acessada pelo usu�rio, apenas alterando de HTTP para HTTPS
'			URL especificada = o redirecionamento ser� feito para a URL especificada no par�metro
	Const HTTPS_OBRIGATORIO_HABILITADO = False
	Const HTTPS_OBRIGATORIO_URL_REDIRECT = ""


'	TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO
'		Se habilitado (True), nas rotinas em que o acesso concorrente causa problemas s�rios, � executado um Update em um campo dummy nos registros pertinentes para obter o lock exclusivo
	Const TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = True



'======================================================================================
'*******   GERAL   *******
'======================================================================================
	Dim IS_SGBD_ACCESS
	IS_SGBD_ACCESS = (Instr(Ucase(NOME_BD), ".MDB") <> 0)

%>
