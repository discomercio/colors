﻿#region [ using ]
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Drawing;
using System.Configuration;
#endregion

namespace ADM2
{
	public class Global
	{
		#region [ Constantes ]
		public class Cte
		{
			#region[ Versão do Aplicativo ]
			public class Aplicativo
			{
				public const string NOME_OWNER = "Artven";
				public const string NOME_SISTEMA = "ADM2";
				public const string VERSAO_NUMERO = "1.13";
				public const string VERSAO_DATA = "30.JAN.2023";
				public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
				public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
				public const string M_DESCRICAO = "Módulo Administrativo";
				public const string ID_APLICATIVO_CTRL_VERSAO = "ADM2";
			}
			#endregion

			#region[ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 01.06.2013 - por HHO
			 *        Início.
			 *        Este programa realiza rotinas de manutenção.
			 *        A versão inicial foi desenvolvida para realizar a carga dos dados do IBPT que são
			 *        usados para calcular o valor estimado total dos tributos exibido na NFe.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - 11.06.2013 - por HHO
			 *		  Alterações no painel de carga do arquivo do IBPT devido ao novo layout do arquivo a
			 *		  partir da versão 0.0.2
			 *		  Foi adicionado o campo 'descricao' entre os campos 'tabela' e 'aliqNac'.
			 *		  Além disso, há um novo tipo de registro: LC 116
			 * -----------------------------------------------------------------------------------------------
			 * v 1.02 - 23.01.2015 - por HHO
			 *		  Alteração dos dados de conexão ao BD devido à migração do servidor, pois o SQL Server
			 *		  não está mais usando a porta padrão por questões de segurança.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.03 - 18.04.2016 - por HHO
			 *		  Implementação de ajustes para alterar a cor de fundo dos painéis de acordo com o
			 *		  ambiente acessado.
			 *		  A cor inicialmente é obtida a partir do arquivo de configuração e, após realizar a
			 *		  conexão com o banco de dados, a cor é obtida através do campo 'cor_fundo_padrao' da
			 *		  tabela t_VERSAO. Caso a cor definida no banco de dados seja diferente da do arquivo,
			 *		  o parâmetro do arquivo é atualizado para respeitar a cor especificada no BD.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.04 - 31.07.2018 - por HHO
			 *		  Implementação de funcionalidade para atualizar dados na planilha do estoque: quantidade
			 *		  disponível no estoque de venda e custo médio de aquisição.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.05 - 18.01.2019 - por HHO
			 *		  Implementação de ajustes na atualização da planilha do estoque:
			 *			1) Ignorar o campo que indica se o produto está indisponível para a loja.
			 *			2) Se o estoque estiver zerado, zerar o valor do custo médio na planilha
			 *			3) Calcular e preencher o preço de lista na planilha
			 * -----------------------------------------------------------------------------------------------
			 * v 1.06 - 02.03.2020 - por HHO
			 *		  Desenvolvimento de rotina para processar arquivo CSV da SSW para cadastrar a data de
			 *		  recebimento do pedido pelo cliente (campo: PedidoRecebidoData). No caso de pedidos de
			 *		  marketplace, também é atualizado o campo MarketplacePedidoRecebidoRegistrarDataRecebido.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.07 - 07.03.2020 - por HHO
			 *		  Ajustes na rotina de processamento do arquivo CSV que cadastra a data de recebimento
			 *		  do pedido pelo cliente para verificar a existência de registros repetidos, pois nesse
			 *		  caso o processamento do registro repetido iria gerar um erro devido a data já estar
			 *		  preenchida no banco de dados. Além disso, foi implementado um tratamento para as
			 *		  situações em que o CPF do cliente vem formatado como sendo um CNPJ (ex: 123.456.789/09
			 *		  sendo informado como 00.012.345/6789-09)
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - 10.03.2020 - por HHO
			 *		  Correção de bug na rotina que tenta localizar o pedido através do número da NF.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - 10.03.2020 - por HHO
			 *		  Ajustes na rotina que registra o recebimento do pedido pelo cliente para tratar o caso
			 *		  em que a entrega é cancelada e a mercadoria é retornada.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - 22.04.2020 - por HHO
			 *		  Ajustes na rotina que registra o recebimento do pedido pelo cliente para tratar a
			 *		  eventual situação de não encontrar a nota fiscal no banco de dados.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10(B) - 20.08.2020 - por HHO
			 *		  Ajustes na rotina que registra o recebimento do pedido pelo cliente para tratar a
			 *		  eventual situação do número de série da NF ser informado com dado inválido no arquivo.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - 18.11.2021 - por HHO
			 *		  Implementação de tratamento para a data de previsão de entrega informada pela trans-
			 *		  portadora no mesmo arquivo CSV usado para processar o recebimento do pedido pelo
			 *		  cliente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11(B) - 02.12.2021 - por HHO
			 *		  Correção de bug na rotina calculaTimeSpanMiliSegundos() e calculaTimeSpanSegundos()
			 *		  para que os cálculos fossem realizados com datatype 'long' e não com 'int'.
			 *		  Ao realizar os cálculos com 'int', poderia ocorrer overflow, como de fato estava
			 *		  ocorrendo ao calcular o tempo decorrido entre DateTime.Now - dtHrUltProgresso,
			 *		  pois dtHrUltProgresso estava inicializada com DateTime.MinValue.
			 *		  Ao ocorrer o overflow, o resultado ficava negativo e a rotina que controlava a exibição
			 *		  do progresso do processamento não funcionava devido à verificação do intervalo entre
			 *		  atualizações (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - 17.12.2021 - por HHO
			 *		  Ajustes na rotina que registra o recebimento do pedido pelo cliente para tratar a
			 *		  situação em que é informada a "data entrega", mas o evento não é "MERCADORIA ENTREGUE".
			 *		  Criação de parâmetro no banco de dados para definir os códigos aceitos no campo
			 *		  "Situacao" do arquivo CSV de rastreio da transportadora para considerar que a mercadoria
			 *		  foi entregue ao destinatário (tabela t_PARAMETRO, registro
			 *		  "ADM2_RastreioPedidoRecebidoCliente_SituacaoMercadoriaEntregue"). Esse tratamento foi
			 *		  desenvolvido para tratar a situação "MERCADORIA PRE-ENTREGUE (MOBILE)" que também
			 *		  representa que a mercadoria foi entregue ao destinatário. Mas nesse caso, o campo
			 *		  "Data Entrega" não está sendo informado, sendo necessário usar a data da ocorrência
			 *		  como a data da entrega.
			 *		  Implementação de tratamento para que o módulo aceite múltiplas versões permitidas
			 *		  na tabela t_VERSAO durante a validação da versão na inicialização do módulo (caso
			 *		  exista mais de uma versão permitida, deve estar separada pelo caractere pipe "|").
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - 30.01.2023 - por HHO
			 *		  Implementação de funcionalidade para registrar o frete no pedido a partir de arquivo
			 *		  CSV.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.15 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.18 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.19 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.20 - XX.XX.20XX - por XXX
			 *		  
			 * ===============================================================================================
			 */
			#endregion

			#region[ Data/Hora ]
			public class DataHora
			{
				public const string FmtDia = "dd";
				public const string FmtDiaAbreviado = "ddd";
				public const string FmtDiaExtenso = "dddd";
				public const string FmtMes = "MM";
				public const string FmtMesAbreviado = "MMM";
				public const string FmtMesExtenso = "MMMM";
				public const string FmtAno = "yyyy";
				public const string FmtAnoCom2Digitos = "yy";
				public const string FmtHora = "HH";
				public const string FmtMin = "mm";
				public const string FmtSeg = "ss";
				public const string FmtMiliSeg = "fff";
				public const string FmtYYYYMMDD = FmtAno + FmtMes + FmtDia;
				public const string FmtHHMMSS = FmtHora + FmtMin + FmtSeg;
				public const string FmtHhMmComSeparador = FmtHora + ":" + FmtMin;
				public const string FmtHhMmSsComSeparador = FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtDdMmYyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAnoCom2Digitos;
				public const string FmtDdMmYyyyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno;
				public const string FmtDdMmYyyyHhMmComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin;
				public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
				public const string FmtYyyyMmDdHhMmSsComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
			}
			#endregion

			#region [ Etc ]
			public class Etc
			{
				public const String SIMBOLO_MONETARIO = "R$";
				public const byte FLAG_NAO_SETADO = 255;
				public const int TAM_MIN_PRODUTO = 6;
				public const int TAM_MIN_FABRICANTE = 3;
				public const int TAM_MIN_LOJA = 2;
				public const int TAM_MIN_NUM_PEDIDO = 6;    // SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
				public const int TAM_MIN_ID_PEDIDO = 7; // PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
				public const char COD_SEPARADOR_FILHOTE = '-';
				public const int MAX_TAM_BOLETO_CAMPO_ENDERECO = 40;
				public const int MAX_TAM_BOLETO_CAMPO_NOME_SACADO = 40;
				public const String ID_PF = "PF";
				public const String ID_PJ = "PJ";
				public const int TAMANHO_CPF = 11;
				public const int TAMANHO_CNPJ = 14;
				public const int TAMANHO_RAIZ_CNPJ = 8;
				public const String PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE = "TFBI";
				public const String SQL_COLLATE_CASE_ACCENT = " COLLATE Latin1_General_CI_AI";
			}
			#endregion

			#region [ Log ]
			public class LogAtividade
			{
				public static string PathLogAtividade = Application.StartupPath + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 365;
				public const string ExtensaoArqLog = "LOG";
			}
			#endregion

			#region [ Classe ADM2 ]
			public class ADM2
			{
				#region [ LogOperacao - Códigos de operação para o log ]
				public class LogOperacao
				{
					// Texto com 20 posições
					public const String LOGON = "ADM2-Logon";
					public const String LOGOFF = "ADM2-Logoff";
					public const String RECONEXAO_BD = "ADM2-Reconexao-BD";
					public const String IBPT_CARGA_ARQUIVO_CSV = "ADM2-IbptCargaArqCsv";
					public const String IBPT_CARGA_ARQUIVO_CSV_AUTORIZACAO = "ADM2-IbptCgArqCsvAut";
					public const String PROCESSAMENTO_PLANILHA_ESTOQUE = "ADM2-ProcXlsEstoque";
					public const String OP_LOG_PEDIDO_RECEBIDO_VIA_ADM2 = "ADM2-PEDIDO RECEBIDO";
					public const String OP_LOG_PEDIDO_RECEBIDO = "PEDIDO RECEBIDO";
					public const String OP_LOG_PEDIDO_RECEBIDO_MARKETPLACE = "PEDIDO RECEBIDO MKTP";
					public const String OP_LOG_PEDIDO_ATUALIZA_PREVISAO_ENTREGA_TRANSP = "PED PREV ETG TRANSP";
					public const string OP_LOG_ANOTA_FRETE_PEDIDO_VIA_CSV = "ADM2-FretePedViaCsv";
					public const string OP_LOG_ANOTA_FRETE_PEDIDO = "ANOTA FRETE PEDIDO";
				}
				#endregion

				#region [ Tabelas ]
				public class Tabelas
				{
					public const String PREFIXO_TABELA_TEMPORARIA = "tmp__";
					public const String T_IBPT = "t_IBPT";
				}
				#endregion

				#region [ ID_T_PARAMETRO ]
				public static class ID_T_PARAMETRO
				{
					public const string VERSAO_ULT_ARQ_IBPT_CSV_CARREGADO = "ADM2_VersaoUltArqIbptCsvCarregado";
					public const string ADM2_PROCESSAR_DATA_PREVISAO_ENTREGA_TRANSPORTADORA = "ADM2_ProcessarDataPrevisaoEntregaTransportadora";
					public const string ADM2_RASTREIO_PEDIDO_RECEBIDO_CLIENTE_SITUACAO_MERCADORIA_ENTREGUE = "ADM2_RastreioPedidoRecebidoCliente_SituacaoMercadoriaEntregue";
				}
				#endregion
			}
			#endregion

			#region [ Classe FIN ]
			public class FIN
			{
				public const String ID_USUARIO_SISTEMA = "SISTEMA";

				#region [ StAtivo ]
				public class StAtivo
				{
					public const byte INATIVO = 0;
					public const byte ATIVO = 1;
				}
				#endregion

				#region [ StSistema ]
				public class StSistema
				{
					public const byte NAO = 0;
					public const byte SIM = 1;
				}
				#endregion

				#region [ StCampoFlag ]
				public class StCampoFlag
				{
					public const byte FLAG_DESLIGADO = 0;
					public const byte FLAG_LIGADO = 1;
				}
				#endregion

				#region [ StSemEfeito ]
				public class StSemEfeito
				{
					public const byte FLAG_DESLIGADO = 0;
					public const byte FLAG_LIGADO = 1;
				}
				#endregion

				#region [ StConfirmacaoPendente ]
				public class StConfirmacaoPendente
				{
					public const byte FLAG_DESLIGADO = 0;
					public const byte FLAG_LIGADO = 1;
				}
				#endregion

				#region [ ST_T_FIN_NF_PARCELA_PAGTO ]
				public class ST_T_FIN_NF_PARCELA_PAGTO
				{
					public const byte INICIAL = 0;
					public const byte CANCELADO = 1;
					public const byte TRATADO = 2;
				}
				#endregion

				#region [ ST_T_FIN_PEDIDO_HIST_PAGTO ]
				public class ST_T_FIN_PEDIDO_HIST_PAGTO
				{
					public const byte PREVISAO = 1;
					public const byte QUITADO = 2;
					public const byte CANCELADO = 3;
				}
				#endregion

				#region [ T_PEDIDO__BOLETO_CONFECCIONADO_STATUS ]
				public class T_PEDIDO__BOLETO_CONFECCIONADO_STATUS
				{
					public const byte NAO = 0;
					public const byte SIM = 1;
					public const byte NAO_DEFINIDO = 10;
				}
				#endregion

				#region [ T_PEDIDO__GARANTIA_INDICADOR_STATUS ]
				public class T_PEDIDO__GARANTIA_INDICADOR_STATUS
				{
					public const byte NAO = 0;
					public const byte SIM = 1;
					public const byte NAO_DEFINIDO = 10;
				}
				#endregion

				#region [ T_PEDIDO__ANALISE_CREDITO_STATUS ]
				public class T_PEDIDO__ANALISE_CREDITO_STATUS
				{
					public const int ST_INICIAL = 0;
					public const int CREDITO_PENDENTE = 1;
					public const int CREDITO_OK = 2;
					public const int PENDENTE_VENDAS = 8;
					public const int CREDITO_OK_AGUARDANDO_DEPOSITO = 9;
					public const int NAO_ANALISADO = 10; // PEDIDOS ANTIGOS QUE JÁ ESTAVAM NA BASE
				}
				#endregion

				#region [ T_PEDIDO_FRETE__TIPO_PREENCHIMENTO ]
				public class T_PEDIDO_FRETE__TIPO_PREENCHIMENTO
				{
					public const int MANUAL = 1;
					public const int ANOTACAO_VIA_CSV_ADM2 = 4;
				}
				#endregion

				#region [ CtrlPagtoModulo ]
				public class CtrlPagtoModulo
				{
					public const byte BOLETO = 1;
					public const byte CHEQUE = 2;
					public const byte VISA = 3;
					public const byte BRASPAG_CARTAO = 4;
				}
				#endregion

				#region [ CtrlPagtoStatus ]
				public enum eCtrlPagtoStatus
				{
					// IMPORTANTE: NUNCA usar o valor reservado FLAG_NAO_SETADO = 255
					CONTROLE_MANUAL = 0,
					CADASTRADO_INICIAL = 1,
					BOLETO_BAIXADO = 3,
					CHEQUE_DEVOLVIDO = 4,
					VISA_CANCELADO = 5,
					BOLETO_PAGO_CHEQUE_VINCULADO = 6,
					BOLETO_COM_PAGAMENTO_CANCELADO = 7,
					PAGO = 10
				}
				#endregion

				#region [ FormaPagto ]
				public class FormaPagto
				{
					public const byte ID_FORMA_PAGTO_DINHEIRO = 1;
					public const byte ID_FORMA_PAGTO_DEPOSITO = 2;
					public const byte ID_FORMA_PAGTO_CHEQUE = 3;
					public const byte ID_FORMA_PAGTO_BOLETO = 4;
					public const byte ID_FORMA_PAGTO_CARTAO = 5;
					public const byte ID_FORMA_PAGTO_BOLETO_AV = 6;
					public const byte ID_FORMA_PAGTO_CARTAO_MAQUINETA = 7;
				}
				#endregion

				#region [ TipoCadastro ]
				public class TipoCadastro
				{
					public const char MANUAL = 'M';
					public const char SISTEMA = 'S';
				}
				#endregion

				#region [ EditadoManual ]
				public class EditadoManual
				{
					public const char NAO = 'N';
					public const char SIM = 'S';
				}
				#endregion

				#region [ Modulo ]
				public class Modulo
				{
					public const String FLUXO_CAIXA = "FLC";
					public const String BOLETO = "BOL";
					public const String CHEQUE = "CHQ";
					public const String VISA = "VIS";
					public const String HIST_PAGTO_PEDIDO = "HPP";
					public const String SERASA_RECIPROCIDADE = "SER";
					public const String FINANCEIRO_SERVICE = "FSV";
				}
				#endregion

				#region [ ID_T_PARAMETRO ]
				public static class ID_T_PARAMETRO
				{
					public const string OwnerPedido_ModoSelecao = "OwnerPedido_ModoSelecao";
					public const string ID_PARAMETRO_ADM2_ValidacaoVersao_VerificacaoPeriodica_FlagHabilitacao = "ADM2_ValidacaoVersao_VerificacaoPeriodica_FlagHabilitacao";
					public const string ID_PARAMETRO_ADM2_ValidacaoVersao_TempoMinEntreValidacoesPeriodicasEmSeg = "ADM2_ValidacaoVersao_TempoMinEntreValidacoesPeriodicasEmSeg";
				}
				#endregion

				#region [ ID_T_FIN_CONTROLE ]
				public static class ID_T_FIN_CONTROLE
				{
					public const string T_PEDIDO_FRETE = "T_PEDIDO_FRETE";
				}
				#endregion

				#region [ PARAMETRO_OPCOES ]
				public static class PARAMETRO_OPCOES
				{
					#region [ OwnerPedido_ModoSelecao ]
					public static class OwnerPedido_ModoSelecao
					{
						public const string Loja = "Loja";
						public const string NFeEmitente = "NFeEmitente";
					}
					#endregion
				}
				#endregion
			}
			#endregion

			#region [ t_CODIGO_DESCRICAO - Grupos ]
			public class GruposCodigoDescricao
			{
				public const string ID_GRUPO__PEDIDO_TIPO_FRETE = "Pedido_TipoFrete";
				public const string ID_GRUPO__PEDIDO_FRETE_TIPO_PREENCHIMENTO = "Pedido_Frete_TipoPreenchimento";
			}
			#endregion

			#region [ Códigos para formas de pagamento do pedido ]
			public class TipoParcelamentoPedido
			{
				public const short COD_FORMA_PAGTO_A_VISTA = 1;
				public const short COD_FORMA_PAGTO_PARCELADO_CARTAO = 2;
				public const short COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA = 3;
				public const short COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA = 4;
				public const short COD_FORMA_PAGTO_PARCELA_UNICA = 5;
				public const short COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA = 6;
			}
			#endregion

			#region [ Códigos para opções de forma de pagamento do pedido ]
			public class CodFormaPagtoPedido
			{
				public const short ID_FORMA_PAGTO_DINHEIRO = 1;
				public const short ID_FORMA_PAGTO_DEPOSITO = 2;
				public const short ID_FORMA_PAGTO_CHEQUE = 3;
				public const short ID_FORMA_PAGTO_BOLETO = 4;
				public const short ID_FORMA_PAGTO_CARTAO = 5;
				public const short ID_FORMA_PAGTO_BOLETO_AV = 6;
				public const short ID_FORMA_PAGTO_CARTAO_MAQUINETA = 7;
			}
			#endregion

			#region [ Status de Entrega do Pedido ]
			public class StEntregaPedido
			{
				public const String ST_ENTREGA_ESPERAR = "ESP";
				public const String ST_ENTREGA_SPLIT_POSSIVEL = "SPL";
				public const String ST_ENTREGA_SEPARAR = "SEP";
				public const String ST_ENTREGA_A_ENTREGAR = "AET";
				public const String ST_ENTREGA_ENTREGUE = "ETG";
				public const String ST_ENTREGA_CANCELADO = "CAN";
			}
			#endregion

			#region [ Status de Pagamento do Pedido ]
			public class StPagtoPedido
			{
				public const String ST_PAGTO_PAGO = "S";
				public const String ST_PAGTO_NAO_PAGO = "N";
				public const String ST_PAGTO_PARCIAL = "P";
			}
			#endregion

			#region [ Status de Pedido Recebido ]
			public class StPedidoRecebido
			{
				public const short COD_ST_PEDIDO_RECEBIDO_NAO = 0;
				public const short COD_ST_PEDIDO_RECEBIDO_SIM = 1;
				public const short COD_ST_PEDIDO_RECEBIDO_NAO_DEFINIDO = 10;
			}
			#endregion
		}
		#endregion

		#region [ AssemblyInfo ]
		public class AssemblyInfo
		{
			#region [ Assembly Attribute Accessors ]

			#region [ AssemblyTitle ]
			public static string AssemblyTitle
			{
				get
				{
					object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
					if (attributes.Length > 0)
					{
						AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
						if (titleAttribute.Title != "")
						{
							return titleAttribute.Title;
						}
					}
					return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
				}
			}
			#endregion

			#region [ AssemblyVersion ]
			public static string AssemblyVersion
			{
				get
				{
					return Assembly.GetExecutingAssembly().GetName().Version.ToString();
				}
			}
			#endregion

			#region [ AssemblyDescription ]
			public static string AssemblyDescription
			{
				get
				{
					object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
					if (attributes.Length == 0)
					{
						return "";
					}
					return ((AssemblyDescriptionAttribute)attributes[0]).Description;
				}
			}
			#endregion

			#region [ AssemblyProduct ]
			public static string AssemblyProduct
			{
				get
				{
					object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
					if (attributes.Length == 0)
					{
						return "";
					}
					return ((AssemblyProductAttribute)attributes[0]).Product;
				}
			}
			#endregion

			#region [ AssemblyCopyright ]
			public static string AssemblyCopyright
			{
				get
				{
					object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
					if (attributes.Length == 0)
					{
						return "";
					}
					return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
				}
			}
			#endregion

			#region [ AssemblyCompany ]
			public static string AssemblyCompany
			{
				get
				{
					object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
					if (attributes.Length == 0)
					{
						return "";
					}
					return ((AssemblyCompanyAttribute)attributes[0]).Company;
				}
			}
			#endregion

			#endregion
		}
		#endregion

		#region [ Atributos ]
		public static DateTime dtHrInicioRefRelogioServidor;
		public static DateTime dtHrInicioRefRelogioLocal;
		public static Color BackColorPainelPadrao = SystemColors.Control;
		public static List<string> listaCodigosRastreioSituacaoMercadoriaEntregue = new List<string>();
		#endregion

		#region [ Parâmetros ]
		public static class Parametro
		{
			public static int ValidacaoVersao_VerificacaoPeriodica_FlagHabilitacao;
			public static int ValidacaoVersao_TempoMinEntreValidacoesPeriodicasEmSeg;
		}
		#endregion

		#region [ enum ]

		#region [ Filtro Flag st_inativo ]
		public enum eFiltroFlagStInativo : byte
		{
			FLAG_DESLIGADO = 0,
			FLAG_LIGADO = 1,
			FLAG_IGNORADO = 255
		}
		#endregion

		#endregion

		#region [ Classe Acesso ]
		public class Acesso
		{
			#region [ Constantes ]
			public const String OP_CEN_ADM2_APP_ADMINISTRATIVO_ACESSO_AO_MODULO = "25200";
			public const string OP_CEN_ADM2_ANOTAR_PEDIDOS_RECEBIDOS_PELO_CLIENTE = "29700";
			public const string OP_CEN_ADM2_ANOTAR_FRETE_PEDIDO_VIA_CSV = "29900";
			#endregion

			#region [ Atributos ]
			public static List<String> listaOperacoesPermitidas = new List<String>();
			#endregion

			#region [ Métodos ]

			#region [ operacaoPermitida ]
			/// <summary>
			/// Indica se a operação especificada no parâmetro consta na lista de operações permitidas do usuário
			/// </summary>
			/// <param name="idOperacao">
			/// Operação a ser pesquisada na lista de operações permitidas
			/// </param>
			/// <returns>
			/// true: a operação pesquisada consta na lista de operações permitidas
			/// false: a operação pesquisada não consta na lista de operações permitidas
			/// </returns>
			public static bool operacaoPermitida(String idOperacao)
			{
				if (idOperacao == null) return false;
				if (idOperacao.Trim().Length == 0) return false;

				for (int i = 0; i < listaOperacoesPermitidas.Count; i++)
				{
					if (listaOperacoesPermitidas[i].ToString().Equals(idOperacao)) return true;
				}
				// Operação não consta da lista de operações permitidas
				return false;
			}
			#endregion

			#endregion
		}
		#endregion

		#region [ Classe Usuario ]
		public class Usuario
		{
			#region [ Atributos ]
			public static String usuario = "";
			public static String senhaDigitada = "";
			public static String senhaCriptografada = "";
			public static String senhaDescriptografada = "";
			public static String nome = "";
			public static bool cadastrado = false;
			public static bool bloqueado = false;
			public static bool senhaExpirada = false;
			public static String fin_email_remetente;
			public static String fin_display_name_remetente;
			public static String fin_servidor_smtp_endereco;
			public static int fin_servidor_smtp_porta;
			public static String fin_usuario_smtp;
			public static String fin_senha_smtp;
			#endregion

			#region [ Defaults ]
			public class Defaults
			{
				public class FIbpt
				{
					public static String pathIbptArquivoCsv = "";
				}

				public class FAtualizaPlanilhaEstoque
				{
					public static String pathArquivoPlanilhaEstoque = "";
					public static String fileNameArquivoPlanilhaEstoque = "";
				}

				public class FAnotarPedidoRecebidoCliente
				{
					public static string pathArquivoRastreio = "";
					public static string fileNameArquivoRastreio = "";
				}

				public class FAnotarFretePedidoViaCsv
				{
					public static string pathArquivoCsv = "";
					public static string fileNameArquivoCsv = "";
				}
			}
			#endregion
		}
		#endregion

		#region [ RegistryApp ]
		public class RegistryApp
		{
			public const string REGISTRY_BASE_PATH = "Software\\" + Cte.Aplicativo.NOME_OWNER + "\\" + Cte.Aplicativo.NOME_SISTEMA;

			#region [ Chaves ]
			public class Chaves
			{
				public static String left = "Left";
				public static String top = "Top";
				public static String usuario = "Usuario";
				public class FIbpt
				{
					public static String pathIbptArquivoCsv = "FIbpt_pathIbptArquivoCsv";
				}
				public class FAtualizaPlanilhaEstoque
				{
					public static String pathArquivoPlanilhaEstoque = "FAtualizaPlanilhaEstoque_pathArquivoPlanilhaEstoque";
					public static String fileNameArquivoPlanilhaEstoque = "FAtualizaPlanilhaEstoque_fileNameArquivoPlanilhaEstoque";
				}
				public class FAnotarPedidoRecebidoCliente
				{
					public static String pathArquivoRastreio = "FAnotarPedidoRecebidoCliente_pathArquivoRastreio";
					public static String fileNameArquivoRastreio = "FAnotarPedidoRecebidoCliente_fileNameArquivoRastreio";
				}
			}
			#endregion

			#region [ Métodos ]

			#region [ criaRegistryKey ]
			public static RegistryKey criaRegistryKey(String subKey)
			{
				RegistryKey regKey = Registry.CurrentUser;
				regKey = regKey.CreateSubKey(subKey);
				return regKey;
			}
			#endregion

			#endregion
		}
		#endregion

		#region[ ReaderWriterLock ]
		public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
		#endregion

		#region [ Métodos ]

		#region [ arredondaParaMonetario ]
		public static decimal arredondaParaMonetario(decimal numero)
		{
			return converteNumeroDecimal(formataMoeda(numero));
		}
		#endregion

		#region[ barraInvertidaAdd ]
		public static string barraInvertidaAdd(string path)
		{
			if (path == null) return "";
			string strResp = path.TrimEnd();
			if (strResp.Length == 0) return "";
			if (strResp[strResp.Length - 1] == (char)92) return strResp;
			return strResp + (char)92;
		}
		#endregion

		#region[ barraInvertidaDel ]
		public static string barraInvertidaDel(string path)
		{
			if (path == null) return "";
			string strResp = path.TrimEnd();
			while (true)
			{
				if (strResp.Length == 0) return "";
				if (strResp[strResp.Length - 1] != (char)92) return strResp;
				strResp = strResp.Substring(0, strResp.Length - 1).TrimEnd();
			}
		}
		#endregion

		#region [ calculaTimeSpanDias ]
		/// <summary>
		/// Calcula a quantidade de dias.
		/// Exemplo de uso:
		///		calculaDateTimeDias(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade de dias.
		/// </returns>
		public static int calculaTimeSpanDias(TimeSpan ts)
		{
			return ts.Days;
		}
		#endregion

		#region [ calculaTimeSpanHoras ]
		/// <summary>
		/// Calcula a quantidade de horas.
		/// Exemplo de uso:
		///		calculaDateTimeHoras(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade de horas.
		/// </returns>
		public static int calculaTimeSpanHoras(TimeSpan ts)
		{
			return ts.Hours + (24 * ts.Days);
		}
		#endregion

		#region [ calculaTimeSpanMiliSegundos ]
		/// <summary>
		/// Calcula a quantidade de milisegundos.
		/// Exemplo de uso:
		///		calculaDateTimeMiliSegundos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade milisegundos.
		/// </returns>
		public static long calculaTimeSpanMiliSegundos(TimeSpan ts)
		{
			return (long)ts.Milliseconds + (long)1000 * ((long)ts.Seconds + ((long)60 * ((long)ts.Minutes + ((long)60 * ((long)ts.Hours + ((long)24 * (long)ts.Days))))));
		}
		#endregion

		#region [ calculaTimeSpanMinutos ]
		/// <summary>
		/// Calcula a quantidade de minutos.
		/// Exemplo de uso:
		///		calculaDateTimeMinutos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade minutos.
		/// </returns>
		public static int calculaTimeSpanMinutos(TimeSpan ts)
		{
			return ts.Minutes + (60 * (ts.Hours + (24 * ts.Days)));
		}
		#endregion

		#region [ calculaTimeSpanSegundos ]
		/// <summary>
		/// Calcula a quantidade de segundos.
		/// Exemplo de uso:
		///		calculaDateTimeSegundos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade segundos.
		/// </returns>
		public static long calculaTimeSpanSegundos(TimeSpan ts)
		{
			return (long)ts.Seconds + ((long)60 * ((long)ts.Minutes + ((long)60 * ((long)ts.Hours + ((long)24 * (long)ts.Days)))));
		}
		#endregion

		#region [ contagemLetras ]
		public static int contagemLetras(string texto)
		{
			#region [ Declarações ]
			int qtdeLetras = 0;
			#endregion

			if ((texto ?? "").Length == 0) return 0;

			for (int i = 0; i < texto.Length; i++)
			{
				if (isLetra(texto[i])) qtdeLetras++;
			}

			return qtdeLetras;
		}
		#endregion

		#region [ converteColorFromHtml ]
		public static Color? converteColorFromHtml(string htmlColor)
		{
			#region [ Declarações ]
			Color cor;
			#endregion

			if (htmlColor == null) return null;
			if (htmlColor.Trim().Length == 0) return null;

			try
			{
				htmlColor = htmlColor.Trim();
				if (!htmlColor.StartsWith("#")) htmlColor = "#" + htmlColor;
				cor = ColorTranslator.FromHtml(htmlColor);
				return cor;
			}
			catch (Exception)
			{
				return null;
			}

		}
		#endregion

		#region[ converteDdMmYyParaDateTime ]
		/// <summary>
		/// Converte um texto no formato DDMMYY (ano c/ 2 dígitos) com ou sem separadores para o tipo DateTime.
		/// O pivotamento do ano é feito com base de ano 80.
		/// </summary>
		/// <param name="strDdMmYy">Texto representando uma data no formato DDMMYY (ano com 2 dígitos) com ou sem separadores</param>
		/// <returns>
		/// Retorna a data representada no tipo DateTime
		/// </returns>
		public static DateTime converteDdMmYyParaDateTime(string strDdMmYy)
		{
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			string strDdMmYyyy;
			String strDdMm;
			String strYyyy;
			string strFormato;

			strDdMm = Texto.leftStr(digitos(strDdMmYy), 4);

			strYyyy = Texto.rightStr(digitos(strDdMmYy), 2);
			if (converteInteiro(strYyyy) >= 80) strYyyy = "19" + strYyyy; else strYyyy = "20" + strYyyy;

			strDdMmYyyy = strDdMm + strYyyy;

			strFormato = Cte.DataHora.FmtDia +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtAno;
			if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteDdMmYyyyParaDateTime ]
		public static DateTime converteDdMmYyyyParaDateTime(string strDdMmYyyy)
		{
			string strFormato;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			strFormato = Cte.DataHora.FmtDia +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtAno;
			if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteYyyyMmDdParaDateTime ]
		public static DateTime converteYyyyMmDdParaDateTime(string strYyyyMmDd)
		{
			string strYyyyMmDdAux;
			string strDdMmYyyy;
			string strFormato;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			strYyyyMmDdAux = digitos(strYyyyMmDd);
			if (strYyyyMmDdAux.Length == 0) return DateTime.MinValue;
			strDdMmYyyy = strYyyyMmDdAux.Substring(6, 2) + strYyyyMmDdAux.Substring(4, 2) + strYyyyMmDdAux.Substring(0, 4);
			strFormato = Cte.DataHora.FmtDia +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtAno;
			if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteYyyyMmDdHhMmSsParaDateTime ]
		/// <summary>
		/// Converte o texto que representa uma data/hora para DateTime
		/// </summary>
		/// <param name="strYyyyMmDdHhMmSs">
		/// Texto representando uma data/hora, com ou sem separadores, sendo que a parte da hora é opcional.
		/// </param>
		/// <returns>
		/// Retorna a data/hora como DateTime, se não for possível fazer a conversão, retorna DateTime.MinValue
		/// </returns>
		public static DateTime converteYyyyMmDdHhMmSsParaDateTime(string strYyyyMmDdHhMmSs)
		{
			#region [ Declarações ]
			char c;
			string strDia = "";
			string strMes = "";
			string strAno = "";
			string strHora = "";
			string strMinuto = "";
			string strSegundo = "";
			string strFormato;
			string strDataHoraAConverter;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			#endregion

			#region [ Ano ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strAno += c;
				if (strAno.Length == 4) break;
			}
			if (strAno.Length == 2)
			{
				if (converteInteiro(strAno) >= 80)
					strAno = "19" + strAno;
				else
					strAno = "20" + strAno;
			}
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Mês ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMes += c;
				if (strMes.Length == 2) break;
			}
			while (strMes.Length < 2) strMes = '0' + strMes;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Dia ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strDia += c;
				if (strDia.Length == 2) break;
			}
			while (strDia.Length < 2) strDia = '0' + strDia;
			#endregion

			#region [ Remove separador(es) entre a data e hora, se houver ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				if (!isDigit(strYyyyMmDdHhMmSs[0]))
					strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				else
					break;
			}
			#endregion

			#region [ Hora ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Segundo ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strSegundo += c;
				if (strSegundo.Length == 2) break;
			}
			while (strSegundo.Length < 2) strSegundo = '0' + strSegundo;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtAno +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtDia +
						 ' ' +
						 Cte.DataHora.FmtHora +
						 Cte.DataHora.FmtMin +
						 Cte.DataHora.FmtSeg;
			#endregion

			#region [ Monta data/hora normalizada ]
			strDataHoraAConverter = strAno +
									strMes +
									strDia +
									' ' +
									strHora +
									strMinuto +
									strSegundo;
			#endregion

			if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteInteiro ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo inteiro
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="valor">
		/// Texto representando um número inteiro
		/// </param>
		/// <returns>
		/// Retorna um número do tipo inteiro
		/// </returns>
		public static Int64 converteInteiro(string valor)
		{
			Int64 intResultado = 0;

			if (valor == null) return 0;

			string strValor = valor.Trim();
			if (strValor.Length == 0) return 0;

			try
			{
				intResultado = Int64.Parse(strValor);
			}
			catch (Exception)
			{
				intResultado = 0;
			}

			return intResultado;
		}
		#endregion

		#region[ converteInteiro ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo inteiro
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="valor">
		/// Texto representando um número inteiro
		/// </param>
		/// <param name="valorDefault">
		/// Valor que será retornado no caso da conversão falhar
		/// </param>
		/// <returns>
		/// Retorna um número do tipo inteiro
		/// </returns>
		public static Int64 converteInteiro(string valor, Int64 valorDefault)
		{
			Int64 intResultado = 0;

			if (valor == null) return valorDefault;

			string strValor = valor.Trim();
			if (strValor.Length == 0) return valorDefault;

			try
			{
				intResultado = Int64.Parse(strValor);
			}
			catch (Exception)
			{
				intResultado = valorDefault;
			}

			return intResultado;
		}
		#endregion

		#region [ converteNumeroDecimal ]

		#region [ converteNumeroDecimal ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo decimal
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="numero">
		/// Texto representando um número decimal
		/// </param>
		/// <returns>
		/// Retorna um número do tipo decimal
		/// </returns>
		public static decimal converteNumeroDecimal(String numero)
		{
			#region [ Declarações ]
			int i;
			char c_separador_decimal;
			String s_numero_aux;
			String s_inteiro = "";
			String s_centavos = "";
			int intSinal = 1;
			decimal decFracionario;
			decimal decInteiro;
			decimal decResultado;
			#endregion

			if (numero == null) return 0;
			if (numero.Trim().Length == 0) return 0;

			numero = numero.Trim();

			if (numero.IndexOf('-') != -1) intSinal = -1;

			c_separador_decimal = retornaSeparadorDecimal(numero);

			#region [ Separa parte inteira e os centavos ]
			s_numero_aux = numero.Replace(c_separador_decimal, 'V');
			String[] v = s_numero_aux.Split('V');
			for (i = 0; i < v.Length; i++)
			{
				if (v[i] == null) v[i] = "";
			}
			// Falha ao determinar o separador de decimal, então calcula como se não houvesse centavos
			if (v.Length > 2)
			{
				s_inteiro = digitos(numero);
			}
			else
			{
				if (v.Length >= 1) s_inteiro = digitos(v[0]);
				if (v.Length >= 2) s_centavos = digitos(v[1]);
			}
			if (s_inteiro.Length == 0) s_inteiro = "0";
			s_centavos = s_centavos.PadRight(2, '0');
			#endregion

			decInteiro = (decimal)converteInteiro(s_inteiro);
			decFracionario = (decimal)converteInteiro(s_centavos) / (decimal)Math.Pow(10, s_centavos.Length);
			decResultado = intSinal * (decInteiro + decFracionario);
			return decResultado;
		}
		#endregion

		#region [ converteNumeroDecimal ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo decimal
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="numero">
		/// Texto representando um número decimal
		/// </param>
		/// <param name="separadorDecimal">
		/// Caractere usado como separador decimal
		/// </param>
		/// <returns>
		/// Retorna um número do tipo decimal
		/// </returns>
		public static decimal converteNumeroDecimal(string numero, char separadorDecimal)
		{
			#region [ Declarações ]
			int i;
			char c_separador_decimal;
			String s_numero_aux;
			String s_inteiro = "";
			String s_centavos = "";
			int intSinal = 1;
			decimal decFracionario;
			decimal decInteiro;
			decimal decResultado;
			#endregion

			if (numero == null) return 0;
			if (numero.Trim().Length == 0) return 0;

			numero = numero.Trim();

			if (numero.IndexOf('-') != -1) intSinal = -1;

			c_separador_decimal = separadorDecimal;
			if (c_separador_decimal.ToString().Trim().Length == 0) c_separador_decimal = retornaSeparadorDecimal(numero);

			#region [ Separa parte inteira e os centavos ]
			s_numero_aux = numero.Replace(c_separador_decimal, 'V');
			String[] v = s_numero_aux.Split('V');
			for (i = 0; i < v.Length; i++)
			{
				if (v[i] == null) v[i] = "";
			}
			// Falha ao determinar o separador de decimal, então calcula como se não houvesse centavos
			if (v.Length > 2)
			{
				s_inteiro = digitos(numero);
			}
			else
			{
				if (v.Length >= 1) s_inteiro = digitos(v[0]);
				if (v.Length >= 2) s_centavos = digitos(v[1]);
			}
			if (s_inteiro.Length == 0) s_inteiro = "0";
			s_centavos = s_centavos.PadRight(2, '0');
			#endregion

			decInteiro = (decimal)converteInteiro(s_inteiro);
			decFracionario = (decimal)converteInteiro(s_centavos) / (decimal)Math.Pow(10, s_centavos.Length);
			decResultado = intSinal * (decInteiro + decFracionario);
			return decResultado;
		}
		#endregion

		#endregion

		#region [ converteNumeroDouble ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo double
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="numero">
		/// Texto representando um número real
		/// </param>
		/// <returns>
		/// Retorna um número do tipo double
		/// </returns>
		public static double converteNumeroDouble(String numero)
		{
			#region [ Declarações ]
			int i;
			char c_separador_decimal;
			String s_numero_aux;
			String s_inteiro = "";
			String s_decimal = "";
			int intSinal = 1;
			double dblFracionario;
			double dblInteiro;
			double dblResultado;
			#endregion

			if (numero == null) return 0;
			if (numero.Trim().Length == 0) return 0;

			numero = numero.Trim();

			if (numero.IndexOf('-') != -1) intSinal = -1;

			#region [ Obtém o separador decimal ]
			c_separador_decimal = '.';
			for (int j = numero.Length - 1; j >= 0; j--)
			{
				if (!isDigit(numero[j]))
				{
					c_separador_decimal = numero[j];
					break;
				}
			}
			#endregion

			#region [ Separa parte inteira e decimal ]
			s_numero_aux = numero.Replace(c_separador_decimal, 'V');
			String[] v = s_numero_aux.Split('V');
			for (i = 0; i < v.Length; i++)
			{
				if (v[i] == null) v[i] = "";
			}
			// Falha ao determinar o separador de decimal, então calcula como se não houvesse decimal
			if (v.Length > 2)
			{
				s_inteiro = digitos(numero);
			}
			else
			{
				if (v.Length >= 1) s_inteiro = digitos(v[0]);
				if (v.Length >= 2) s_decimal = digitos(v[1]);
			}
			if (s_inteiro.Length == 0) s_inteiro = "0";
			s_decimal = s_decimal.PadRight(1, '0');
			#endregion

			dblInteiro = (double)converteInteiro(s_inteiro);
			dblFracionario = (double)converteInteiro(s_decimal) / (double)Math.Pow(10, s_decimal.Length);
			dblResultado = intSinal * (dblInteiro + dblFracionario);
			return dblResultado;
		}
		#endregion

		#region[ digitos ]
		public static string digitos(string texto)
		{
			StringBuilder d = new StringBuilder();
			if (texto == null) return "";
			for (int i = 0; i < texto.Length; i++)
			{
				if ((texto[i] >= '0') && (texto[i] <= '9')) d.Append(texto[i]);
			}
			return d.ToString();
		}
		#endregion

		#region [ excel_converte_numeracao_digito_para_letra ]
		public static string excel_converte_numeracao_digito_para_letra(int numeracao_digito)
		{
			#region [ Declarações ]
			const int TOTAL_LETRAS_ALFABETO = 26;
			string strResp;
			int intQuoc;
			int intResto;
			#endregion

			strResp = "";
			if (numeracao_digito <= 0) return "";
			intQuoc = (int)(numeracao_digito - 1) / TOTAL_LETRAS_ALFABETO;
			intResto = numeracao_digito - (intQuoc * TOTAL_LETRAS_ALFABETO);
			if (intQuoc > TOTAL_LETRAS_ALFABETO) return "";
			if (intQuoc > 0) strResp = ((char)(65 - 1 + intQuoc)).ToString();
			strResp += ((char)(65 - 1 + intResto)).ToString();
			return strResp;
		}
		#endregion

		#region [ executaManutencaoArqLogAtividade ]
		/// <summary>
		/// Apaga os arquivos de log de atividade antigos
		/// </summary>
		public static void executaManutencaoArqLogAtividade()
		{
			#region [ Declarações ]
			DateTime dtCorte = DateTime.Now.AddDays(-Global.Cte.LogAtividade.CorteArqLogEmDias);
			string strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtYYYYMMDD);
			string[] ListaArqLog;
			string strNomeArq;
			int i;
			#endregion

			#region[ Apaga arquivos de log de atividade antigos ]
			ListaArqLog = Directory.GetFiles(Global.Cte.LogAtividade.PathLogAtividade, "*." + Global.Cte.LogAtividade.ExtensaoArqLog, SearchOption.TopDirectoryOnly);
			for (i = 0; i < ListaArqLog.Length; i++)
			{
				strNomeArq = Global.extractFileName(ListaArqLog[i]);
				strNomeArq = strNomeArq.Substring(0, strDataCorte.Length);
				if (string.Compare(strNomeArq, strDataCorte) < 0) File.Delete(ListaArqLog[i]);
			}
			#endregion
		}
		#endregion

		#region[ extractFileName ]
		public static string extractFileName(string fileName)
		{
			string strResp = "";
			for (int i = (fileName.Length - 1); i >= 0; i--)
			{
				if (fileName[i] == (char)92) return strResp;
				if (fileName[i] == (char)47) return strResp;
				if (fileName[i] == (char)58) return strResp;
				strResp = fileName[i] + strResp;
			}
			return strResp;
		}
		#endregion

		#region [ filtraAcentuacao ]
		public static String filtraAcentuacao(String texto)
		{
			#region [ Declarações ]
			String strResp;
			#endregion

			if (texto == null) return texto;
			if (texto.Length == 0) return texto;

			strResp = texto.ToString();
			if (strResp.IndexOf('á') != -1) strResp = strResp.Replace('á', 'a');
			if (strResp.IndexOf('à') != -1) strResp = strResp.Replace('à', 'a');
			if (strResp.IndexOf('ã') != -1) strResp = strResp.Replace('ã', 'a');
			if (strResp.IndexOf('â') != -1) strResp = strResp.Replace('â', 'a');
			if (strResp.IndexOf('ä') != -1) strResp = strResp.Replace('ä', 'a');
			if (strResp.IndexOf('é') != -1) strResp = strResp.Replace('é', 'e');
			if (strResp.IndexOf('è') != -1) strResp = strResp.Replace('è', 'e');
			if (strResp.IndexOf('ê') != -1) strResp = strResp.Replace('ê', 'e');
			if (strResp.IndexOf('ë') != -1) strResp = strResp.Replace('ë', 'e');
			if (strResp.IndexOf('í') != -1) strResp = strResp.Replace('í', 'i');
			if (strResp.IndexOf('ì') != -1) strResp = strResp.Replace('ì', 'i');
			if (strResp.IndexOf('î') != -1) strResp = strResp.Replace('î', 'i');
			if (strResp.IndexOf('ï') != -1) strResp = strResp.Replace('ï', 'i');
			if (strResp.IndexOf('ó') != -1) strResp = strResp.Replace('ó', 'o');
			if (strResp.IndexOf('ò') != -1) strResp = strResp.Replace('ò', 'o');
			if (strResp.IndexOf('õ') != -1) strResp = strResp.Replace('õ', 'o');
			if (strResp.IndexOf('ô') != -1) strResp = strResp.Replace('ô', 'o');
			if (strResp.IndexOf('ö') != -1) strResp = strResp.Replace('ö', 'o');
			if (strResp.IndexOf('ú') != -1) strResp = strResp.Replace('ú', 'u');
			if (strResp.IndexOf('ù') != -1) strResp = strResp.Replace('ù', 'u');
			if (strResp.IndexOf('û') != -1) strResp = strResp.Replace('û', 'u');
			if (strResp.IndexOf('ü') != -1) strResp = strResp.Replace('ü', 'u');
			if (strResp.IndexOf('ç') != -1) strResp = strResp.Replace('ç', 'c');
			if (strResp.IndexOf('ñ') != -1) strResp = strResp.Replace('ñ', 'n');
			if (strResp.IndexOf('ÿ') != -1) strResp = strResp.Replace('ÿ', 'y');

			if (strResp.IndexOf('Á') != -1) strResp = strResp.Replace('Á', 'A');
			if (strResp.IndexOf('À') != -1) strResp = strResp.Replace('À', 'A');
			if (strResp.IndexOf('Ã') != -1) strResp = strResp.Replace('Ã', 'A');
			if (strResp.IndexOf('Â') != -1) strResp = strResp.Replace('Â', 'A');
			if (strResp.IndexOf('Ä') != -1) strResp = strResp.Replace('Ä', 'A');
			if (strResp.IndexOf('É') != -1) strResp = strResp.Replace('É', 'E');
			if (strResp.IndexOf('È') != -1) strResp = strResp.Replace('È', 'E');
			if (strResp.IndexOf('Ê') != -1) strResp = strResp.Replace('Ê', 'E');
			if (strResp.IndexOf('Ë') != -1) strResp = strResp.Replace('Ë', 'E');
			if (strResp.IndexOf('Í') != -1) strResp = strResp.Replace('Í', 'I');
			if (strResp.IndexOf('Ì') != -1) strResp = strResp.Replace('Ì', 'I');
			if (strResp.IndexOf('Î') != -1) strResp = strResp.Replace('Î', 'I');
			if (strResp.IndexOf('Ï') != -1) strResp = strResp.Replace('Ï', 'I');
			if (strResp.IndexOf('Ó') != -1) strResp = strResp.Replace('Ó', 'O');
			if (strResp.IndexOf('Ò') != -1) strResp = strResp.Replace('Ò', 'O');
			if (strResp.IndexOf('Õ') != -1) strResp = strResp.Replace('Õ', 'O');
			if (strResp.IndexOf('Ô') != -1) strResp = strResp.Replace('Ô', 'O');
			if (strResp.IndexOf('Ö') != -1) strResp = strResp.Replace('Ö', 'O');
			if (strResp.IndexOf('Ú') != -1) strResp = strResp.Replace('Ú', 'U');
			if (strResp.IndexOf('Ù') != -1) strResp = strResp.Replace('Ù', 'U');
			if (strResp.IndexOf('Û') != -1) strResp = strResp.Replace('Û', 'U');
			if (strResp.IndexOf('Ü') != -1) strResp = strResp.Replace('Ü', 'U');
			if (strResp.IndexOf('Ç') != -1) strResp = strResp.Replace('Ç', 'C');
			if (strResp.IndexOf('Ñ') != -1) strResp = strResp.Replace('Ñ', 'N');

			return strResp;
		}
		#endregion

		#region [ filtraDigitacaoCep ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de CEP
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoCep(char c)
		{
			if (!(isDigit(c) || (c == '-') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoCnpjCpf ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de CNPJ/CPF
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoCnpjCpf(char c)
		{
			if (!(isDigit(c) || (c == '.') || (c == '-') || (c == '/') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoData ]
		/// <summary>
		/// Filtra os caracteres durante a digitação da data
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoData(char c)
		{
			// Deixa passar somente dígitos, o caracter separador de data e o backspace,
			// caso contrário, retorna o caracter nulo.
			if (!(((c >= '0') && (c <= '9')) || (c == '/') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoEmail ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de endereço de e-mail, aceitando também os
		/// seguintes caracteres separadores quando é digitada uma lista de e-mails: espaço em branco,
		/// vírgula e ponto e vírgula
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoEmail(char c)
		{
			if (!(isDigit(c) || isLetra(c) || (c == '@') || (c == '.') || (c == '_') || (c == '-') || (c == ' ') || (c == ',') || (c == ';') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoMoeda ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de valor monetário
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoMoeda(char c)
		{
			// Deixa passar somente dígitos, o sinal negativo, os caracteres separadores de milhar e 
			// decimal e o backspace, caso contrário, retorna o caracter nulo.
			if (!(((c >= '0') && (c <= '9')) || (c == '.') || (c == ',') || (c == '-') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoNumeroInteiro ]
		public static char filtraDigitacaoNumeroInteiro(char c)
		{
			// Deixa passar somente dígitos e o backspace, caso contrário, retorna o caracter nulo.
			if (!(((c >= '0') && (c <= '9')) || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoPercentual ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de número percentual
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoPercentual(char c)
		{
			// Deixa passar somente dígitos, o caracter separador de decimal e o backspace, caso contrário,
			// retorna o caracter nulo.
			if (!(((c >= '0') && (c <= '9')) || (c == '.') || (c == ',') || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoSomenteLetras ]
		public static char filtraDigitacaoSomenteLetras(char c)
		{
			// Deixa passar somente letras e o backspace, caso contrário, retorna o caracter nulo.
			if (!(((c >= 'a') && (c <= 'z')) || ((c >= 'A') && (c <= 'Z')) || (c == '\b'))) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraDigitacaoTexto ]
		/// <summary>
		/// Filtra os caracteres durante a digitação de campo texto livre
		/// </summary>
		/// <param name="c">
		/// Caracter digitado
		/// </param>
		/// <returns>
		/// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
		/// </returns>
		public static char filtraDigitacaoTexto(char c)
		{
			// Filtra os caracteres Ascii 34 e 39 (aspas duplas e aspas simples, respectivamente)
			if ((c == '\x0022') || (c == '\x0027') || (c == '|')) c = '\0';
			return c;
		}
		#endregion

		#region [ filtraTexto ]
		/// <summary>
		/// Filtra caracteres inválidos para um campo texto livre. Ex: aspas simples, aspas duplas, etc.
		/// </summary>
		/// <param name="texto">
		/// Conteúdo de um campo do tipo texto livre.
		/// </param>
		/// <returns>
		/// Retorna o texto sem conter nenhum caracter inválido para um campo do tipo texto livre.
		/// </returns>
		public static String filtraTexto(String texto)
		{
			StringBuilder sb = new StringBuilder("");
			for (int i = 0; i < texto.Length; i++)
			{
				if ((texto[i] != '\x0022') &&
					(texto[i] != '\x0027'))
				{
					sb.Append(texto[i]);
				}
			}
			return sb.ToString();
		}
		#endregion

		#region [ formataCep ]
		public static String formataCep(String cep)
		{
			String strCep;
			if (cep == null) return "";
			strCep = digitos(cep);
			if (strCep.Length != 8) return cep;
			strCep = strCep.Substring(0, 5) + '-' + strCep.Substring(5, 3);
			return strCep;
		}
		#endregion

		#region [ formataCnpjCpf ]
		/// <summary>
		/// Formata os dígitos de CNPJ/CPF informados aplicando a máscara de formatação
		/// </summary>
		/// <param name="cnpj_cpf">
		/// Dígitos do CNPJ/CPF
		/// </param>
		/// <returns>
		/// Retorna o CNPJ/CPF formatado
		/// </returns>
		public static String formataCnpjCpf(String cnpj_cpf)
		{
			String s;
			String s_aux;
			String s_resp;

			if (cnpj_cpf == null) return "";

			s = digitos(cnpj_cpf);

			#region [ Verifica se é um CNPJ mesmo ou se é um CPF c/ zeros p/ normalizar à esquerda ]
			if (s.Length == 14)
			{
				if (!isCnpjOk(s))
				{
					if (Texto.leftStr(s, 3).Equals("000"))
					{
						s_aux = Texto.rightStr(s, 11);
						if (isCpfOk(s_aux)) s = s_aux;
					}
				}
			}
			#endregion

			// CPF
			if (s.Length == 11)
			{
				s_resp = s.Substring(0, 3) + '.' + s.Substring(3, 3) + '.' + s.Substring(6, 3) + '/' + s.Substring(9, 2);
			}
			// CNPJ
			else if (s.Length == 14)
			{
				s_resp = s.Substring(0, 2) + '.' + s.Substring(2, 3) + '.' + s.Substring(5, 3) + '/' + s.Substring(8, 4) + '-' + s.Substring(12, 2);
			}
			// Desconhecido
			else
			{
				s_resp = cnpj_cpf;
			}
			return s_resp;
		}
		#endregion

		#region [ formataDataDdMmYyComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YY
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YY
		/// </returns>
		public static String formataDataDdMmYyComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyComSeparador);
		}
		#endregion

		#region [ formataDataDdMmYyyyComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY
		/// </returns>
		public static String formataDataDdMmYyyyComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);
		}
		#endregion

		#region [ formataDataDdMmYyyyHhMmComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY HH:MM
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY HH:MM
		/// </returns>
		public static String formataDataDdMmYyyyHhMmComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmComSeparador);
		}
		#endregion

		#region [ formataDataDdMmYyyyHhMmSsComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY HH:MM:SS
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY HH:MM
		/// </returns>
		public static String formataDataDdMmYyyyHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmSsComSeparador);
		}
		#endregion

		#region [ formataDataDigitadaParaDDMMYYYYComSeparador ]
		/// <summary>
		/// A partir de uma data digitada pelo usuário, tenta normalizar e retornar uma data no formato DD/MM/YYYY
		/// </summary>
		/// <param name="data">
		/// Texto com a data digitada pelo usuário a ser normalizado
		/// Formatos aceitos: DDMMYY, DDMMYYYY, DD/MM/YY, DD/MM/YYYY
		/// </param>
		/// <returns>
		/// Retorna a data no formato DD/MM/YYYY caso a data informada esteja em um formato válido, caso contrário, retorna o próprio texto informado
		/// </returns>
		public static String formataDataDigitadaParaDDMMYYYYComSeparador(String data)
		{
			String strDia;
			String strMes;
			String strAno;

			if (data == null) return "";
			if (data.Trim().Length == 0) return "";

			if (data.IndexOf('/') == -1)
			{
				// A data foi digitada sem os separadores
				data = digitos(data);
				// Neste caso, aceita somente se tiver sido digitada no formado DDMM ou DDMMYY ou DDMMYYYY
				if ((data.Length != 4) && (data.Length != 6) && (data.Length != 8)) return data;
				strDia = data.Substring(0, 2);
				strMes = data.Substring(2, 2);
				if (data.Length > 4)
					strAno = data.Substring(4, data.Length - 4);
				else
					strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
			}
			else
			{
				String[] v = data.Split('/');
				// É necessário que a data tenha vindo separada em 2 ou 3 partes: dia/mês ou dia/mês/ano
				if ((v.Length != 2) && (v.Length != 3)) return data;
				for (int i = 0; i < v.Length; i++)
				{
					if (v[i] == null) return data;
					v[i] = digitos(v[i]);
					if (v[i].Trim().Length == 0) return data;
				}
				strDia = v[0].PadLeft(2, '0');
				strMes = v[1].PadLeft(2, '0');
				if (v.Length > 2)
					strAno = v[2];
				else
					strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
			}

			if (strAno.Length == 3)
			{
				if (converteInteiro(strAno) >= 900) strAno = "1" + strAno; else strAno = "2" + strAno;
			}
			else if (strAno.Length == 2)
			{
				if (converteInteiro(strAno) >= 80) strAno = "19" + strAno; else strAno = "20" + strAno;
			}
			else if (strAno.Length == 1)
			{
				strAno = DateTime.Now.Year.ToString().Substring(0, 3) + strAno;
			}
			else if (strAno.Length != 4) return data;

			return strDia + "/" + strMes + "/" + strAno;
		}
		#endregion

		#region [ formataDataYyyyMmDdComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato YYYY-MM-DD
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato YYYY-MM-DD
		/// </returns>
		public static String formataDataYyyyMmDdComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtYyyyMmDdComSeparador);
		}
		#endregion

		#region [ formataDataYyyyMmDdComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato YYYY-MM-DD
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato YYYY-MM-DD
		/// </returns>
		public static String formataDataYyyyMmDdComSeparador(DateTime data, string separador)
		{
			#region [ Declarações ]
			string sData;
			string sResp;
			string[] vData;
			#endregion

			if (separador == null) separador = "";

			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			sData = data.ToString(Global.Cte.DataHora.FmtYyyyMmDdComSeparador);
			vData = sData.Split('-');
			sResp = vData[0] + separador + vData[1] + separador + vData[2];
			return sResp;
		}
		#endregion

		#region [ formataDataYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato YYYY-MM-DD HH:MM:SS
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato YYYY-MM-DD HH:MM:SS
		/// </returns>
		public static String formataDataYyyyMmDdHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtYyyyMmDdHhMmSsComSeparador);
		}
		#endregion

		#region[ formataDuracaoHMS ]
		public static string formataDuracaoHMS(TimeSpan ts)
		{
			StringBuilder sb = new StringBuilder();
			// Dias
			if (ts.Days > 0) sb.Append(ts.Days.ToString() + 'd');
			// Horas
			if (sb.ToString().Length == 0)
			{
				if (ts.Hours > 0) sb.Append(ts.Hours.ToString() + 'h');
			}
			else
			{
				sb.Append(ts.Hours.ToString().PadLeft(2, '0') + 'h');
			}
			// Minutos
			if (sb.ToString().Length == 0)
			{
				if (ts.Minutes > 0) sb.Append(ts.Minutes.ToString() + 'm');
			}
			else
			{
				sb.Append(ts.Minutes.ToString().PadLeft(2, '0') + 'm');
			}
			// Segundos
			if (sb.ToString().Length == 0)
			{
				sb.Append(ts.Seconds.ToString() + 's');
			}
			else
			{
				sb.Append(ts.Seconds.ToString().PadLeft(2, '0') + 's');
			}
			return sb.ToString();
		}
		#endregion

		#region [ formataHoraHhMmSsComSeparador ]
		public static String formataHoraHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtHhMmSsComSeparador);
		}
		#endregion

		#region [ formataHoraHhMmSsComSimbolo ]
		public static String formataHoraHhMmSsComSimbolo(DateTime data)
		{
			#region [ Declarações ]
			string sHora;
			string sResp;
			string[] vHora;
			#endregion

			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			sHora = data.ToString(Global.Cte.DataHora.FmtHhMmSsComSeparador);
			vHora = sHora.Split(':');
			sResp = vHora[0] + 'h' + vHora[1] + 'm' + vHora[2] + 's';
			return sResp;
		}
		#endregion

		#region [ formataInteiro ]
		public static String formataInteiro(int numero)
		{
			String strResp = "";
			String strNumero;
			int intPonto = 0;

			strNumero = digitos(numero.ToString());
			for (int i = strNumero.Length - 1; i >= 0; i--)
			{
				intPonto++;
				strResp = strNumero[i] + strResp;
				if ((intPonto % 3 == 0) && (i != 0)) strResp = '.' + strResp;
			}
			return strResp;
		}
		#endregion

		#region [ formataMoeda ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato monetário
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um valor monetário
		/// </param>
		/// <returns>
		/// Retorna um texto com formato monetário
		/// </returns>
		public static String formataMoeda(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataMoedaComSimbolo ]
		public static string formataMoedaComSimbolo(decimal valor)
		{
			return Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(valor);
		}
		#endregion

		#region [ formataMoedaDigitada ]
		/// <summary>
		/// A partir de um valor digitado pelo usuário, tentar normalizar e retornar um valor monetário
		/// formatado com separador de milhar e de decimais
		/// </summary>
		/// <param name="numero">
		/// Texto com o valor monetário digitado a ser normalizado, positivo ou negativo
		/// </param>
		/// <returns>
		/// Retorna o valor formatado com separador de milhar e de decimais: 999.999,99
		/// </returns>
		public static String formataMoedaDigitada(String numero)
		{
			#region [ Declarações ]
			int i;
			int j;
			char c_separador_decimal;
			String s_numero_aux;
			String s_inteiro = "";
			String s_centavos = "";
			String s_valor_formatado;
			String s_sinal = "";
			#endregion

			if (numero == null) return "";
			if (numero.Trim().Length == 0) return "";

			numero = numero.Trim();

			if (numero.IndexOf('-') != -1) s_sinal = "-";

			c_separador_decimal = retornaSeparadorDecimal(numero);

			#region [ Formata o valor monetário ]
			s_numero_aux = numero.Replace(c_separador_decimal, 'V');
			String[] v = s_numero_aux.Split('V');
			for (i = 0; i < v.Length; i++)
			{
				if (v[i] == null) v[i] = "";
			}
			// Falha ao determinar o separador de decimal, então retorna o próprio valor informado
			if (v.Length > 2) return numero;

			if (v.Length >= 1) s_inteiro = digitos(v[0]);
			if (v.Length >= 2) s_centavos = digitos(v[1]);
			if (s_inteiro.Length == 0) s_inteiro = "0";
			s_centavos = Texto.leftStr(s_centavos, 2);
			s_centavos = s_centavos.PadRight(2, '0');

			// Coloca os separadores de milhar
			s_numero_aux = "";
			j = 0;
			for (i = s_inteiro.Length - 1; i >= 0; i--)
			{
				j++;
				s_numero_aux = s_inteiro[i] + s_numero_aux;
				if (((j % 3) == 0) && (i != s_inteiro.Length - 1) && (i != 0)) s_numero_aux = "." + s_numero_aux;
			}
			s_inteiro = s_numero_aux;

			s_valor_formatado = s_sinal + s_inteiro + "," + s_centavos;
			#endregion

			return s_valor_formatado;
		}
		#endregion

		#region [ formataPercentual ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentual(double valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataPercentualCom1Decimal ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentualCom1Decimal(double valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.0");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 2), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataPercentualCom2Decimais ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentualCom2Decimais(double valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataTelefone ]
		public static String formataTelefone(String telefone)
		{
			int i;
			String strTel = "";

			if (telefone != null) strTel = digitos(telefone);
			if ((strTel.Length == 0) || (strTel.Length > 8) || (!isTelefoneOk(strTel))) return strTel;

			i = strTel.Length - 4;
			strTel = strTel.Substring(0, i) + "-" + strTel.Substring(i);
			return strTel;
		}

		public static String formataTelefone(String ddd, String telefone)
		{
			String strDDD = "";
			String strTel;
			strTel = formataTelefone(telefone);
			if (ddd != null) strDDD = digitos(ddd);
			if ((strTel.Length > 0) && (strDDD.Length > 0)) strTel = "(" + strDDD + ") " + strTel;
			return strTel;
		}

		public static String formataTelefone(String ddd, String telefone, String ramal)
		{
			String strRamal = "";
			String strTel;
			strTel = formataTelefone(ddd, telefone);
			if (ramal != null) strRamal = digitos(ramal);
			if ((strTel.Length > 0) && (strRamal.Length > 0)) strTel += " R:" + strRamal;
			return strTel;
		}
		#endregion

		#region [ getBackColorFromAppConfig ]
		public static Color? getBackColorFromAppConfig()
		{
			#region[ Declarações ]
			string sBackColor;
			#endregion

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			sBackColor = ConfigurationManager.AppSettings["backgroundColorPainel"];
			return converteColorFromHtml(sBackColor);
			#endregion
		}
		#endregion

		#region [ GetConfigurationValue ]
		public static string GetConfigurationValue(string key)
		{
			Assembly service = Assembly.GetAssembly(typeof(Global));
			Configuration config = ConfigurationManager.OpenExeConfiguration(service.Location);
			if (config.AppSettings.Settings[key] != null)
			{
				return config.AppSettings.Settings[key].Value;
			}
			else
			{
				return null;
			}
		}
		#endregion

		#region [ getVScrollBarWidth ]
		/// <summary>
		/// Dado um componente (ex: DataGridView) que contém um vertical scroll bar, retorna a largura do scroll bar
		/// </summary>
		/// <param name="control">
		/// Objeto que contém o scroll bar
		/// </param>
		/// <returns>
		/// Retorna a largura do scroll bar
		/// </returns>
		public static int getVScrollBarWidth(Control control)
		{
			foreach (Control c in control.Controls)
			{
				if (c.GetType().Equals(typeof(VScrollBar)))
				{
					return c.Width;
				}
			}
			return 0;
		}
		#endregion

		#region[ gravaLogAtividade ]
		/// <summary>
		/// Grava a informação do parâmetro no arquivo de log, junto com a data/hora
		/// Se o parâmetro for 'null', será gravada uma linha em branco no arquivo
		/// Se o parâmetro uma string vazia, será gravada uma linha apenas com a data/hora
		/// </summary>
		/// <param name="mensagem"></param>
		public static void gravaLogAtividade(string mensagem)
		{
			string linha;
			DateTime dataHora = DateTime.Now;
			const string FmtHHMMSS = Cte.DataHora.FmtHora + ":" + Cte.DataHora.FmtMin + ":" + Cte.DataHora.FmtSeg + "." + Cte.DataHora.FmtMiliSeg;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			const string FmtYYYYMMDD = Cte.DataHora.FmtAno + Cte.DataHora.FmtMes + Cte.DataHora.FmtDia;
			string strArqLog = Global.barraInvertidaAdd(Global.Cte.LogAtividade.PathLogAtividade) +
							   DateTime.Now.ToString(FmtYYYYMMDD) +
							   "." +
							   Global.Cte.LogAtividade.ExtensaoArqLog;
			if (mensagem == null)
				linha = "";
			else
				linha = dataHora.ToString(FmtHHMMSS) + ": " + mensagem;

			try
			{
				rwlArqLogAtividade.AcquireWriterLock(60 * 1000);
				try
				{
					using (StreamWriter sw = new StreamWriter(strArqLog, true, encode))
					{
						sw.WriteLine(linha);
						sw.Flush();
						sw.Close();
					}
				}
				finally
				{
					rwlArqLogAtividade.ReleaseWriterLock();
				}
			}
			catch (Exception)
			{
				// Nop
			}
		}
		#endregion

		#region[ haOutraInstanciaEmExecucao ]
		public static bool haOutraInstanciaEmExecucao()
		{
			Process current = Process.GetCurrentProcess();
			Process[] processes = Process.GetProcessesByName(current.ProcessName);

			//Loop through the running processes in with the same name 
			foreach (Process process in processes)
			{
				//Ignore the current process 
				if (process.Id != current.Id)
				{
					return true;
				}
			}
			return false;
		}
		#endregion

		#region [ isAlfaNumerico ]
		public static bool isAlfaNumerico(char c)
		{
			if (isDigit(c) || isLetra(c)) return true;
			return false;
		}
		#endregion

		#region [ isCepOk ]
		public static bool isCepOk(String cep)
		{
			String strCep;
			if (cep == null) return false;
			strCep = digitos(cep);
			if ((strCep.Length == 5) || (strCep.Length == 8)) return true;
			return false;
		}
		#endregion

		#region [ isCnpjOk ]
		/// <summary>
		/// Indica se o CNPJ está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cnpj">
		/// CNPJ a testar
		/// </param>
		/// <returns>
		/// true: CNPJ válido
		/// false: CNPJ inválido
		/// </returns>
		public static bool isCnpjOk(String cnpj)
		{
			String s_cnpj;
			String p1 = "543298765432";
			String p2 = "6543298765432";
			bool tudo_igual;
			int i;
			int d;

			if (cnpj == null) return false;

			s_cnpj = digitos(cnpj);
			if (s_cnpj.Length != 14) return false;

			// Dígitos são todos iguais?
			tudo_igual = true;
			for (i = 0; i < (s_cnpj.Length - 1); i++)
			{
				if (!s_cnpj.Substring(i, 1).Equals(s_cnpj.Substring(i + 1, 1)))
				{
					tudo_igual = false;
					break;
				}
			}
			if (tudo_igual) return false;

			// Verifica o primeiro check digit
			d = 0;
			for (i = 0; i < 12; i++)
			{
				d = d + int.Parse(p1.Substring(i, 1)) * int.Parse(s_cnpj.Substring(i, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cnpj.Substring(12, 1))) return false;

			// Verifica o segundo check digit
			d = 0;
			for (i = 0; i < 13; i++)
			{
				d = d + int.Parse(p2.Substring(i, 1)) * int.Parse(s_cnpj.Substring(i, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cnpj.Substring(13, 1))) return false;

			// Ok
			return true;
		}
		#endregion

		#region [ isCnpjCpfOk ]
		/// <summary>
		/// Indica se o CNPJ/CPF está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cnpj_cpf">
		/// CNPJ/CPF a testar
		/// </param>
		/// <returns>
		/// true: CNPJ/CPF válido
		/// false: CNPJ/CPF inválido
		/// </returns>
		public static bool isCnpjCpfOk(String cnpj_cpf)
		{
			String s;
			if (cnpj_cpf == null) return false;
			s = digitos(cnpj_cpf);
			if (s.Length == 11)
			{
				return isCpfOk(s);
			}
			else if (s.Length == 14)
			{
				return isCnpjOk(s);
			}
			return false;
		}
		#endregion

		#region [ isCpfOk ]
		/// <summary>
		/// Indica se o CPF está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cpf">
		/// CPF a testar
		/// </param>
		/// <returns>
		/// true: CPF válido
		/// false: CPF inválido
		/// </returns>
		public static bool isCpfOk(String cpf)
		{
			int i;
			int d;
			bool tudo_igual;
			String s_cpf;

			if (cpf == null) return false;

			s_cpf = digitos(cpf);
			if (s_cpf.Length != 11) return false;

			// Dígitos todos iguais?
			tudo_igual = true;
			for (i = 0; i < (s_cpf.Length - 1); i++)
			{
				if (!s_cpf.Substring(i, 1).Equals(s_cpf.Substring(i + 1, 1)))
				{
					tudo_igual = false;
					break;
				}
			}
			if (tudo_igual) return false;

			// Verifica o primeiro check digit
			d = 0;
			for (i = 1; i <= 9; i++)
			{
				d = d + (11 - i) * int.Parse(s_cpf.Substring(i - 1, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cpf.Substring(9, 1))) return false;

			// Verifica o segundo check digit
			d = 0;
			for (i = 2; i <= 10; i++)
			{
				d = d + (12 - i) * int.Parse(s_cpf.Substring(i - 1, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cpf.Substring(10, 1))) return false;

			// Ok
			return true;
		}
		#endregion

		#region [ isDataOk ]
		/// <summary>
		/// Indica se a data representada pelo texto no formato DD/MM/YYYY é uma data válida
		/// </summary>
		/// <param name="data">
		/// Texto representando uma data no formato DD/MM/YYYY
		/// </param>
		/// <returns>
		/// true: data válida
		/// false: data inválida
		/// </returns>
		public static bool isDataOk(String data)
		{
			bool blnDataOk;
			string strFormato;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			data = formataDataDigitadaParaDDMMYYYYComSeparador(data);
			if (data.Length != 10) return false;
			strFormato = Cte.DataHora.FmtDia +
						 "/" +
						 Cte.DataHora.FmtMes +
						 "/" +
						 Cte.DataHora.FmtAno;
			blnDataOk = DateTime.TryParseExact(data, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp);
			return blnDataOk;
		}
		#endregion

		#region[ isDigit ]
		public static bool isDigit(char c)
		{
			if ((c >= '0') && (c <= '9')) return true;
			return false;
		}
		#endregion

		#region [ isEmailOk ]
		/// <summary>
		/// Indica se o e-mail possui sintaxe válida. Se for uma lista de e-mails, testa cada um dos e-mails.
		/// </summary>
		/// <param name="email">
		/// Um ou mais e-mails que devem ser analisados. Os e-mails podem ser separados por espaço em branco,
		/// vírgula ou ponto e vírgula.
		/// </param>
		/// <param name="relacaoEmailInvalido">
		/// Informa os e-mails inválidos separados por espaço em branco.
		/// </param>
		/// <returns>
		/// true: todos os e-mails são válidos
		/// false: um ou mais e-mails inválidos
		/// </returns>
		public static bool isEmailOk(String email, ref String relacaoEmailInvalido)
		{
			string strRegExEmailValidacao = "^([0-9a-zA-Z]([-.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";
			bool blnSucesso;
			int intQtdeEmail = 0;
			String[] v;
			String strEmail;
			Regex rgex = new Regex(strRegExEmailValidacao);

			relacaoEmailInvalido = "";
			if (email == null) return false;
			if (email.Trim().Length == 0) return false;

			blnSucesso = true;
			strEmail = email.Trim();
			strEmail = strEmail.Replace(',', ' ');
			strEmail = strEmail.Replace(';', ' ');
			strEmail = strEmail.Replace("\n", " ");
			strEmail = strEmail.Replace("\r", " ");
			v = strEmail.Split(' ');
			for (int i = 0; i < v.Length; i++)
			{
				if (v[i].Trim().Length > 0)
				{
					intQtdeEmail++;
					if (!rgex.IsMatch(v[i].Trim()))
					{
						if (relacaoEmailInvalido.Length > 0) relacaoEmailInvalido += " ";
						relacaoEmailInvalido += v[i];
						blnSucesso = false;
					}
				}
			}
			if (intQtdeEmail <= 0) return false;
			return blnSucesso;
		}
		#endregion

		#region [ IsFileLocked ]
		public static bool IsFileLocked(string fullFileName)
		{
			FileInfo fi;

			if (fullFileName == null) return false;
			if (fullFileName.Trim().Length == 0) return false;
			if (!File.Exists(fullFileName)) return false;

			fi = new FileInfo(fullFileName);
			return IsFileLocked(fi);
		}

		public static bool IsFileLocked(FileInfo file)
		{
			FileStream stream = null;

			try
			{
				stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
			}
			catch (IOException)
			{
				//the file is unavailable because it is:
				//still being written to
				//or being processed by another thread
				//or does not exist (has already been processed)
				return true;
			}
			finally
			{
				if (stream != null)
					stream.Close();
			}

			//file is not locked
			return false;
		}
		#endregion

		#region [ isLetra ]
		public static bool isLetra(char c)
		{
			return ((Char.ToUpper(c) >= 'A') && (Char.ToUpper(c) <= 'Z'));
		}
		#endregion

		#region [ isLetra ]
		public static bool isLetra(String c)
		{
			if (c == null) return false;
			if (c.Trim().Length == 0) return false;

			for (int i = 0; i < c.Length; i++)
			{
				if (!isLetra(c[i])) return false;
			}
			return true;
		}
		#endregion

		#region [ isNumeroPedido ]
		public static bool isNumeroPedido(String numeroPedido)
		{
			String strParteNumerica;
			if (numeroPedido == null) return false;
			if (numeroPedido.Trim().Length == 0) return false;

			strParteNumerica = digitos(Texto.leftStr(numeroPedido, Cte.Etc.TAM_MIN_NUM_PEDIDO));
			if (strParteNumerica.Length != Cte.Etc.TAM_MIN_NUM_PEDIDO) return false;
			if (!isLetra(numeroPedido.Substring(Cte.Etc.TAM_MIN_NUM_PEDIDO, 1))) return false;
			return true;
		}
		#endregion

		#region [ isPedidoFilhote ]
		/// <summary>
		/// Analisa se o número do pedido é de um pedido-base ou de um pedido-filhote
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido a ser analisado
		/// </param>
		/// <returns>
		/// true: trata-se de um número de pedido-filhote
		/// false: trata-se de um número de pedido-base
		/// </returns>
		public static bool isPedidoFilhote(String numeroPedido)
		{
			if (numeroPedido == null) return false;
			numeroPedido = numeroPedido.Trim();
			numeroPedido = normalizaNumeroPedido(numeroPedido);
			if (numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE) > -1) return true;
			return false;
		}
		#endregion

		#region [ isTelefoneOk ]
		public static bool isTelefoneOk(String telefone)
		{
			String strTelefone;
			if (telefone == null) return false;
			strTelefone = digitos(telefone);
			if (strTelefone.Length < 7) return false;
			return true;
		}
		#endregion

		#region [ isUfOk ]
		public static bool isUfOk(String uf)
		{
			String strListaUf = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO";
			String strUf;
			String[] v;
			if (uf == null) return false;
			strUf = uf.Trim().ToUpper();
			if (strUf.Length != 2) return false;
			v = strListaUf.Split(' ');
			for (int i = 0; i < v.Length; i++)
			{
				if (strUf.Equals(v[i].Trim()))
				{
					return true;
				}
			}
			return false;
		}
		#endregion

		#region [ isVScrollBarVisible ]
		/// <summary>
		/// Indica se o Vertical Scroll Bar de um componente (ex: DataGridView) está visível
		/// </summary>
		/// <param name="control">
		/// Objeto que contém o scroll bar
		/// </param>
		/// <returns>
		/// true: o vertical scroll bar está visível
		/// false: o vertical scroll bar não está visível
		/// </returns>
		public static bool isVScrollBarVisible(Control control)
		{
			foreach (Control c in control.Controls)
			{
				if (c.GetType().Equals(typeof(VScrollBar))) return c.Visible;
			}
			return false;
		}
		#endregion

		#region [ normalizaCodigo ]
		public static string normalizaCodigo(string codigo, int tamanhoDefault)
		{
			#region [ Declarações ]
			StringBuilder sbCodigoNormalizado;
			#endregion

			if (codigo == null) return null;
			if (codigo.Trim().Length == 0) return "";

			sbCodigoNormalizado = new StringBuilder(codigo.Trim());
			while (sbCodigoNormalizado.Length < tamanhoDefault)
			{
				sbCodigoNormalizado.Insert(0, '0');
			}

			return sbCodigoNormalizado.ToString();
		}
		#endregion

		#region [ normalizaCodigoFabricante ]
		public static string normalizaCodigoFabricante(string codigoFabricante)
		{
			return normalizaCodigo(codigoFabricante, Cte.Etc.TAM_MIN_FABRICANTE);
		}
		#endregion

		#region [ normalizaCodigoProduto ]
		public static string normalizaCodigoProduto(string codigoProduto)
		{
			return normalizaCodigo(codigoProduto, Cte.Etc.TAM_MIN_PRODUTO);
		}
		#endregion

		#region [ normalizaNumeroLoja ]
		public static string normalizaNumeroLoja(string numeroLoja)
		{
			#region [ Declarações ]
			int numLoja;
			#endregion

			numLoja = (int)converteInteiro(numeroLoja);
			return normalizaCodigo(numLoja.ToString(), Cte.Etc.TAM_MIN_LOJA);
		}
		#endregion

		#region [ normalizaNumeroPedido ]
		public static String normalizaNumeroPedido(String pedido)
		{
			String id_pedido;
			String s = "";
			String s_ano = "";
			String s_num = "";
			String s_filhote = "";
			char c;

			if (pedido == null) return "";
			id_pedido = pedido.Trim().ToUpper();
			if (id_pedido.Length == 0) return "";

			for (int i = 0; i < id_pedido.Length; i++)
			{
				if (isDigit(id_pedido[i]))
					s_num += id_pedido[i];
				else
					break;
			}
			if (s_num.Length == 0) return "";

			for (int i = 0; i < id_pedido.Length; i++)
			{
				c = id_pedido[i];
				if (isLetra(c))
				{
					if (s_ano.Length == 0)
					{
						s_ano = c.ToString();
					}
					else
					{
						if (s_filhote.Length == 0) s_filhote = c.ToString();
					}
				}
			}
			if (s_ano.Length == 0) return "";
			s_num = s_num.PadLeft(Cte.Etc.TAM_MIN_NUM_PEDIDO, '0');
			s = s_num + s_ano;
			if (s_filhote.Length > 0) s += Cte.Etc.COD_SEPARADOR_FILHOTE + s_filhote;
			return s;
		}
		#endregion

		#region [ retornaNumeroPedidoBase ]
		/// <summary>
		/// No caso do número do pedido ser de um pedido-filhote, retorna apenas a parte do número
		/// correspondente ao pedido-base.
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido a ser analisado.
		/// </param>
		/// <returns>
		/// Retorna apenas a parte do número que identifica o pedido-base.
		/// </returns>
		public static String retornaNumeroPedidoBase(String numeroPedido)
		{
			if (numeroPedido == null) return "";
			numeroPedido = numeroPedido.Trim();
			if (numeroPedido.Length == 0) return "";
			numeroPedido = normalizaNumeroPedido(numeroPedido);
			if (numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE) == -1) return numeroPedido;
			return numeroPedido.Substring(0, numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE));
		}
		#endregion

		#region [ retornaSeparadorDecimal ]
		/// <summary>
		/// Analisa o texto do parâmetro que representa um valor monetário para determinar se o separador decimal é ponto ou vírgula
		/// </summary>
		/// <param name="numero">
		/// Texto representando um valor monetário
		/// </param>
		/// <returns>
		/// Retorna o caracter usado para representação do separador decimal (de centavos)
		/// </returns>
		private static char retornaSeparadorDecimal(String valorMonetario)
		{
			int i;
			int n_ponto = 0;
			int n_virgula = 0;
			int n_digitos_finais = 0;
			int n_digitos_iniciais = 0;
			char c;
			String s_numero;
			char c_ult_sep = '\0';
			char c_separador_decimal;

			if (valorMonetario == null) return ',';
			if (valorMonetario.Trim().Length == 0) return ',';

			s_numero = valorMonetario.Trim();
			for (i = s_numero.Length - 1; i >= 0; i--)
			{
				c = s_numero[i];
				if (c == '.')
				{
					n_ponto++;
					if (c_ult_sep == '\0') c_ult_sep = c;
				}
				else if (c == ',')
				{
					n_virgula++;
					if (c_ult_sep == '\0') c_ult_sep = c;
				}
				if (isDigit(c) && (n_ponto == 0) && (n_virgula == 0)) n_digitos_finais++;
				if (isDigit(c) && ((n_ponto > 0) || (n_virgula > 0))) n_digitos_iniciais++;
			}

			// Default
			c_separador_decimal = ',';
			if (c_ult_sep == '.')
			{
				if ((n_ponto == 1) && (n_virgula == 0) && (n_digitos_iniciais <= 3) && (n_digitos_finais == 3))
				{
					// NOP: Considera 123.456 como cento e vinte e três mil e quatrocentos e cinquenta e seis
				}
				else if (n_ponto == 1)
				{
					c_separador_decimal = '.';
				}
			}
			else if (c_ult_sep == ',')
			{
				if ((n_virgula > 1) && (n_ponto == 0)) c_separador_decimal = '.';
			}
			return c_separador_decimal;
		}
		#endregion

		#region [ setBackColorToAppConfig ]
		public static bool setBackColorToAppConfig(string htmlColor)
		{
			try
			{
				Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
				config.AppSettings.Settings["backgroundColorPainel"].Value = (htmlColor == null ? "" : htmlColor);
				config.Save(ConfigurationSaveMode.Modified);
				ConfigurationManager.RefreshSection("appSettings");

				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
		#endregion

		#region [ sqlFormataDdMmYyyyParaSqlYyyyMmDd ]
		/// <summary>
		/// A partir de um texto representando uma data no formato DD/MM/YYYY, com ou sem separadores, formata para um texto representando a data no formato 'YYYY-MM-DD' que é entendido pelo SQL Server como uma data
		/// </summary>
		/// <param name="dataDdMmYyyy">
		/// Texto representando uma data no formato DD/MM/YYYY, com ou sem separadores
		/// </param>
		/// <returns>
		/// Retorna um texto representando a data no formato 'YYYY-MM-DD' que é entendido pelo SQL Server como uma data
		/// </returns>
		public static String sqlFormataDdMmYyyyParaSqlYyyyMmDd(String dataDdMmYyyy)
		{
			string strData;

			if (dataDdMmYyyy == null) return "NULL";
			if (dataDdMmYyyy.Trim().Length == 0) return "NULL";

			strData = digitos(dataDdMmYyyy);
			if (strData.Length != 8) return "NULL";
			strData = strData.Substring(4, 4) + "-" + strData.Substring(2, 2) + "-" + strData.Substring(0, 2);
			return strData;
		}
		#endregion

		#region [ sqlFormataDecimal ]
		/// <summary>
		/// Dado um número do tipo decimal, formata um texto representando esse número de forma adequada para usá-lo em uma expressão SQL
		/// </summary>
		/// <param name="valor">
		/// Número do tipo decimal que se deseja representar em um texto para ser usado em expressão SQL
		/// </param>
		/// <returns>
		/// Retorna um texto representando o número em um formato adequado para ser usado em expressão SQL
		/// </returns>
		public static String sqlFormataDecimal(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal = "";
			decimal decNumeroAuxiliar = .5M;
			String strNumeroAuxiliar;

			strNumeroAuxiliar = decNumeroAuxiliar.ToString();

			if (strNumeroAuxiliar.IndexOf(".") > -1)
				strSeparadorDecimal = ".";
			else if (strNumeroAuxiliar.IndexOf(",") > -1)
				strSeparadorDecimal = ",";

			strValorFormatado = valor.ToString();
			if (strSeparadorDecimal.Length > 0)
			{
				strValorFormatado = strValorFormatado.Replace(strSeparadorDecimal, "V");
				strValorFormatado = strValorFormatado.Replace(".", "");
				strValorFormatado = strValorFormatado.Replace(",", "");
				strValorFormatado = strValorFormatado.Replace("V", ".");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ sqlMontaCaseWhenParametroStringVaziaComoNull ]
		/// <summary>
		/// Para parâmetros de objetos SqlCommand que são usados para datas expressas como
		/// string no formato YYYY-MM-DD, monta uma expressão CASE WHEN para gravar NULL
		/// quando o valor do parâmetro for uma string vazia.
		/// Lembrando que o SQL Server grava automaticamente a data de 1900-01-01 quando
		/// converte uma string vazia para um campo datetime.
		/// </summary>
		/// <param name="nomeParametroDoCommand">Nome do parâmetro (ex: @dtVencto)</param>
		/// <returns>Retorna um texto contendo uma expressão CASE WHEN, ex: CASE WHEN @dt_vencto='' THEN NULL ELSE @dt_vencto END</returns>
		public static String sqlMontaCaseWhenParametroStringVaziaComoNull(String nomeParametroDoCommand)
		{
			String strResp;
			strResp = "CASE WHEN " + nomeParametroDoCommand + " = '' THEN NULL ELSE " + nomeParametroDoCommand + " END";
			return strResp;
		}
		#endregion

		#region [ sqlMontaDdMmYyyyParaSqlDateTime ]
		/// <summary>
		/// A partir de um texto representando uma data no formato DD/MM/YYYY, com ou sem separadores, monta uma expressão SQL para converter para o tipo de dados DataTime do SQL Server
		/// </summary>
		/// <param name="dataDdMmYyyy">
		/// Texto representando uma data no formato DD/MM/YYYY, com ou sem separadores
		/// </param>
		/// <returns>
		/// Retorna uma expressão SQL para converter para o tipo de dados DateTime do SQL Server
		/// </returns>
		public static string sqlMontaDdMmYyyyParaSqlDateTime(String dataDdMmYyyy)
		{
			string strData;

			if (dataDdMmYyyy == null) return "NULL";
			if (dataDdMmYyyy.Trim().Length == 0) return "NULL";

			strData = digitos(dataDdMmYyyy);
			if (strData.Length != 8) return "NULL";
			strData = strData.Substring(4, 4) + "-" + strData.Substring(2, 2) + "-" + strData.Substring(0, 2);
			return "Convert(datetime, '" + strData + "', 120)";
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTime ]
		public static string sqlMontaDateTimeParaSqlDateTime(DateTime dtReferencia)
		{
			string strDataHora;
			string strSql;

			if (dtReferencia == null) return "NULL";
			if (dtReferencia == DateTime.MinValue) return "NULL";

			strDataHora = dtReferencia.ToString(Cte.DataHora.FmtAno) +
						  "-" +
						  dtReferencia.ToString(Cte.DataHora.FmtMes) +
						  "-" +
						  dtReferencia.ToString(Cte.DataHora.FmtDia) +
						  " " +
						  dtReferencia.ToString(Cte.DataHora.FmtHora) +
						  ":" +
						  dtReferencia.ToString(Cte.DataHora.FmtMin) +
						  ":" +
						  dtReferencia.ToString(Cte.DataHora.FmtSeg);
			strSql = "Convert(datetime, '" + strDataHora + "', 120)";
			return strSql;
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTimeSomenteData ]
		public static string sqlMontaDateTimeParaSqlDateTimeSomenteData(DateTime dtReferencia)
		{
			string strData;
			string strSql;
			strData = dtReferencia.ToString(Cte.DataHora.FmtAno) +
					  "-" +
					  dtReferencia.ToString(Cte.DataHora.FmtMes) +
					  "-" +
					  dtReferencia.ToString(Cte.DataHora.FmtDia);
			strSql = "Convert(datetime, '" + strData + "', 120)";
			return strSql;
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <param name="strAlias">
		/// Informa o nome do Alias, caso seja informado uma string vazia, então será usado o nome do próprio campo.
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdComSeparador(string strNomeCampo, string strAlias)
		{
			string strResposta;
			if (strAlias.Trim().Length == 0) strAlias = strNomeCampo;
			strResposta = "Coalesce(Convert(varchar(19), " + strNomeCampo + ", 121), '')";
			if (strAlias.Length > 0) strResposta += " AS " + strAlias;
			return strResposta;
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdComSeparador(string strNomeCampo)
		{
			return sqlMontaDateTimeParaYyyyMmDdComSeparador(strNomeCampo, "");
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30 14:27:01
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <param name="strAlias">
		/// Informa o nome do Alias, caso seja informado uma string vazia, então será usado o nome do próprio campo.
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(string strNomeCampo, string strAlias)
		{
			string strResposta;
			if (strAlias.Trim().Length == 0) strAlias = strNomeCampo;
			strResposta = "Coalesce(Convert(varchar(19), " + strNomeCampo + ", 121), '')";
			if (strAlias.Length > 0) strResposta += " AS " + strAlias;
			return strResposta;
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30 14:27:01
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(string strNomeCampo)
		{
			return sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(strNomeCampo, "");
		}
		#endregion

		#region [ sqlMontaGetdateSomenteData ]
		/// <summary>
		/// Monta uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </summary>
		/// <returns>
		/// Retorna uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </returns>
		public static string sqlMontaGetdateSomenteData()
		{
			string strResposta;
			String strNomeCampo = "getdate()";
			strResposta = "Coalesce(Convert(varchar(10), " + strNomeCampo + ", 121), '')";
			return strResposta;
		}
		#endregion

		#region [ sqlMontaGetdateSomenteData ]
		/// <summary>
		/// Monta uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </summary>
		/// <returns>
		/// Retorna uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </returns>
		public static string sqlMontaGetdateSomenteData(string strAlias)
		{
			string strResposta;
			strResposta = sqlMontaGetdateSomenteData();
			if (strAlias.Length > 0) strResposta += " AS " + strAlias;
			return strResposta;
		}
		#endregion

		#region [ sqlMontaPadLeftCampoNumerico ]
		/// <summary>
		/// Monta uma expressão SQL (sintaxe do SQL Server) para realizar a função de PadLeft() em um campo do tipo numérico que será convertido para varchar
		/// </summary>
		/// <param name="nomeCampo">
		/// Nome do campo no banco de dados
		/// </param>
		/// <param name="preenchimento">
		/// Caracter para preenchimento no padding
		/// </param>
		/// <param name="tamanhoCampo">
		/// Tamanho que o texto deve ficar após execução do padding
		/// </param>
		/// <returns>
		/// Expressão SQL (sintaxe do SQL Server) para realizar a função PadLeft()
		/// </returns>
		public static String sqlMontaPadLeftCampoNumerico(String nomeCampo, char preenchimento, int tamanhoCampo)
		{
			String strResp;
			strResp = " Coalesce(Replicate('" + preenchimento + "'," + tamanhoCampo.ToString() + "-Len(Convert(varchar," + nomeCampo + "))), '') + Convert(varchar," + nomeCampo + ")";
			return strResp;
		}
		#endregion

		#region [ sqlMontaPadLeftCampoTexto ]
		/// <summary>
		/// Monta uma expressão SQL (sintaxe do SQL Server) para realizar a função de PadLeft() em um campo do tipo texto
		/// </summary>
		/// <param name="nomeCampo">
		/// Nome do campo no banco de dados
		/// </param>
		/// <param name="preenchimento">
		/// Caracter para preenchimento no padding
		/// </param>
		/// <param name="tamanhoCampo">
		/// Tamanho que o texto deve ficar após execução do padding
		/// </param>
		/// <returns>
		/// Expressão SQL (sintaxe do SQL Server) para realizar a função PadLeft()
		/// </returns>
		public static String sqlMontaPadLeftCampoTexto(String nomeCampo, char preenchimento, int tamanhoCampo)
		{
			String strResp;
			strResp = " Coalesce(Replicate('" + preenchimento + "'," + tamanhoCampo.ToString() + "-Len(" + nomeCampo + ")), '') + " + nomeCampo;
			return strResp;
		}
		#endregion

		#region [ stEntregaPedidoCor ]
		/// <summary>
		/// Obtém a cor de exibição do status de entrega do pedido
		/// </summary>
		/// <param name="status">
		/// Código do status de entrega do pedido
		/// </param>
		/// <returns>
		/// Retorna uma cor para exibição do status de entrega do pedido
		/// </returns>
		public static Color stEntregaPedidoCor(String status)
		{
			Color cor = Color.Black;

			if (status == null) return cor;
			status = status.Trim();
			if (status.Length == 0) return cor;

			if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ESPERAR))
				cor = Color.DeepPink;
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SPLIT_POSSIVEL))
				cor = Color.DarkOrange;
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SEPARAR))
				cor = Color.Maroon;
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR))
				cor = Color.Blue;
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
				cor = Color.Green;
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
				cor = Color.Red;

			return cor;
		}
		#endregion

		#region [ stEntregaPedidoCor ]
		/// <summary>
		/// Obtém a cor de exibição do status de entrega do pedido
		/// </summary>
		/// <param name="status">
		/// Código do status de entrega do pedido
		/// </param>
		/// <param name="qtdeItensDevolvidos">
		/// Quantidade de itens devolvidos que o pedido já teve
		/// </param>
		/// <returns>
		/// Retorna uma cor para exibição do status de entrega do pedido
		/// </returns>
		public static Color stEntregaPedidoCor(String status, int qtdeItensDevolvidos)
		{
			Color cor;

			cor = stEntregaPedidoCor(status);

			if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
			{
				if (qtdeItensDevolvidos > 0) cor = Color.Red;
			}
			return cor;
		}
		#endregion

		#region [ stEntregaPedidoDescricao ]
		/// <summary>
		/// Obtém a descrição do status de entrega do pedido
		/// </summary>
		/// <param name="status">
		/// Código do status de entrega do pedido
		/// </param>
		/// <returns>
		/// Retorna uma descrição do status de entrega do pedido
		/// </returns>
		public static String stEntregaPedidoDescricao(String status)
		{
			String strResp = "";

			if (status == null) return "";
			status = status.Trim();
			if (status.Length == 0) return "";

			if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ESPERAR))
				strResp = "Esperar Mercadoria";
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SPLIT_POSSIVEL))
				strResp = "Split Possível";
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SEPARAR))
				strResp = "Separar Mercadoria";
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR))
				strResp = "A Entregar";
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
				strResp = "Entregue";
			else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
				strResp = "Cancelado";
			else
				strResp = "Desconhecido (" + status + ")";

			return strResp;
		}
		#endregion

		#region [ textBoxPosicionaCursorNoFinal ]
		public static void textBoxPosicionaCursorNoFinal(object sender)
		{
			TextBox c;
			c = (System.Windows.Forms.TextBox)sender;
			c.SelectionLength = 0;
			if (c.Text.Length > 0) c.SelectionStart = c.Text.Length;
		}
		#endregion

		#region [ textBoxSelecionaConteudo ]
		public static void textBoxSelecionaConteudo(object sender)
		{
			((System.Windows.Forms.TextBox)sender).Select(0, ((System.Windows.Forms.TextBox)sender).Text.Length);
		}
		#endregion

		#region [ trataComboBoxKeyDown ]
		/// <summary>
		/// Trata o evento KeyDown de um campo ComboBox
		/// </summary>
		/// <param name="sender">
		/// O próprio parâmetro "sender" do evento "KeyDown"
		/// </param>
		/// <param name="e">
		/// O próprio parâmetro "e" do evento "KeyDown"
		/// </param>
		/// <param name="proximo">
		/// O próximo para o qual deve ser passado o foco no caso de teclar "Enter" no campo atual
		/// </param>
		public static void trataComboBoxKeyDown(object sender, KeyEventArgs e, Control proximo)
		{
			ComboBox cb = null;

			if (sender.GetType() == typeof(ComboBox)) cb = (ComboBox)sender;

			#region [ Enter ]
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (proximo != null) proximo.Focus();
				return;
			}
			#endregion

			#region [ Delete ]
			if (e.KeyCode == Keys.Delete)
			{
				e.SuppressKeyPress = true;
				if (cb != null)
				{
					if (cb.DroppedDown) cb.DroppedDown = false;
					cb.SelectedIndex = -1;
				}
				return;
			}
			#endregion
		}
		#endregion

		#region [ trataTextBoxKeyDown ]
		/// <summary>
		/// Trata o evento KeyDown de um campo TextBox
		/// </summary>
		/// <param name="sender">
		/// O próprio parâmetro "sender" do evento "KeyDown"
		/// </param>
		/// <param name="e">
		/// O próprio parâmetro "e" do evento "KeyDown"
		/// </param>
		/// <param name="proximo">
		/// O próximo para o qual deve ser passado o foco no caso de teclar "Enter" no campo atual
		/// </param>
		public static void trataTextBoxKeyDown(object sender, KeyEventArgs e, Control proximo)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				proximo.Focus();
				return;
			}
		}
		#endregion

		#endregion
	}
}
