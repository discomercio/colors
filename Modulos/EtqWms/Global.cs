#region [ using ]
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
using System.Collections.Specialized;
#endregion

namespace EtqWms
{
	class Global
	{
		#region [ Constantes ]
		public class Cte
		{
			#region[ Versão do Aplicativo ]
			public class Aplicativo
			{
				public const string NOME_OWNER = "Artven";
				public const string NOME_SISTEMA = "EtqWms";
				public const string VERSAO_NUMERO = "1.12";
				public const string VERSAO_DATA = "31.AGO.2020";
				public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
				public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
				public const string M_DESCRICAO = "Módulo Etiqueta (WMS)";
			}
			#endregion

			#region[ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 01.10.2013 - por HHO
			 *        Início.
			 *        Este programa realiza a impressão das etiquetas c/ os dados do destinatário e que
			 *        serão coladas nas caixas dos produtos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - 24.10.2013 - por HHO
			 *		  Inclusão de impressão de etiqueta de um volume específico (ou faixa de volumes). Isso
			 *		  se mostrou necessário devido a pedidos que vendem grandes quantidades de um produto
			 *		  (ex: 50 unidades de um produto composto por 2 volumes). Caso fosse necessário reimprimir
			 *		  uma única etiqueta, isso não seria possível, pois a impressão parcial faria a impressão
			 *		  de 100 etiquetas.
			 *		  
			 *		  Adicionada a informação referente a campo t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO.id
			 *		  para ser impressa junto (logo após) os campos t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO.id
			 *		  e Nº Sequência
			 *		  
			 *		  O formato fica assim:
			 *		  AAA-BBBB / CCCC
			 *		  AAA = t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO.id
			 *		  BBBB = Nº Sequência
			 *		  CCCC = t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO.id
			 *		  
			 *		  A inclusão dessa informação é para viabilizar a edição dos dados da etiqueta através
			 *		  da página ASP desenvolvida para isso.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.02 - 27.11.2013 - por HHO
			 *		  Houve a necessidade de criar um novo campo ('obs_3') no pedido p/ armazenar o nº da NF
			 *		  de simples remessa, sendo que é essa a NF que acompanha o pedido no transporte.
			 *		  Portanto, sempre que esse campo estiver preenchido, deve ser usado p/ imprimir o nº da
			 *		  NF na etiqueta, caso contrário, será o conteúdo do campo 'obs_2'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.03 - 21.02.2014 - por HHO
			 *		  Implementação de consistência p/ impedir a impressão de etiquetas de pedidos que este-
			 *		  jam c/ os seguintes status de entrega: A Entregar, Entregue e Cancelado.
			 *		  O objetivo é tentar minimizar os erros de operação que ocasionam o envio de mercadorias
			 *		  em duplicidade.
			 *		  Na operação de reimpressão por nº do volume, foi corrigida a lógica ao verificar se
			 *		  o nº do volume escolhido p/ reimpressão está dentro do intervalo válido.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.04 - 28.11.2014 - por HHO
			 *		  Inclusão de uma nova linha na etiqueta para imprimir a UF e cidade do destino da 
			 *		  entrega.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.05 - 23.01.2015 - por HHO
			 *		  Alteração dos dados de conexão ao BD devido à migração do servidor, pois o SQL Server
			 *		  não está mais usando a porta padrão por questões de segurança.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.06 - 04.10.2015 - por LHGX
			 *		  Implementação do funcionamento em diferentes Centros de Distribuição (CD)
			 * -----------------------------------------------------------------------------------------------
			 * v 1.07 - 27.01.2016 - por LHGX
			 *		  Alterações referentes ao desmembramento da tabela t_FIN_BOLETO_CEDENTE em
			 *		  t_NFE_EMITENTE
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - 18.04.2016 - por HHO
			 *		  Implementação de ajustes para alterar a cor de fundo dos painéis de acordo com o
			 *		  ambiente acessado.
			 *		  A cor inicialmente é obtida a partir do arquivo de configuração e, após realizar a
			 *		  conexão com o banco de dados, a cor é obtida através do campo 'cor_fundo_padrao' da
			 *		  tabela t_VERSAO. Caso a cor definida no banco de dados seja diferente da do arquivo,
			 *		  o parâmetro do arquivo é atualizado para respeitar a cor especificada no BD.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - 05.05.2016 - por LHGX
			 *		  Alteração para funcionamento em diversos ambientes (entrada da DIS)
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - 31.03.2017 - por LHGX
			 *		  Remoção do controle antigo por CD's
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - 10.08.2020 - por HHO
			 *		  Impressão do código de barras na etiqueta com as informações: nº NF, qtde total de
			 *		  volumes, ID da etiqueta (t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO.id),
			 *		  número do volume (individual).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - 31.08.2020 - por HHO
			 *		  Ajustes para tratar a memorização do endereço de cobrança no pedido, pois, a partir de
			 *		  agora, ao invés de obter os dados do endereço no cadastro do cliente (t_CLIENTE), deve-se
			 *		  usar os dados que estão gravados no próprio pedido. O tratamento que já ocorria com o
			 *		  endereço de entrega deve passar a ser feito p/ o endereço de cobrança/cadastro.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - XX.XX.20XX - por XXX
			 *		  
			 * ===============================================================================================
			 */
			#endregion

			#region [ Etc ]
			public class Etc
			{
				public const String SIMBOLO_MONETARIO = "R$";
				public const byte FLAG_NAO_SETADO = 255;
				public const int TAM_MIN_LOJA = 2;
				public const int TAM_MIN_NUM_PEDIDO = 6;    // SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
				public const int TAM_MIN_ID_PEDIDO = 7; // PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
				public const char COD_SEPARADOR_FILHOTE = '-';
				public const int MAX_TAM_BOLETO_CAMPO_ENDERECO = 40;
				public const int MAX_TAM_BOLETO_CAMPO_NOME_SACADO = 40;
				public const String ID_PF = "PF";
				public const String ID_PJ = "PJ";
				public const String PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE = "TFBI";
				public const String SQL_COLLATE_CASE_ACCENT = " COLLATE Latin1_General_CI_AI";
			}
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
				public const string FmtDdMmYyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAnoCom2Digitos;
				public const string FmtDdMmYyyyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno;
				public const string FmtDdMmYyyyHhMmComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin;
				public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
				public const string FmtYyyyMmDdHhMmSsComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
			}
			#endregion

			#region [ ID_T_PARAMETRO ]
			public static class ID_T_PARAMETRO
			{
				public const string ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS = "Flag_Pedido_MemorizacaoCompletaEnderecos";
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

			#region [ Classe EtqWms ]
			public class EtqWms
			{
				#region [ LogOperacao - Códigos de operação para o log ]
				public class LogOperacao
				{
					// Texto com 20 posições
					public const String LOGON = "EtqWms-Logon";
					public const String LOGOFF = "EtqWms-Logoff";
					public const String ETIQUETA_WMS_IMPRESSAO_COMPLETA = "EtqWms-Completa";
					public const String ETIQUETA_WMS_IMPRESSAO_PARCIAL = "EtqWms-Parcial";
					public const String ETIQUETA_WMS_IMPRESSAO_VOLUME = "EtqWms-Volume";
					public const String RECONEXAO_BD = "EtqWms-Reconexao-BD";
				}
				#endregion
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
		public static String strModoSelecao = "";
		#endregion

		#region [ Classe Acesso ]
		public class Acesso
		{
			#region [ Constantes ]
			public const String OP_CEN_ETQWMS_APP_ETIQUETA_WMS_ACESSO_AO_MODULO = "25500";
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
			public static String emit;
			public static String emit_uf;
			public static String emit_id;
			public static String txtEspecifico;

			public struct InfoEmitentes
			{
				public String emit;
				public String emit_uf;
				public String emit_id;
				public String emit_texto_especifico;
				public string emitente_nome;
			}

			public static List<InfoEmitentes> listaEmitentes = new List<InfoEmitentes>();


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
				public static String usuEmit = "UsuEmit";
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
		public static int calculaTimeSpanMiliSegundos(TimeSpan ts)
		{
			return ts.Milliseconds + 1000 * (ts.Seconds + (60 * (ts.Minutes + (60 * (ts.Hours + (24 * ts.Days))))));
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
		public static int calculaTimeSpanSegundos(TimeSpan ts)
		{
			return ts.Seconds + (60 * (ts.Minutes + (60 * (ts.Hours + (24 * ts.Days)))));
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

		#region[ converteDdMmYyyyHhMmSsParaDateTime ]
		public static DateTime converteDdMmYyyyHhMmSsParaDateTime(string strDdMmYyyyHhMmSs)
		{
			string strFormato;
			string strDigitosDdMmYyyyHhMmSs;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			strDigitosDdMmYyyyHhMmSs = digitos(strDdMmYyyyHhMmSs);
			if (strDigitosDdMmYyyyHhMmSs.Length == 12)
			{
				strFormato = Cte.DataHora.FmtDia +
							 Cte.DataHora.FmtMes +
							 Cte.DataHora.FmtAno +
							 Cte.DataHora.FmtHora +
							 Cte.DataHora.FmtMin;
			}
			else
			{
				strFormato = Cte.DataHora.FmtDia +
							 Cte.DataHora.FmtMes +
							 Cte.DataHora.FmtAno +
							 Cte.DataHora.FmtHora +
							 Cte.DataHora.FmtMin +
							 Cte.DataHora.FmtSeg;
			}
			if (DateTime.TryParseExact(strDigitosDdMmYyyyHhMmSs, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
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

		#region [ filtraDigitacaoNumeroPedido ]
		public static char filtraDigitacaoNumeroPedido(char c)
		{
			char letra;
			if (c == '\b') return c;
			letra = Char.ToUpper(c);
			if ((!isDigit(letra)) && (!isLetra(letra)) && (letra != Cte.Etc.COD_SEPARADOR_FILHOTE)) c = '\0';
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

		#region [ isStEntregaBloqueadoParaImpressaoEtiqueta ]
		public static bool isStEntregaPedidoBloqueadoParaImpressaoEtiqueta(string st_entrega)
		{
			if (st_entrega == null) return false;

			if (st_entrega.Equals(Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR)) return true;
			if (st_entrega.Equals(Cte.StEntregaPedido.ST_ENTREGA_CANCELADO)) return true;
			if (st_entrega.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE)) return true;

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
