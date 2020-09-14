using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models;
using System.Threading;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Configuration;

namespace ART3WebAPI.Models.Domains
{
	public static class Global
	{
		#region [ Construtor estático ]
		static Global()
		{
			#region [ Declarações ]
			string msg_erro;
			#endregion

			gravaLogAtividade(Cte.Versao.M_ID);
			executaManutencaoArqLogAtividade(out msg_erro);
		}
		#endregion

		#region[ Constantes ]
		public static class Cte
		{
			#region [ Versao ]
			public static class Versao
			{
				public const string NomeSistema = "WebAPI";
				public const string Numero = "2.23";
				public const string Data = "12.SET.2020";
				public const string M_ID = NomeSistema + " - " + Numero + " - " + Data;
			}
			#endregion

			#region [ Comentário sobre as versões ]
			/*================================================================================================
			 * v 2.00 - 08.09.2017 - por TRR
			 *		Ajuste no relatório Farol para aceitar lista de lojas no filtro.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.01 - 09.09.2017 - por HHO
			 *		Correção do relatório Farol para incluir a restrição do filtro por loja ao calcular a
			 *		quantidade de produtos devolvidos no período.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.02 - 07.11.2017 - por TRR
			 *		Implementação de API para atualizar os dados na tabela t_RELATORIO_PRODUTO_FLAG à medida
			 *		em que o usuário marca ou desmarca os produtos no Relatório de Estoque II para contabi-
			 *		lizar a quantidade dos produtos assinalados.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.03 - 09.01.2018 - por HHO
			 *		Implementação de tratamento para consultar pedidos do Magento através da API.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.04 - 11.01.2018 - por HHO
			 *		Ajustes nas mensagens de erro quando não há parâmetros de login cadastrados para a API do
			 *		Magento na consulta de pedidos.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.05 - 12.01.2018 - por HHO
			 *		Ajustes na consulta do pedido Magento: inclusão do nº fax na gravação dos dados no BD.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.06 - 03.03.2018 - por HHO
			 *		Ajustes na consulta do pedido Magento: inclusão dos campos 'product_type', 'has_children'
			 *		e 'parent_item_id' na gravação dos dados no BD devido a produtos como o ventilador.
			 *		O ventilador é cadastrado como um produto configurável, ou seja, um ventilador exibido
			 *		no site possui várias opções (ex: 110V e 220V). Esse tipo de produto é retornado na
			 *		consulta da API como dois itens (product_type 'configurable' e 'simple'). Deve-se observar
			 *		que os valores são retornados no item 'configurable' e a descrição específica no item
			 *		'simple'.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.07 - 10.03.2018 - por HHO
			 *		Ajuste em MagentoApiController para obter o primeiro número ao invés do segundo durante
			 *		a extração do nº pedido marketplace no caso do Walmart.
			 *		Ex: Skyhub code: Walmart-76954296-1796973
			 *			Retornar 76954296 ao invés de 1796973
			 * -----------------------------------------------------------------------------------------------
			 * v 2.08 - 18.04.2018 - por HHO
			 *		Ajuste em GetDataController para adicionar consulta somente pelo código do produto (sem
			 *		código do fabricante) e para incluir autenticação do usuário pela sessionToken.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.09 - 25.04.2018 - por HHO
			 *		Ajustes nas rotinas de upload de arquivos para tratar o novo campo
			 *		'st_confirmation_required' que indica se é necessária uma confirmação para considerar
			 *		o arquivo como sendo válido. O objetivo desse campo é identificar arquivos órfãos decor-
			 *		rentes de operações em que ocorreram erros antes de concluir a operação.
			 *		Arquivos que estejam com t_UPLOAD_FILE.st_confirmation_required = 1 precisam ter o campo
			 *		t_UPLOAD_FILE.st_confirmation_ok = 1, caso contrário, serão excluídos automaticamente.
			 *		Ajuste em UploadFileController.PostFile() para o nome do arquivo armazenado no servidor
			 *		ser no formato: yyyyMMdd_HHmmss_fff__GUID.ext, onde GUID é um global unique identifier
			 *		(ex: 926B1C67-85E1-4434-A06C-EFF1A36B40BE).
			 * -----------------------------------------------------------------------------------------------
			 * v 2.10 - 22.05.2018 - por HHO
			 *		Ajustes em MagentoApiController para salvar o nº de pedido marketplace completo devido
			 *		aos casos em que somente a parte significativa do número é salva no campo
			 *		'pedido_marketplace' (ex: Walmart).
			 *		Esta alteração foi realizada devido à possibilidade de se necessitar do nº completo caso
			 *		seja realizada uma integração com a Intelipost.
			 *		Implementação de tratamento para salvar o conteúdo de arquivos enviados via UploadFile
			 *		diretamente no banco de dados (t_UPLOAD_FILE.file_content ou
			 *		t_UPLOAD_FILE.file_content_text).
			 * -----------------------------------------------------------------------------------------------
			 * v 2.11 - 30.05.2018 - por TRR
			 *		Ajustes no relatório de Compras 2 para tratar opção de saída por nº NF.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.12 - 08.06.2018 - por TRR
			 *		Desenvolvimento da saída em Excel para o relatório Devolução de Produtos II.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.13 - 21.06.2018 - por HHO
			 *		Ajustes na consulta de dados do pedido Magento para armazenar informações necessárias
			 *		para a implantação da fase piloto no novo site da Bonshop. Ex: installer_document,
			 *		commission_value, etc
			 * -----------------------------------------------------------------------------------------------
			 * v 2.14 - 13.07.2018 - por HHO
			 *		Ajustes na consulta do pedido Magento para retornar os dados do usuário responsável pelo
			 *		pedido já cadastrado anteriormente no ERP, ao invés de retornar a identificação do
			 *		vendedor. Até então, o vendedor sempre era o responsável pelo cadastramento, mas a partir
			 *		da implantação do site Magento da Bonshop, o responsável pelo cadastramento pode ser
			 *		um operador.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.15 - 16.01.2019 - por HHO
			 *		Ajustes na consulta de dados do pedido Magento para interpretar os dados referentes ao
			 *		nº pedido marketplace devido às alterações provocadas pela atualização do módulo da
			 *		Skyhub.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.16 - 26.02.2019 - por HHO
			 *		Ajustes em MagentoApiController para retirar o sufixo do nº pedido marketplace dos
			 *		pedidos do Carrefour.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.17 - 11.07.2019 - por TRR
			 *		Correção do relatório Farol Resumido (FarolV3Controller e DataFarol.GetV3()).
			 * -----------------------------------------------------------------------------------------------
			 * v 2.18 - 30.10.2019 - por HHO
             *      Desenvolvimento de tratamento em MagentoApiController para pedidos do marketplace Leroy
             *      Merlin.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.19 - 21.01.2020 - por LHGX
             *      Desenvolvimento do download XLS do relatório de ocorrências.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.20 - 08.04.2020 - por HHO
			 *      Implementação do filtro para o campo 'subgrupo' no relatório Farol Resumido.
			 * -----------------------------------------------------------------------------------------------
			 * v 2.21 - 21.05.2020 - por HHO
			 *      Implementação da consulta por período de entrega no relatório Farol Resumido em
			 *      FarolV3Controller.GetXLSReport().
			 * -----------------------------------------------------------------------------------------------
			 * v 2.22 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.23 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.24 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.25 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.26 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.27 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.28 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.29 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.30 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 2.XX - XX.XX.20XX - por XXX
			* ===============================================================================================
			*/
			#endregion

			#region [ Usuario ]
			public static class Usuario
			{
				public const string ID_USUARIO_SISTEMA = "SISTEMA";
			}
			#endregion

			#region [ Ocorrências em pedidos ]
			public const string COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA = "CE->LJ";
			public const string GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA = "OcorrenciasEmPedidos_TipoOcorrencia";
			public const string GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA = "OcorrenciaPedido_MotivoAbertura";
			#endregion

			#region [ Parâmetros ]
			public static class Parametros
			{
				public static class ID_T_PARAMETRO
				{
					public const string FLAG_HABILITACAO_UPLOAD_FILE_BACKUP_RECENT_FILES = "WebAPI_UploadFile_FlagHabilitacao_BackupRecentFiles";
					public const string UPLOAD_FILE_SAVE_FILE_CONTENT_IN_DB_MAX_SIZE_IN_BYTES = "WebAPI_UploadFile_SaveFileContentInDb_MaxSizeInBytes";
					public const string UPLOAD_FILE_SAVE_FILE_CONTENT_IN_DB_AS_TEXT_MAX_SIZE_IN_CHARS = "WebAPI_UploadFile_SaveFileContentInDbAsText_MaxSizeInChars";
				}
			}
			#endregion

			#region [ MagentoSoapApi ]
			public static class MagentoSoapApi
			{
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				public static readonly int REQUEST_TIMEOUT_EM_MS = 3 * 60 * 1000;
				public static readonly int TIMEOUT_READER_WRITER_LOCK_EM_MS = 60 * 1000;

				public static readonly string TIPO_ENDERECO__COBRANCA = "COB";
				public static readonly string TIPO_ENDERECO__ENTREGA = "ETG";

				#region [ Transacao ]
				public sealed class Transacao
				{
					// Type safe enum pattern
					private readonly string methodName;
					private readonly string codOpLog;
					private readonly string soapAction;

					public static readonly Transacao login = new Transacao("login", "login", "urn:Mage_Api_Model_Server_HandlerAction");
					public static readonly Transacao call = new Transacao("call", "call", "urn:Mage_Api_Model_Server_HandlerAction");
					public static readonly Transacao endSession = new Transacao("endSession", "endSession", "urn:Mage_Api_Model_Server_HandlerAction");

					private Transacao(string methodName, string codOpLog, string soapAction)
					{
						this.methodName = methodName;
						this.codOpLog = codOpLog;
						this.soapAction = soapAction;
					}

					public string GetMethodName()
					{
						return methodName;
					}

					public string GetSoapAction()
					{
						return soapAction;
					}

					public string GetCodOpLog()
					{
						return codOpLog;
					}

					public override string ToString()
					{
						return codOpLog;
					}
				}
				#endregion
			}
			#endregion

			#region[ Data/Hora ]
			public static class DataHora
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
				public const string FmtHora12 = "hh";
				public const string FmtMin = "mm";
				public const string FmtSeg = "ss";
				public const string FmtMiliSeg = "fff";
				public const string FmtAmPm = "tt";
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
			public static class LogAtividade
			{
				// System.Reflection.Assembly.GetExecutingAssembly().CodeBase retorna o nome do arquivo, ex: file:///C:/inetpub/wwwroot/Teste/WebAPI/bin/WebAPI.DLL
				public static string PathLogAtividade = Path.GetDirectoryName(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).Substring(6)) + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 90;
				public const string ExtensaoArqLog = "LOG";
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

			#region [ Código/Descrição ]
			public static class CodigoDescricao
			{
				public const string PedidoECommerce_Origem = "PedidoECommerce_Origem";
				public const string PedidoECommerce_Origem_Grupo = "PedidoECommerce_Origem_Grupo";
			}
			#endregion

			#region [ Loja ]
			public static class Loja
			{
				public static readonly string ArClube = getConfigurationValue("LojaArclube");
				public static readonly string Bonshop = getConfigurationValue("LojaBonshop");
			}
			#endregion
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

		#region[ ReaderWriterLock ]
		public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
		#endregion

		#region [ arredondaParaMonetario ]
		public static decimal arredondaParaMonetario(decimal numero)
		{
			return converteNumeroDecimal(formataMoeda(numero));
		}
		#endregion

		#region[Asc]
		public static int Asc(string letra)
		{
			return (int)(Convert.ToChar(letra));
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

		#region[ converteDdMmYyyyParaDateTime ]
		public static DateTime converteDdMmYyyyParaDateTime(string strDdMmYyyy)
		{
			string strFormato;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			strFormato = DataHora.FmtDia +
						 DataHora.FmtMes +
						 DataHora.FmtAno;
			if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
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

		#region [ decodificaUfExtensoParaSigla ]
		public static string decodificaUfExtensoParaSigla(string ufExtenso)
		{
			#region [ Declarações ]
			string ufSigla = "";
			#endregion

			ufExtenso = (ufExtenso ?? "").Trim().ToUpper();
			ufExtenso = filtraAcentuacao(ufExtenso);

			if (ufExtenso.Length == 0) return "";

			if (ufExtenso.Equals("ACRE"))
				ufSigla = "AC";
			else if (ufExtenso.Equals("ALAGOAS"))
				ufSigla = "AL";
			else if (ufExtenso.Equals("AMAZONAS"))
				ufSigla = "AM";
			else if (ufExtenso.Equals("AMAPA"))
				ufSigla = "AP";
			else if (ufExtenso.Equals("BAHIA"))
				ufSigla = "BA";
			else if (ufExtenso.Equals("CEARA"))
				ufSigla = "CE";
			else if (ufExtenso.Equals("DISTRITO FEDERAL"))
				ufSigla = "DF";
			else if (ufExtenso.Equals("ESPIRITO SANTO"))
				ufSigla = "ES";
			else if (ufExtenso.Equals("GOIAS"))
				ufSigla = "GO";
			else if (ufExtenso.Equals("MARANHAO"))
				ufSigla = "MA";
			else if (ufExtenso.Equals("MINAS GERAIS"))
				ufSigla = "MG";
			else if (ufExtenso.Equals("MATO GROSSO DO SUL"))
				ufSigla = "MS";
			else if (ufExtenso.Equals("MATO GROSSO"))
				ufSigla = "MT";
			else if (ufExtenso.Equals("PARA"))
				ufSigla = "PA";
			else if (ufExtenso.Equals("PARAIBA"))
				ufSigla = "PB";
			else if (ufExtenso.Equals("PERNAMBUCO"))
				ufSigla = "PE";
			else if (ufExtenso.Equals("PIAUI"))
				ufSigla = "PI";
			else if (ufExtenso.Equals("PARANA"))
				ufSigla = "PR";
			else if (ufExtenso.Equals("RIO DE JANEIRO"))
				ufSigla = "RJ";
			else if (ufExtenso.Equals("RIO GRANDE DO NORTE"))
				ufSigla = "RN";
			else if (ufExtenso.Equals("RONDONIA"))
				ufSigla = "RO";
			else if (ufExtenso.Equals("RORAIMA"))
				ufSigla = "RR";
			else if (ufExtenso.Equals("RIO GRANDE DO SUL"))
				ufSigla = "RS";
			else if (ufExtenso.Equals("SANTA CATARINA"))
				ufSigla = "SC";
			else if (ufExtenso.Equals("SERGIPE"))
				ufSigla = "SE";
			else if (ufExtenso.Equals("SAO PAULO"))
				ufSigla = "SP";
			else if (ufExtenso.Equals("TOCANTINS"))
				ufSigla = "TO";

			return ufSigla;
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

		#region [ executaManutencaoArqLogAtividade ]
		/// <summary>
		/// Apaga os arquivos de log de atividade antigos
		/// </summary>
		public static bool executaManutencaoArqLogAtividade(out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaManutencaoArqLogAtividade()";
			String strMsg;
			DateTime dtCorte = DateTime.Now.AddDays(-Global.Cte.LogAtividade.CorteArqLogEmDias);
			string strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtYYYYMMDD);
			string[] ListaArqLog;
			string strNomeArq;
			int i;
			int intQtdeApagada = 0;
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			strMsgErro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada (data de corte: " + formataDataDdMmYyyyComSeparador(dtCorte) + ")";
				Global.gravaLogAtividade(strMsg);

				#region[ Apaga arquivos de log de atividade antigos ]
				ListaArqLog = Directory.GetFiles(Global.Cte.LogAtividade.PathLogAtividade, "*." + Global.Cte.LogAtividade.ExtensaoArqLog, SearchOption.TopDirectoryOnly);
				for (i = 0; i < ListaArqLog.Length; i++)
				{
					strNomeArq = Global.extractFileName(ListaArqLog[i]);
					strNomeArq = strNomeArq.Substring(0, strDataCorte.Length);
					if (string.Compare(strNomeArq, strDataCorte) < 0)
					{
						File.Delete(ListaArqLog[i]);
						intQtdeApagada++;
					}
				}
				#endregion

				strMsg = "Rotina " + NOME_DESTA_ROTINA + " concluída com sucesso: " + intQtdeApagada.ToString() + " arquivos excluídos (duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio) + ")";
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				return false;
			}
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
					if (leftStr(s, 3).Equals("000"))
					{
						s_aux = rightStr(s, 11);
						if (isCpfOk(s_aux)) s = s_aux;
					}
				}
			}
			#endregion

			// CPF
			if (s.Length == 11)
			{
				s_resp = s.Substring(0, 3) + '.' + s.Substring(3, 3) + '.' + s.Substring(6, 3) + '-' + s.Substring(9, 2);
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

		public static string formataDataDdMmYyyyComSeparador(DateTime? data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return ((DateTime)data).ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);
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

		public static String formataDataYyyyMmDdHhMmSsComSeparador(DateTime? data)
		{
			if (data == null) return "";
			return formataDataYyyyMmDdHhMmSsComSeparador((DateTime)data);
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
			strSeparadorDecimal = leftStr(rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ getConfigurationValue ]
		public static string getConfigurationValue(string key)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "getConfigurationValue()";
			string msg;
			System.Configuration.Configuration rootWebConfig1;
			#endregion

			try
			{
				rootWebConfig1 = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
				if (rootWebConfig1.AppSettings.Settings.Count > 0)
				{
					System.Configuration.KeyValueConfigurationElement customSetting = rootWebConfig1.AppSettings.Settings[key];
					if (customSetting != null)
					{
						msg = NOME_DESTA_ROTINA + " - Parâmetro '" + key + "' = " + customSetting.Value;
						gravaLogAtividade(msg);
						return customSetting.Value;
					}
				}
				msg = NOME_DESTA_ROTINA + " - Parâmetro '" + key + "' não encontrado no arquivo de configuração!";
				gravaLogAtividade(msg);
				return null;
			}
			catch (Exception ex)
			{
				msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
				gravaLogAtividade(msg);
				return null;
			}
		}
		#endregion

		#region [ getDescFabricante ]
		public static string getDescFabricante(string fabricante)
		{
			string descricao = "";
			SqlConnection cn = new SqlConnection(Repository.BD.getConnectionString());
			string sqlString;

			sqlString = "SELECT * FROM t_FABRICANTE WHERE (fabricante = '" + fabricante + "')";
			cn.Open();

			try
			{

				SqlCommand cmd = new SqlCommand(sqlString, cn);
				IDataReader reader = cmd.ExecuteReader();
				int idxDescricao = reader.GetOrdinal("razao_social");
				try
				{
					if (reader.Read())
					{
						descricao = reader.GetString(idxDescricao);
					}

				}
				finally
				{
					reader.Close();
				}
			}
			finally
			{
				cn.Close();
			}

			return descricao;
		}
		#endregion

		#region [ getDescProduto ]
		public static string getDescProduto(string produto)
		{
			string descricao = "";
			SqlConnection cn = new SqlConnection(Repository.BD.getConnectionString());
			string sqlString;

			sqlString = "SELECT * FROM t_PRODUTO WHERE (produto = '" + produto + "')";
			cn.Open();

			try
			{

				SqlCommand cmd = new SqlCommand(sqlString, cn);
				IDataReader reader = cmd.ExecuteReader();
				int idxDescricao = reader.GetOrdinal("descricao");
				try
				{
					if (reader.Read())
					{
						descricao = reader.GetString(idxDescricao);
					}

				}
				finally
				{
					reader.Close();
				}
			}
			finally
			{
				cn.Close();
			}

			return descricao;
		}
		#endregion

		#region [ getDetalhamento ]
		public static string getDetalhamento(string detalhamento)
		{
			string descricao = "";
			switch (detalhamento)
			{
				case "SINTETICO_FABR":
					descricao = descricao + "Sintético por Fabricante";
					break;
                case "SINTETICO_NF":
                    descricao = descricao + "Sintético por Nota Fiscal";
                    break;
                case "SINTETICO_PROD":
					descricao = descricao + "Sintético por Produto";
					break;
				case "CUSTO_MEDIO":
					descricao = descricao + "Valor Referência Médio";
					break;
				case "CUSTO_INDIVIDUAL":
					descricao = descricao + "Valor Referência Individual";
					break;
				default:
					descricao = "";
					break;
			}
			return descricao;

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

		#region[ isDigit ]
		public static bool isDigit(char c)
		{
			if ((c >= '0') && (c <= '9')) return true;
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

		#region [ leftStr ]
		/// <summary>
		/// Retorna a quantidade de caracteres especificada mais à esquerda do texto.
		/// Se o texto for null, retorna um texto vazio.
		/// Se a quantidade for menor ou igual a zero, retorna um texto vazio.
		/// Se a quantidade for maior que o tamanho do texto, retorna o próprio texto.
		/// </summary>
		/// <param name="texto">Texto a partir do qual será retornado um trecho</param>
		/// <param name="qtde">Quantidade de caracteres a ser retornada</param>
		/// <returns>
		/// Retorna um trecho mais à esquerda do texto especificado
		/// </returns>
		public static String leftStr(String texto, int qtde)
		{
			if (texto == null) return "";
			if (qtde <= 0) return "";

			if (qtde >= texto.Length) return texto;
			return texto.Substring(0, qtde);
		}
		#endregion

		#region [ mesPorExtenso ]
		public static string mesPorExtenso(int mes)
		{
			string mesPorExtenso;

			switch (mes)
			{
				case 1:
					mesPorExtenso = "jan";
					break;
				case 2:
					mesPorExtenso = "fev";
					break;
				case 3:
					mesPorExtenso = "mar";
					break;
				case 4:
					mesPorExtenso = "abr";
					break;
				case 5:
					mesPorExtenso = "mai";
					break;
				case 6:
					mesPorExtenso = "jun";
					break;
				case 7:
					mesPorExtenso = "jul";
					break;
				case 8:
					mesPorExtenso = "ago";
					break;
				case 9:
					mesPorExtenso = "set";
					break;
				case 10:
					mesPorExtenso = "out";
					break;
				case 11:
					mesPorExtenso = "nov";
					break;
				case 12:
					mesPorExtenso = "dez";
					break;
				default:
					mesPorExtenso = "";
					break;
			}
			return mesPorExtenso;
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

		#region [ obtemXmlChildNodeValue ]
		public static string obtemXmlChildNodeValue(XmlNode xmlNode, string xmlNodeName)
		{
			return obtemXmlChildNodeValue(xmlNode, xmlNodeName, "");
		}

		public static string obtemXmlChildNodeValue(XmlNode xmlNode, string nodeName, string valorDefault)
		{
			string strResp;

			if (xmlNode == null) return valorDefault;
			if (xmlNode.ChildNodes.Count == 0) return valorDefault;
			try
			{
				strResp = xmlNode[nodeName].InnerText;
			}
			catch (Exception)
			{
				return valorDefault;
			}

			return strResp;
		}
		#endregion

		#region [ retornaEcDescricaoState ]
		public static string retornaEcDescricaoState(string state, string loja)
		{
			#region [ Declarações ]
			string sResp;
			#endregion

			sResp = state;

			if (loja.Equals(Cte.Loja.ArClube))
			{
				#region [ State (Arclube) ]
				switch (state)
				{
					case "separando":
						sResp = "separando";
						break;
					case "processing":
						sResp = "Processando";
						break;
					case "pending_payment":
						sResp = "Pagamento Pendente";
						break;
					case "payment_review":
						sResp = "Análise de Pagamento";
						break;
					case "new":
						sResp = "Novo";
						break;
					case "holded":
						sResp = "Segurado";
						break;
					case "complete":
						sResp = "Completo";
						break;
					case "closed":
						sResp = "Fechado";
						break;
					case "canceled":
						sResp = "Cancelado";
						break;
					default:
						sResp = state;
						break;
				}
				#endregion
			}

			return sResp;
		}
		#endregion

		#region [ retornaEcDescricaoStatus ]
		public static string retornaEcDescricaoStatus(string status, string loja)
		{
			#region [ Declarações ]
			string sResp;
			#endregion

			sResp = status;

			if (loja.Equals(Cte.Loja.ArClube))
			{
				#region [ Status (Arclube) ]
				switch (status)
				{
					case "separando":
						sResp = "Liberar";
						break;
					case "pagto_aprovado_integra":
						sResp = "Pagto Aprovado IC";
						break;
					case "boleto_pago":
						sResp = "Pago";
						break;
					case "processing":
						sResp = "Boleto Emitido";
						break;
					case "separando2":
						sResp = "Separando";
						break;
					case "aguardando_nf_ic":
						sResp = "Aguardando NF IC";
						break;
					case "pgto_auth":
						sResp = "Pagamento Autorizado";
						break;
					case "pending_payment":
						sResp = "Análise de Pagamento";
						break;
					case "payment_review":
						sResp = "Análise de Pagamento";
						break;
					case "fraud":
						sResp = "Suspected Fraud";
						break;
					case "pending":
						sResp = "Pedido Realizado";
						break;
					case "holded":
						sResp = "On Hold";
						break;
					case "delivered":
						sResp = "Entregue Integracommerce";
						break;
					case "despachado":
						sResp = "Enviado";
						break;
					case "rastreio_ic":
						sResp = "Rastreio IC";
						break;
					case "shipexception":
						sResp = "Falha no Envio Integracommerce";
						break;
					case "complete":
						sResp = "Completo";
						break;
					case "closed":
						sResp = "Estornado";
						break;
					case "canceled":
						sResp = "Cancelado";
						break;
					case "ip_in_transit":
						sResp = "Em trânsito";
						break;
					case "ip_delivered":
						sResp = "Entregue";
						break;
					case "ip_shipped":
						sResp = "Despachado";
						break;
					case "paypal_canceled_reversal":
						sResp = "PayPal Canceled Reversal";
						break;
					case "ip_delivery_failed":
						sResp = "Entrega Falhou";
						break;
					case "pending_paypal":
						sResp = "Pending PayPal";
						break;
					case "ip_to_be_delivered":
						sResp = "Saiu para Entrega";
						break;
					case "aprovado":
						sResp = "Pagamento Aprovado";
						break;
					case "paypal_reversed":
						sResp = "PayPal Reversed";
						break;
					case "ip_delivery_late":
						sResp = "Atraso na entrega";
						break;
					default:
						sResp = status;
						break;
				}
				#endregion
			}

			return sResp;
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

		#region [ rightStr ]
		/// <summary>
		/// Retorna a quantidade de caracteres especificada mais à direita do texto.
		/// Se o texto for null, retorna um texto vazio.
		/// Se a quantidade for menor ou igual a zero, retorna um texto vazio.
		/// Se a quantidade for maior que o tamanho do texto, retorna o próprio texto.
		/// </summary>
		/// <param name="texto">Texto a partir do qual será retornado um trecho</param>
		/// <param name="qtde">Quantidade de caracteres a ser retornada</param>
		/// <returns>
		/// Retorna um trecho mais à direita do texto especificado
		/// </returns>
		public static String rightStr(String texto, int qtde)
		{
			if (texto == null) return "";
			if (qtde <= 0) return "";

			if (qtde >= texto.Length) return texto;
			return texto.Substring(texto.Length - qtde, qtde);
		}
		#endregion

		#region [ setDefaultBD ]
		public static bool setDefaultBD(string usuario, string nome_chave, string valor_texto)
		{
			string sqlString;
			bool salvou = false;
			SqlCommand cmdInsere;
			SqlConnection cn = new SqlConnection(Repository.BD.getConnectionString());

			sqlString = "SELECT * FROM t_DEFAULT WHERE (usuario = '" + usuario + "') AND (nome_chave = '" + nome_chave + "')";
			cn.Open();

			try
			{

				SqlCommand cmd = new SqlCommand(sqlString, cn);
				IDataReader reader = cmd.ExecuteReader();
				try
				{
					if (reader.Read())
					{
						cmdInsere = new SqlCommand("UPDATE t_DEFAULT SET valor_default_texto = @valor_texto, dt_hr_ult_atualizacao = GETDATE() WHERE usuario=@usuario AND nome_chave=@nome_chave", cn);
						cmdInsere.Parameters.AddWithValue("@usuario", usuario);
						cmdInsere.Parameters.AddWithValue("@nome_chave", nome_chave);
						cmdInsere.Parameters.AddWithValue("@valor_texto", valor_texto);
					}
					else
					{
						cmdInsere = new SqlCommand("INSERT INTO t_DEFAULT (usuario, nome_chave, valor_default_texto, dt_hr_cadastro, dt_hr_ult_atualizacao)" +
							" VALUES (@usuario, @nome_chave, @valor_texto, GETDATE(), GETDATE())", cn);
						cmdInsere.Parameters.AddWithValue("@usuario", usuario);
						cmdInsere.Parameters.AddWithValue("@nome_chave", nome_chave);
						cmdInsere.Parameters.AddWithValue("@valor_texto", valor_texto);
					}
				}
				finally
				{
					reader.Close();
				}

				int i = cmdInsere.ExecuteNonQuery();
				salvou = i > 0;
			}
			finally
			{
				cn.Close();
			}

			return salvou;
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTime ]
		public static string sqlMontaDateTimeParaSqlDateTime(DateTime dtReferencia)
		{
			string strDataHora;
			string strSql;

			if (dtReferencia == null) return "NULL";
			if (dtReferencia == DateTime.MinValue) return "NULL";

			strDataHora = dtReferencia.ToString(DataHora.FmtAno) +
						  "-" +
						  dtReferencia.ToString(DataHora.FmtMes) +
						  "-" +
						  dtReferencia.ToString(DataHora.FmtDia) +
						  " " +
						  dtReferencia.ToString(DataHora.FmtHora) +
						  ":" +
						  dtReferencia.ToString(DataHora.FmtMin) +
						  ":" +
						  dtReferencia.ToString(DataHora.FmtSeg);
			strSql = "Convert(datetime, '" + strDataHora + "', 120)";
			return strSql;
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTimeSomenteData ]
		public static string sqlMontaDateTimeParaSqlDateTimeSomenteData(DateTime dtReferencia)
		{
			string strData;
			string strSql;
			strData = dtReferencia.ToString(DataHora.FmtAno) +
					  "-" +
					  dtReferencia.ToString(DataHora.FmtMes) +
					  "-" +
					  dtReferencia.ToString(DataHora.FmtDia);
			strSql = "Convert(datetime, '" + strData + "', 120)";
			return strSql;
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
			if ((strAlias.Trim().Length == 0) && (strNomeCampo.IndexOf('(') == -1)) strAlias = strNomeCampo;
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
			strResposta = "Convert(datetime, Convert(varchar(10), getdate(), 121), 121)";
			return strResposta;
		}
		#endregion

		#region [ sqlMontaGetdateSomenteHora ]
		/// <summary>
		/// Monta uma expressão para obter a hora do Sql Server com a hora no formato hh:mm:ss, sem a parte da data
		/// </summary>
		/// <returns>
		/// Retorna uma expressão para obter a hora do Sql Server com a hora no formato hh:mm:ss, sem a parte da data
		/// </returns>
		public static string sqlMontaGetdateSomenteHora()
		{
			string strResposta;
			strResposta = "Convert(varchar(8), getdate(), 108)";
			return strResposta;
		}
		#endregion
	}
}
 