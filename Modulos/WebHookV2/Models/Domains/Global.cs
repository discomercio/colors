using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using System.Xml.Serialization;

namespace WebHookV2.Models.Domains
{
	public static class Global
	{
		#region [ Construtor estático ]
		static Global()
		{
			#region [ Declarações ]
			string strValue;
			string msg_erro;
			#endregion

			if ((Global.Cte.LogAtividade.PathLogAtividade ?? "").Length > 0)
			{
				if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);
			}

			gravaLogAtividade(new string('=', 80));
			gravaLogAtividade(Cte.Versao.M_ID);
			gravaLogAtividade(new string('=', 80));

			#region [ Configuração do parâmetro que define o tratamento para evitar acesso concorrente nas operações com o BD ]
			strValue = getConfigurationValue("TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO");
			if ((strValue ?? "").Length > 0)
			{
				if (strValue.ToUpper().Equals("TRUE")
					|| strValue.ToUpper().Equals("1")
					|| strValue.ToUpper().Equals("YES")
					|| strValue.ToUpper().Equals("Y")
					|| strValue.ToUpper().Equals("SIM")
					|| strValue.ToUpper().Equals("S"))
				{
					Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = true;
				}
				else
				{
					Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = false;
				}
			}
			#endregion

			executaManutencaoArqLogAtividade(out msg_erro);
		}
		#endregion

		#region [ Constantes ]
		public static class Cte
		{
			#region [ Versão ]
			public static class Versao
			{
				public const string NomeSistema = "WebHookV2";
				public const string Numero = "1.00";
				public const string Data = "30.JUN.2021";
				public const string M_ID = NomeSistema + " - " + Numero + " - " + Data;
			}
			#endregion

			#region [ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 30.06.2021 - por HHO
			 *    Início.
			 *    Esta nova API foi criada devido à desativação do antigo serviço de 2º post da Braspag e
			 *    substituição pelo Post de Notificação.
			 *    Detalhes técnicos em:
			 *    https://braspag.github.io/manual/braspag-pagador?json#post-de-notifica%C3%A7%C3%A3o
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.02 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.03 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.04 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.05 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.06 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.07 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - XX.XX.20XX - por XXX
			 *    
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - XX.XX.20XX - por XXX
			 *    
			* ================================================================================================
			*/
			#endregion

			#region [ Usuario ]
			public static class Usuario
			{
				public const string ID_USUARIO_SISTEMA = "SISTEMA";
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
				public const int TAM_MAX_NSU = 12;
			}
			#endregion

			#region [ Sistema Responsável Cadastro ]
			public class SistemaResponsavelCadastro
			{
				public const int COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP = 1;
				public const int COD_SISTEMA_RESPONSAVEL_CADASTRO__ITS = 2;
				public const int COD_SISTEMA_RESPONSAVEL_CADASTRO__UNIS = 3;
				public const int COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP_WEBAPI = 4;
			}
			#endregion

			#region [ Log ]
			public static class LogAtividade
			{
				// System.Reflection.Assembly.GetExecutingAssembly().CodeBase retorna o nome do arquivo, ex: file:///C:/inetpub/wwwroot/Teste/WebAPI/bin/WebAPI.DLL
				public static string PathLogAtividade = Path.GetDirectoryName(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).Substring(6)) + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 365;
				public const string ExtensaoArqLog = "LOG";
			}
			#endregion

		}
		#endregion

		#region [ Parâmetros ]
		public static class Parametros
		{
			#region [ Geral ]
			public static class Geral
			{
				public static bool TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = false;
			}
			#endregion

			#region [ Braspag ]
			public static class Braspag
			{
				public static readonly string BraspagWebHookV2Empresa = getConfigurationValue("BraspagWebHookV2Empresa");
			}
			#endregion
		}
		#endregion

		#region[ ReaderWriterLock ]
		public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
		#endregion

		#region [ Métodos ]

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

		#region[ formataDuracaoHMSMs ]
		public static string formataDuracaoHMSMs(TimeSpan ts)
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
			// Milissegundos
			sb.Append(ts.Milliseconds.ToString() + "ms");

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

		#region [ serializaObjectToXml ]
		public static string serializaObjectToXml(object obj)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Global.serializaObjectToXml()";
			XmlSerializer xmlWriter;
			StringWriter stringWriter = new System.IO.StringWriter();
			#endregion

			if (obj == null) return "";

			try
			{
				xmlWriter = new XmlSerializer(obj.GetType());
				xmlWriter.Serialize(stringWriter, obj);
				return stringWriter.ToString();
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception\n" + ex.ToString());
				return "";
			}
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

		#endregion
	}
}