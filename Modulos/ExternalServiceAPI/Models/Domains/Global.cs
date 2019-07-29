#region [ usings ]
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
#endregion

namespace ExternalServiceAPI.Models.Domains
{
	public static class Global
	{
		#region [ Construtor estático ]
		static Global()
		{
			#region [ Declarações ]
			string msg_erro;
			#endregion

			gravaLogAtividade(Cte.Aplicativo.M_ID);
			executaManutencaoArqLogAtividade(out msg_erro);
		}
		#endregion

		#region [ Constantes ]
		public class Cte
		{
			#region[ Versão da API ]
			public class Aplicativo
			{
				public const string NOME_OWNER = "Artven";
				public const string NOME_SISTEMA = "ExternalServiceAPI";
				public const string VERSAO_NUMERO = "1.02";
				public const string VERSAO_DATA = "04.AGO.2018";
				public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
				public const string M_ID = NOME_SISTEMA + " - " + VERSAO;
				public const string M_DESCRICAO = "";
			}
			#endregion

			#region [ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 02.08.2018 - por HHO
			 *		Início: desenvolvimento de WebAPI exclusivamente para prestar serviços a terceiros.
			 *		Neste primeiro momento, a funcionalidade disponibilizada é uma consulta para validar
			 *		o dígito verificador da IE (inscrição estadual).
			 *		O principal objetivo de segregar esse tipo de funcionalidade em uma WebAPI específica
			 *		é poder restringir o acesso a IP's pré-cadastrados na configuração do IIS.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - 03.08.2018 - por HHO
			 *		Implementação de gravação do log de atividade em arquivo, registrando as requisições
			 *		e o respectivo IP de origem.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.02 - 04.08.2018 - por HHO
			 *		Ajustes no registro do log de atividade.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.03 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.04 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.05 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.06 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.07 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - XX.XX.20XX - por XXX
			 *		
			 *		
			 * -----------------------------------------------------------------------------------------------
			 * v 1.15 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.18 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.19 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.20 - XX.XX.20XX - por XXX
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
				public const string FmtDdMmYyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAnoCom2Digitos;
				public const string FmtDdMmYyyyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno;
				public const string FmtDdMmYyyyHhMmComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin;
				public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
				public const string FmtYyyyMmDdHhMmSsComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
			}
			#endregion

			#region [ LogAtividade ]
			public static class LogAtividade
			{
				// System.Reflection.Assembly.GetExecutingAssembly().CodeBase retorna o nome do arquivo, ex: file:///C:/inetpub/wwwroot/Teste/WebAPI/bin/WebAPI.DLL
				public static string PathLogAtividade = Path.GetDirectoryName(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).Substring(6)) + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 90;
				public const string ExtensaoArqLog = "LOG";
			}
			#endregion
		}
		#endregion

		#region[ ReaderWriterLock ]
		public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
		#endregion

		#region [ Métodos ]

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
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception: " + ex.Message);

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

		#endregion
	}
}