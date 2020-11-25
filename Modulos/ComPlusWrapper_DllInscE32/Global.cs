#region [ using ]
using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Threading;
#endregion

namespace ComPlusWrapper_DllInscE32
{
	#region[ Global ]
	class Global
	{
		#region[ Cte ]
		public class Cte
		{
			#region[ Versão ]
			public class Versao
			{
				public const string strNomeSistema = "ComPlusWrapper_DllInscE32";
				public const string strVersao = "1.01 - 10.NOV.2019";

				#region[ Comentário sobre as versões ]
				/*================================================================================================
				 * v 1.00 - 01.12.2010 - por HHO
				 *        Início.
				 * -----------------------------------------------------------------------------------------------
				 * v 1.01 - 10.11.2019 - por HHO
				 *	Recompilação do projeto usando Visual Studio 2015 e usando .NET 4.6.1 ao invés do Visual
				 *	Studio 2005 e .NET 2.0
				 * -----------------------------------------------------------------------------------------------
				 * v 1.02 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.03 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.04 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.05 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.06 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.07 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.08 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.09 - XX.XX.20XX - por XXX
				 * -----------------------------------------------------------------------------------------------
				 * v 1.10 - XX.XX.20XX - por XXX
				 * ===============================================================================================
				 */
				#endregion
			}
			#endregion

			#region[ Data / Hora ]
			public class DataHora
			{
				public const string FmtDia = "dd";
				public const string FmtDiaAbreviado = "ddd";
				public const string FmtDiaExtenso = "dddd";
				public const string FmtMes = "MM";
				public const string FmtMesAbreviado = "MMM";
				public const string FmtMesExtenso = "MMMM";
				public const string FmtAno = "yyyy";
				public const string FmtHora = "HH";
				public const string FmtMin = "mm";
				public const string FmtSeg = "ss";
				public const string FmtMiliSeg = "fff";
				public const string FmtYYYYMMDD = FmtAno + FmtMes + FmtDia;
				public const string FmtHHMMSS = FmtHora + FmtMin + FmtSeg;
				public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
			}
			#endregion
		}
		#endregion

		#region [ Declarações ]
		public static ReaderWriterLock rwlDllInscE32 = new ReaderWriterLock();
		#endregion

		#region[ Funções ]

		#region [ SqlQuotedStr ]
		public static string SqlQuotedStr(string strTexto)
		{
			string strTextoFiltrado = string.Empty;
			for (int i = 0; i < strTexto.Length; i++)
			{
				if (strTexto[i].ToString() == "'") strTextoFiltrado += "''"; else strTextoFiltrado += strTexto[i];
			}
			return "'" + strTextoFiltrado + "'";
		}
		#endregion

		#region[ SqlMontaDateTimeParaSqlDateTime ]
		public static string SqlMontaDateTimeParaSqlDateTime(DateTime dtReferencia)
		{
			string strDataHora;
			string strSql;
			strDataHora = dtReferencia.ToString(Global.Cte.DataHora.FmtAno) +
						  "-" +
						  dtReferencia.ToString(Global.Cte.DataHora.FmtMes) +
						  "-" +
						  dtReferencia.ToString(Global.Cte.DataHora.FmtDia) +
						  " " +
						  dtReferencia.ToString(Global.Cte.DataHora.FmtHora) +
						  ":" +
						  dtReferencia.ToString(Global.Cte.DataHora.FmtMin) +
						  ":" +
						  dtReferencia.ToString(Global.Cte.DataHora.FmtSeg);
			strSql = "Convert(datetime, '" + strDataHora + "', 120)";
			return strSql;
		}
		#endregion

		#region[ SqlMontaDateTimeParaYyyyMmDdHhMmSsCommSeparador ]
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
		public static string SqlMontaDateTimeParaYyyyMmDdHhMmSsCommSeparador(string strNomeCampo, string strAlias)
		{
			string strResposta;
			if (strAlias.Trim().Length == 0) strAlias = strNomeCampo;
			strResposta = "Coalesce(Convert(varchar(19), " + strNomeCampo + ", 121), '')";
			if (strAlias.Length > 0) strResposta += " AS " + strAlias;
			return strResposta;
		}
		#endregion

		#region[ SqlMontaDateTimeParaYyyyMmDdHhMmSsCommSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30 14:27:01
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <returns></returns>
		public static string SqlMontaDateTimeParaYyyyMmDdHhMmSsCommSeparador(string strNomeCampo)
		{
			return SqlMontaDateTimeParaYyyyMmDdHhMmSsCommSeparador(strNomeCampo, "");
		}
		#endregion

		#region[ ConverteYyyyMmDdHhMmSsParaDateTime ]
		public static DateTime ConverteYyyyMmDdHhMmSsParaDateTime(string strYyyyMmDdHhMmSs)
		{
			string strFormato;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			strFormato = Cte.DataHora.FmtAno +
						 "-" +
						 Cte.DataHora.FmtMes +
						 "-" +
						 Cte.DataHora.FmtDia +
						 " " +
						 Cte.DataHora.FmtHora +
						 ":" +
						 Cte.DataHora.FmtMin +
						 ":" +
						 Cte.DataHora.FmtSeg;
			if (DateTime.TryParseExact(strYyyyMmDdHhMmSs, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ GravaEventLog ]
		public static void GravaEventLog(string strSource, string strMessage, EventLogEntryType eTipoMensagem)
		{
			if (strMessage.Length > 32000)
			{
				// Tamanho máximo do log do Event Viewer é de 32766 caracteres !!
				strMessage = strMessage.Substring(0, 32000) + " (truncado ...)";
			}

			try
			{
				EventLog.WriteEntry(strSource, strMessage, eTipoMensagem);
			}
			catch(Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}
		#endregion

		#endregion
	}
	#endregion
}
