#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Configuration;
#endregion

namespace EmailSenderService
{
	public class Global
	{
		#region [ Constantes ]
		public static class Cte
		{
			#region[ Versão do Aplicativo ]
			public static class Aplicativo
			{
				public const string NOME_OWNER = "Artven";
				public const string NOME_SISTEMA = "Email Sender Service";
				public static string ID_SISTEMA_EMAILSENDER = GetConfigurationValue("ServiceName");
				public const string VERSAO_NUMERO = "1.01";
				public const string VERSAO_DATA = "15.NOV.2016";
				public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
				public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
				public const string M_DESCRICAO = "Serviço do Windows para envio automático de e-mails";
			}
			#endregion

			#region[ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 08.06.2016 - por LHGX
			 *		Início.
			 *		Este serviço do Windows realiza envio automático de e-mails.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - 15.11.2016 - por LHGX
			 *		Implementação de tratamento para os novos campos 'replyToMsg' e 'st_replyToMsg' da
			 *		tabela t_EMAILSNDSVC_MENSAGEM, pois o uso de um endereço para ReplyTo passou a ser
			 *		individualizado por mensagem de e-mail.
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
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - XX.XX.20XX - por XXX
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - XX.XX.20XX - por XXX
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

			#region [ EMAILSND ]
			public static class EMAILSND
			{
				#region [ ID_T_PARAMETRO ]
				public static class ID_T_PARAMETRO
				{
					public const string DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE = "EmailSndSvc_DtHrUltManutencaoArqLogAtividade";
					public const string FLAG_HABILITACAO_ENVIO_EMAILS = "EmailSndSvc_FlagHabilitacao";
				}
				#endregion
			}
			#endregion


			#region [ Log ]
			public static class LogAtividade
			{
				public static string PathLogAtividade = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 365;
				public const string ExtensaoArqLog = "LOG";
			}
			#endregion

			#region [ LogBd ]
			public static class LogBd
			{
				#region [ Usuario ]
				public static class Usuario
				{
					public const string ID_USUARIO_SISTEMA = "SISTEMA";
					public const string ID_USUARIO_LOG = "EMAILSNDSVC";
				}
				#endregion

				#region [ Operacao ]
				public static class Operacao
				{
					public const string OP_LOG_EMAILSENDERSERVICE_INICIADO = "EMAILSNDSVC INICIADO";
					public const string OP_LOG_EMAILSENDERSERVICE_ENCERRADO = "EMAILSNDSVC ENCERRADO";
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
				public const string FmtMin = "mm";
				public const string FmtSeg = "ss";
				public const string FmtMiliSeg = "fff";
				public const string FmtYYYYMMDD = FmtAno + FmtMes + FmtDia;
				public const string FmtHHMMSS = FmtHora + FmtMin + FmtSeg;
				public const string FmtHhMmComSeparador = FmtHora + ":" + FmtMin;
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
				public const int COD_NEGATIVO_UM = -1;
				public const int TAM_MAX_NSU = 12;
			}
			#endregion

			#region [ Nsu ]
			public class Nsu
			{
				public const string T_EMAILSNDSVC_MENSAGEM = "T_EMAILSNDSVC_MENSAGEM";
				public const string T_EMAILSNDSVC_REMETENTE = "T_EMAILSNDSVC_REMETENTE";
				public const string T_EMAILSNDSVC_LOG = "T_EMAILSNDSVC_LOG";
				public const string T_EMAILSNDSVC_LOG_ERRO = "T_EMAILSNDSVC_LOG_ERRO";
			}
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
			return (long)ts.Milliseconds + 1000L * ((long)ts.Seconds + (60L * ((long)ts.Minutes + (60L * ((long)ts.Hours + (24L * (long)ts.Days))))));
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
			return (long)ts.Seconds + (60L * ((long)ts.Minutes + (60L * ((long)ts.Hours + (24L * (long)ts.Days)))));
		}
		#endregion

		#region[ converteDdMmYyyyHhMmSsParaDateTime ]
		/// <summary>
		/// Converte o texto que representa uma data/hora para DateTime
		/// </summary>
		/// <param name="strDdMmYyyyHhMmSs">
		/// Texto representando uma data/hora, com ou sem separadores, sendo que a parte da hora é opcional.
		/// </param>
		/// <returns>
		/// Retorna a data/hora como DateTime, se não for possível fazer a conversão, retorna DateTime.MinValue
		/// </returns>
		public static DateTime converteDdMmYyyyHhMmSsParaDateTime(string strDdMmYyyyHhMmSs)
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

			#region [ Dia ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strDia += c;
				if (strDia.Length == 2) break;
			}
			while (strDia.Length < 2) strDia = '0' + strDia;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Mês ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMes += c;
				if (strMes.Length == 2) break;
			}
			while (strMes.Length < 2) strMes = '0' + strMes;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Ano ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
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

			#region [ Remove separador(es) entre a data e hora, se houver ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				if (!isDigit(strDdMmYyyyHhMmSs[0]))
					strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				else
					break;
			}
			#endregion

			#region [ Hora ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Segundo ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strSegundo += c;
				if (strSegundo.Length == 2) break;
			}
			while (strSegundo.Length < 2) strSegundo = '0' + strSegundo;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtDia +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtAno +
						 ' ' +
						 Cte.DataHora.FmtHora +
						 Cte.DataHora.FmtMin +
						 Cte.DataHora.FmtSeg;
			#endregion

			#region [ Monta data/hora normalizada ]
			strDataHoraAConverter = strDia +
									strMes +
									strAno +
									' ' +
									strHora +
									strMinuto +
									strSegundo;
			#endregion

			if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
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

		#region [ executaManutencaoArqLogAtividade ]
		/// <summary>
		/// Apaga os arquivos de log de atividade antigos
		/// </summary>
		public static bool executaManutencaoArqLogAtividade(out string strMsgErro)
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "executaManutencaoArqLogAtividade()";
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
				strMsg = "Rotina " + strNomeDestaRotina + " iniciada";
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

				strMsg = "Rotina " + strNomeDestaRotina + " concluída com sucesso: " + intQtdeApagada.ToString() + " arquivos excluídos (duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio) + ")";
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = strNomeDestaRotina + "\n" + ex.ToString();
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

		#region [ GetConfigurationValue ]
		private static string GetConfigurationValue(string key)
		{
			Assembly service = Assembly.GetAssembly(typeof(EmailSenderInstaller));
			Configuration config = ConfigurationManager.OpenExeConfiguration(service.Location);
			if (config.AppSettings.Settings[key] != null)
			{
				return config.AppSettings.Settings[key].Value;
			}
			else
			{
				throw new IndexOutOfRangeException("Settings collection does not contain the requested key:" + key);
			}
		}
		#endregion

		#region[ gravaEventLog ]
		public static void gravaEventLog(string strSource, string strMessage, EventLogEntryType eTipoMensagem)
		{
			if (strMessage.Length > 32000)
			{
				// Tamanho máximo do log do Event Viewer é de 32766 caracteres !!
				strMessage = strMessage.Substring(0, 32000) + " (truncado ...)";
			}

			try
			{
				gravaLogAtividade(strMessage);
				System.Diagnostics.EventLog.WriteEntry(strSource, strMessage, eTipoMensagem);
			}
			catch
			{
				// NOP
			}
		}
		#endregion

		#region[ gravaEventLog (overload) ]
		public static void gravaEventLog(string strMessage, EventLogEntryType eTipoMensagem)
		{
			gravaEventLog(Cte.Aplicativo.ID_SISTEMA_EMAILSENDER, strMessage, eTipoMensagem);
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

		#endregion
	}
}
