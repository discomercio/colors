#region [ using ]
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace ComPlusCalcCedulas
{
	class Global
	{
		#region [ Cte ]
		public class Cte
		{
			#region [ Versão ]
			public class Versao
			{
				public const string strNomeSistema = "ComPlusCalcCedulas";
				public const string ID_SISTEMA_EVENTLOG = "ComPlusCalcCedulas";
				public const string strVersao = "1.00 - 15.JUL.2015";

				#region[ Comentário sobre as versões ]
				/*================================================================================================
				 * v 1.00 - 15.07.2015 - por HHO
				 *        Início.
				 *        Preparação da estrutura básica do projeto.
				 * -----------------------------------------------------------------------------------------------
				 * v 1.01 - XX.XX.20XX - por XXX
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

		#region [ Funções ]

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

		#region[ isDigit ]
		public static bool isDigit(char c)
		{
			if ((c >= '0') && (c <= '9')) return true;
			return false;
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
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}
		#endregion

		#region [ GravaEventLogAvisoErro ]
		public static void GravaEventLogAvisoErro(string strMessage)
		{
			if (strMessage.Length > 32000)
			{
				// Tamanho máximo do log do Event Viewer é de 32766 caracteres !!
				strMessage = strMessage.Substring(0, 32000) + " (truncado ...)";
			}

			try
			{
				EventLog.WriteEntry(Global.Cte.Versao.ID_SISTEMA_EVENTLOG, strMessage, EventLogEntryType.Error);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}
		#endregion

		#region [ GravaEventLogAvisoInformativo ]
		public static void GravaEventLogAvisoInformativo(string strMessage)
		{
			if (strMessage.Length > 32000)
			{
				// Tamanho máximo do log do Event Viewer é de 32766 caracteres !!
				strMessage = strMessage.Substring(0, 32000) + " (truncado ...)";
			}

			try
			{
				EventLog.WriteEntry(Global.Cte.Versao.ID_SISTEMA_EVENTLOG, strMessage, EventLogEntryType.Information);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
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


		#endregion
	}
}
