#region [ using ]
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
#endregion

namespace EtqFinanceiro
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
                public const string NOME_SISTEMA = "EtqFinanceiro";
                public const string VERSAO_NUMERO = "1.00";
                public const string VERSAO_DATA = "14.JAN.2019";
                public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
                public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
                public const string M_DESCRICAO = "Módulo Etiquetas (Financeiro)";
            }
			#endregion

			#region[ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 30.05.2017 - por TRR
			 *        Início.
			 *        Este programa realiza a impressão das etiquetas para uso do departamento financeiro.
             *        (Impressão de etiquetas dos dados das comissões dos indicadores)
			 * -----------------------------------------------------------------------------------------------
			 * v 1.00(B) - 14.01.2019 - por HHO
			 *		  Implementação de novo painel para imprimir as etiquetas usando os valores com desconto
			 *		  na comissão.
			 * ===============================================================================================
			 */
			#endregion

			#region [ Etc ]
			public class Etc
            {
                public const string SIMBOLO_MONETARIO = "R$";
                public const string ID_PF = "PF";
                public const string ID_PJ = "PJ";
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

            #region [ Log ]
            public class LogAtividade
            {
                public static string PathLogAtividade = Application.StartupPath + "\\LOG_ATIVIDADE";
                public const int CorteArqLogEmDias = 365;
                public const string ExtensaoArqLog = "LOG";
            }
            #endregion

            #region [ Classe EtqFinanceiro ]
            public class EtqFinanceiro
            {
                #region [ LogOperacao - Códigos de operação para o log ]
                public class LogOperacao
                {
                    // Texto com 20 posições
                    public const string LOGON = "EtqFin-Logon";
                    public const string LOGOFF = "EtqFin-Logoff";
                    public const string ETIQUETA_FIN_IMPRESSAO_COMPLETA = "EtqFin-Completa";
                    public const string ETIQUETA_FIN_IMPRESSAO_SELECIONADO = "EtqFin-Selecionado";
                    public const string RECONEXAO_BD = "EtqFin-Reconexao-BD";
                }
                #endregion
            }
            #endregion
        }
        #endregion

        #region [ Atributos ]
        public static DateTime dtHrInicioRefRelogioServidor;
        public static DateTime dtHrInicioRefRelogioLocal;
        public static Color BackColorPainelPadrao = SystemColors.Control;
        #endregion

        #region [ Classe Acesso ]
        public class Acesso
        {
            #region [ Constantes ]
            public const String OP_CEN_ETQFIN_APP_ETIQUETA_FIN_ACESSO_AO_MODULO = "28900";
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
            public static bool operacaoPermitida(string idOperacao)
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
            public static string usuario = "";
            public static string senhaDigitada = "";
            public static string senhaCriptografada = "";
            public static string senhaDescriptografada = "";
            public static string nome = "";
            public static bool cadastrado = false;
            public static bool bloqueado = false;
            public static bool senhaExpirada = false;
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
        public static string formataDataDdMmYyyyHhMmSsComSeparador(DateTime data)
        {
            if (data == null) return "";
            if (data == DateTime.MinValue) return "";
            return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmSsComSeparador);
        }
        #endregion

        #region [ formataDataDigitadaParaMMYYYYComSeparador ]
        /// <summary>
        /// A partir de uma data digitada pelo usuário, tenta normalizar e retornar uma data no formato MM/YYYY
        /// </summary>
        /// <param name="data">
        /// Texto com a data digitada pelo usuário a ser normalizado
        /// Formatos aceitos: MMYY, MMYYYY, MM/YY, MM/YYYY
        /// </param>
        /// <returns>
        /// Retorna a data no formato MM/YYYY caso a data informada esteja em um formato válido, caso contrário, retorna o próprio texto informado
        /// </returns>
        public static String formataDataDigitadaParaMMYYYYComSeparador(String data)
        {
            String strMes;
            String strAno;

            if (data == null) return "";
            if (data.Trim().Length == 0) return "";

            if (data.IndexOf('/') == -1)
            {
                // A data foi digitada sem os separadores
                data = digitos(data);
                // Neste caso, aceita somente se tiver sido digitada no formado MM ou MMYY ou MMYYYY
                if ((data.Length != 4) && (data.Length != 6)) return data;
                strMes = data.Substring(0, 2);
                strAno = data.Substring(2, data.Length - 2);
            }
            else
            {
                String[] v = data.Split('/');
                // É necessário que a data tenha vindo separada em 2 partes: mês/ano
                if ((v.Length != 2)) return data;
                for (int i = 0; i < v.Length; i++)
                {
                    if (v[i] == null) return data;
                    v[i] = digitos(v[i]);
                    if (v[i].Trim().Length == 0) return data;
                }
                strMes = v[0].PadLeft(2, '0');
                strAno = v[1];

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

            return strMes + "/" + strAno;
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

        #region [ isDataMMYYYYOk ]
        /// <summary>
        /// Indica se a data representada pelo texto no formato MM/YYYY é uma data válida
        /// </summary>
        /// <param name="data">
        /// Texto representando uma data no formato MM/YYYY
        /// </param>
        /// <returns>
        /// true: data válida
        /// false: data inválida
        /// </returns>
        public static bool isDataMMYYYYOk(String data)
        {
            if (string.IsNullOrWhiteSpace(data)) return true;

            bool blnDataOk;
            string strFormato;
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            data = formataDataDigitadaParaMMYYYYComSeparador(data);
            if (data.Length != 7) return false;
            strFormato = Cte.DataHora.FmtMes + "/" + Cte.DataHora.FmtAno;
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
