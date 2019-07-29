#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
#endregion

namespace EtqFinanceiro
{
    class BD
    {
        #region [ Atributos ]
        public static SqlConnection cnConexao;
        private static SqlTransaction _sqlTransacao;
        private static bool _transacaoEmAndamento;

        public static bool isTransacaoEmAndamento
        {
            get { return _transacaoEmAndamento; }
        }

        #region [ Parâmetros de conexão ]
        public static string ServidorBD = ConfigurationManager.ConnectionStrings["ServidorBanco"].ConnectionString;
        public static string NomeBD = ConfigurationManager.ConnectionStrings["NomeBanco"].ConnectionString;
        public static string LoginBD = ConfigurationManager.ConnectionStrings["LoginBanco"].ConnectionString;
        public static string SenhaCriptografadaBD = ConfigurationManager.ConnectionStrings["SenhaBanco"].ConnectionString;
        public static string Schema = "dbo";
        #endregion

        #region [ Constantes ]
        public const int MAX_TAMANHO_VARCHAR = 8000;
        public const int MAX_TENTATIVAS_INSERT_BD = 3;
        public const int MAX_TENTATIVAS_UPDATE_BD = 2;
        public const int MAX_TENTATIVAS_DELETE_BD = 2;
        public const int intCommandTimeoutEmSegundos = 5 * 60;
        public const char CARACTER_CURINGA_TODOS = '%';
        #endregion

        #endregion

        #region [ Métodos ]

        #region[ montaStringConexaoBd ]
        private static String montaStringConexaoBd()
        {
            String strStringConexaoBd;
            strStringConexaoBd = "Data Source=" + ServidorBD + ";" +
                                 "Initial Catalog=" + NomeBD + ";" +
                                 "User Id=" + LoginBD + ";" +
                                 "Password=" + Criptografia.Descriptografa(SenhaCriptografadaBD) + ";";
            return strStringConexaoBd;
        }
        #endregion

        #region [ abreConexao ]
        public static void abreConexao()
        {
            BD.cnConexao = abreNovaConexao();
        }
        #endregion

        #region [ abreNovaConexao ]
        public static SqlConnection abreNovaConexao()
        {
            SqlConnection cn;
            String strConnection;

            strConnection = montaStringConexaoBd();
            cn = new SqlConnection(strConnection);
            cn.Open();

            return cn;
        }
        #endregion

        #region [ fechaConexao ]
        public static void fechaConexao()
        {
            try
            {
                fechaConexao(ref cnConexao);
            }
            catch (Exception)
            {
                // Nop
            }
        }

        public static void fechaConexao(ref SqlConnection cn)
        {
            try
            {
                if (cn == null) return;
                if (cn.State != ConnectionState.Closed) cn.Close();
            }
            catch (Exception)
            {
                // Nop
            }
        }
        #endregion

        #region [ criaSqlCommand ]
        public static SqlCommand criaSqlCommand()
        {
            SqlCommand cmCommand;
            cmCommand = criaSqlCommand(ref cnConexao);
            if (_transacaoEmAndamento) cmCommand.Transaction = _sqlTransacao;
            return cmCommand;
        }

        public static SqlCommand criaSqlCommand(ref SqlConnection cn)
        {
            SqlCommand cmCommand = new SqlCommand();
            cmCommand.Connection = cn;
            cmCommand.CommandTimeout = 0;
            cmCommand.CommandType = CommandType.Text;
            return cmCommand;
        }
        #endregion

        #region [ criaSqlDataAdapter ]
        public static SqlDataAdapter criaSqlDataAdapter()
        {
            SqlDataAdapter daDataAdapter = new SqlDataAdapter();
            return daDataAdapter;
        }
        #endregion

        #region [ iniciaTransacao ]
        public static void iniciaTransacao()
        {
            _transacaoEmAndamento = true;
            _sqlTransacao = cnConexao.BeginTransaction();
        }
        #endregion

        #region [ commitTransacao ]
        public static void commitTransacao()
        {
            _transacaoEmAndamento = false;
            _sqlTransacao.Commit();
        }
        #endregion

        #region [ rollbackTransacao ]
        public static void rollbackTransacao()
        {
            _transacaoEmAndamento = false;
            _sqlTransacao.Rollback();
        }
        #endregion

        #region [ executaNonQuery ]
        public static int executaNonQuery(ref SqlCommand cmComando)
        {
            if (_transacaoEmAndamento)
            {
                if (cmComando.Transaction != _sqlTransacao) cmComando.Transaction = _sqlTransacao;
            }
            return cmComando.ExecuteNonQuery();
        }
        #endregion

        public static string getBancoDescricao(string codigo, out string strMsgErro)
        {
            #region [ Declarações ]
            string strResp;
            string strSql;
            SqlCommand cmCommand;
            SqlDataReader dr;
            #endregion

            strMsgErro = "";
            codigo = codigo.Trim();
            try
            {
                cmCommand = BD.criaSqlCommand();

                strSql = "SELECT descricao FROM t_BANCO WHERE (CONVERT(smallint, codigo) = " + codigo + ")";
                cmCommand.CommandText = strSql;
                dr = cmCommand.ExecuteReader();
                try
                {
                    if (dr.Read())
                    {
                        strResp = readToString(dr["descricao"]);
                        return strResp;
                    }
                    else
                    {
                        strMsgErro = "Banco de código '" + codigo + "' não encontrado!!";
                        return "";
                    }
                }
                finally
                {
                    dr.Close();
                }

            }
            catch (Exception ex)
            {
                strMsgErro = ex.ToString();
                return "";
            }
        }

        #region [ getVersaoModulo ]
        public static VersaoModulo getVersaoModulo(string modulo, out string strMsgErro)
        {
            #region [ Declarações ]
            VersaoModulo versaoModulo = new VersaoModulo();
            String strSql;
            SqlCommand cmCommand;
            SqlDataReader drVersao;
            #endregion

            strMsgErro = "";
            try
            {
                cmCommand = BD.criaSqlCommand();

                strSql = "SELECT " +
                            "*" +
                        " FROM t_VERSAO" +
                        " WHERE" +
                            " (modulo = '" + modulo + "')";
                cmCommand.CommandText = strSql;
                drVersao = cmCommand.ExecuteReader();
                try
                {
                    if (drVersao.Read())
                    {
                        versaoModulo.modulo = readToString(drVersao["modulo"]);
                        versaoModulo.versao = readToString(drVersao["versao"]);
                        versaoModulo.mensagem = readToString(drVersao["mensagem"]);
                        versaoModulo.cor_fundo_padrao = readToString(drVersao["cor_fundo_padrao"]);
                        return versaoModulo;
                    }
                    else
                    {
                        strMsgErro = "Módulo '" + modulo + "' não cadastrado no controle de versões do sistema!!";
                        return null;
                    }
                }
                finally
                {
                    drVersao.Close();
                }
            }
            catch (Exception ex)
            {
                strMsgErro = ex.ToString();
                return null;
            }
        }
        #endregion

        #region [ obtemDataHoraServidor ]
        public static DateTime obtemDataHoraServidor()
        {
            #region [ Declarações ]
            DateTime dataHoraResposta = DateTime.MinValue;
            String strSql;
            SqlCommand cmCommand;
            SqlDataReader drVersao;
            #endregion

            try
            {
                cmCommand = BD.criaSqlCommand();
                strSql = "SELECT getdate() AS data_hora";
                cmCommand.CommandText = strSql;
                drVersao = cmCommand.ExecuteReader();
                try
                {
                    if (drVersao.Read())
                    {
                        dataHoraResposta = readToDateTime(drVersao["data_hora"]);
                    }
                }
                finally
                {
                    drVersao.Close();
                }

                return dataHoraResposta;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }
        #endregion

        #region [ readToString ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo texto
        /// </param>
        /// <returns>
        /// Retorna o texto armazenado no campo. Caso o conteúdo seja DBNull, retorna uma String vazia.
        /// </returns>
        public static String readToString(object campo)
        {
            return !Convert.IsDBNull(campo) ? campo.ToString() : "";
        }
        #endregion

        #region [ readToDateTime ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo data
        /// </param>
        /// <returns>
        /// Retorna a data armazenada no campo. Caso o conteúdo seja DBNull, retorna DateTime.MinValue
        /// </returns>
        public static DateTime readToDateTime(object campo)
        {
            return !Convert.IsDBNull(campo) ? (DateTime)campo : DateTime.MinValue;
        }
        #endregion

        #region [ readToByte ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo byte
        /// </param>
        /// <returns>
        /// Retorna o número armazenado no campo
        /// </returns>
        public static byte readToByte(object campo)
        {
            return (byte)campo;
        }
        #endregion

        #region [ readToShort ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo short
        /// </param>
        /// <returns>
        /// Retorna o número armazenado no campo
        /// </returns>
        public static short readToShort(object campo)
        {
            return (short)campo;
        }
        #endregion

        #region [ readToInt ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo int
        /// </param>
        /// <returns>
        /// Retorna o número armazenado no campo
        /// </returns>
        public static int readToInt(object campo)
        {
            if (campo.GetType().Name.Equals("Int16"))
            {
                return (int)(Int16)campo;
            }
            else
            {
                return (int)campo;
            }
        }
        #endregion

        #region [ readToInt16 ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo System.Int16
        /// </param>
        /// <returns>
        /// Retorna o número armazenado no campo
        /// </returns>
        public static Int16 readToInt16(object campo)
        {
            return (Int16)campo;
        }
        #endregion

        #region [ readToChar ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo char
        /// </param>
        /// <returns>
        /// Retorna o caracter armazenado no campo. Caso o conteúdo seja DBNull, retorna um caracter nulo.
        /// </returns>
        public static char readToChar(object campo)
        {
            String s;
            char c = '\0';

            if (!Convert.IsDBNull(campo))
            {
                s = campo.ToString();
                if (s.Length > 0) c = s[0];
            }

            return c;
        }
        #endregion

        #region [ readToDecimal ]
        /// <summary>
        /// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
        /// </summary>
        /// <param name="campo">
        /// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo decimal
        /// </param>
        /// <returns>
        /// Retorna o número armazenado no campo
        /// </returns>
        public static decimal readToDecimal(object campo)
        {
            return (decimal)campo;
        }
        #endregion

        #region [ isConexaoOk ]
        public static bool isConexaoOk()
        {
            #region [ Declarações ]
            DateTime dtHrServidor = DateTime.MinValue;
            #endregion

            try
            {
                dtHrServidor = obtemDataHoraServidor();
                if (dtHrServidor != DateTime.MinValue) return true;
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #endregion
    }

    #region [ Classe VersaoModulo ]
    public class VersaoModulo
    {
        private string _modulo;
        public string modulo
        {
            get { return _modulo; }
            set { _modulo = value; }
        }

        private string _versao;
        public string versao
        {
            get { return _versao; }
            set { _versao = value; }
        }

        private string _mensagem;
        public string mensagem
        {
            get { return _mensagem; }
            set { _mensagem = value; }
        }

        private string _cor_fundo_padrao;
        public string cor_fundo_padrao
        {
            get { return _cor_fundo_padrao; }
            set { _cor_fundo_padrao = value; }
        }
    }
    #endregion
}
