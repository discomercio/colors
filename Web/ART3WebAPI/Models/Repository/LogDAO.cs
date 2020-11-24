#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading;
using ART3WebAPI.Models.Domains;
#endregion

namespace ART3WebAPI.Models.Repository
{
    class LogDAO
    {

        #region [ inicializaConstrutorEstatico ]
        public static void inicializaConstrutorEstatico()
        {
            // NOP
            // 1) The static constructor for a class executes before any instance of the class is created.
            // 2) The static constructor for a class executes before any of the static members for the class are referenced.
            // 3) The static constructor for a class executes after the static field initializers (if any) for the class.
            // 4) The static constructor for a class executes at most one time during a single program instantiation
            // 5) A static constructor does not take access modifiers or have parameters.
            // 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
            // 7) A static constructor cannot be called directly.
            // 8) The user has no control on when the static constructor is executed in the program.
            // 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
        }
        #endregion

        #region [ Construtor estático ]
        static LogDAO()
        {
            
        }
        #endregion

        #region [ Métodos ]

        #region [ insere ]
        /// <summary>
        /// Grava novo registro no log
        /// </summary>
        /// <param name="usuario">
        /// Identificação do usuário que realizou a operação
        /// </param>
        /// <param name="log">
        /// Objeto que representa um registro do log contendo os dados para gravar
        /// </param>
        /// <param name="strMsgErro">
        /// Retorna a mensagem de erro no caso de ocorrer exception
        /// </param>
        /// <returns>
        /// true: gravação efetuada com sucesso
        /// false: falha na gravação
        /// </returns>
        public static bool insere(String usuario, Entities.Log log, out String strMsgErro)
        {
            const string NOME_DESTA_ROTINA = "LogDAO.insere()";
            String strSql;

            strMsgErro = "";

            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            cn.Open();
            try
            {

                #region [ cmInsereFinLog ]
                strSql = "INSERT INTO t_LOG (" +
                            "data, " +
                            "usuario, " +
                            "loja, " +
                            "pedido, " +
                            "id_cliente, " +
                            "operacao, " +
                            "complemento" +
                        ") VALUES (" +
                            "getdate(), " +
                            "@usuario, " +
                            "@loja, " +
                            "@pedido, " +
                            "@id_cliente, " +
                            "@operacao, " +
                            "@complemento" +
                        ")";
                SqlCommand cmInsereLog = new SqlCommand();
                cmInsereLog.Connection = cn;
                cmInsereLog.CommandText = strSql;
                cmInsereLog.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
                cmInsereLog.Parameters.Add("@loja", SqlDbType.VarChar, 3);
                cmInsereLog.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
                cmInsereLog.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
                cmInsereLog.Parameters.Add("@operacao", SqlDbType.VarChar, 20);
                cmInsereLog.Parameters.Add("@complemento", SqlDbType.Text, -1);
                cmInsereLog.Prepare();
                #endregion


                #region [ Declarações ]
                bool blnSucesso = false;
                int intQtdeTentativas = 0;
                int intRetorno;
                StringBuilder sbLog = new StringBuilder("");
                #endregion

                try
                {
                    #region [ Laço de tentativas de inserção no banco de dados ]
                    do
                    {
                        intQtdeTentativas++;

                        strMsgErro = "";

                        #region [ Preenche o valor dos parâmetros ]
                        cmInsereLog.Parameters["@usuario"].Value = log.usuario;
                        cmInsereLog.Parameters["@loja"].Value = (log.loja == null ? "" : log.loja);
                        cmInsereLog.Parameters["@pedido"].Value = (log.pedido == null ? "" : log.pedido);
                        cmInsereLog.Parameters["@id_cliente"].Value = (log.id_cliente == null ? "" : log.id_cliente);
                        cmInsereLog.Parameters["@operacao"].Value = (log.operacao == null ? "" : log.operacao);
                        cmInsereLog.Parameters["@complemento"].Value = (log.complemento == null ? "" : log.complemento);
                        #endregion

                        #region [ Monta texto para o log em arquivo ]
                        // Se houver conteúdo de alguma tentativa anterior, descarta
                        sbLog = new StringBuilder("");
                        foreach (SqlParameter item in cmInsereLog.Parameters)
                        {
                            if (!item.ParameterName.Equals("@complemento"))
                            {
                                if (sbLog.Length > 0) sbLog.Append("; ");
                                sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
                            }
                        }
                        #endregion

                        #region [ Tenta inserir o registro ]
                        try
                        {
                            intRetorno = cmInsereLog.ExecuteNonQuery();
                        }
                        catch 
                        {
                            intRetorno = 0;
                        }
                        #endregion

                        #region [ Processamento para sucesso ou falha desta tentativa de inserção ]
                        if (intRetorno == 1)
                        {
                            blnSucesso = true;
                        }
                        else
                        {
                            Thread.Sleep(100);
                        }
                        #endregion

                    } while ((!blnSucesso) && (intQtdeTentativas < 3));
                    #endregion

                    #region [ Processamento final de sucesso ou falha ]
                    if (blnSucesso)
                    {
                        Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Log gravado com sucesso no BD:\n" + sbLog.ToString());
                        return true;
                    }
                    else
                    {
                        strMsgErro = "Falha ao gravar no banco de dados o log após " + intQtdeTentativas.ToString() + " tentativas!!";
                        Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsgErro);
                        return false;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    // Para o usuário, exibe uma mensagem mais sucinta
                    strMsgErro = ex.Message;
                    Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception: " + strMsgErro);
                    return false;
                }
            }
            finally
            {
                cn.Close();
            }
        }
        #endregion

        #endregion
    }
}