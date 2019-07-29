using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;

namespace Financeiro
{
    class PlanilhaRepasseMktplaceDAO
    {
        #region [ Atributos ]
        private static SqlCommand cmdPlanilhaRepasseMktplceN1SelectChecksum;
        private static SqlCommand cmdPlanilhaRepasseMktplaceN1Insert;
        private static SqlCommand cmdPlanilhaRepasseMktplaceN2Insert;
        private static SqlCommand cmdPlanilhaRepasseMktplaceN3Insert;
        private static SqlCommand cmdPlanilhaRepasseMktplaceN4Insert;
        #endregion

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
        static PlanilhaRepasseMktplaceDAO()
        {
            inicializaObjetosEstaticos();
        }
        #endregion

        public static void inicializaObjetosEstaticos()
        {
            #region [ Declarações ]
            string strSql;
            #endregion

            #region [ Insert t_MKTP_REPASSE_N1 ]
            strSql = "INSERT INTO t_MKTP_REPASSE_N1 (" +
                        "Checksum, " +
                        "Usuario, " +
                        "NomeArquivo, " +
                        "Path, " +
                        "OrigemGrupo, " +
                        "DataCadastro, " +
                        "DataHoraCadastro" +
                    ") VALUES (" +
                        "@Checksum, " +
                        "@Usuario, " +
                        "@NomeArquivo, " +
                        "@Path, " +
                        "@OrigemGrupo, " +
                        Global.sqlMontaGetdateSomenteData() + ", " +
                        "getdate()" +
                    ") SET @Id = SCOPE_IDENTITY();";
            cmdPlanilhaRepasseMktplaceN1Insert = BD.criaSqlCommand();
            cmdPlanilhaRepasseMktplaceN1Insert.CommandText = strSql;
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@Checksum", SqlDbType.VarChar, 64);
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@Usuario", SqlDbType.VarChar, 10);
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@NomeArquivo", SqlDbType.VarChar, 256);
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@Path", SqlDbType.VarChar, 1024);
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@OrigemGrupo", SqlDbType.VarChar, 20);
            cmdPlanilhaRepasseMktplaceN1Insert.Parameters.Add("@Id", SqlDbType.Int).Direction = ParameterDirection.Output;
            #endregion

            #region [ cmdPlanilhaRepasseMktplceN1SelectChecksum ]
            strSql = "SELECT * FROM t_MKTP_REPASSE_N1 WHERE (Checksum = @Checksum)";
            cmdPlanilhaRepasseMktplceN1SelectChecksum = BD.criaSqlCommand();
            cmdPlanilhaRepasseMktplceN1SelectChecksum.CommandText = strSql;
            cmdPlanilhaRepasseMktplceN1SelectChecksum.Parameters.Add("@Checksum", SqlDbType.VarChar, 64);
            #endregion

            #region [ Insert t_MKTP_REPASSE_N2 ]
            strSql = "INSERT INTO t_MKTP_REPASSE_N2 (" +
                        "MktpRepasseN1Id, " +
                        "LojaOrigemId, " +
                        "Pedido, " +
                        "DataPedido, " +
                        "TipoPagamentoId, " +
                        "MeioPagamentoId, " +
                        "ValorTotalPedido, " +
                        "Linha" +
                    ") VALUES (" +
                        "@MktpRepasseN1Id, " +
                        "@LojaOrigemId, " +
                        "@Pedido, " +
                        "@DataPedido, " +
                        "@TipoPagamentoId, " +
                        "@MeioPagamentoId, " +
                        "@ValorTotalPedido, " +
                        "@Linha" +
                    ") SET @Id = SCOPE_IDENTITY();";
            cmdPlanilhaRepasseMktplaceN2Insert = BD.criaSqlCommand();
            cmdPlanilhaRepasseMktplaceN2Insert.CommandText = strSql;
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@MktpRepasseN1Id", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@LojaOrigemId", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@Pedido", SqlDbType.VarChar, 20);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@DataPedido", SqlDbType.VarChar, 10);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@TipoPagamentoId", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@MeioPagamentoId", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@ValorTotalPedido", SqlDbType.Money);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@Linha", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN2Insert.Parameters.Add("@Id", SqlDbType.Int).Direction = ParameterDirection.Output;
            #endregion

            #region [ Insert t_MKTP_REPASSE_N3 ]
            strSql = "INSERT INTO t_MKTP_REPASSE_N3 (" +
                        "MktpRepasseN2Id, " +
                        "ProdutoId, " +
                        "ValorFrete, " +
                        "ValorItem, " +
                        "ValorItemBruto, " +
                        "PercComissao, " +
                        "Linha" +
                    ") VALUES (" +
                        "@MktpRepasseN2Id, " +
                        "@ProdutoId, " +
                        "@ValorFrete, " +
                        "@ValorItem, " +
                        "@ValorItemBruto, " +
                        "@PercComissao, " +
                        "@Linha" +
                    ") SET @Id = SCOPE_IDENTITY();";
            cmdPlanilhaRepasseMktplaceN3Insert = BD.criaSqlCommand();
            cmdPlanilhaRepasseMktplaceN3Insert.CommandText = strSql;
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@MktpRepasseN2Id", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@ProdutoId", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@ValorFrete", SqlDbType.Money);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@ValorItem", SqlDbType.Money);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@ValorItemBruto", SqlDbType.Money);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@PercComissao", SqlDbType.Decimal);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@Linha", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN3Insert.Parameters.Add("@Id", SqlDbType.Int).Direction = ParameterDirection.Output;
            #endregion

            #region [ Insert t_MKTP_REPASSE_N4 ]
            strSql = "INSERT INTO t_MKTP_REPASSE_N4 (" +
                        "MktpRepasseN3Id, " +
                        "TipoTransacao, " +
                        "StatusTransacaoId, " +
                        "Valor, " +
                        "DataPagamento, " +
                        "DataLiberacao, " +
                        "DataEstorno, " +
                        "Linha" +
                    ") VALUES (" +
                        "@MktpRepasseN3Id, " +
                        "@TipoTransacao, " +
                        "@StatusTransacaoId, " +
                        "@Valor, " +
                        "@DataPagamento, " +
                        "@DataLiberacao, " +
                        "@DataEstorno, " +
                        "@Linha" +
                    ")";
            cmdPlanilhaRepasseMktplaceN4Insert = BD.criaSqlCommand();
            cmdPlanilhaRepasseMktplaceN4Insert.CommandText = strSql;
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@MktpRepasseN3Id", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@TipoTransacao", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@StatusTransacaoId", SqlDbType.Int);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@Valor", SqlDbType.Money);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@DataPagamento", SqlDbType.VarChar, 10);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@DataLiberacao", SqlDbType.VarChar, 10);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@DataEstorno", SqlDbType.VarChar, 10);
            cmdPlanilhaRepasseMktplaceN4Insert.Parameters.Add("@Linha", SqlDbType.Int);
            #endregion
        }

        #region [ Métodos ]

        #region [ insereN1 ]
        public static bool insereN1(string usuario, PlanilhaRepasseMktplaceN1 planilhaN1, ref string strMsgErro)
        {
            #region [ Declarações ]
            bool blnSucesso = false;
            int intQtdeTentativas = 0;
            int intRetorno;
            string strOperacao = "Gravar planilha de repasse Marketplace (" + planilhaN1.OrigemGrupo + ")";
            StringBuilder sbLog = new StringBuilder("");
            #endregion

            try
            {
                #region [ Laço de tentativas de inserção no banco de dados ]
                do
                {
                    intQtdeTentativas++;
                    strMsgErro = "";

                    #region [ Preenche os valores dos parâmetros ]
                    cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@Checksum"].Value = planilhaN1.Checksum;
                    cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@Usuario"].Value = usuario;
                    cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@NomeArquivo"].Value = planilhaN1.NomeArquivo;
                    cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@Path"].Value = planilhaN1.Path;
                    cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@OrigemGrupo"].Value = planilhaN1.OrigemGrupo;
                    #endregion

                    #region [ Monta o texto para o log em arquivo ]
                    // Se houver conteúdo de alguma tentativa anterior, descarta
                    sbLog = new StringBuilder("");
                    foreach (SqlParameter param in cmdPlanilhaRepasseMktplaceN1Insert.Parameters)
                    {
                        if (sbLog.Length > 0) sbLog.Append("; ");
                        sbLog.Append(param.ParameterName + "=" + (param.Value ?? ""));
                    }
                    #endregion

                    #region [ Tenta inserir o registro ]
                    try
                    {
                        intRetorno = BD.executaNonQuery(ref cmdPlanilhaRepasseMktplaceN1Insert);
                        if (intRetorno == 1) planilhaN1.Id = (int)cmdPlanilhaRepasseMktplaceN1Insert.Parameters["@Id"].Value;
                    }
                    catch (Exception ex)
                    {
                        intRetorno = 0;
                        Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
                    }
                    #endregion

                    #region [ Processamento para sucesso ou falha desta tentativa de inserção ]
                    if(intRetorno == 1)
                    {
                        Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
                        blnSucesso = true;
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                    #endregion

                } while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
                #endregion

                #region [ Processamento final de sucesso ou falha ]
                if (blnSucesso)
                {
                    return true;
                }
                else
                {
                    strMsgErro = "Falha ao tentar gravar no BD os dados do arquivo de planilha de repasse Marketplace após " + intQtdeTentativas.ToString() + " tentativas!";
                    return false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
                return false;
            }
        }
        #endregion

        #region [ t_MKTP_REPASSE_N1: Select Checksum ]
        /// <summary>
        /// Método que verifica através do checksum do arquivo, se uma planilha já foi processada.
        /// </summary>
        /// <param name="checksum">checksum a ser verificado</param>
        /// <returns>Classe genérica Tuple
        /// bool Item1: retorna 'true' se o arquivo já foi processado
        /// string Item2: retorna o usuário que processou o arquivo. Caso não tenha sido processado retorna string vazia
        /// string Item3: retorna a data em que o arquivo foi processado. Caso não tenha sido processado retorna string vazia
        /// </returns>
        public static Tuple<bool, string, string> isArquivoProcessado(string checksum)
        {
            #region [ Declarações ]
            SqlDataReader reader;
            string usuarioCadastro;
            string dataCadastro;
            #endregion

            cmdPlanilhaRepasseMktplceN1SelectChecksum.Parameters["@Checksum"].Value = checksum;
            using (reader = cmdPlanilhaRepasseMktplceN1SelectChecksum.ExecuteReader())
            {
                if (reader.Read())
                {
                    usuarioCadastro = reader["Usuario"].ToString();
                    dataCadastro = reader["DataHoraCadastro"].ToString();
                    return new Tuple<bool, string, string>(true, usuarioCadastro, dataCadastro);
                }
                else
                    return new Tuple<bool, string, string>(false, string.Empty, string.Empty);
            }
        }
        #endregion

        #region [ insereN2 ]
        public static bool insereN2(PlanilhaRepasseMktplaceN2 planilhaN2, ref string strMsgErro)
        {
            #region [ Declarações ]
            bool blnSucesso = false;
            int intQtdeTentativas = 0;
            int intRetorno;
            string strOperacao = "Gravar planilha de repasse Marketplace N2 (ID N1: " + planilhaN2.MktplceRepasseN1Id + ", Linha: " + planilhaN2.Linha + ")";
            StringBuilder sbLog = new StringBuilder("");
            #endregion

            try
            {
                #region [ Laço de tentativas de inserção no banco de dados ]
                do
                {
                    intQtdeTentativas++;
                    strMsgErro = "";
                    
                    #region [ Preenche os valores dos parâmetros ]
                    cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@MktpRepasseN1Id"].Value = planilhaN2.MktplceRepasseN1Id;
                    cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@LojaOrigemId"].Value = planilhaN2.LojaOrigemId;
                    cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@Pedido"].Value = planilhaN2.Pedido;
                    if (planilhaN2.DataPedido != DateTime.MinValue)
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@DataPedido"].Value = Global.formataDataYyyyMmDdComSeparador(planilhaN2.DataPedido);
                    else
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@DataPedido"].Value = DBNull.Value;
                    if (planilhaN2.TipoPagamentoId > 0)
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@TipoPagamentoId"].Value = planilhaN2.TipoPagamentoId;
                    else
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@TipoPagamentoId"].Value = DBNull.Value;
                    cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@MeioPagamentoId"].Value = planilhaN2.MeioPagamentoId;
                    if (planilhaN2.ValorTotalPedido > 0)
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@ValorTotalPedido"].Value = planilhaN2.ValorTotalPedido;
                    else
                        cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@ValorTotalPedido"].Value = DBNull.Value;
                    cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@Linha"].Value = planilhaN2.Linha;
                    #endregion

                    #region [ Tenta inserir o registro ]
                    try
                    {
                        intRetorno = BD.executaNonQuery(ref cmdPlanilhaRepasseMktplaceN2Insert);
                        if (intRetorno == 1) planilhaN2.Id = (int)cmdPlanilhaRepasseMktplaceN2Insert.Parameters["@Id"].Value;
                    }
                    catch (Exception ex)
                    {
                        intRetorno = 0;
                        Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
                    }
                    #endregion

                    #region [ Processamento para sucesso ou falha desta tentativa de inserção ]
                    if (intRetorno == 1)
                    {
                        Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
                        blnSucesso = true;
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                    #endregion

                } while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
                #endregion

                #region [ Processamento final de sucesso ou falha ]
                if (blnSucesso)
                {
                    return true;
                }
                else
                {
                    strMsgErro = "Falha ao tentar gravar no BD os dados do arquivo de planilha de repasse Marketplace após " + intQtdeTentativas.ToString() + " tentativas!";
                    return false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
                return false;
            }
        }
        #endregion

        #region [ insereN3 ]
        public static bool insereN3(PlanilhaRepasseMktplaceN3 planilhaN3, ref string strMsgErro)
        {
            #region [ Declarações ]
            bool blnSucesso = false;
            int intQtdeTentativas = 0;
            int intRetorno;
            string strOperacao = "Gravar planilha de repasse Marketplace N3 (ID N2: " + planilhaN3.MktplceRepasseN2Id + ", Linha: " + planilhaN3.Linha + ")";
            StringBuilder sbLog = new StringBuilder("");
            #endregion

            try
            {
                #region [ Laço de tentativas de inserção no banco de dados ]
                do
                {
                    intQtdeTentativas++;
                    strMsgErro = "";

                    #region [ Preenche os valores dos parâmetros ]
                    cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@MktpRepasseN2Id"].Value = planilhaN3.MktplceRepasseN2Id;
                    if (planilhaN3.ProdutoId > 0)
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ProdutoId"].Value = planilhaN3.ProdutoId;
                    else
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ProdutoId"].Value = DBNull.Value;
                    if (planilhaN3.ValorFrete > 0)
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorFrete"].Value = planilhaN3.ValorFrete;
                    else
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorFrete"].Value = DBNull.Value;
                    if (planilhaN3.ValorItem > 0)
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorItem"].Value = planilhaN3.ValorItem;
                    else
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorItem"].Value = DBNull.Value;
                    if (planilhaN3.ValorItemBruto > 0)
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorItemBruto"].Value = planilhaN3.ValorItemBruto;
                    else
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@ValorItemBruto"].Value = DBNull.Value;
                    if (planilhaN3.PercComissao > 0)
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@PercComissao"].Value = planilhaN3.PercComissao;
                    else
                        cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@PercComissao"].Value = DBNull.Value;
                    cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@Linha"].Value = planilhaN3.Linha;
                    #endregion

                    #region [ Tenta inserir o registro ]
                    try
                    {
                        intRetorno = BD.executaNonQuery(ref cmdPlanilhaRepasseMktplaceN3Insert);
                        if (intRetorno == 1) planilhaN3.Id = (int)cmdPlanilhaRepasseMktplaceN3Insert.Parameters["@Id"].Value;
                    }
                    catch (Exception ex)
                    {
                        intRetorno = 0;
                        Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
                    }
                    #endregion

                    #region [ Processamento para sucesso ou falha desta tentativa de inserção ]
                    if (intRetorno == 1)
                    {
                        Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
                        blnSucesso = true;
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                    #endregion

                } while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
                #endregion

                #region [ Processamento final de sucesso ou falha ]
                if (blnSucesso)
                {
                    return true;
                }
                else
                {
                    strMsgErro = "Falha ao tentar gravar no BD os dados do arquivo de planilha de repasse Marketplace após " + intQtdeTentativas.ToString() + " tentativas!";
                    return false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
                return false;
            }
        }
        #endregion

        #region [ insereN4 ]
        public static bool insereN4(PlanilhaRepasseMktplaceN4 planilhaN4, ref string strMsgErro)
        {
            #region [ Declarações ]
            bool blnSucesso = false;
            int intQtdeTentativas = 0;
            int intRetorno;
            string strOperacao = "Gravar planilha de repasse Marketplace N4 (ID N3: " + planilhaN4.MktplceRepasseN3Id + ", Linha: " + planilhaN4.Linha + ")";
            StringBuilder sbLog = new StringBuilder("");
            #endregion

            try
            {
                #region [ Laço de tentativas de inserção no banco de dados ]
                do
                { 
                    intQtdeTentativas++;
                    strMsgErro = "";

                    #region [ Preenche os valores dos parâmetros ]
                    cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@MktpRepasseN3Id"].Value = planilhaN4.MktplceRepasseN3Id;
                    cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@TipoTransacao"].Value = planilhaN4.TipoTransacao;
                    if (planilhaN4.StatusTransacaoId > 0)
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@StatusTransacaoId"].Value = planilhaN4.StatusTransacaoId;
                    else
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@StatusTransacaoId"].Value = DBNull.Value;
                    cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@Valor"].Value = planilhaN4.Valor;
                    cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@DataPagamento"].Value = Global.formataDataYyyyMmDdComSeparador(planilhaN4.DataPagamento);
                    if (planilhaN4.DataLiberacao != DateTime.MinValue)
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@DataLiberacao"].Value = Global.formataDataYyyyMmDdComSeparador(planilhaN4.DataLiberacao);
                    else
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@DataLiberacao"].Value = DBNull.Value;
                    if (planilhaN4.DataEstorno != DateTime.MinValue)
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@DataEstorno"].Value = Global.formataDataYyyyMmDdComSeparador(planilhaN4.DataEstorno);
                    else
                        cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@DataEstorno"].Value = DBNull.Value;
                    cmdPlanilhaRepasseMktplaceN4Insert.Parameters["@Linha"].Value = planilhaN4.Linha;
                    #endregion

                    #region [ Tenta inserir o registro ]
                    try
                    {
                        intRetorno = BD.executaNonQuery(ref cmdPlanilhaRepasseMktplaceN4Insert);
                    }
                    catch (Exception ex)
                    {
                        intRetorno = 0;
                        Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
                    }
                    #endregion

                    #region [ Processamento para sucesso ou falha desta tentativa de inserção ]
                    if (intRetorno == 1)
                    {
                        Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
                        blnSucesso = true;
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                    #endregion

                } while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
                #endregion

                #region [ Processamento final de sucesso ou falha ]
                if (blnSucesso)
                {                    
                    return true;
                }
                else
                {
                    strMsgErro = "Falha ao tentar gravar no BD os dados do arquivo de planilha de repasse Marketplace após " + intQtdeTentativas.ToString() + " tentativas!";
                    return false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
                return false;
            }
        }
        #endregion

        #region [ getLojaOrigemId ]
        public static int getLojaOrigemId(string descricao, string empresaOrigem)
        {
            #region [ Declarações ]
            string strSql;
            int id;
            SqlCommand cmdSelect;
            SqlCommand cmdInsert;
            #endregion

            strSql = "SELECT * FROM t_MKTP_REPASSE_LOJA_ORIGEM WHERE (Descricao=@Descricao AND EmpresaOrigem=@EmpresaOrigem)";
            cmdSelect = BD.criaSqlCommand();
            cmdSelect.CommandText = strSql;
            cmdSelect.Parameters.AddWithValue("@Descricao", descricao);
            cmdSelect.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
            id = (int)(cmdSelect.ExecuteScalar() ?? 0);
            if (id == 0)
            {
                strSql = "SET NOCOUNT ON; INSERT INTO t_MKTP_REPASSE_LOJA_ORIGEM (Descricao, EmpresaOrigem, DataHoraCadastro)" +
                    " VALUES (@Descricao, @EmpresaOrigem, getdate()); SELECT SCOPE_IDENTITY() AS Id;";
                cmdInsert = BD.criaSqlCommand();
                cmdInsert.CommandText = strSql;
                cmdInsert.Parameters.AddWithValue("@Descricao", descricao);
                cmdInsert.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
                id = Convert.ToInt32((cmdInsert.ExecuteScalar() ?? 0));
            }

            return id;
        }
        #endregion

        #region [ getMeioPagamentoId ]
        public static int getMeioPagamentoId(string descricao, string empresaOrigem)
        {
            #region [ Declarações ]
            string strSql;
            int id;
            SqlCommand cmdSelect;
            SqlCommand cmdInsert;
            #endregion

            strSql = "SELECT * FROM t_MKTP_REPASSE_MEIO_PAGAMENTO WHERE (Descricao=@Descricao AND EmpresaOrigem=@EmpresaOrigem)";
            cmdSelect = BD.criaSqlCommand();
            cmdSelect.CommandText = strSql;
            cmdSelect.Parameters.AddWithValue("@Descricao", descricao);
            cmdSelect.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
            id = (int)(cmdSelect.ExecuteScalar() ?? 0);
            if (id == 0)
            {
                strSql = "SET NOCOUNT ON; INSERT INTO t_MKTP_REPASSE_MEIO_PAGAMENTO (Descricao, EmpresaOrigem, DataHoraCadastro)" +
                    " VALUES (@Descricao, @EmpresaOrigem, getdate()); SELECT SCOPE_IDENTITY() AS Id;";
                cmdInsert = BD.criaSqlCommand();
                cmdInsert.CommandText = strSql;
                cmdInsert.Parameters.AddWithValue("@Descricao", descricao);
                cmdInsert.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
                id = Convert.ToInt32((cmdInsert.ExecuteScalar() ?? 0));
            }

            return id;
        }
        #endregion

        #region [ getTipoPagamentoId ]
        public static int getTipoPagamentoId(string descricao, string empresaOrigem)
        {
            #region [ Declarações ]
            string strSql;
            int id;
            SqlCommand cmdSelect;
            SqlCommand cmdInsert;
            #endregion

            strSql = "SELECT * FROM t_MKTP_REPASSE_TIPO_PAGAMENTO WHERE (Descricao=@Descricao AND EmpresaOrigem=@EmpresaOrigem)";
            cmdSelect = BD.criaSqlCommand();
            cmdSelect.CommandText = strSql;
            cmdSelect.Parameters.AddWithValue("@Descricao", descricao);
            cmdSelect.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
            id = (int)(cmdSelect.ExecuteScalar() ?? 0);
            if (id == 0)
            {
                strSql = "SET NOCOUNT ON; INSERT INTO t_MKTP_REPASSE_TIPO_PAGAMENTO (Descricao, EmpresaOrigem, DataHoraCadastro)" +
                    " VALUES (@Descricao, @EmpresaOrigem, getdate()); SELECT SCOPE_IDENTITY() AS Id;";
                cmdInsert = BD.criaSqlCommand();
                cmdInsert.CommandText = strSql;
                cmdInsert.Parameters.AddWithValue("@Descricao", descricao);
                cmdInsert.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
                id = Convert.ToInt32((cmdInsert.ExecuteScalar() ?? 0));
            }

            return id;
        }
        #endregion

        #region [ getStatusTransacaoId ]
        public static int getStatusTransacaoId(string descricao, string empresaOrigem)
        {
            #region [ Declarações ]
            string strSql;
            int id;
            SqlCommand cmdSelect;
            SqlCommand cmdInsert;
            #endregion

            strSql = "SELECT * FROM t_MKTP_REPASSE_STATUS_DESCRICAO WHERE (Descricao=@Descricao AND EmpresaOrigem=@EmpresaOrigem)";
            cmdSelect = BD.criaSqlCommand();
            cmdSelect.CommandText = strSql;
            cmdSelect.Parameters.AddWithValue("@Descricao", descricao);
            cmdSelect.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
            id = (int)(cmdSelect.ExecuteScalar() ?? 0);
            if (id == 0)
            {
                strSql = "SET NOCOUNT ON; INSERT INTO t_MKTP_REPASSE_STATUS_DESCRICAO (Descricao, EmpresaOrigem, DataHoraCadastro)" +
                    " VALUES (@Descricao, @EmpresaOrigem, getdate()); SELECT SCOPE_IDENTITY() AS Id;";
                cmdInsert = BD.criaSqlCommand();
                cmdInsert.CommandText = strSql;
                cmdInsert.Parameters.AddWithValue("@Descricao", descricao);
                cmdInsert.Parameters.AddWithValue("@EmpresaOrigem", empresaOrigem);
                id = Convert.ToInt32((cmdInsert.ExecuteScalar() ?? 0));
            }

            return id;
        }
        #endregion

        #region [ checkPedidoDados ]
        public static void checkPedidoDados(PlanilhaRepasseMktplaceN2 planilhaN2, out string strMsgErro)
        {
            #region [ Declarações ]
            string strSql;
            SqlCommand cmdSelect;
            SqlDataReader reader;
            #endregion

            strMsgErro = "";
            if (planilhaN2 == null)
            {
                strMsgErro = "Pedido não informado para checagem";
                return;
            }
            if (planilhaN2.Pedido.Length == 0) return;

            try
            {
                strSql = "SELECT TOP 1 vl_total_NF, st_entrega FROM t_PEDIDO WHERE (pedido_bs_x_marketplace=@pedido_bs_x_marketplace) ORDER BY data DESC";
                cmdSelect = BD.criaSqlCommand();
                cmdSelect.CommandText = strSql;
                cmdSelect.Parameters.AddWithValue("@pedido_bs_x_marketplace", planilhaN2.Pedido);
                reader = cmdSelect.ExecuteReader();
                if (!reader.Read())
                {
                    planilhaN2.PedidoExiste = false;
                    reader.Close();
                    return;
                }
                else
                {
                    planilhaN2.PedidoExiste = true;
                    planilhaN2.ValorTotalPedidoCorreto = (decimal)reader["vl_total_NF"];
                }
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                return;
            }

            reader.Close();
        } 
        #endregion

        public static bool isPedidoCancelado(string numeroMktplace, out string strMsgErro)
        {
            #region [ Declarações ]
            int intItensDevolvidos;
            string strSql;
            string stEntrega;
            string numPedido = "";
            SqlCommand cmdSelectPedido;
            SqlCommand cmdSelectPedidoItemDevolvido;
            SqlDataReader reader;
            #endregion

            strMsgErro = "";
            if (numeroMktplace == null || numeroMktplace.Trim().Length == 0)
            {
                strMsgErro = "Pedido não informado para checagem";
                return false;
            }

            try
            {
                strSql = "SELECT pedido, st_entrega FROM t_PEDIDO WHERE (pedido_bs_x_marketplace=@pedido_bs_x_marketplace)";
                cmdSelectPedido = BD.criaSqlCommand();
                cmdSelectPedido.CommandText = strSql;
                cmdSelectPedido.Parameters.AddWithValue("@pedido_bs_x_marketplace", numeroMktplace);
                reader = cmdSelectPedido.ExecuteReader();
                if (reader.Read())
                {
                    stEntrega = (string)reader["st_entrega"];
                    if (stEntrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
                    {
                        reader.Close();
                        return true;
                    }
                    else
                    {
                        numPedido = (string)reader["pedido"];
                    }
                }

                reader.Close();
                strSql = "SELECT COUNT(*) FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido=@pedido)";
                cmdSelectPedidoItemDevolvido = BD.criaSqlCommand();
                cmdSelectPedidoItemDevolvido.CommandText = strSql;
                cmdSelectPedidoItemDevolvido.Parameters.AddWithValue("@pedido", numPedido);
                intItensDevolvidos = (int)cmdSelectPedidoItemDevolvido.ExecuteScalar();
                if (intItensDevolvidos > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message;
                return false;
            }
        }

        #endregion
    }
}
