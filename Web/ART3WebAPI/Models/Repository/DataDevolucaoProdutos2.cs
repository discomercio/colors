using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace ART3WebAPI.Models.Repository
{
    public class DataDevolucaoProdutos2
    {
        public List<DevolucaoProduto2Entity> Get(string usuario, string dt_devolucao_inicio, string dt_devolucao_termino, string fabricante, string produto, string pedido, string vendedor, string indicador, string captador, string lojas)
        {
            #region [ Declarações ]
            string strSql;
            string dtDevolucaoInicioSqlDateTime = "";
            string dtDevolucaoTerminoSqlDateTime = "";
            string[] vLojas;
            string[] vAux;
            int intParametroFlagPedidoMemorizacaoCompletaEnderecos;
            StringBuilder sbWhere = new StringBuilder("");
            StringBuilder sbWhereLoja = new StringBuilder("");
            StringBuilder sbAux = new StringBuilder("");
            DevolucaoProduto2Entity DevolProd2Entity;
            List<DevolucaoProduto2Entity> DevolProd2Lista = new List<DevolucaoProduto2Entity>(); ;
            SqlConnection cn;
            SqlCommand cmd;
            SqlDataReader reader;
            #endregion

            intParametroFlagPedidoMemorizacaoCompletaEnderecos = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.FLAG_PEDIDO_MEMORIZACAO_COMPLETA_ENDERECOS);

            #region [ Prepara acesso ao BD ]
            cn = new SqlConnection(BD.getConnectionString());
            cn.Open();
            cmd = new SqlCommand();
            cmd.Connection = cn;
			#endregion

			try // Finally: BD.fechaConexao(ref cn)
			{
                #region [ Formata datas para consulta no BD ]
                if (!string.IsNullOrEmpty(dt_devolucao_inicio))
                    dtDevolucaoInicioSqlDateTime = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(Global.converteDdMmYyyyParaDateTime(dt_devolucao_inicio));
                if (!string.IsNullOrEmpty(dt_devolucao_termino))
                    dtDevolucaoTerminoSqlDateTime = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(Global.converteDdMmYyyyParaDateTime(dt_devolucao_termino).AddDays(1));
                #endregion

                #region [ Filtro: Data Devolução ]
                if (dtDevolucaoInicioSqlDateTime != "")
                {
                    if (sbAux.Length > 0) sbAux.Append(" AND");
                    sbAux.Append(" (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " + dtDevolucaoInicioSqlDateTime + ")");
                }
                if (dtDevolucaoTerminoSqlDateTime != "")
                {
                    if (sbAux.Length > 0) sbAux.Append(" AND");
                    sbAux.Append(" (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " + dtDevolucaoTerminoSqlDateTime + ")");
                }

                if (sbAux.Length > 0)
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (" + sbAux.ToString() + ")");
                }
                #endregion

                #region [ Filtro: Fabricante ]
                if (!string.IsNullOrEmpty(fabricante))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" + fabricante + "')");
                }
                #endregion

                #region [ Filtro: Produto ]
                if (!string.IsNullOrEmpty(produto))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" + produto + "')");
                }
                #endregion

                #region [ Filtro: Pedido ]
                if (!string.IsNullOrEmpty(pedido))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_PEDIDO_ITEM_DEVOLVIDO.pedido = '" + pedido + "')");
                }
                #endregion

                #region [ Filtro: Vendedor ]
                if (!string.IsNullOrEmpty(vendedor))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_PEDIDO.vendedor = '" + vendedor + "')");
                }
                #endregion

                #region [ Filtro: Indicador ]
                if (!string.IsNullOrEmpty(indicador))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_PEDIDO__BASE.indicador = '" + indicador + "')");
                }
                #endregion

                #region [ Filtro: Captador ]
                if (!string.IsNullOrEmpty(captador))
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (t_ORCAMENTISTA_E_INDICADOR.captador = '" + captador + "')");
                }
                #endregion

                #region [ Filtro: Lojas ]
                if (!string.IsNullOrEmpty(lojas))
                {
                    vLojas = lojas.Split('_');
                    for (int i = 0; i < vLojas.Length; i++)
                    {
                        if (vLojas[i] != "")
                        {
                            vAux = vLojas[i].Split('-');
                            if (vAux.Length == 1)
                            {
                                if (sbWhereLoja.Length > 0) sbWhereLoja.Append(" OR");
                                sbWhereLoja.Append(" (t_PEDIDO.numero_loja = " + vLojas[i] + ")");
                            }
                            else
                            {
                                sbAux.Clear();
                                if (vAux[0] != "")
                                {
                                    if (sbAux.Length > 0) sbAux.Append(" AND");
                                    sbAux.Append(" (t_PEDIDO.numero_loja >= " + vAux[0] + ")");
                                }
                                if (vAux[1] != "")
                                {
                                    if (sbAux.Length > 0) sbAux.Append(" AND");
                                    sbAux.Append(" (t_PEDIDO.numero_loja <= " + vAux[1] + ")");
                                }
                                if (sbAux.Length > 0)
                                {
                                    if (sbWhereLoja.Length > 0) sbWhereLoja.Append(" OR");
                                    sbWhereLoja.Append(" (" + sbAux.ToString() + ")");
                                }
                            }
                        }
                    }
                }
                if (sbWhereLoja.Length > 0)
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(" (" + sbWhereLoja.ToString() + ")");
                }
                #endregion

                if (sbWhere.Length > 0) sbWhere.Insert(0, " WHERE");

                #region [ Monta consulta ]
                strSql = "SELECT " +
                    "*" +
                    "FROM (" +
                        "SELECT" +
                    " t_PEDIDO.data AS data_pedido," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data_devolucao," +
                    " t_PEDIDO.loja," +
                    " t_PEDIDO.numero_loja," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.id AS id_item_devolvido," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.fabricante," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.produto," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.descricao," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.descricao_html," +
                    " t_PEDIDO.pedido," +
                    " t_PEDIDO.obs_2," +
                    " t_PEDIDO__BASE.vendedor," +
                    " t_PEDIDO__BASE.indicador,";
                
                if (intParametroFlagPedidoMemorizacaoCompletaEnderecos == 1)
                {
                    strSql += " dbo.SqlClrUtilIniciaisEmMaiusculas(t_PEDIDO.endereco_nome) AS nome_cliente, ";
                }
                else
                {
                    strSql += " t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente,";
                }

                strSql +=
                    " t_PEDIDO_ITEM_DEVOLVIDO.motivo," +
                    " t_PEDIDO_ITEM_DEVOLVIDO.qtde," +
                    " (SELECT Count(*) FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS tAuxPIDBN INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO tAuxPID ON (tAuxPIDBN.id_item_devolvido=tAuxPID.id) WHERE (tAuxPID.pedido=t_PEDIDO.pedido) AND (anulado_status = 0)) AS qtde_msgs," +
                    " (" +
                        "SELECT" +
                            " TOP 1 t_ESTOQUE_MOVIMENTO.anulado_data" +
                        " FROM t_ESTOQUE INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque=t_ESTOQUE_MOVIMENTO.id_estoque)" +
                        " WHERE" +
                            " (t_ESTOQUE.devolucao_id_item_devolvido=t_PEDIDO_ITEM_DEVOLVIDO.id)" +
                            " AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM_DEVOLVIDO.fabricante)" +
                            " AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM_DEVOLVIDO.produto)" +
                            " AND (t_ESTOQUE_MOVIMENTO.estoque='DEV')" +
                            " AND (operacao='DEV')" +
                        " ORDER BY" +
                            " id_movimento" +
                    ") AS data_baixa" +
                " FROM t_PEDIDO" +
                    " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" +
                        " ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" +
                    " INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" +
                    " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" +
                        " ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
                    " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" +
                    sbWhere.ToString() +
                    ") t" +
                " ORDER BY" +
                    " data_pedido," +
                    " data_devolucao," +
                    " data_baixa," +
                    " fabricante," +
                    " produto," +
                    " descricao," +
                    " pedido";
                #endregion

                #region [ Executa consulta ]
                cmd.CommandText = strSql;
                reader = cmd.ExecuteReader();
                #endregion

                try
                {
                    #region [ Recupera os dados e adiciona na lista ]
                    while (reader.Read())
                    {
                        DevolProd2Entity = new DevolucaoProduto2Entity();
                        DevolProd2Entity.Cliente = reader["nome_cliente"].ToString();
                        DevolProd2Entity.Pedido = reader["pedido"].ToString();
                        DevolProd2Entity.Fabricante = reader["fabricante"].ToString();
                        DevolProd2Entity.Produto = reader["produto"].ToString();
                        DevolProd2Entity.Descricao = reader["descricao"].ToString();
                        DevolProd2Entity.Vendedor = reader["vendedor"].ToString();
                        DevolProd2Entity.Indicador = reader["indicador"].ToString();
                        DevolProd2Entity.Motivo = reader["motivo"].ToString();
                        DevolProd2Entity.Qtde = (int)Global.converteInteiro(reader["qtde"].ToString());
                        DevolProd2Entity.DataBaixa = BD.readToDateTime(reader["data_baixa"]);
                        DevolProd2Entity.DataDevolvido = BD.readToDateTime(reader["data_devolucao"]);
                        DevolProd2Entity.DataPedido = BD.readToDateTime(reader["data_pedido"]);

                        DevolProd2Lista.Add(DevolProd2Entity);
                    } 
                    #endregion
                }
                finally
                {
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                BD.fechaConexao(ref cn);
            }

            return DevolProd2Lista;
        }
    }
}