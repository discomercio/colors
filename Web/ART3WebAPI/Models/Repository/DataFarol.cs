using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using System;
using System.Linq;

namespace ART3WebAPI.Models.Repository
{
    public class DataFarol
    {
        public Farol[] Get(string dt_inicio, string dt_termino, string fabricante, string grupo, string btu, string ciclo, string pos_mercado, string loja)
        {
            List<Farol> listaFarol = new List<Farol>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());

            DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
            DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);

            int totalMeses = ((dt2.Year - dt1.Year) * 12) + dt2.Month - dt1.Month;


            DateTime dtInicioDateType = Global.converteDdMmYyyyParaDateTime(dt_inicio);
            DateTime dtTerminoDateType = Global.converteDdMmYyyyParaDateTime(dt_termino);
            dtTerminoDateType = dtTerminoDateType.AddDays(1);
            string dtInicioFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtInicioDateType);
            string dtTerminoFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtTerminoDateType);
            string s_sql_lista_base, s_sql_qtde_vendida, s_sql_qtde_devolvida, s_sql_qtde_estoque_venda, s_sql_qtde_vendida_mes, s_where_temp, s_where_loja, sqlString;
            string[] vLojas, vAux;

            if (loja == null) loja = "";
            vLojas = loja.Split(',');

            #region [ Relação de Produtos ]
            /* MONTA O SQL QUE SELECIONA A RELAÇÃO DE PRODUTOS
             * A LÓGICA CONSISTE EM SELECIONAR:
             * 1) PRODUTOS QUE TENHAM SALDO NO ESTOQUE DE VENDA E NO ESTOQUE DE SHOW ROOM
             * 2) PRODUTOS QUE CONSTEM COMO 'VENDÁVEIS'
            */
            s_sql_lista_base = "SELECT DISTINCT fabricante, produto FROM t_ESTOQUE_ITEM WHERE ((qtde - qtde_utilizada) > 0) ";

            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT fabricante " +
                    ", produto " +
                "FROM t_ESTOQUE_MOVIMENTO " +
                "WHERE (qtde > 0) " +
                "   AND (estoque = 'SHR')");
            if (!string.IsNullOrEmpty(fabricante))
            {
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }

            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT t_PRODUTO.fabricante, t_PRODUTO.produto FROM t_PRODUTO INNER JOIN (" +
                "SELECT DISTINCT fabricante, produto FROM t_PRODUTO_LOJA WHERE (vendavel = 'S') ");
            if (!string.IsNullOrEmpty(fabricante))
            {
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            s_sql_lista_base = string.Concat(s_sql_lista_base, ") tPL_AUX ON (t_PRODUTO.fabricante=tPL_AUX.fabricante) AND (t_PRODUTO.produto=tPL_AUX.produto) WHERE (excluido_status = 0) ");

            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (t_PRODUTO.fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }

            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT fabricante, produto FROM t_PRODUTO WHERE (farol_qtde_comprada > 0) ");
            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            #endregion

            #region [ Produtos Vendidos no Período ]

            s_sql_qtde_vendida = "SELECT SUM(qtde) FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido) " +
            "WHERE (t_PEDIDO_ITEM.fabricante=t_PROD_LISTA_BASE.fabricante) AND (t_PEDIDO_ITEM.produto=t_PROD_LISTA_BASE.produto) AND (st_entrega <> 'CAN')";

            if (!string.IsNullOrEmpty(loja))
            {
                s_where_loja = "";
                for (int i = vLojas.GetLowerBound(0); i <= vLojas.GetUpperBound(0); i++)
                {
                    if (vLojas[i] != "")
                    {
                        vAux = vLojas[i].Split('-');
                        if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                        {
                            if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                            s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[i] + ")";
                        }
                        else
                        {
                            s_where_temp = "";
                            if (vAux[vAux.GetLowerBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                            }
                            if (vAux[vAux.GetUpperBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                            }
                            if (s_where_temp != "")
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                            }
                        }
                    }
                }
                if (s_where_loja != "")
                {
                    s_sql_qtde_vendida = s_sql_qtde_vendida + " AND (" + s_where_loja + ")";
                }

            }

            s_sql_qtde_vendida = string.Concat(s_sql_qtde_vendida, string.Concat(" AND (t_PEDIDO.data >= ", dtInicioFormatado) + ")");

            s_sql_qtde_vendida = string.Concat(s_sql_qtde_vendida, string.Concat(" AND (t_PEDIDO.data < ", dtTerminoFormatado) + ")");
            #endregion

            #region [ Produtos Devolvidos no Período ]

            s_sql_qtde_devolvida = "SELECT" +
                                        " SUM(qtde)" +
                                    " FROM t_PEDIDO_ITEM_DEVOLVIDO" +
                                        " INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)" +
                                    " WHERE" +
                                        " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PROD_LISTA_BASE.fabricante)" +
                                        " AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PROD_LISTA_BASE.produto)";

            if (!string.IsNullOrEmpty(loja))
            {
                s_where_loja = "";
                for (int i = vLojas.GetLowerBound(0); i <= vLojas.GetUpperBound(0); i++)
                {
                    if (vLojas[i] != "")
                    {
                        vAux = vLojas[i].Split('-');
                        if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                        {
                            if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                            s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[i] + ")";
                        }
                        else
                        {
                            s_where_temp = "";
                            if (vAux[vAux.GetLowerBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                            }
                            if (vAux[vAux.GetUpperBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                            }
                            if (s_where_temp != "")
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                            }
                        }
                    }
                }
                if (s_where_loja != "")
                {
                    s_sql_qtde_devolvida = s_sql_qtde_devolvida + " AND (" + s_where_loja + ")";
                }
            }

            s_sql_qtde_devolvida = string.Concat(s_sql_qtde_devolvida, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= ", dtInicioFormatado) + ")");

            s_sql_qtde_devolvida = string.Concat(s_sql_qtde_devolvida, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < ", dtTerminoFormatado) + ")");
            #endregion

            #region [ Estoque de Vendas + Show-Room ]

            s_sql_qtde_estoque_venda = "SELECT SUM(qtde_total) FROM ( " +
                "SELECT SUM(qtde - qtde_utilizada) AS qtde_total " +
                "FROM t_ESTOQUE_ITEM " +
                "WHERE (t_ESTOQUE_ITEM.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                "	AND (t_ESTOQUE_ITEM.produto = t_PROD_LISTA_BASE.produto) " +
                "	AND ((qtde - qtde_utilizada) > 0) " +
                "UNION ALL " +
                "SELECT SUM(qtde) AS qtde_total " +
                "FROM t_ESTOQUE_MOVIMENTO " +
                "WHERE (t_ESTOQUE_MOVIMENTO.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                "	AND (t_ESTOQUE_MOVIMENTO.produto = t_PROD_LISTA_BASE.produto) " +
                "	AND (t_ESTOQUE_MOVIMENTO.anulado_status = 0) " +
                "	AND (t_ESTOQUE_MOVIMENTO.estoque = 'SHR') " +
                "   AND (qtde > 0) " +
                ") tCALC ";
            #endregion

            #region [ Qtde Vendida Mes a Mes ]
            s_sql_qtde_vendida_mes = "";

            for (int i = 0; i <= totalMeses; i++)
            {
                string mesDtInicial = "", mesDtFinal = "";
                if (i == 0)
                {
                    mesDtInicial = dt1.ToString("dd/MM/yyyy");
                }
                else
                {
                    mesDtInicial = (dt1.AddMonths(i)).AddDays(-dt1.Day + 1).ToString("dd/MM/yyyy");
                }
                if (i == totalMeses)
                {
                    mesDtFinal = dt2.AddDays(1).ToString("dd/MM/yyyy");
                }
                else
                {
                    mesDtFinal = (dt1.AddMonths(i + 1).AddDays(-dt1.Day + 1)).ToString("dd/MM/yyyy");
                }
                DateTime mesDtInicalDateTime = Global.converteDdMmYyyyParaDateTime(mesDtInicial);
                DateTime mesDtFinalDateTime = Global.converteDdMmYyyyParaDateTime(mesDtFinal);
                string dt1Formatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(mesDtInicalDateTime);
                string dt2Formatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(mesDtFinalDateTime);

                s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + "(SELECT SUM(qtde) FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido) " +
                    "WHERE (t_PEDIDO_ITEM.fabricante=t_PROD_LISTA_BASE.fabricante) AND (t_PEDIDO_ITEM.produto=t_PROD_LISTA_BASE.produto) AND (st_entrega <> 'CAN')";

                if (!string.IsNullOrEmpty(loja))
                {
                    s_where_loja = "";
                    for (int j = vLojas.GetLowerBound(0); j <= vLojas.GetUpperBound(0); j++)
                    {
                        if (vLojas[j] != "")
                        {
                            vAux = vLojas[j].Split('-');
                            if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[j] + ")";
                            }
                            else
                            {
                                s_where_temp = "";
                                if (vAux[vAux.GetLowerBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                                }
                                if (vAux[vAux.GetUpperBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                                }
                                if (s_where_temp != "")
                                {
                                    if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                    s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                                }
                            }
                        }
                    }
                    if (s_where_loja != "")
                    {
                        s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + " AND (" + s_where_loja + ")";
                    }
                }

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO.data >= ", dt1Formatado) + ")");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO.data < ", dt2Formatado) + ")) AS mes" + i + ", ");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, "(SELECT SUM(qtde) FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido) WHERE (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                    "AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PROD_LISTA_BASE.produto) ");

                if (!string.IsNullOrEmpty(loja))
                {
                    s_where_loja = "";
                    for (int j = vLojas.GetLowerBound(0); j <= vLojas.GetUpperBound(0); j++)
                    {
                        if (vLojas[j] != "")
                        {
                            vAux = vLojas[j].Split('-');
                            if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[j] + ")";
                            }
                            else
                            {
                                s_where_temp = "";
                                if (vAux[vAux.GetLowerBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                                }
                                if (vAux[vAux.GetUpperBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                                }
                                if (s_where_temp != "")
                                {
                                    if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                    s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                                }
                            }
                        }
                    }
                    if (s_where_loja != "")
                    {
                        s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + " AND (" + s_where_loja + ")";
                    }
                }

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= ", dt1Formatado) + ")");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < ", dt2Formatado) + ")");
                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, ") AS dev_mes" + i);

                if (i < totalMeses) s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, ", ");
            }
            #endregion

            #region [ Consulta Completa ]

            sqlString = "SELECT fabricante, produto, descricao, descricao_html, grupo, potencia_BTU, ciclo, posicao_mercado, descontinuado, Coalesce(farol_qtde_comprada, 0) AS farol_qtde_comprada, Coalesce(qtde_vendida, 0) AS qtde_vendida, Coalesce(qtde_devolvida, 0) AS qtde_devolvida, Coalesce(qtde_estoque_venda, 0) AS qtde_estoque_venda";

            for (int i = 0; i <= totalMeses; i++)
            {
                string mesDtInicial = (dt1.AddMonths(i)).ToString("yyyy-MM-dd");
                string mesDtFinal = (dt1.AddMonths(i + 1)).ToString("yyyy-MM-dd");

                sqlString = string.Concat(sqlString, ", Coalesce(mes" + i + ", 0) AS mes" + i);
                sqlString = string.Concat(sqlString, ", Coalesce(dev_mes" + i + ", 0) AS dev_mes" + i);

            }

            sqlString = string.Concat(sqlString, " FROM (SELECT t_PROD_LISTA_BASE.fabricante, t_PROD_LISTA_BASE.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Coalesce(t_PRODUTO.grupo, '') AS grupo, Coalesce(t_PRODUTO.potencia_BTU, '') AS potencia_BTU, Coalesce(t_PRODUTO.ciclo, '') AS ciclo, Coalesce(t_PRODUTO.posicao_mercado, '') AS posicao_mercado, Coalesce(t_PRODUTO.descontinuado, '') AS descontinuado," +
                                " t_PRODUTO.farol_qtde_comprada, (" + s_sql_qtde_vendida + ") AS qtde_vendida, (" + s_sql_qtde_devolvida + ") AS qtde_devolvida, (" + s_sql_qtde_estoque_venda + ") AS qtde_estoque_venda, " + s_sql_qtde_vendida_mes +
                                " FROM (" + s_sql_lista_base + ") t_PROD_LISTA_BASE" +
                                " LEFT JOIN t_PRODUTO ON (t_PROD_LISTA_BASE.fabricante = t_PRODUTO.fabricante) AND (t_PROD_LISTA_BASE.produto = t_PRODUTO.produto)" +
                                " WHERE" +
                                " (descricao <> '.')" +
                                " AND (descricao <> '*')" +
                                ") tREL" +
                                " WHERE" +
                                " (UPPER(descontinuado) <> 'S') ");

            s_where_temp = "";
            if (!string.IsNullOrEmpty(grupo))
            {
                string[] v_grupo = grupo.Split('_');
                for (int i = 0; i < v_grupo.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (grupo = '" + v_grupo[i] + "')");
                }
                sqlString = string.Concat(sqlString, string.Concat("AND (", s_where_temp) + ")");
            }

            if (!string.IsNullOrEmpty(btu)) sqlString = string.Concat(sqlString, string.Concat(" AND (potencia_BTU = ", Global.digitos(btu)) + ")");

            if (!string.IsNullOrEmpty(ciclo)) sqlString = string.Concat(sqlString, string.Concat(" AND (ciclo = '", ciclo) + "')");

            if (!string.IsNullOrEmpty(pos_mercado)) sqlString = string.Concat(sqlString, string.Concat(" AND (posicao_mercado = '", pos_mercado) + "')");

            sqlString = string.Concat(string.Concat("SELECT * FROM (", sqlString) + ") tRelFinal ");
            sqlString = string.Concat(sqlString, "WHERE (NOT (((qtde_vendida - qtde_devolvida) <= 0) AND (qtde_estoque_venda = 0) AND (farol_qtde_comprada = 0))) ");
            sqlString = string.Concat(sqlString, "ORDER BY fabricante, produto");
            #endregion

            cn.Open();

            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();

                try
                {
                    int idxFabricante = reader.GetOrdinal("fabricante");
                    int idxProduto = reader.GetOrdinal("produto");
                    int idxDescricao = reader.GetOrdinal("descricao");
                    int idxDescricaoHtml = reader.GetOrdinal("descricao_html");
                    int idxGrupo = reader.GetOrdinal("grupo");
                    int idxPontenciaBTU = reader.GetOrdinal("potencia_BTU");
                    int idxCiclo = reader.GetOrdinal("ciclo");
                    int idxPosicaoMercado = reader.GetOrdinal("posicao_mercado");
                    int idxDescontinuado = reader.GetOrdinal("descontinuado");
                    int idxQtdeComprada = reader.GetOrdinal("farol_qtde_comprada");
                    int idxQtdeVendida = reader.GetOrdinal("qtde_vendida");
                    int idxQtdeDevolvida = reader.GetOrdinal("qtde_devolvida");
                    int idxQtdeEstoqueVenda = reader.GetOrdinal("qtde_estoque_venda");

                    while (reader.Read())
                    {
                        int[] qtdeVendidaMeses = new int[totalMeses + 1];
                        int[] qtdeDevolvidaMeses = new int[totalMeses + 1];
                        for (int i = 0; i < totalMeses + 1; i++)
                        {
                            qtdeVendidaMeses[i] = int.Parse(reader["mes" + i].ToString());
                            qtdeDevolvidaMeses[i] = int.Parse(reader["dev_mes" + i].ToString());
                            qtdeVendidaMeses[i] = qtdeVendidaMeses[i] - qtdeDevolvidaMeses[i];
                        }

                        Farol _novo = new Farol(reader.GetString(idxFabricante), reader.GetString(idxProduto), reader.GetString(idxDescricao), reader.GetString(idxDescricaoHtml),
                            reader.GetString(idxGrupo), reader.GetInt32(idxPontenciaBTU), reader.GetString(idxCiclo), reader.GetString(idxPosicaoMercado),
                            reader.GetInt32(idxQtdeComprada), reader.GetInt32(idxQtdeVendida), reader.GetInt32(idxQtdeDevolvida), reader.GetInt32(idxQtdeEstoqueVenda), 0, qtdeVendidaMeses);

                        listaFarol.Add(_novo);
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            finally
            {
                cn.Close();
            }

            return listaFarol.ToArray();
        }

        public Farol[] GetV3(string dt_inicio, string dt_termino, string fabricante, string grupo, string btu, string ciclo, string pos_mercado, string loja)
        {
            List<Farol> listaFarol = new List<Farol>();
            List<Farol> listaFarolUnificados = new List<Farol>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());

            DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
            DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);
            SqlDataReader readerComposto;
            SqlCommand cmdComposto = new SqlCommand();
            SqlDataAdapter daDataAdapter = new SqlDataAdapter();
            DataTable dtbCompostos = new DataTable();
            DataTable dtbCompostosItem = new DataTable();


            int totalMeses = ((dt2.Year - dt1.Year) * 12) + dt2.Month - dt1.Month;


            DateTime dtInicioDateType = Global.converteDdMmYyyyParaDateTime(dt_inicio);
            DateTime dtTerminoDateType = Global.converteDdMmYyyyParaDateTime(dt_termino);
            dtTerminoDateType = dtTerminoDateType.AddDays(1);
            string dtInicioFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtInicioDateType);
            string dtTerminoFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtTerminoDateType);
            string s_sql_lista_base, s_sql_qtde_vendida, s_sql_qtde_devolvida, s_sql_qtde_estoque_venda, s_sql_qtde_vendida_mes, s_where_temp, s_where_loja, sqlString;
            string[] vLojas, vAux;
            int countItensComposicao;

            if (loja == null) loja = "";
            vLojas = loja.Split(',');

            #region [ Relação de Produtos ]
            /* MONTA O SQL QUE SELECIONA A RELAÇÃO DE PRODUTOS
             * A LÓGICA CONSISTE EM SELECIONAR:
             * 1) PRODUTOS QUE TENHAM SALDO NO ESTOQUE DE VENDA E NO ESTOQUE DE SHOW ROOM
             * 2) PRODUTOS QUE CONSTEM COMO 'VENDÁVEIS'
            */
            s_sql_lista_base = "SELECT DISTINCT fabricante, produto FROM t_ESTOQUE_ITEM WHERE ((qtde - qtde_utilizada) > 0) ";

            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT fabricante " +
                    ", produto " +
                "FROM t_ESTOQUE_MOVIMENTO " +
                "WHERE (qtde > 0) " +
                "   AND (estoque = 'SHR')");
            if (!string.IsNullOrEmpty(fabricante))
            {
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }

            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT t_PRODUTO.fabricante, t_PRODUTO.produto FROM t_PRODUTO INNER JOIN (" +
                "SELECT DISTINCT fabricante, produto FROM t_PRODUTO_LOJA WHERE (vendavel = 'S') ");
            if (!string.IsNullOrEmpty(fabricante))
            {
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            s_sql_lista_base = string.Concat(s_sql_lista_base, ") tPL_AUX ON (t_PRODUTO.fabricante=tPL_AUX.fabricante) AND (t_PRODUTO.produto=tPL_AUX.produto) WHERE (excluido_status = 0) ");

            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (t_PRODUTO.fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }

            s_sql_lista_base = string.Concat(s_sql_lista_base, " UNION " +
                "SELECT DISTINCT fabricante, produto FROM t_PRODUTO WHERE (farol_qtde_comprada > 0) ");
            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (fabricante = '" + v_fabricante[i] + "')");
                }
                s_sql_lista_base = string.Concat(s_sql_lista_base, "AND (" + s_where_temp + ")");
            }
            #endregion

            #region [ Produtos Vendidos no Período ]

            s_sql_qtde_vendida = "SELECT SUM(qtde) FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido) " +
            "WHERE (t_PEDIDO_ITEM.fabricante=t_PROD_LISTA_BASE.fabricante) AND (t_PEDIDO_ITEM.produto=t_PROD_LISTA_BASE.produto) AND (st_entrega <> 'CAN')";

            if (!string.IsNullOrEmpty(loja))
            {
                s_where_loja = "";
                for (int i = vLojas.GetLowerBound(0); i <= vLojas.GetUpperBound(0); i++)
                {
                    if (vLojas[i] != "")
                    {
                        vAux = vLojas[i].Split('-');
                        if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                        {
                            if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                            s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[i] + ")";
                        }
                        else
                        {
                            s_where_temp = "";
                            if (vAux[vAux.GetLowerBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                            }
                            if (vAux[vAux.GetUpperBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                            }
                            if (s_where_temp != "")
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                            }
                        }
                    }
                }
                if (s_where_loja != "")
                {
                    s_sql_qtde_vendida = s_sql_qtde_vendida + " AND (" + s_where_loja + ")";
                }

            }

            s_sql_qtde_vendida = string.Concat(s_sql_qtde_vendida, string.Concat(" AND (t_PEDIDO.data >= ", dtInicioFormatado) + ")");

            s_sql_qtde_vendida = string.Concat(s_sql_qtde_vendida, string.Concat(" AND (t_PEDIDO.data < ", dtTerminoFormatado) + ")");
            #endregion

            #region [ Produtos Devolvidos no Período ]

            s_sql_qtde_devolvida = "SELECT" +
                                        " SUM(qtde)" +
                                    " FROM t_PEDIDO_ITEM_DEVOLVIDO" +
                                        " INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)" +
                                    " WHERE" +
                                        " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PROD_LISTA_BASE.fabricante)" +
                                        " AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PROD_LISTA_BASE.produto)";

            if (!string.IsNullOrEmpty(loja))
            {
                s_where_loja = "";
                for (int i = vLojas.GetLowerBound(0); i <= vLojas.GetUpperBound(0); i++)
                {
                    if (vLojas[i] != "")
                    {
                        vAux = vLojas[i].Split('-');
                        if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                        {
                            if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                            s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[i] + ")";
                        }
                        else
                        {
                            s_where_temp = "";
                            if (vAux[vAux.GetLowerBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                            }
                            if (vAux[vAux.GetUpperBound(0)] != "")
                            {
                                if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                            }
                            if (s_where_temp != "")
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                            }
                        }
                    }
                }
                if (s_where_loja != "")
                {
                    s_sql_qtde_devolvida = s_sql_qtde_devolvida + " AND (" + s_where_loja + ")";
                }
            }

            s_sql_qtde_devolvida = string.Concat(s_sql_qtde_devolvida, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= ", dtInicioFormatado) + ")");

            s_sql_qtde_devolvida = string.Concat(s_sql_qtde_devolvida, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < ", dtTerminoFormatado) + ")");
            #endregion

            #region [ Estoque de Vendas + Show-Room ]

            s_sql_qtde_estoque_venda = "SELECT SUM(qtde_total) FROM ( " +
                "SELECT SUM(qtde - qtde_utilizada) AS qtde_total " +
                "FROM t_ESTOQUE_ITEM " +
                "WHERE (t_ESTOQUE_ITEM.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                "	AND (t_ESTOQUE_ITEM.produto = t_PROD_LISTA_BASE.produto) " +
                "	AND ((qtde - qtde_utilizada) > 0) " +
                "UNION ALL " +
                "SELECT SUM(qtde) AS qtde_total " +
                "FROM t_ESTOQUE_MOVIMENTO " +
                "WHERE (t_ESTOQUE_MOVIMENTO.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                "	AND (t_ESTOQUE_MOVIMENTO.produto = t_PROD_LISTA_BASE.produto) " +
                "	AND (t_ESTOQUE_MOVIMENTO.anulado_status = 0) " +
                "	AND (t_ESTOQUE_MOVIMENTO.estoque = 'SHR') " +
                "   AND (qtde > 0) " +
                ") tCALC ";
            #endregion

            #region [ Qtde Vendida Mes a Mes ]
            s_sql_qtde_vendida_mes = "";

            for (int i = 0; i <= totalMeses; i++)
            {
                string mesDtInicial = "", mesDtFinal = "";
                if (i == 0)
                {
                    mesDtInicial = dt1.ToString("dd/MM/yyyy");
                }
                else
                {
                    mesDtInicial = (dt1.AddMonths(i)).AddDays(-dt1.Day + 1).ToString("dd/MM/yyyy");
                }
                if (i == totalMeses)
                {
                    mesDtFinal = dt2.AddDays(1).ToString("dd/MM/yyyy");
                }
                else
                {
                    mesDtFinal = (dt1.AddMonths(i + 1).AddDays(-dt1.Day + 1)).ToString("dd/MM/yyyy");
                }
                DateTime mesDtInicalDateTime = Global.converteDdMmYyyyParaDateTime(mesDtInicial);
                DateTime mesDtFinalDateTime = Global.converteDdMmYyyyParaDateTime(mesDtFinal);
                string dt1Formatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(mesDtInicalDateTime);
                string dt2Formatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(mesDtFinalDateTime);

                s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + "(SELECT SUM(qtde) FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido) " +
                    "WHERE (t_PEDIDO_ITEM.fabricante=t_PROD_LISTA_BASE.fabricante) AND (t_PEDIDO_ITEM.produto=t_PROD_LISTA_BASE.produto) AND (st_entrega <> 'CAN')";

                if (!string.IsNullOrEmpty(loja))
                {
                    s_where_loja = "";
                    for (int j = vLojas.GetLowerBound(0); j <= vLojas.GetUpperBound(0); j++)
                    {
                        if (vLojas[j] != "")
                        {
                            vAux = vLojas[j].Split('-');
                            if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[j] + ")";
                            }
                            else
                            {
                                s_where_temp = "";
                                if (vAux[vAux.GetLowerBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                                }
                                if (vAux[vAux.GetUpperBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                                }
                                if (s_where_temp != "")
                                {
                                    if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                    s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                                }
                            }
                        }
                    }
                    if (s_where_loja != "")
                    {
                        s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + " AND (" + s_where_loja + ")";
                    }
                }

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO.data >= ", dt1Formatado) + ")");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO.data < ", dt2Formatado) + ")) AS mes" + i + ", ");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, "(SELECT SUM(qtde) FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido) WHERE (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PROD_LISTA_BASE.fabricante) " +
                    "AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PROD_LISTA_BASE.produto) ");

                if (!string.IsNullOrEmpty(loja))
                {
                    s_where_loja = "";
                    for (int j = vLojas.GetLowerBound(0); j <= vLojas.GetUpperBound(0); j++)
                    {
                        if (vLojas[j] != "")
                        {
                            vAux = vLojas[j].Split('-');
                            if (vAux.GetLowerBound(0) == vAux.GetUpperBound(0))
                            {
                                if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                s_where_loja = s_where_loja + " (t_PEDIDO.numero_loja = " + vLojas[j] + ")";
                            }
                            else
                            {
                                s_where_temp = "";
                                if (vAux[vAux.GetLowerBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja >= " + vAux[vAux.GetLowerBound(0)] + ")";
                                }
                                if (vAux[vAux.GetUpperBound(0)] != "")
                                {
                                    if (s_where_temp != "") s_where_temp = s_where_temp + " AND";
                                    s_where_temp = s_where_temp + " (t_PEDIDO.numero_loja <= " + vAux[vAux.GetUpperBound(0)] + ")";
                                }
                                if (s_where_temp != "")
                                {
                                    if (s_where_loja != "") s_where_loja = s_where_loja + " OR";
                                    s_where_loja = string.Concat(s_where_loja, " (" + s_where_temp + ")");
                                }
                            }
                        }
                    }
                    if (s_where_loja != "")
                    {
                        s_sql_qtde_vendida_mes = s_sql_qtde_vendida_mes + " AND (" + s_where_loja + ")";
                    }
                }

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= ", dt1Formatado) + ")");

                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, string.Concat(" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < ", dt2Formatado) + ")");
                s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, ") AS dev_mes" + i);

                if (i < totalMeses) s_sql_qtde_vendida_mes = string.Concat(s_sql_qtde_vendida_mes, ", ");
            }
            #endregion

            #region [ Consulta Completa ]

            sqlString = "SELECT fabricante, produto, descricao, descricao_html, custo, grupo, potencia_BTU, ciclo, posicao_mercado, descontinuado, Coalesce(farol_qtde_comprada, 0) AS farol_qtde_comprada, Coalesce(qtde_vendida, 0) AS qtde_vendida, Coalesce(qtde_devolvida, 0) AS qtde_devolvida, Coalesce(qtde_estoque_venda, 0) AS qtde_estoque_venda";

            for (int i = 0; i <= totalMeses; i++)
            {
                string mesDtInicial = (dt1.AddMonths(i)).ToString("yyyy-MM-dd");
                string mesDtFinal = (dt1.AddMonths(i + 1)).ToString("yyyy-MM-dd");

                sqlString = string.Concat(sqlString, ", Coalesce(mes" + i + ", 0) AS mes" + i);
                sqlString = string.Concat(sqlString, ", Coalesce(dev_mes" + i + ", 0) AS dev_mes" + i);

            }

            sqlString = string.Concat(sqlString, " FROM (SELECT t_PROD_LISTA_BASE.fabricante, t_PROD_LISTA_BASE.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PRODUTO.preco_fabricante AS custo, Coalesce(t_PRODUTO.grupo, '') AS grupo, Coalesce(t_PRODUTO.potencia_BTU, '') AS potencia_BTU, Coalesce(t_PRODUTO.ciclo, '') AS ciclo, Coalesce(t_PRODUTO.posicao_mercado, '') AS posicao_mercado, Coalesce(t_PRODUTO.descontinuado, '') AS descontinuado," +
                                " t_PRODUTO.farol_qtde_comprada, (" + s_sql_qtde_vendida + ") AS qtde_vendida, (" + s_sql_qtde_devolvida + ") AS qtde_devolvida, (" + s_sql_qtde_estoque_venda + ") AS qtde_estoque_venda, " + s_sql_qtde_vendida_mes +
                                " FROM (" + s_sql_lista_base + ") t_PROD_LISTA_BASE" +
                                " LEFT JOIN t_PRODUTO ON (t_PROD_LISTA_BASE.fabricante = t_PRODUTO.fabricante) AND (t_PROD_LISTA_BASE.produto = t_PRODUTO.produto)" +
                                " WHERE" +
                                " (descricao <> '.')" +
                                " AND (descricao <> '*')" +
                                ") tREL" +
                                " WHERE" +
                                " (UPPER(descontinuado) <> 'S') ");

            s_where_temp = "";
            if (!string.IsNullOrEmpty(grupo))
            {
                string[] v_grupo = grupo.Split('_');
                for (int i = 0; i < v_grupo.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (grupo = '" + v_grupo[i] + "')");
                }
                sqlString = string.Concat(sqlString, string.Concat("AND (", s_where_temp) + ")");
            }

            if (!string.IsNullOrEmpty(btu)) sqlString = string.Concat(sqlString, string.Concat(" AND (potencia_BTU = ", Global.digitos(btu)) + ")");

            if (!string.IsNullOrEmpty(ciclo)) sqlString = string.Concat(sqlString, string.Concat(" AND (ciclo = '", ciclo) + "')");

            if (!string.IsNullOrEmpty(pos_mercado)) sqlString = string.Concat(sqlString, string.Concat(" AND (posicao_mercado = '", pos_mercado) + "')");

            sqlString = string.Concat(string.Concat("SELECT * FROM (", sqlString) + ") tRelFinal ");
            sqlString = string.Concat(sqlString, "WHERE (NOT (((qtde_vendida - qtde_devolvida) <= 0) AND (qtde_estoque_venda = 0) AND (farol_qtde_comprada = 0))) ");
            sqlString = string.Concat(sqlString, "ORDER BY fabricante, produto");
            #endregion

            cn.Open();

            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();

                try
                {
                    int idxFabricante = reader.GetOrdinal("fabricante");
                    int idxProduto = reader.GetOrdinal("produto");
                    int idxDescricao = reader.GetOrdinal("descricao");
                    int idxDescricaoHtml = reader.GetOrdinal("descricao_html");
                    int idxGrupo = reader.GetOrdinal("grupo");
                    int idxPontenciaBTU = reader.GetOrdinal("potencia_BTU");
                    int idxCiclo = reader.GetOrdinal("ciclo");
                    int idxPosicaoMercado = reader.GetOrdinal("posicao_mercado");
                    int idxDescontinuado = reader.GetOrdinal("descontinuado");
                    int idxQtdeComprada = reader.GetOrdinal("farol_qtde_comprada");
                    int idxQtdeVendida = reader.GetOrdinal("qtde_vendida");
                    int idxQtdeDevolvida = reader.GetOrdinal("qtde_devolvida");
                    int idxQtdeEstoqueVenda = reader.GetOrdinal("qtde_estoque_venda");
                    int idxCusto = reader.GetOrdinal("custo");

                    while (reader.Read())
                    {
                        int[] qtdeVendidaMeses = new int[totalMeses + 1];
                        int[] qtdeDevolvidaMeses = new int[totalMeses + 1];
                        for (int i = 0; i < totalMeses + 1; i++)
                        {
                            qtdeVendidaMeses[i] = int.Parse(reader["mes" + i].ToString());
                            qtdeDevolvidaMeses[i] = int.Parse(reader["dev_mes" + i].ToString());
                            qtdeVendidaMeses[i] = qtdeVendidaMeses[i] - qtdeDevolvidaMeses[i];
                        }

                        Farol _novo = new Farol(reader.GetString(idxFabricante), reader.GetString(idxProduto), reader.GetString(idxDescricao), reader.GetString(idxDescricaoHtml),
                            reader.GetString(idxGrupo), reader.GetInt32(idxPontenciaBTU), reader.GetString(idxCiclo), reader.GetString(idxPosicaoMercado),
                            reader.GetInt32(idxQtdeComprada), reader.GetInt32(idxQtdeVendida), reader.GetInt32(idxQtdeDevolvida), reader.GetInt32(idxQtdeEstoqueVenda), reader.GetDecimal(idxCusto), qtdeVendidaMeses);

                        listaFarol.Add(_novo);
                    }
                }
                finally
                {
                    reader.Close();
                }

                #region [ Agrupar por código unificado ]
                if (listaFarol.Count > 0)
                {
                    sqlString = "SELECT" +
                                " tECPC.fabricante_composto," +
                                " tECPC.produto_composto," +
                                " tECPC.descricao," +
                                " tF.nome AS nome_fabricante," +
                                " tP.grupo," +
                                " tP.potencia_BTU," +
                                " tP.ciclo," +
                                " tP.posicao_mercado" +
                            " FROM t_EC_PRODUTO_COMPOSTO tECPC" +
                                " LEFT JOIN t_PRODUTO tP ON ((tECPC.fabricante_composto = tP.fabricante) AND (tECPC.produto_composto = tP.produto))" +
                                " LEFT JOIN t_FABRICANTE tF ON (tECPC.fabricante_composto = tF.fabricante)" +
                            " ORDER BY" +
                                " tECPC.fabricante_composto," +
                                " tECPC.produto_composto";

                    cmd.CommandText = sqlString;
                    daDataAdapter.SelectCommand = cmd;
                    daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                    daDataAdapter.Fill(dtbCompostos);

                    for (int i = 0; i < dtbCompostos.Rows.Count; i++)
                    {
                        sqlString = "SELECT " +
                                    " fabricante_item," +
                                    " produto_item," +
                                    " qtde" +
                                    " FROM t_EC_PRODUTO_COMPOSTO_ITEM" +
                                    " WHERE" +
                                    " (fabricante_composto = '" + BD.readToString(dtbCompostos.Rows[i]["fabricante_composto"]) + "')" +
                                    " AND (produto_composto = '" + BD.readToString(dtbCompostos.Rows[i]["produto_composto"]) + "')";

                        countItensComposicao = 0;
                        cmd.CommandText = sqlString;
                        reader = cmd.ExecuteReader();

                        try
                        {
                            List<Farol> fListAux = new List<Farol>();
                            while (reader.Read())
                            {
                                countItensComposicao++;

                                Farol fAux;
                                Farol fCompostoItem = new Farol();
                                fAux = listaFarol.FirstOrDefault(p => p.Produto == (string)reader["produto_item"] && p.Fabricante == (string)reader["fabricante_item"]);
                                if (fAux != null)
                                {
                                    int[] qm = new int[totalMeses + 1];
                                    for (int n = 0; n <= totalMeses; n++)
                                        qm[n] = fAux.Meses[n];
                                    fCompostoItem.Produto = BD.readToString(dtbCompostos.Rows[i]["produto_composto"]);
                                    fCompostoItem.Fabricante = BD.readToString(dtbCompostos.Rows[i]["fabricante_composto"]);
                                    fCompostoItem.Descricao = BD.readToString(dtbCompostos.Rows[i]["descricao"]);
                                    fCompostoItem.Custo = fAux.Custo;
                                    fCompostoItem.Qtde_composto_item = (short)reader["qtde"];
                                    fCompostoItem.Saldo = fAux.Saldo;
                                    fCompostoItem.Qtde_vendida = fAux.Qtde_vendida;
                                    fCompostoItem.Qtde_estoque_venda = fAux.Qtde_estoque_venda;
                                    fCompostoItem.Qtde_devolvida = fAux.Qtde_devolvida;
                                    fCompostoItem.Ciclo = BD.readToString(dtbCompostos.Rows[i]["ciclo"]);
                                    fCompostoItem.Farol_qtde_comprada = fAux.Farol_qtde_comprada;
                                    fCompostoItem.Grupo = BD.readToString(dtbCompostos.Rows[i]["grupo"]);
                                    fCompostoItem.Meses = qm;
                                    fCompostoItem.Posicao_mercado = BD.readToString(dtbCompostos.Rows[i]["posicao_mercado"]);
                                    fCompostoItem.Potencia_BTU = BD.readToInt(!Convert.IsDBNull(dtbCompostos.Rows[i]["potencia_BTU"]) ? dtbCompostos.Rows[i]["potencia_BTU"] : 0);
                                    fListAux.Add(fCompostoItem);
                                    fAux.IsItemComposicao = true;
                                }
                            }

                            if (fListAux.Count > 0)
                            {
                                if (fListAux.Count == countItensComposicao)
                                {
                                    Farol fComposto = new Farol();
                                    fComposto = fListAux.First();

                                    fComposto.Qtde_estoque_venda = fListAux.Min(x => (x.Qtde_estoque_venda / x.Qtde_composto_item));
                                    fComposto.Qtde_vendida = fListAux.Min(x => (x.Qtde_vendida / x.Qtde_composto_item));
                                    fComposto.Custo = fListAux.Sum(x => x.Custo * x.Qtde_composto_item);

                                    for (int j = 0; j <= totalMeses; j++)
                                    {
                                        fComposto.Meses[j] = fListAux.Min(x => (x.Meses[j] / x.Qtde_composto_item));
                                    }

                                    listaFarolUnificados.Add(fComposto);
                                }

                                fListAux.Clear();
                            }
                        }
                        finally
                        {
                            reader.Close();
                        }
                    }
                    listaFarolUnificados.AddRange(listaFarol.Where(p => p.IsItemComposicao == false));
                }
                #endregion
            }
            finally
            {
                cn.Close();
            }

            return listaFarolUnificados.ToArray();
        }
    }
}