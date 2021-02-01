using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using System;

namespace ART3WebAPI.Models.Repository
{
    public class DataCompras2
    {
        public Compras[] Get(string tipo_periodo, string dt_inicio, string dt_termino, string fabricante, string produto, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string nf, string dt_nf_inicio, string dt_nf_termino, string visao, string detalhamento)
        {
            #region [ Declarações ]
            List<Compras> listaProduto = new List<Compras>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            DateTime dt1;
            DateTime dt2;
            int totalMeses;
            DateTime dtInicioDateType;
            DateTime dtTerminoDateType;
            string dtInicioFormatado;
            string dtTerminoFormatado;
            string s_sql_mes, s_where_temp, sqlString = "";
			#endregion

			if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
				dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);
			}
			else
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(dt_nf_inicio);
				dt2 = Global.converteDdMmYyyyParaDateTime(dt_nf_termino);
			}

			totalMeses = ((dt2.Year - dt1.Year) * 12) + dt2.Month - dt1.Month;
			dtInicioDateType = dt1;
			dtTerminoDateType = dt2;
			dtTerminoDateType = dtTerminoDateType.AddDays(1);
			dtInicioFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtInicioDateType);
			dtTerminoFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtTerminoDateType);

            #region [ Sintético por NF ]
            if (detalhamento == "SINTETICO_NF")
            {
                sqlString = "SELECT" +
                    " s_mes.documento," +
                    " s_mes.fabricante," +
                    " Sum(qtde) AS qtde_total," +
                    " Sum(qtde* s_mes.vl_custo2) AS valor_total";
            }
            #endregion

            #region [ Sintético por Fabricante ]
            if (detalhamento == "SINTETICO_FABR")
            {
                sqlString = "SELECT" +
                    " s_mes.fabricante," +
                    " Sum(qtde* s_mes.vl_custo2) AS valor";             
            }
            #endregion

            #region [ Sintético por Produto ]

            else if (detalhamento == "SINTETICO_PROD")
            {
                sqlString = "SELECT" +
                " s_mes.fabricante," +
                " s_mes.produto," +
                " Coalesce(Sum(qtde),0) AS qtde," +
                " Coalesce(Sum(qtde* s_mes.vl_custo2),0) AS valor";
            }

            #endregion

            #region [ Valor Referência Médio ]

            else if (detalhamento == "CUSTO_MEDIO")
            {
                sqlString = "SELECT" +
                " s_mes.fabricante," +
                " s_mes.produto," +
                " Coalesce(Sum(qtde),0) AS qtde," +
                " Coalesce(Sum(qtde* s_mes.vl_custo2),0) AS valor";
            }
            #endregion

            #region [ Valor Referência Individual ]

            else if (detalhamento == "CUSTO_INDIVIDUAL")
            {
                sqlString = "SELECT" +
                " s_mes.fabricante," +
                " s_mes.produto," +
                " s_mes.vl_custo2," +
                " Coalesce(Sum(qtde),0) AS qtde";
            }
            #endregion

            #region [ Qtde Vendida Mes a Mes ]
            s_sql_mes = "";

            if (detalhamento != "SINTETICO_NF")
            {
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

                    #region [ Sintético por Fabricante ]
                    if (detalhamento == "SINTETICO_FABR")
                    {
                        s_sql_mes = s_sql_mes + " (SELECT Coalesce(Sum(qtde * tei.vl_custo2),0) AS valor " +
                   " FROM t_ESTOQUE es " +
                   " INNER JOIN t_ESTOQUE_ITEM tei ON (es.id_estoque = tei.id_estoque) " +
                   " INNER JOIN t_PRODUTO pr on (tei.fabricante = pr.fabricante) AND (tei.produto = pr.produto) " +
                   " WHERE (es.fabricante = e.fabricante) " +
                       " AND (kit = 0) " +
                       " AND (entrada_especial = 0)" +
                       " AND (devolucao_status = 0) ";

						if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada < ", dt2Formatado) + ")");
						}
						else
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada < ", dt2Formatado) + ")");
						}

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(fabricante))
                        {
                            string[] v_fabricante = fabricante.Split('_');
                            for (int x = 0; x < v_fabricante.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (tei.fabricante = '" + v_fabricante[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(grupo))
                        {
                            string[] v_grupo = grupo.Split('_');
                            for (int x = 0; x < v_grupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (grupo = '" + v_grupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(subgrupo))
                        {
                            string[] v_subgrupo = subgrupo.Split('_');
                            for (int x = 0; x < v_subgrupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (subgrupo = '" + v_subgrupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        if (!string.IsNullOrEmpty(produto)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (tei.produto = '", produto) + "')");

                        if (!string.IsNullOrEmpty(btu)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (potencia_BTU = ", btu) + ")");

                        if (!string.IsNullOrEmpty(ciclo)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (ciclo = '", ciclo) + "')");

                        if (!string.IsNullOrEmpty(pos_mercado)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (posicao_mercado = '", pos_mercado) + "')");

                        if (!string.IsNullOrEmpty(nf)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (es.documento = '", nf) + "')");

                        s_sql_mes = string.Concat(s_sql_mes, string.Concat("GROUP BY tei.fabricante ") + ") AS mes" + i);

                        if (i < totalMeses) s_sql_mes = string.Concat(s_sql_mes, ", ");
                    }
                    #endregion

                    #region [ Valor Referência Individual ]
                    else if (detalhamento == "CUSTO_INDIVIDUAL")
                    {
                        s_sql_mes = s_sql_mes + "(SELECT Coalesce(Sum(qtde), 0) AS qtde " +
                   " FROM t_ESTOQUE es " +
                   " INNER JOIN t_ESTOQUE_ITEM tei ON (es.id_estoque = tei.id_estoque) " +
                   " INNER JOIN t_PRODUTO pr on (tei.fabricante = pr.fabricante) AND (tei.produto = pr.produto) " +
                   " WHERE (es.fabricante = e.fabricante) " +
                       " AND (pr.produto = p.produto) " +
                       " AND (tei.vl_custo2 = i.vl_custo2) " +
                       " AND (kit = 0) " +
                       " AND (entrada_especial = 0)" +
                       " AND (devolucao_status = 0) ";

						if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada < ", dt2Formatado) + ")");
						}
						else
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada < ", dt2Formatado) + ")");
						}

						s_where_temp = "";
                        if (!string.IsNullOrEmpty(fabricante))
                        {
                            string[] v_fabricante = fabricante.Split('_');
                            for (int x = 0; x < v_fabricante.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (tei.fabricante = '" + v_fabricante[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(grupo))
                        {
                            string[] v_grupo = grupo.Split('_');
                            for (int x = 0; x < v_grupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (grupo = '" + v_grupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(subgrupo))
                        {
                            string[] v_subgrupo = subgrupo.Split('_');
                            for (int x = 0; x < v_subgrupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (subgrupo = '" + v_subgrupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        if (!string.IsNullOrEmpty(produto)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (tei.produto = '", produto) + "')");

                        if (!string.IsNullOrEmpty(btu)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (potencia_BTU = ", btu) + ")");

                        if (!string.IsNullOrEmpty(ciclo)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (ciclo = '", ciclo) + "')");

                        if (!string.IsNullOrEmpty(pos_mercado)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (posicao_mercado = '", pos_mercado) + "')");

                        if (!string.IsNullOrEmpty(nf)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (es.documento = '", nf) + "')");

                        s_sql_mes = string.Concat(s_sql_mes, string.Concat("GROUP BY tei.fabricante,pr.produto,tei.vl_custo2") + ") AS mes" + i);

                        if (i < totalMeses) s_sql_mes = string.Concat(s_sql_mes, ", ");
                    }
                    #endregion

                    #region [ Sintético por Produto/Valor Referência Médio ]
                    else
                    {
                        s_sql_mes = s_sql_mes + "(SELECT Coalesce(Sum(qtde), 0) AS qtde " +
                  " FROM t_ESTOQUE es " +
                  " INNER JOIN t_ESTOQUE_ITEM tei ON (es.id_estoque = tei.id_estoque) " +
                  " INNER JOIN t_PRODUTO pr on (i.fabricante = pr.fabricante) AND (tei.produto = pr.produto) " +
                  " WHERE (es.fabricante = e.fabricante) " +
                      " AND (pr.produto = p.produto) " +
                      " AND (kit = 0) " +
                      " AND (entrada_especial = 0)" +
                      " AND (devolucao_status = 0) ";

						if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_entrada < ", dt2Formatado) + ")");
						}
						else
						{
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada >= ", dt1Formatado) + ")");
							s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (data_emissao_NF_entrada < ", dt2Formatado) + ")");
						}

						s_where_temp = "";
                        if (!string.IsNullOrEmpty(fabricante))
                        {
                            string[] v_fabricante = fabricante.Split('_');
                            for (int x = 0; x < v_fabricante.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (tei.fabricante = '" + v_fabricante[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(grupo))
                        {
                            string[] v_grupo = grupo.Split('_');
                            for (int x = 0; x < v_grupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (grupo = '" + v_grupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        s_where_temp = "";
                        if (!string.IsNullOrEmpty(subgrupo))
                        {
                            string[] v_subgrupo = subgrupo.Split('_');
                            for (int x = 0; x < v_subgrupo.GetLength(0); x++)
                            {
                                if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                                s_where_temp = string.Concat(s_where_temp, " (subgrupo = '" + v_subgrupo[x] + "')");
                            }
                            s_sql_mes = string.Concat(s_sql_mes, string.Concat("AND (", s_where_temp) + ")");
                        }

                        if (!string.IsNullOrEmpty(produto)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (tei.produto = '", produto) + "')");

                        if (!string.IsNullOrEmpty(btu)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (potencia_BTU = ", btu) + ")");

                        if (!string.IsNullOrEmpty(ciclo)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (ciclo = '", ciclo) + "')");

                        if (!string.IsNullOrEmpty(pos_mercado)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (posicao_mercado = '", pos_mercado) + "')");

                        if (!string.IsNullOrEmpty(nf)) s_sql_mes = string.Concat(s_sql_mes, string.Concat(" AND (es.documento = '", nf) + "')");

                        s_sql_mes = string.Concat(s_sql_mes, string.Concat("GROUP BY tei.fabricante,pr.produto") + ") AS mes" + i);

                        if (i < totalMeses) s_sql_mes = string.Concat(s_sql_mes, ", ");
                    }
                    #endregion

                } 
            }
            #endregion

            #region [ Consulta Completa ]

            if (detalhamento != "SINTETICO_NF")
            {
                for (int i = 0; i <= totalMeses; i++)
                {
                    sqlString = string.Concat(sqlString, ", Coalesce(mes" + i + ", 0) AS mes" + i);
                } 
            }

            sqlString = string.Concat(sqlString, " FROM (" +
                " SELECT i.fabricante" +
                " ,i.produto" +
                " ,e.kit" +
                " ,e.entrada_especial" +
                " ,e.devolucao_status" +
                " ,e.documento" +
                " ,grupo" +
                " ,subgrupo" +
                " ,potencia_BTU" +
                " ,ciclo" +
                " ,posicao_mercado" +
                " ,i.vl_custo2" +
                " ,i.preco_fabricante" +
                " ,qtde" +
                " ,data_entrada" +
                " ,data_emissao_NF_entrada");
            if (s_sql_mes != "") sqlString = sqlString + " ," + s_sql_mes;
            sqlString = sqlString + " FROM t_ESTOQUE e" +
		        " INNER JOIN t_ESTOQUE_ITEM i ON (e.id_estoque = i.id_estoque)" +
		        " INNER JOIN t_PRODUTO p ON (i.fabricante = p.fabricante)" +
			    " AND (i.produto = p.produto)";

			sqlString = string.Concat(sqlString, " WHERE");
			if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
			{
				sqlString = string.Concat(sqlString, string.Concat(" (data_entrada >= ", dtInicioFormatado) + ")");
				sqlString = string.Concat(sqlString, string.Concat(" AND (data_entrada < ", dtTerminoFormatado) + ")");
			}
			else
			{
				sqlString = string.Concat(sqlString, string.Concat(" (data_emissao_NF_entrada >= ", dtInicioFormatado) + ")");
				sqlString = string.Concat(sqlString, string.Concat(" AND (data_emissao_NF_entrada < ", dtTerminoFormatado) + ")");
			}

			sqlString = string.Concat(sqlString, ") s_mes " +
            " WHERE (s_mes.kit = 0) " +
            " AND (s_mes.entrada_especial = 0) " +
            " AND (s_mes.devolucao_status = 0)");

            s_where_temp = "";
            if (!string.IsNullOrEmpty(fabricante))
            {
                string[] v_fabricante = fabricante.Split('_');
                for (int i = 0; i < v_fabricante.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (s_mes.fabricante = '" + v_fabricante[i] + "')");
                }
                sqlString = string.Concat(sqlString, string.Concat("AND (", s_where_temp) + ")");
            }

            s_where_temp = "";
            if (!string.IsNullOrEmpty(grupo))
            {
                string[] v_grupo = grupo.Split('_');
                for (int i = 0; i < v_grupo.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (s_mes.grupo = '" + v_grupo[i] + "')");
                }
                sqlString = string.Concat(sqlString, string.Concat("AND (", s_where_temp) + ")");
            }

            s_where_temp = "";
            if (!string.IsNullOrEmpty(subgrupo))
            {
                string[] v_subgrupo = subgrupo.Split('_');
                for (int i = 0; i < v_subgrupo.GetLength(0); i++)
                {
                    if (s_where_temp != "") s_where_temp = string.Concat(s_where_temp, " OR");
                    s_where_temp = string.Concat(s_where_temp, " (s_mes.subgrupo = '" + v_subgrupo[i] + "')");
                }
                sqlString = string.Concat(sqlString, string.Concat("AND (", s_where_temp) + ")");
            }

            if (!string.IsNullOrEmpty(produto)) sqlString = string.Concat(sqlString, string.Concat(" AND (s_mes.produto = '", produto) + "')");

            if (!string.IsNullOrEmpty(btu)) sqlString = string.Concat(sqlString, string.Concat(" AND (s_mes.potencia_BTU = ", btu) + ")");

            if (!string.IsNullOrEmpty(ciclo)) sqlString = string.Concat(sqlString, string.Concat(" AND (s_mes.ciclo = '", ciclo) + "')");

            if (!string.IsNullOrEmpty(pos_mercado)) sqlString = string.Concat(sqlString, string.Concat(" AND (s_mes.posicao_mercado = '", pos_mercado) + "')");

            if (!string.IsNullOrEmpty(nf)) sqlString = string.Concat(sqlString, string.Concat(" AND (s_mes.documento = '", nf) + "')");

            if (detalhamento == "SINTETICO_FABR")
            {
                sqlString = string.Concat(sqlString, string.Concat(" GROUP BY s_mes.fabricante"));
                for (int i = 0; i <= totalMeses; i++)
                {
                    sqlString = string.Concat(sqlString, ", mes" + i);
                }
                sqlString = string.Concat(sqlString, string.Concat(" ORDER BY s_mes.fabricante"));
            }
            else if (detalhamento == "SINTETICO_NF")
            {
                sqlString = string.Concat(sqlString, string.Concat(" GROUP BY s_mes.documento, s_mes.fabricante"));
                sqlString = string.Concat(sqlString, string.Concat(" ORDER BY s_mes.fabricante, s_mes.documento"));
            }
            else if (detalhamento == "CUSTO_INDIVIDUAL")
            {
                sqlString = string.Concat(sqlString, string.Concat(" GROUP BY s_mes.fabricante, s_mes.produto, s_mes.vl_custo2"));
                for (int i = 0; i <= totalMeses; i++)
                {
                    sqlString = string.Concat(sqlString, ", mes" + i);
                }
                sqlString = string.Concat(sqlString, string.Concat(" ORDER BY s_mes.fabricante, s_mes.produto, s_mes.vl_custo2"));
            }
            else
            {
                sqlString = string.Concat(sqlString, string.Concat(" GROUP BY s_mes.fabricante, s_mes.produto"));
                for (int i = 0; i <= totalMeses; i++)
                {
                    sqlString = string.Concat(sqlString, ", mes" + i);
                }
                sqlString = string.Concat(sqlString, string.Concat(" ORDER BY s_mes.fabricante, s_mes.produto"));
            }

            cn.Open();
            #endregion

            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();

                try
                {
                    #region [ Sintético por NF ]
                    if (detalhamento == "SINTETICO_NF")
                    {
                        int idxNF = reader.GetOrdinal("documento");
                        int idxFabricante = reader.GetOrdinal("fabricante");
                        int idxQtde = reader.GetOrdinal("qtde_total");
                        int idxValor = reader.GetOrdinal("valor_total");

                        while (reader.Read())
                        {
                            
                            Compras _novo = new Compras();
                            _novo.NF = reader.GetString(idxNF);
                            _novo.Fabricante = reader.GetString(idxFabricante);
                            _novo.Qtde = reader.IsDBNull(idxQtde) ? 0 : reader.GetInt32(idxQtde);
                            _novo.Valor = reader.IsDBNull(idxValor) ? 0 : reader.GetDecimal(idxValor);

                            listaProduto.Add(_novo);
                        }
                    }
                    #endregion

                    #region [ Sintético por Fabricante ]
                    if (detalhamento == "SINTETICO_FABR")
                    {
                        int idxFabricante = reader.GetOrdinal("fabricante");
                        int idxValor = reader.GetOrdinal("valor");

                        while (reader.Read())
                        {
                            decimal[] qtdeMeses = new decimal[totalMeses + 1];

                            for (int i = 0; i < totalMeses + 1; i++)
                            {
                                qtdeMeses[i] = decimal.Parse(reader["mes" + i].ToString());
                            }
                            Compras _novo = new Compras();
                            _novo.Fabricante = reader.GetString(idxFabricante);
                            _novo.Valor = reader.IsDBNull(idxValor) ? 0 : reader.GetDecimal(idxValor);

                            _novo.Meses = qtdeMeses;
                            listaProduto.Add(_novo);
                        }
                    } 
                    #endregion

                    #region [ Sintético por Produto ]
                    else if (detalhamento == "SINTETICO_PROD")
                    {
                        int idxFabricante = reader.GetOrdinal("fabricante");
                        int idxProduto = reader.GetOrdinal("produto");
                        int idxQtde = reader.GetOrdinal("qtde");

                        while (reader.Read())
                        {
                            decimal[] qtdeMeses = new decimal[totalMeses + 1];

                            for (int i = 0; i < totalMeses + 1; i++)
                            {
                                qtdeMeses[i] = decimal.Parse(reader["mes" + i].ToString());
                            }

                            Compras _novo = new Compras();
                            _novo.Fabricante = reader.GetString(idxFabricante);
                            _novo.Produto = reader.IsDBNull(idxProduto) ? "" : reader.GetString(idxProduto);
                            _novo.Qtde = reader.IsDBNull(idxQtde) ? 0 : reader.GetInt32(idxQtde);
                            _novo.Meses = qtdeMeses;
                            listaProduto.Add(_novo);
                        }
                    } 
                    #endregion

                    #region [ Valor Refência Médio ]
                    else if (detalhamento == "CUSTO_MEDIO")
                    {
                        int idxFabricante = reader.GetOrdinal("fabricante");
                        int idxProduto = reader.GetOrdinal("produto");
                        int idxQtde = reader.GetOrdinal("qtde");
                        int idxValor = reader.GetOrdinal("valor");

                        while (reader.Read())
                        {
                            decimal[] qtdeMeses = new decimal[totalMeses + 1];

                            for (int i = 0; i < totalMeses + 1; i++)
                            {
                                qtdeMeses[i] = decimal.Parse(reader["mes" + i].ToString());
                            }

                            Compras _novo = new Compras();
                            _novo.Fabricante = reader.GetString(idxFabricante);
                            _novo.Produto = reader.IsDBNull(idxProduto) ? "" : reader.GetString(idxProduto);
                            _novo.Qtde = reader.IsDBNull(idxQtde) ? 0 : reader.GetInt32(idxQtde);
                            _novo.Valor = reader.IsDBNull(idxValor) ? 0 : reader.GetDecimal(idxValor);
                            _novo.Meses = qtdeMeses;
                            listaProduto.Add(_novo);
                        }

                    } 
                    #endregion

                    #region [ Valor Refência Individual ]
                    else if (detalhamento == "CUSTO_INDIVIDUAL")
                    {

                        int idxFabricante = reader.GetOrdinal("fabricante");
                        int idxProduto = reader.GetOrdinal("produto");
                        int idxQtde = reader.GetOrdinal("qtde");
                        int idxValor = reader.GetOrdinal("vl_custo2");


                        while (reader.Read())
                        {
                            decimal[] qtdeMeses = new decimal[totalMeses + 1];

                            for (int i = 0; i < totalMeses + 1; i++)
                            {
                                qtdeMeses[i] = decimal.Parse(reader["mes" + i].ToString());
                            }

                            Compras _novo = new Compras();
                            _novo.Fabricante = reader.GetString(idxFabricante);
                            _novo.Produto = reader.IsDBNull(idxProduto) ? "" : reader.GetString(idxProduto);
                            _novo.Qtde = reader.IsDBNull(idxQtde) ? 0 : reader.GetInt32(idxQtde);
                            _novo.Valor = reader.IsDBNull(idxValor) ? 0 : reader.GetDecimal(idxValor);
                            _novo.Meses = qtdeMeses;
                            listaProduto.Add(_novo);
                        }
                    } 
                    #endregion
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
            return listaProduto.ToArray();
        }

    }
}