using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using ART3WebAPI.Models.Repository;
using System;

namespace ART3WebAPI.Models.Repository
{
    public class DataEstatisticasOcorrencias
    {
        public Ocorrencias[] Get(string dt_inicio,string dt_termino,string motivo_ocorrencia,string tp_ocorrencia,string transportadora,string vendedor,string indicador, string UF,string loja)
        {
            List<Ocorrencias> ListaOcorrencias = new List<Ocorrencias>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            
            
            DateTime dtInicioDateType = Global.converteDdMmYyyyParaDateTime(dt_inicio);
            DateTime dtTerminoDateType = Global.converteDdMmYyyyParaDateTime(dt_termino).AddDays(1);         
            string dtInicioFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtInicioDateType);
            string dtTerminoFormatado = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtTerminoDateType);
            string s_where_temp, sqlString, s;
            int intParametroFlagPedidoMemorizacaoCompletaEnderecos;

            intParametroFlagPedidoMemorizacaoCompletaEnderecos = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.FLAG_PEDIDO_MEMORIZACAO_COMPLETA_ENDERECOS);

            #region [s_where e Consulta]
            s_where_temp = "";
            if (!string.IsNullOrEmpty(dt_inicio))
            {
                s_where_temp = s_where_temp + " AND (tPO.dt_cadastro >= " + dtInicioFormatado + ")";
					
            }
            if (!string.IsNullOrEmpty(dt_termino))
            {
                s_where_temp = s_where_temp + " AND (tPO.dt_cadastro < " + dtTerminoFormatado + ")";
					
            }

            if (!string.IsNullOrEmpty(tp_ocorrencia))
            {
                s_where_temp = s_where_temp + " AND (tPO.tipo_ocorrencia = '" + tp_ocorrencia + "')";

            }
            if (!string.IsNullOrEmpty(motivo_ocorrencia))
            {
                s_where_temp = s_where_temp + " AND (tPO.cod_motivo_abertura = '" + motivo_ocorrencia + "')";

            }
            if ((!string.IsNullOrEmpty(transportadora)) && (transportadora != "0"))
            {
                s_where_temp = s_where_temp + " AND (tP.transportadora_id = '" + transportadora + "')";

            }
            if (!string.IsNullOrEmpty(vendedor))
            {
                s_where_temp = s_where_temp + " AND (tP.vendedor = '" + vendedor + "')";

            }
            if (!string.IsNullOrEmpty(indicador))
            {
                s_where_temp = s_where_temp + " AND (tP.indicador = '" + indicador + "')";

            }
            if (!string.IsNullOrEmpty(UF))
            {
                if (intParametroFlagPedidoMemorizacaoCompletaEnderecos == 1)
                {
                    s_where_temp = s_where_temp + " AND " +
                            "(" +
                                "((tP.st_end_entrega <> 0) And (tP.EndEtg_uf = '" + UF + "'))" +
                                " OR " +
                                "((tP.st_end_entrega = 0) And (tP.endereco_uf = '" + UF + "'))" +
                            ")";
                }
                else
                {
                    s_where_temp = s_where_temp + " AND " +
                            "(" +
                                "((tP.st_end_entrega <> 0) And (tP.EndEtg_uf = '" + UF + "'))" +
                                " OR " +
                                "((tP.st_end_entrega = 0) And (tC.uf = '" + UF + "'))" +
                            ")";
                }
            }
            string s_where_loja = "";
            if (!string.IsNullOrEmpty(loja))
            {
                string[] v_loja = loja.Split('_');
                for (int i = 0; i < v_loja.Length; i++)
                {
                    if (v_loja[i] != "")
                    {
                        string[] v = v_loja[i].Split('-');
                        if (v.Length - 1 == 0)
                        {
                            if (s_where_loja != "") { s_where_loja = s_where_loja + " OR"; }
                            s_where_loja = s_where_loja + " (tP.numero_loja = " + v_loja[i] + ")";
                        }
                        else
                        {
                            s = "";
                            if (v[0] != "")
                            {
                                if (s != "") { s = s + " AND"; }
                                s = s + " (tP.numero_loja >= " + v[0] + ")";
                            }
                            if (v[v.Length - 1] != "")
                            {
                                if (s != "") { s = s + " AND"; }
                                s = s + " (tP.numero_loja <= " + v[v.Length - 1] + ")";
                            }
                            if (s != "")
                            {
                                if (s_where_loja != "") { s_where_loja = s_where_loja + " OR"; }
                                s_where_loja = s_where_loja + " (" + s + ")";
                            }
                        }
                    }
                }
            }
            

            if (s_where_loja != "")
            {
                s_where_temp = s_where_temp + " AND (" + s_where_loja + ")";
            }

            sqlString = "SELECT" +
                            " tPO.id," +
                            " tPO.pedido," +
                            " tPO.usuario_cadastro," +
                            " tPO.dt_cadastro," +
                            " tPO.dt_hr_cadastro," +
                            " tPO.contato," +
                            " tPO.texto_ocorrencia," +
                            " tPO.tipo_ocorrencia," +
                            " tPO.cod_motivo_abertura," +
                            " tPO.texto_finalizacao," +
                            " tP.transportadora_id,";

            if (intParametroFlagPedidoMemorizacaoCompletaEnderecos == 1)
            {
                sqlString += " dbo.SqlClrUtilIniciaisEmMaiusculas(tP.endereco_nome) AS nome_cliente, ";
            }
            else
            {
                sqlString += " tC.nome_iniciais_em_maiusculas AS nome_cliente, ";
            }

            sqlString += " (" +
                                "SELECT" +
                                    " TOP 1 NFe_numero_NF" +
                                " FROM t_NFe_EMISSAO tNE" +
                                " WHERE" +
                                    " (tNE.pedido=tPO.pedido)" +
                                    " AND (tipo_NF = '1')" +
                                    " AND (st_anulado = 0)" +
                                    " AND (codigo_retorno_NFe_T1 = 1)" +
                                " ORDER BY" +
                                    " id DESC" +
                            ") AS numeroNFe," +
                            " (" +
                                "SELECT" +
                                    " Count(*)" +
                                " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" +
                                " WHERE" +
                                    " (id_ocorrencia=tPO.id)" +
                                    " AND (fluxo_mensagem='" + Global.Cte.COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA + "')" +
                            ") AS qtde_msg_central" +
                        " FROM t_PEDIDO_OCORRENCIA tPO" +
                            " INNER JOIN t_PEDIDO tP ON (tPO.pedido=tP.pedido)" +
                            " INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" +
                        " WHERE" +
                            " (tPO.finalizado_status <> 0)" +
                             s_where_temp;

            sqlString = "SELECT * FROM (" + sqlString + ") t ORDER BY dt_hr_cadastro, id";
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
                    int idxPedido = reader.GetOrdinal("pedido");
                    int idxNF = reader.GetOrdinal("numeroNFe");
                    int idxTransportadora = reader.GetOrdinal("transportadora_id");
                    int idxOcorrencia = reader.GetOrdinal("texto_ocorrencia");
                    int idxTipoOcorrencia = reader.GetOrdinal("tipo_ocorrencia");
                    int idxCod_motivo_abertura = reader.GetOrdinal("cod_motivo_abertura");

                    while (reader.Read())
                    {
                        string consulta;
                        Ocorrencias _novo = new Ocorrencias();
                        _novo.Pedido = reader.IsDBNull(idxPedido) ? "" : reader.GetString(idxPedido);
                        _novo.NF = reader.IsDBNull(idxNF) ? 0 : reader.GetInt32(idxNF);
                        _novo.Transportadora = reader.IsDBNull(idxTransportadora) ? "" : reader.GetString(idxTransportadora);

                        consulta = BD.obtem_descricao_tabela_t_codigo_descricao(Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, reader.IsDBNull(idxCod_motivo_abertura) ? "" : reader.GetString(idxCod_motivo_abertura));
                        if (!string.IsNullOrEmpty(consulta))
                        {
                            _novo.Ocorrencia = consulta + " " + reader.GetString(idxOcorrencia);
                            
                        }
                        else
                        {
                            _novo.Ocorrencia = reader.IsDBNull(idxOcorrencia) ? "" : reader.GetString(idxOcorrencia);
                        }

                        consulta = BD.obtem_descricao_tabela_t_codigo_descricao(Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, reader.IsDBNull(idxTipoOcorrencia) ? "" : reader.GetString(idxTipoOcorrencia));
                        
                        if (!string.IsNullOrEmpty(consulta))
                        {
                            _novo.TipoOcorrencia = consulta;

                        }
                        else
                        {
                            _novo.TipoOcorrencia = reader.IsDBNull(idxTipoOcorrencia) ? "" : reader.GetString(idxTipoOcorrencia);
                        }
                        
                        ListaOcorrencias.Add(_novo);
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

            return ListaOcorrencias.ToArray();
        }
        

    }
}