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
    public class DataOcorrencias
    {
        public OcorrenciasStatus[] Get(string oc_status,string transportadora,string loja)
        {
            List<OcorrenciasStatus> ListaOcorrenciasStatus = new List<OcorrenciasStatus>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            
            
            //string s_where_temp, sqlString, s;
            string sqlString;



            #region [s_where e Consulta]

            sqlString = "SELECT" +
						" tPO.id," +
						" tPO.pedido," +
						" tPO.usuario_cadastro," +
						" tPO.dt_cadastro," +
						" tPO.dt_hr_cadastro," +
						" tPO.contato," +
						" tPO.ddd_1," +
						" tPO.tel_1," +
						" tPO.ddd_2," +
						" tPO.tel_2," +
                        "'('+ tPO.ddd_1+')'+tPO.tel_1 AS telefone," +
                        " tPO.texto_ocorrencia," +
                        " tPO.tipo_ocorrencia," +
                        " tP.loja," +
						" tP.loja AS pedido_loja," +
						" tP.transportadora_id," +
						" tC.nome_iniciais_em_maiusculas AS nome_cliente, tCD.codigo, tCD.descricao," +
						" tEmit.apelido as CD, " +
                        " CASE WHEN tP.st_end_entrega <> 0 THEN tP.EndEtg_uf ELSE tC.uf END AS UF, " +
                        " CASE WHEN tP.st_end_entrega <> 0 THEN tP.EndEtg_cidade ELSE tC.cidade END AS Cidade, " +
						" (" +
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
						") AS qtde_msg_central," +
						" (" +
							" SELECT Count(*)" +
							   " FROM t_PEDIDO_OCORRENCIA_MENSAGEM INNER JOIN t_PEDIDO_OCORRENCIA ON (t_PEDIDO_OCORRENCIA_MENSAGEM.id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" +
							   " INNER JOIN t_PEDIDO ON (t_PEDIDO_OCORRENCIA.pedido=t_PEDIDO.pedido)" + 
							   " WHERE (id_ocorrencia = tPO.id)" +
							   " AND (t_PEDIDO.loja = '" + Global.Cte.Loja.ArClube + "')" + 
						") AS qtde_msg," +
						" CASE WHEN (" +
                                     " (" +
                                        "SELECT" +
                                            " Count(*)" +
                                        " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" +
                                        " WHERE" +
                                            " (id_ocorrencia=tPO.id)" +
                                            " AND (fluxo_mensagem='" + Global.Cte.COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA + "')" +
                                    ") = 0 AND " +
                                    " (" +
                                        " SELECT Count(*)" +
                                           " FROM t_PEDIDO_OCORRENCIA_MENSAGEM INNER JOIN t_PEDIDO_OCORRENCIA ON (t_PEDIDO_OCORRENCIA_MENSAGEM.id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" +
                                           " INNER JOIN t_PEDIDO ON (t_PEDIDO_OCORRENCIA.pedido=t_PEDIDO.pedido)" +
                                           " WHERE (id_ocorrencia = tPO.id)" +
                                           " AND (t_PEDIDO.loja = '" + Global.Cte.Loja.ArClube + "')" +
                                    ") = 0" +
                        ") THEN 'Aberta' ELSE 'Em Andamento' END AS status_ocorrencia" +
				   " FROM t_PEDIDO_OCORRENCIA tPO" +
						" INNER JOIN t_PEDIDO tP ON (tPO.pedido=tP.pedido)" +
						" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" +
						" INNER JOIN t_NFe_EMITENTE tEmit ON (tP.id_nfe_emitente=tEmit.id)" +
						" LEFT JOIN t_CODIGO_DESCRICAO tCD ON (tPO.cod_motivo_abertura=tCD.codigo) AND (tCD.grupo='" + Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA + "')" +
				   " WHERE" +
						" (finalizado_status = 0)";
						
						
					if ((!string.IsNullOrEmpty(transportadora)) && (transportadora != "0"))
					{						
						sqlString = sqlString + " AND (tP.transportadora_id = < " + transportadora + ")";					
					}

					if (!string.IsNullOrEmpty(loja))
					{
						sqlString = sqlString + " AND (tP.numero_loja = < " + loja + ")";					
					}

					if (oc_status == "ABERTA") 
					{
						sqlString = "SELECT * FROM (" + sqlString + ") t WHERE (qtde_msg_central = 0 AND qtde_msg = 0) ORDER BY dt_hr_cadastro, id";
					}
					else if (oc_status == "EM_ANDAMENTO")
					{
						sqlString = "SELECT * FROM (" + sqlString + ") t WHERE (qtde_msg_central > 0 OR qtde_msg > 0) ORDER BY dt_hr_cadastro, id";
					}
					else
					{
						sqlString = "SELECT * FROM (" + sqlString + ") t ORDER BY dt_hr_cadastro, id";
					}

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
                    int idxLoja = reader.GetOrdinal("loja");
                    int idxCD = reader.GetOrdinal("CD");
                    int idxPedido = reader.GetOrdinal("pedido");
                    int idxNF = reader.GetOrdinal("numeroNFe");
                    int idxCliente = reader.GetOrdinal("nome_cliente");
                    int idxUF = reader.GetOrdinal("UF");
                    int idxCidade = reader.GetOrdinal("cidade");
                    int idxTransportadora = reader.GetOrdinal("transportadora_id");
                    int idxContato = reader.GetOrdinal("contato");
                    int idxTelefone = reader.GetOrdinal("telefone");
                    int idxOcorrencia = reader.GetOrdinal("texto_ocorrencia");
                    int idxTipoOcorrencia = reader.GetOrdinal("tipo_ocorrencia");
                    int idxStatus = reader.GetOrdinal("status_ocorrencia");

                    while (reader.Read())
                    {
                        string consulta;
                        OcorrenciasStatus _novo = new OcorrenciasStatus();
                        _novo.Loja = reader.IsDBNull(idxLoja) ? "" : reader.GetString(idxLoja);
						_novo.CD = reader.IsDBNull(idxCD) ? "" : reader.GetString(idxCD);
						_novo.Pedido = reader.IsDBNull(idxPedido) ? "" : reader.GetString(idxPedido);
                        _novo.NF = reader.IsDBNull(idxNF) ? 0 : reader.GetInt32(idxNF);
                        _novo.Cliente = reader.IsDBNull(idxCliente) ? "" : reader.GetString(idxCliente);
						_novo.UF = reader.IsDBNull(idxUF) ? "" : reader.GetString(idxUF);
						_novo.Cidade = reader.IsDBNull(idxCidade) ? "" : reader.GetString(idxCidade);
                        _novo.Transportadora = reader.IsDBNull(idxTransportadora) ? "" : reader.GetString(idxTransportadora);
						_novo.Contato = reader.IsDBNull(idxContato) ? "" : reader.GetString(idxContato);
						_novo.Telefone = reader.IsDBNull(idxTelefone) ? "" : reader.GetString(idxTelefone);
						_novo.Ocorrencia = reader.IsDBNull(idxOcorrencia) ? "" : reader.GetString(idxOcorrencia);
						_novo.TipoOcorrencia = reader.IsDBNull(idxTipoOcorrencia) ? "" : reader.GetString(idxTipoOcorrencia);
						_novo.Status = reader.IsDBNull(idxStatus) ? "" : reader.GetString(idxStatus);

                        /* consulta = BD.obtem_descricao_tabela_t_codigo_descricao(Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, reader.IsDBNull(idxCod_motivo_abertura) ? "" : reader.GetString(idxCod_motivo_abertura));
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
                        } */
                        
                        ListaOcorrenciasStatus.Add(_novo);
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

            return ListaOcorrenciasStatus.ToArray();
        }
        

    }
}