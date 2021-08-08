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
	public class DataRelPreDevolucao
	{
		public List<RelPreDevolucaoEntity> Get(Guid? httpRequestId, string usuario, string loja, string filtro_status, string filtro_data_inicio, string filtro_data_termino, string filtro_lojas)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DataRelPreDevolucao.Get()";
			DateTime dti = DateTime.MinValue;
			DateTime dtf = DateTime.MinValue;
			int intParametroFlagPedidoMemorizacaoCompletaEnderecos;
			int idx_id_devolucao;
			int idx_pedido;
			int idx_usuario_cadastro;
			int idx_dt_cadastro;
			int idx_dt_hr_cadastro;
			int idx_status;
			int idx_status_data_hora;
			int idx_cod_procedimento;
			int idx_cod_devolucao_motivo;
			int idx_descricao_devolucao_motivo;
			int idx_vl_devolucao;
			int idx_cod_credito_transacao;
			int idx_loja;
			int idx_data_pedido;
			int idx_vendedor;
			int idx_transportadora_id;
			int idx_indicador;
			int idx_cliente_nome;
			int idx_vl_pedido;
			StringBuilder sbSql = new StringBuilder("");
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbWhereLoja = new StringBuilder("");
			StringBuilder sbAux = new StringBuilder("");
			string s_where_campo_dt_periodo = "";
			string[] vLojas;
			string[] vAux;
			RelPreDevolucaoEntity linhaRel;
			List<RelPreDevolucaoEntity> resultado = new List<RelPreDevolucaoEntity>();
			SqlConnection cn;
			SqlCommand cmd;
			SqlDataReader reader;
			#endregion

			#region [ Prepara campos de filtragem ]
			if ((filtro_data_inicio ?? "").Length > 0) dti = Global.converteDdMmYyyyParaDateTime(filtro_data_inicio);
			if ((filtro_data_termino ?? "").Length > 0) dtf = Global.converteDdMmYyyyParaDateTime(filtro_data_termino);
			#endregion

			#region [ Prepara SQL ]
			intParametroFlagPedidoMemorizacaoCompletaEnderecos = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.FLAG_PEDIDO_MEMORIZACAO_COMPLETA_ENDERECOS);

			sbSql.Append("SELECT " +
						" tPD.id AS id_devolucao," +
						" tPD.pedido," +
						" tPD.usuario_cadastro," +
						" tPD.dt_cadastro," +
						" tPD.dt_hr_cadastro," +
						" tPD.status," +
						" tPD.status_data_hora," +
						" tPD.cod_procedimento," +
						" tPD.cod_devolucao_motivo," +
						" tCD.descricao AS descricao_devolucao_motivo," +
						" tPD.vl_devolucao," +
						" tPD.cod_credito_transacao," +
						" tP.loja," +
						" tP.data AS data_pedido," +
						" tP.vendedor," +
						" tP.transportadora_id," +
						" tP.indicador,");
			if (intParametroFlagPedidoMemorizacaoCompletaEnderecos == 1)
			{
				sbSql.Append(" tP.endereco_nome_iniciais_em_maiusculas AS cliente_nome");
			}
			else
			{
				sbSql.Append(" tC.nome_iniciais_em_maiusculas AS cliente_nome");
			}

			sbSql.Append(", (SELECT SUM(qtde * preco_NF) FROM t_PEDIDO_ITEM WHERE t_PEDIDO_ITEM.pedido = tP.pedido) AS vl_pedido");

			sbSql.Append(" FROM t_PEDIDO_DEVOLUCAO tPD" +
							" INNER JOIN t_PEDIDO tP ON (tPD.pedido = tP.pedido)" +
							" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" +
							" LEFT JOIN t_CODIGO_DESCRICAO tCD ON (tCD.grupo = '" + Global.Cte.GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO + "') AND (tCD.codigo = tPD.cod_devolucao_motivo)");

			#region [ Filtro: status ]
			if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.CADASTRADA))
			{
				s_where_campo_dt_periodo = "tPD.dt_cadastro";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.CADASTRADA.ToString() + ")");
			}
			else if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.EM_ANDAMENTO))
			{
				s_where_campo_dt_periodo = "tPD.dt_cadastro";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.EM_ANDAMENTO.ToString() + ")");
			}
			else if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.MERCADORIA_RECEBIDA))
			{
				s_where_campo_dt_periodo = "tPD.dt_mercadoria_recebida";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.MERCADORIA_RECEBIDA.ToString() + ")");
			}
			else if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.REPROVADA))
			{
				s_where_campo_dt_periodo = "tPD.dt_reprovado";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.REPROVADA.ToString() + ")");
			}
			else if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.FINALIZADA))
			{
				s_where_campo_dt_periodo = "tPD.dt_finalizado";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.FINALIZADA.ToString() + ")");
			}
			else if (filtro_status.Equals(Global.Cte.Relatorio.RelPreDevolucao.FILTRO_STATUS.CANCELADA))
			{
				s_where_campo_dt_periodo = "tPD.dt_cancelado";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tPD.status = " + Global.Cte.Relatorio.RelPreDevolucao.ST_PEDIDO_DEVOLUCAO.CANCELADA.ToString() + ")");
			}
			#endregion

			#region [ Filtro: período ]
			if ((dti > DateTime.MinValue) && (s_where_campo_dt_periodo.Length>0))
			{
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (" + s_where_campo_dt_periodo + " >= " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dti) + ")");
			}

			if ((dtf > DateTime.MinValue) && (s_where_campo_dt_periodo.Length > 0))
			{
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (" + s_where_campo_dt_periodo + " < " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtf.AddDays(1)) + ")");
			}
			#endregion

			#region [ Filtro: loja ]
			// Se o relatório está sendo solicitado através do módulo Loja, assegura que ocorra a filtragem pela loja
			if (((filtro_lojas ?? "").Length == 0) && ((loja ?? "").Length > 0)) filtro_lojas = loja;

			if ((filtro_lojas ?? "").Length > 0)
			{
				filtro_lojas = filtro_lojas.Replace('_', ',');
				vLojas = filtro_lojas.Split(',');

				for (int i = 0; i < vLojas.Length; i++)
				{
					if (vLojas[i] != "")
					{
						vAux = vLojas[i].Split('-');
						if (vAux.Length == 1)
						{
							if (sbWhereLoja.Length > 0) sbWhereLoja.Append(" OR");
							sbWhereLoja.Append(" (tP.numero_loja = " + vLojas[i] + ")");
						}
						else
						{
							sbAux.Clear();
							if (vAux[0] != "")
							{
								if (sbAux.Length > 0) sbAux.Append(" AND");
								sbAux.Append(" (tP.numero_loja >= " + vAux[0] + ")");
							}
							if (vAux[1] != "")
							{
								if (sbAux.Length > 0) sbAux.Append(" AND");
								sbAux.Append(" (tP.numero_loja <= " + vAux[1] + ")");
							}
							if (sbAux.Length > 0)
							{
								if (sbWhereLoja.Length > 0) sbWhereLoja.Append(" OR");
								sbWhereLoja.Append(" (" + sbAux.ToString() + ")");
							}
						}
					}
				}

				if (sbWhereLoja.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (" + sbWhereLoja.ToString() + ")");
				}
			}
			#endregion

			if (sbWhere.Length > 0) sbWhere.Insert(0, " WHERE");

			sbSql.Append(sbWhere.ToString());

			sbSql.Append(" ORDER BY tPD.status_data_hora DESC");
			#endregion

			#region [ Abre conexão com o BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			#endregion

			try // Finally: BD.fechaConexao(ref cn)
			{
				#region [ Executa consulta ]
				cmd = new SqlCommand();
				cmd.Connection = cn;
				cmd.CommandText = sbSql.ToString();
				reader = cmd.ExecuteReader();
				#endregion

				try // Finally: reader.Close()
				{
					idx_id_devolucao = reader.GetOrdinal("id_devolucao");
					idx_pedido = reader.GetOrdinal("pedido");
					idx_usuario_cadastro = reader.GetOrdinal("usuario_cadastro");
					idx_dt_cadastro = reader.GetOrdinal("dt_cadastro");
					idx_dt_hr_cadastro = reader.GetOrdinal("dt_hr_cadastro");
					idx_status = reader.GetOrdinal("status");
					idx_status_data_hora = reader.GetOrdinal("status_data_hora");
					idx_cod_procedimento = reader.GetOrdinal("cod_procedimento");
					idx_cod_devolucao_motivo = reader.GetOrdinal("cod_devolucao_motivo");
					idx_descricao_devolucao_motivo = reader.GetOrdinal("descricao_devolucao_motivo");
					idx_vl_devolucao = reader.GetOrdinal("vl_devolucao");
					idx_cod_credito_transacao = reader.GetOrdinal("cod_credito_transacao");
					idx_loja = reader.GetOrdinal("loja");
					idx_data_pedido = reader.GetOrdinal("data_pedido");
					idx_vendedor = reader.GetOrdinal("vendedor");
					idx_transportadora_id = reader.GetOrdinal("transportadora_id");
					idx_indicador = reader.GetOrdinal("indicador");
					idx_cliente_nome = reader.GetOrdinal("cliente_nome");
					idx_vl_pedido = reader.GetOrdinal("vl_pedido");

					while (reader.Read())
					{
						linhaRel = new RelPreDevolucaoEntity();
						linhaRel.id_devolucao = reader.GetInt32(idx_id_devolucao);
						linhaRel.pedido = reader.GetString(idx_pedido);
						linhaRel.usuario_cadastro = reader.GetString(idx_usuario_cadastro);
						linhaRel.dt_cadastro = reader.GetDateTime(idx_dt_cadastro);
						linhaRel.dt_hr_cadastro = reader.GetDateTime(idx_dt_hr_cadastro);
						linhaRel.status = reader.GetByte(idx_status);
						linhaRel.status_data_hora = reader.GetDateTime(idx_status_data_hora);
						linhaRel.cod_procedimento = reader.GetString(idx_cod_procedimento);
						linhaRel.cod_devolucao_motivo = reader.GetString(idx_cod_devolucao_motivo);
						linhaRel.descricao_devolucao_motivo = reader.IsDBNull(idx_descricao_devolucao_motivo) ? "" : reader.GetString(idx_descricao_devolucao_motivo);
						linhaRel.vl_devolucao = reader.GetDecimal(idx_vl_devolucao);
						linhaRel.cod_credito_transacao = reader.GetString(idx_cod_credito_transacao);
						linhaRel.loja = reader.GetString(idx_loja);
						linhaRel.data_pedido = reader.GetDateTime(idx_data_pedido);
						linhaRel.vendedor = reader.GetString(idx_vendedor);
						linhaRel.transportadora_id = reader.IsDBNull(idx_transportadora_id) ? "" : reader.GetString(idx_transportadora_id);
						linhaRel.indicador = reader.IsDBNull(idx_indicador) ? "" : reader.GetString(idx_indicador);
						linhaRel.cliente_nome = reader.GetString(idx_cliente_nome);
						linhaRel.vl_pedido = reader.GetDecimal(idx_vl_pedido);
						resultado.Add(linhaRel);
					}
				}
				finally
				{
					reader.Close();
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + ex.ToString());
				throw new Exception(ex.Message);
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return resultado;
		}
	}
}