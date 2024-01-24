using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Domains;
using System.Threading;

namespace ART3WebAPI.Models.Repository
{
	public class NFeEmitenteDAO
	{
		#region [ NFeEmitenteLoadFromDataRow ]
		public static NFeEmitente NFeEmitenteLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			NFeEmitente emitente = new NFeEmitente();
			#endregion

			emitente.id = BD.readToInt(rowDados["id"]);
			emitente.id_boleto_cedente = BD.readToInt(rowDados["id_boleto_cedente"]);
			emitente.st_ativo = BD.readToByte(rowDados["st_ativo"]);
			emitente.apelido = BD.readToString(rowDados["apelido"]);
			emitente.cnpj = BD.readToString(rowDados["cnpj"]);
			emitente.razao_social = BD.readToString(rowDados["razao_social"]);
			emitente.endereco = BD.readToString(rowDados["endereco"]);
			emitente.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			emitente.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			emitente.bairro = BD.readToString(rowDados["bairro"]);
			emitente.cidade = BD.readToString(rowDados["cidade"]);
			emitente.uf = BD.readToString(rowDados["uf"]);
			emitente.cep = BD.readToString(rowDados["cep"]);
			emitente.NFe_st_emitente_padrao = BD.readToByte(rowDados["NFe_st_emitente_padrao"]);
			emitente.NFe_serie_NF = BD.readToInt(rowDados["NFe_serie_NF"]);
			emitente.NFe_numero_NF = BD.readToInt(rowDados["NFe_numero_NF"]);
			emitente.NFe_T1_servidor_BD = BD.readToString(rowDados["NFe_T1_servidor_BD"]);
			emitente.NFe_T1_nome_BD = BD.readToString(rowDados["NFe_T1_nome_BD"]);
			emitente.NFe_T1_usuario_BD = BD.readToString(rowDados["NFe_T1_usuario_BD"]);
			emitente.NFe_T1_senha_BD = BD.readToString(rowDados["NFe_T1_senha_BD"]);
			emitente.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			emitente.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			emitente.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			emitente.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			emitente.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			emitente.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);
			emitente.st_habilitado_ctrl_estoque = BD.readToByte(rowDados["st_habilitado_ctrl_estoque"]);
			emitente.ordem = BD.readToInt(rowDados["ordem"]);
			emitente.texto_fixo_especifico = BD.readToString(rowDados["texto_fixo_especifico"]);

			return emitente;
		}
		#endregion

		#region [ NFeEmitenteCfgDanfeLoadFromDataRow ]
		public static NFeEmitenteCfgDanfe NFeEmitenteCfgDanfeLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			NFeEmitenteCfgDanfe emitenteCfgDanfe = new NFeEmitenteCfgDanfe();
			#endregion

			emitenteCfgDanfe.id = BD.readToInt(rowDados["id"]);
			emitenteCfgDanfe.id_nfe_emitente = BD.readToInt(rowDados["id_nfe_emitente"]);
			emitenteCfgDanfe.min_tamanho_serie_NFe = BD.readToByte(rowDados["min_tamanho_serie_NFe"]);
			emitenteCfgDanfe.min_tamanho_numero_NFe = BD.readToByte(rowDados["min_tamanho_numero_NFe"]);
			emitenteCfgDanfe.convencao_nome_arq_pdf_danfe = BD.readToString(rowDados["convencao_nome_arq_pdf_danfe"]);
			emitenteCfgDanfe.diretorio_pdf_danfe = BD.readToString(rowDados["diretorio_pdf_danfe"]);
			emitenteCfgDanfe.convencao_nome_arq_xml_nfe = BD.readToString(rowDados["convencao_nome_arq_xml_nfe"]);
			emitenteCfgDanfe.diretorio_xml_nfe = BD.readToString(rowDados["diretorio_xml_nfe"]);
			emitenteCfgDanfe.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			emitenteCfgDanfe.ordenacao = BD.readToInt(rowDados["ordenacao"]);
			emitenteCfgDanfe.observacao = BD.readToString(rowDados["observacao"]);

			return emitenteCfgDanfe;
		}
		#endregion

		#region [ getAllNFeEmitente ]
		public static List<NFeEmitente> getAllNFeEmitente()
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<NFeEmitente> listaEmitente = new List<NFeEmitente>();
			NFeEmitente emitente;
			NFeEmitenteCfgDanfe emitenteCfgDanfe;
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Dados principais dos emitentes ]

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"e.*, " +
							"n.NFe_Numero_NF, " +
							"n.NFe_Serie_NF " +
						" FROM t_NFe_EMITENTE e" +
						" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
						" ORDER BY" +
							" e.id";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return listaEmitente;

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					emitente = NFeEmitenteLoadFromDataRow(dtbResultado.Rows[i]);
					listaEmitente.Add(emitente);
				}
				#endregion

				#region [ Dados sobre os diretórios de armazenamento da DANFE/XML ]
				foreach (NFeEmitente emitenteAux in listaEmitente)
				{
					strSql = "SELECT " +
								"*" +
							" FROM t_NFe_EMITENTE_CFG_DANFE" +
							" WHERE" +
								" (id_nfe_emitente = " + emitenteAux.id.ToString() + ")" +
							" ORDER BY" +
								" ordenacao";

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					dtbResultado.Reset();
					daDataAdapter.Fill(dtbResultado);
					#endregion

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						emitenteCfgDanfe = NFeEmitenteCfgDanfeLoadFromDataRow(dtbResultado.Rows[i]);
						emitenteAux.listaCfgDanfe.Add(emitenteCfgDanfe);
					}
				}
				#endregion

				return listaEmitente;
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}
		}
		#endregion

		#region [ getNFeEmitenteById ]
		public static NFeEmitente getNFeEmitenteById(int id)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			NFeEmitente emitente;
			NFeEmitenteCfgDanfe emitenteCfgDanfe;
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Dados principais do emitente ]

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"e.*, " +
							"n.NFe_Numero_NF, " +
							"n.NFe_Serie_NF " +
						" FROM t_NFe_EMITENTE e" +
						" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
						" WHERE" +
							" (e.id = " + id.ToString() + ")";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				emitente = NFeEmitenteLoadFromDataRow(dtbResultado.Rows[0]);
				#endregion

				#region [ Dados sobre os diretórios de armazenamento da DANFE/XML ]
				strSql = "SELECT " +
							"*" +
						" FROM t_NFe_EMITENTE_CFG_DANFE" +
						" WHERE" +
							" (id_nfe_emitente = " + id.ToString() + ")" +
						" ORDER BY" +
							" ordenacao";

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					emitenteCfgDanfe = NFeEmitenteCfgDanfeLoadFromDataRow(dtbResultado.Rows[i]);
					emitente.listaCfgDanfe.Add(emitenteCfgDanfe);
				}
				#endregion

				return emitente;
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}
		}
		#endregion
	}
}