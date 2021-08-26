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

			return emitente;
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