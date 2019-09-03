#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class NFeEmitenteDAO
	{
		#region [ getNFeEmitenteById ]
		public static NFeEmitente getNFeEmitenteById(int id)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "NFeEmitenteDAO.getNFeEmitenteById()";
			String strSql;
			NFeEmitente emitente;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id == 0)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - O identificador do registro não foi informado!!");
				return null;
			}
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do emitente ]
			strSql = "SELECT " +
						"e.*, " +
                        "n.NFe_Numero_NF, " +
                        "n.NFe_Serie_NF " +
                    " FROM t_NFe_EMITENTE e" +
                    " INNER t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
                    " WHERE" +
						" (e.id = " + id.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Registro id=" + id.ToString() + " não localizado na tabela t_NFe_EMITENTE!!");
				return null;
			}

			rowResultado = dtbResultado.Rows[0];

			emitente = new NFeEmitente();
			emitente.id = BD.readToInt(rowResultado["id"]);
			emitente.id_boleto_cedente = BD.readToInt(rowResultado["id_boleto_cedente"]);
			emitente.st_ativo = BD.readToByte(rowResultado["st_ativo"]);
			emitente.apelido = BD.readToString(rowResultado["apelido"]);
			emitente.cnpj = BD.readToString(rowResultado["cnpj"]);
			emitente.razao_social = BD.readToString(rowResultado["razao_social"]);
			emitente.endereco = BD.readToString(rowResultado["endereco"]);
			emitente.endereco_numero = BD.readToString(rowResultado["endereco_numero"]);
			emitente.endereco_complemento = BD.readToString(rowResultado["endereco_complemento"]);
			emitente.bairro = BD.readToString(rowResultado["bairro"]);
			emitente.cidade = BD.readToString(rowResultado["cidade"]);
			emitente.uf = BD.readToString(rowResultado["uf"]);
			emitente.cep = BD.readToString(rowResultado["cep"]);
			emitente.NFe_st_emitente_padrao = BD.readToByte(rowResultado["NFe_st_emitente_padrao"]);
			emitente.NFe_serie_NF = BD.readToInt(rowResultado["NFe_serie_NF"]);
			emitente.NFe_numero_NF = BD.readToInt(rowResultado["NFe_numero_NF"]);
			emitente.NFe_T1_servidor_BD = BD.readToString(rowResultado["NFe_T1_servidor_BD"]);
			emitente.NFe_T1_nome_BD = BD.readToString(rowResultado["NFe_T1_nome_BD"]);
			emitente.NFe_T1_usuario_BD = BD.readToString(rowResultado["NFe_T1_usuario_BD"]);
			emitente.NFe_T1_senha_BD = BD.readToString(rowResultado["NFe_T1_senha_BD"]);
			emitente.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
			emitente.dt_hr_cadastro = BD.readToDateTime(rowResultado["dt_hr_cadastro"]);
			emitente.usuario_cadastro = BD.readToString(rowResultado["usuario_cadastro"]);
			emitente.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
			emitente.dt_hr_ult_atualizacao = BD.readToDateTime(rowResultado["dt_hr_ult_atualizacao"]);
			emitente.usuario_ult_atualizacao = BD.readToString(rowResultado["usuario_ult_atualizacao"]);
			#endregion

			return emitente;
		}
		#endregion

		#region [ getListaNFeEmitenteByIdBoletoCedente ]
		public static List<NFeEmitente> getListaNFeEmitenteByIdBoletoCedente(int id_boleto_cedente)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<NFeEmitente> listaResultado = new List<NFeEmitente>();
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Pesquisa no BD ]
			strSql = "SELECT" +
						" id" +
					" FROM t_NFe_EMITENTE" +
					" WHERE" +
						" (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				listaResultado.Add(getNFeEmitenteById(BD.readToInt(dtbResultado.Rows[i]["id"])));
			}
			#endregion

			return listaResultado;
		}
		#endregion
	}
}
