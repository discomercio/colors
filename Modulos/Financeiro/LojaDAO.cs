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
	class LojaDAO
	{
		#region [ getLoja ]
		/// <summary>
		/// Retorna um objeto Loja contendo os dados lidos do BD
		/// </summary>
		/// <param name="numeroLoja">
		/// Identificação do registro
		/// </param>
		/// <returns>
		/// Retorna um objeto Loja contendo os dados lidos do BD
		/// </returns>
		public static Loja getLoja(String numeroLoja)
		{
			#region [ Declarações ]
			String strSql;
			Loja loja = new Loja();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroLoja == null) throw new FinanceiroException("O identificador do registro não foi fornecido!!");
			if (Global.converteInteiro(numeroLoja) == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados da loja ]
			strSql = "SELECT " +
						"*" +
					" FROM t_LOJA" +
					" WHERE" +
						" (loja = '" + numeroLoja + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			if (dtbResultado.Rows.Count == 0)
			{
				#region [ Tenta uma nova consulta assegurando que zeros à esquerda serão desprezados ]
				dtbResultado.Reset();
				strSql = "SELECT " +
							"*" +
						" FROM t_LOJA" +
						" WHERE" +
							" (CONVERT(smallint, loja) = " + numeroLoja + ")";
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion
			}

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro da loja " + numeroLoja + " não localizado na tabela t_LOJA!!");

			rowResultado = dtbResultado.Rows[0];

			loja.loja = BD.readToString(rowResultado["loja"]);
			loja.cnpj = BD.readToString(rowResultado["cnpj"]);
			loja.ie = BD.readToString(rowResultado["ie"]);
			loja.nome = BD.readToString(rowResultado["nome"]);
			loja.razao_social = BD.readToString(rowResultado["razao_social"]);
			loja.endereco = BD.readToString(rowResultado["endereco"]);
			loja.endereco_numero = BD.readToString(rowResultado["endereco_numero"]);
			loja.endereco_complemento = BD.readToString(rowResultado["endereco_complemento"]);
			loja.bairro = BD.readToString(rowResultado["bairro"]);
			loja.cidade = BD.readToString(rowResultado["cidade"]);
			loja.uf = BD.readToString(rowResultado["uf"]);
			loja.cep = BD.readToString(rowResultado["cep"]);
			loja.ddd = BD.readToString(rowResultado["ddd"]);
			loja.telefone = BD.readToString(rowResultado["telefone"]);
			loja.fax = BD.readToString(rowResultado["fax"]);
			loja.comissao_indicacao = BD.readToSingle(rowResultado["comissao_indicacao"]);
			loja.percMaxSenhaDesconto = BD.readToSingle(rowResultado["PercMaxSenhaDesconto"]);
			loja.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
			loja.id_plano_contas_grupo = BD.readToShort(rowResultado["id_plano_contas_grupo"]);
			loja.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
			loja.natureza = BD.readToChar(rowResultado["natureza"]);
			#endregion

			return loja;
		}
		#endregion
	}
}
