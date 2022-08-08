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
	class CepDAO
	{
		#region [ getLocalidades ]
		/// <summary>
		/// Retorna uma lista de localidades de uma determinada UF
		/// </summary>
		/// <param name="uf">
		/// UF cujas localidades devem ser retornadas
		/// </param>
		/// <returns>
		/// Retorna uma lista de localidades de uma determinada UF
		/// </returns>
		public static List<String> getLocalidades(String uf)
		{
			#region [ Declarações ]
			String strSql;
			String strLocalidade;
			List<String> listaLocalidades = new List<String>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (uf == null) throw new FinanceiroException("UF a ser pesquisada não foi fornecida!!");
			if (uf.Length == 0) throw new FinanceiroException("UF a ser pesquisada não foi informada!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BDCep.criaSqlCommand();
			daDataAdapter = BDCep.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			uf = uf.Trim();
			#endregion

			#region [ Pesquisa localidades ]

			#region [ Monta Select ]
			strSql = "SELECT DISTINCT" +
						" LOC_NOSUB" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS localidade" +
					" FROM LOG_LOCALIDADE" +
					" WHERE" +
						" (UFE_SG = '" + uf + "')" +
					" ORDER BY" +
						" LOC_NOSUB" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT;
			#endregion

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];

				strLocalidade = !Convert.IsDBNull(rowResultado["localidade"]) ? rowResultado["localidade"].ToString() : "";
				listaLocalidades.Add(strLocalidade);
			}
			#endregion

			return listaLocalidades;
		}
		#endregion

		#region [ getCep ]
		/// <summary>
		/// Retorna uma lista de objetos Cep contendo os dados lidos do BD
		/// </summary>
		/// <param name="numeroCep">
		/// Nº do CEP com 5 ou 8 dígitos
		/// </param>
		/// <returns>
		/// Retorna uma lista de objetos Cep contendo os dados lidos do BD
		/// </returns>
		public static List<Cep> getCep(String numeroCep)
		{
			#region [ Declarações ]
			String strSql;
			String strBairro;
			String strBairroExtenso;
			String strBairroAbreviado;
			String strLogradouro;
			String strLogradouroTipo;
			String strLogradouroNome;
			String strEspaco;
			List<Cep> listaCep = new List<Cep>();
			Cep cep;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroCep == null) throw new FinanceiroException("O CEP a ser pesquisado não foi fornecido!!");
			if (numeroCep.Length == 0) throw new FinanceiroException("O CEP a ser pesquisado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BDCep.criaSqlCommand();
			daDataAdapter = BDCep.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			numeroCep = Global.digitos(numeroCep);
			#endregion

			#region [ Pesquisa CEP ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" 'LOGRADOURO' AS tabela_origem," +
						" Logr.CEP_DIG AS cep," +
						" Logr.UFE_SG AS uf," +
						" Loc.LOC_NOSUB AS localidade," +
						" Bai.BAI_NO AS bairro_extenso," +
						" Bai.BAI_NO_ABREV AS bairro_abreviado," +
						" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," +
						" Logr.LOG_NO AS logradouro_nome," +
						" Logr.LOG_COMPLEMENTO AS logradouro_complemento" +
					" FROM LOG_LOGRADOURO Logr" +
						" LEFT JOIN LOG_BAIRRO Bai ON (Logr.BAI_NU_SEQUENCIAL_INI = Bai.BAI_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_LOCALIDADE Loc ON (Logr.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" +
					" WHERE";

			if (numeroCep.Length == 5)
				strSql += " (Logr.CEP_DIG LIKE '" + numeroCep + BDCep.CARACTER_CURINGA_TODOS + "')";
			else
				strSql += " (Logr.CEP_DIG = '" + numeroCep + "')";

			strSql += " UNION " +
					"SELECT" +
						" 'GRANDE_USUARIO' AS tabela_origem," +
						" GU.CEP_DIG AS cep," +
						" GU.UFE_SG AS uf," +
						" Loc.LOC_NOSUB AS localidade," +
						" Bai.BAI_NO AS bairro_extenso," +
						" Bai.BAI_NO_ABREV AS bairro_abreviado," +
						" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," +
						" Logr.LOG_NO AS logradouro_nome," +
						" GU.GRU_NO AS logradouro_complemento" +
					" FROM LOG_GRANDE_USUARIO GU" +
						" LEFT JOIN LOG_LOGRADOURO Logr ON (GU.LOG_NU_SEQUENCIAL = Logr.LOG_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_BAIRRO Bai ON (GU.BAI_NU_SEQUENCIAL = Bai.BAI_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_LOCALIDADE Loc ON (GU.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" +
					" WHERE";

			if (numeroCep.Length == 5)
				strSql += " (GU.CEP_DIG LIKE '" + numeroCep + BDCep.CARACTER_CURINGA_TODOS + "')";
			else
				strSql += " (GU.CEP_DIG = '" + numeroCep + "')";

			strSql += " UNION " +
					"SELECT" +
						" 'LOCALIDADE' AS tabela_origem," +
						" CEP_DIG AS cep," +
						" UFE_SG AS uf," +
						" LOC_NOSUB AS localidade," +
						" '' AS bairro_extenso," +
						" '' AS bairro_abreviado," +
						" '' AS logradouro_tipo," +
						" '' AS logradouro_nome," +
						" '' AS logradouro_complemento" +
					" FROM LOG_LOCALIDADE" +
					" WHERE";

			if (numeroCep.Length == 5)
				strSql += " (CEP_DIG LIKE '" + numeroCep + BDCep.CARACTER_CURINGA_TODOS + "')";
			else
				strSql += " (CEP_DIG = '" + numeroCep + "')";

			// CONSULTA DADOS DA TABELA ANTIGA, POIS ELA É MANTIDA P/ MANTER FUNCIONANDO O CADASTRAMENTO MANUAL DE CEP'S
			strSql += " UNION " +
					"SELECT" +
						" 'LOGRADOURO' AS tabela_origem," +
						" cep8_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS cep," +
						" uf_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS uf," +
						" nome_local" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS localidade," +
						" extenso_bai" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS bairro_extenso," +
						" abrev_bai" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS bairro_abreviado," +
						" abrev_tipo" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_tipo," +
						" nome_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_nome," +
						" comple_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_complemento" +
					" FROM t_CEP_LOGRADOURO" +
					" WHERE";

			if (numeroCep.Length == 5)
				strSql += " (cep8_log LIKE '" + numeroCep + BDCep.CARACTER_CURINGA_TODOS + "')";
			else
				strSql += " (cep8_log = '" + numeroCep + "')";

			strSql += " ORDER BY cep";
			#endregion

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];

				strLogradouroTipo = !Convert.IsDBNull(rowResultado["logradouro_tipo"]) ? rowResultado["logradouro_tipo"].ToString() : "";
				strLogradouroNome = !Convert.IsDBNull(rowResultado["logradouro_nome"]) ? rowResultado["logradouro_nome"].ToString() : "";
				if ((strLogradouroTipo.Length > 0) && (strLogradouroNome.Length > 0))
					strEspaco = " ";
				else
					strEspaco = "";
				strLogradouro = strLogradouroTipo + strEspaco + strLogradouroNome;

				strBairroExtenso = !Convert.IsDBNull(rowResultado["bairro_extenso"]) ? rowResultado["bairro_extenso"].ToString() : "";
				strBairroAbreviado = !Convert.IsDBNull(rowResultado["bairro_abreviado"]) ? rowResultado["bairro_abreviado"].ToString() : "";
				strBairro = strBairroExtenso;
				if (strBairro.Length == 0) strBairro = strBairroAbreviado;

				cep = new Cep();
				cep.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
				cep.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
				cep.cidade = !Convert.IsDBNull(rowResultado["localidade"]) ? rowResultado["localidade"].ToString() : "";
				cep.bairro = strBairro;
				cep.logradouro = strLogradouro;
				cep.complemento = !Convert.IsDBNull(rowResultado["logradouro_complemento"]) ? rowResultado["logradouro_complemento"].ToString() : "";
				
				listaCep.Add(cep);
			}
			#endregion

			return listaCep;
		}
		#endregion

		#region [ getCep ]
		/// <summary>
		/// Retorna uma lista de objetos Cep contendo os dados lidos do BD
		/// </summary>
		/// <param name="uf">
		/// UF a ser pesquisada
		/// </param>
		/// <param name="localidade">
		/// Localidade a ser pesquisada
		/// </param>
		/// <param name="endereco">
		/// Parâmetro opcional contendo parte do endereço a ser pesquisado
		/// </param>
		/// <returns>
		/// Retorna uma lista de objetos Cep contendo os dados lidos do BD
		/// </returns>
		public static List<Cep> getCep(String uf, String localidade, String endereco)
		{
			#region [ Declarações ]
			String strSql;
			String strBairro;
			String strBairroExtenso;
			String strBairroAbreviado;
			String strLogradouro;
			String strLogradouroTipo;
			String strLogradouroNome;
			String strEspaco;
			List<Cep> listaCep = new List<Cep>();
			Cep cep;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (uf == null) throw new FinanceiroException("UF a ser pesquisada não foi fornecida!!");
			if (uf.Length == 0) throw new FinanceiroException("UF a ser pesquisada não foi informada!!");
			if (localidade == null) throw new FinanceiroException("Localidade a ser pesquisada não foi fornecida!!");
			if (localidade.Length == 0) throw new FinanceiroException("Localidade a ser pesquisada não foi informada!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BDCep.criaSqlCommand();
			daDataAdapter = BDCep.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			uf = uf.Trim();
			localidade = localidade.Trim();
			#endregion

			#region [ Pesquisa CEP ]

			#region [ Monta Select ]
			strSql = "SELECT TOP 500" +
						" 'LOGRADOURO' AS tabela_origem," +
						" Logr.CEP_DIG AS cep," +
						" Logr.UFE_SG AS uf," +
						" Loc.LOC_NOSUB AS localidade," +
						" Bai.BAI_NO AS bairro_extenso," +
						" Bai.BAI_NO_ABREV AS bairro_abreviado," +
						" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," +
						" Logr.LOG_NO AS logradouro_nome," +
						" Logr.LOG_COMPLEMENTO AS logradouro_complemento" +
					" FROM LOG_LOGRADOURO Logr" +
						" LEFT JOIN LOG_BAIRRO Bai ON (Logr.BAI_NU_SEQUENCIAL_INI = Bai.BAI_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_LOCALIDADE Loc ON (Logr.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" +
					" WHERE" +
						"(Logr.UFE_SG = '" + uf + "')" +
						" AND " +
						"(Loc.LOC_NOSUB = '" + localidade + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";

			if (endereco.Length > 0)
			{
				strSql += " AND (Logr.LOG_NO LIKE '" + BDCep.CARACTER_CURINGA_TODOS + endereco + BDCep.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
			}

			strSql += " UNION " +
					"SELECT TOP 500" +
						" 'GRANDE_USUARIO' AS tabela_origem," +
						" GU.CEP_DIG AS cep," +
						" GU.UFE_SG AS uf," +
						" Loc.LOC_NOSUB AS localidade," +
						" Bai.BAI_NO AS bairro_extenso," +
						" Bai.BAI_NO_ABREV AS bairro_abreviado," +
						" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," +
						" Logr.LOG_NO AS logradouro_nome," +
						" GU.GRU_NO AS logradouro_complemento" +
					" FROM LOG_GRANDE_USUARIO GU" +
						" LEFT JOIN LOG_LOGRADOURO Logr ON (GU.LOG_NU_SEQUENCIAL = Logr.LOG_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_BAIRRO Bai ON (GU.BAI_NU_SEQUENCIAL = Bai.BAI_NU_SEQUENCIAL)" +
						" LEFT JOIN LOG_LOCALIDADE Loc ON (GU.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" +
					" WHERE" +
						"(GU.UFE_SG = '" + uf + "')" +
						" AND " +
						"(Loc.LOC_NOSUB = '" + localidade + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";

			if (endereco.Length > 0)
			{
				strSql += " AND (Logr.LOG_NO LIKE '" + BDCep.CARACTER_CURINGA_TODOS + endereco + BDCep.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
			}

			strSql += " UNION " +
					"SELECT" +
						" 'LOCALIDADE' AS tabela_origem," +
						" CEP_DIG AS cep," +
						" UFE_SG AS uf," +
						" LOC_NOSUB AS localidade," +
						" '' AS bairro_extenso," +
						" '' AS bairro_abreviado," +
						" '' AS logradouro_tipo," +
						" '' AS logradouro_nome," +
						" '' AS logradouro_complemento" +
					" FROM LOG_LOCALIDADE" +
					" WHERE" +
						" (UFE_SG = '" + uf + "')" +
						" AND (LOC_NOSUB = '" + localidade + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")" +
						" AND (LEN(Coalesce(CEP_DIG,'')) > 0)";

			// CONSULTA DADOS DA TABELA ANTIGA, POIS ELA É MANTIDA P/ MANTER FUNCIONANDO O CADASTRAMENTO MANUAL DE CEP'S
			strSql += " UNION " +
					"SELECT TOP 500" +
						"'LOGRADOURO' AS tabela_origem," +
						" cep8_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS cep," +
						" uf_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS uf," +
						" nome_local" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS localidade," +
						" extenso_bai" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS bairro_extenso," +
						" abrev_bai" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS bairro_abreviado," +
						" abrev_tipo" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_tipo," +
						" nome_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_nome," +
						" comple_log" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + " AS logradouro_complemento" +
					" FROM t_CEP_LOGRADOURO" +
					" WHERE" +
						" (uf_log = '" + uf + "')" +
						" AND (nome_local = '" + localidade + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";

			if (endereco.Length > 0)
			{
				strSql += " AND (nome_log LIKE '" + BDCep.CARACTER_CURINGA_TODOS + endereco + BDCep.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
			}

			strSql += " ORDER BY" +
						" uf," +
						" localidade," +
						" bairro_extenso," +
						" cep," +
						" logradouro_nome";
			#endregion

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];

				strLogradouroTipo = !Convert.IsDBNull(rowResultado["logradouro_tipo"]) ? rowResultado["logradouro_tipo"].ToString() : "";
				strLogradouroNome = !Convert.IsDBNull(rowResultado["logradouro_nome"]) ? rowResultado["logradouro_nome"].ToString() : "";
				if ((strLogradouroTipo.Length > 0) && (strLogradouroNome.Length > 0))
					strEspaco = " ";
				else
					strEspaco = "";
				strLogradouro = strLogradouroTipo + strEspaco + strLogradouroNome;

				strBairroExtenso = !Convert.IsDBNull(rowResultado["bairro_extenso"]) ? rowResultado["bairro_extenso"].ToString() : "";
				strBairroAbreviado = !Convert.IsDBNull(rowResultado["bairro_abreviado"]) ? rowResultado["bairro_abreviado"].ToString() : "";
				strBairro = strBairroExtenso;
				if (strBairro.Length == 0) strBairro = strBairroAbreviado;

				cep = new Cep();
				cep.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
				cep.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
				cep.cidade = !Convert.IsDBNull(rowResultado["localidade"]) ? rowResultado["localidade"].ToString() : "";
				cep.bairro = strBairro;
				cep.logradouro = strLogradouro;
				cep.complemento = !Convert.IsDBNull(rowResultado["logradouro_complemento"]) ? rowResultado["logradouro_complemento"].ToString() : "";

				listaCep.Add(cep);
			}
			#endregion

			return listaCep;
		}
		#endregion
	}
}
