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
	class ComboDAO
	{
		#region [ Enum ]

		#region [ enum: eFiltraStAtivo ]
		public enum eFiltraStAtivo : byte
		{
			TODOS = 0,
			SOMENTE_ATIVOS = 1,
			SOMENTE_INATIVOS = 2
		}
		#endregion

		#region [ enum: eFiltraStSistema ]
		public enum eFiltraStSistema : byte
		{
			TODOS = 0,
			SOMENTE_CONTAS_SISTEMA = 1,
			SOMENTE_CONTAS_NORMAIS = 2
		}
		#endregion

		#region [ enum: eFiltraNatureza ]
		public enum eFiltraNatureza : byte
		{
			TODOS = 0,
			SOMENTE_CREDITO = 1,
			SOMENTE_DEBITO = 2
		}
		#endregion

		#endregion

		#region [ Atributos Estáticos ]
		private static DsDataSource.DtbEquipeVendasComboDataTable _dtbEquipeVendasComboBuffer;
		private static DsDataSource.DtbVendedorComboDataTable _dtbVendedorComboBuffer;
		private static DsDataSource.DtbIndicadorComboDataTable _dtbIndicadorComboBuffer;
		private static DsDataSource.DtbContaCorrenteComboDataTable _dtbContaCorrenteComboBuffer;
		private static DsDataSource.DtbPlanoContasEmpresaComboDataTable _dtbPlanoContasEmpresaComboBuffer;
		private static DsDataSource.DtbPlanoContasGrupoComboDataTable _dtbPlanoContasGrupoComboBuffer;
		private static DsDataSource.DtbPlanoContasContaComboDataTable _dtbPlanoContasContaComboBuffer;
		private static DsDataSource.DtbBoletoCedenteComboDataTable _dtbBoletoCedenteComboBuffer;
		private static DsDataSource.DtbUnidadeNegocioComboDataTable _dtbUnidadeNegocioComboBuffer;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor estático ]
		static ComboDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			_dtbEquipeVendasComboBuffer = criaDtbEquipeVendasCombo();
			_dtbVendedorComboBuffer = criaDtbVendedor();
			_dtbIndicadorComboBuffer = criaDtbIndicador();

			#region [ Conta Corrente ]
			_dtbContaCorrenteComboBuffer = criaDtbContaCorrenteCombo(eFiltraStAtivo.TODOS);
			_dtbContaCorrenteComboBuffer.TableName = "_dtbContaCorrenteComboBuffer";
			#endregion

			#region [ Plano Contas Empresa ]
			_dtbPlanoContasEmpresaComboBuffer = criaDtbPlanoContasEmpresaCombo(eFiltraStAtivo.TODOS);
			_dtbPlanoContasEmpresaComboBuffer.TableName = "_dtbPlanoContasEmpresaComboBuffer";
			#endregion

			#region [ Plano Contas Grupo ]
			_dtbPlanoContasGrupoComboBuffer = criaDtbPlanoContasGrupoCombo(eFiltraStAtivo.TODOS);
			_dtbPlanoContasGrupoComboBuffer.TableName = "_dtbPlanoContasGrupoComboBuffer";
			#endregion

			#region [ Plano Contas Conta ]
			_dtbPlanoContasContaComboBuffer = criaDtbPlanoContasContaCombo(eFiltraNatureza.TODOS, eFiltraStAtivo.TODOS, eFiltraStSistema.TODOS);
			_dtbPlanoContasContaComboBuffer.TableName = "_dtbPlanoContasContaComboBuffer";
			#endregion

			#region [ Boleto Cedente ]
			_dtbBoletoCedenteComboBuffer = criaDtbBoletoCedenteCombo(eFiltraStAtivo.TODOS);
			_dtbBoletoCedenteComboBuffer.TableName = "_dtbBoletoCedenteComboBuffer";
			#endregion

			#region [ Unidade de Negócio ]
			_dtbUnidadeNegocioComboBuffer = criaDtbUnidadeNegocioCombo(eFiltraStAtivo.TODOS);
			_dtbUnidadeNegocioComboBuffer.TableName = "_dtbUnidadeNegocioComboBuffer";
			#endregion
		}
		#endregion

		#region [ criaDtbContaCorrenteCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbContaCorrenteComboDataTable criaDtbContaCorrenteCombo(eFiltraStAtivo stAtivo)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbContaCorrenteComboDataTable dtbContaCorrente;
			DataView dvContaCorrente;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbContaCorrenteComboBuffer != null)
			{
				dvContaCorrente = new DataView();
				dvContaCorrente.Table = _dtbContaCorrenteComboBuffer;
				dvContaCorrente.Sort = strOrderBy;
				if (strWhere.Length > 0) dvContaCorrente.RowFilter = strWhere;
				dtbContaCorrente = new DsDataSource.DtbContaCorrenteComboDataTable();
				dtbContaCorrente.Merge(dvContaCorrente.ToTable());
				return dtbContaCorrente;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbContaCorrente = new DsDataSource.DtbContaCorrenteComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" conta," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.CONTA_CORRENTE_ID) + " + ' - ' + conta + '  ' + descricao AS idContaDescricao," +
						" conta + ' - ' + descricao AS contaComDescricao" +
					" FROM t_FIN_CONTA_CORRENTE" +
					strWhere +
					" ORDER BY" +
						strOrderBy;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbContaCorrente);
			return dtbContaCorrente;
		}
		#endregion

		#region [ criaDtbContaCorrenteCombo (opção para incluir item BRANCO/TODOS) ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbContaCorrenteComboDataTable criaDtbContaCorrenteCombo(eFiltraStAtivo stAtivo, Global.eOpcaoIncluirItemTodos opcaoIncluir)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbContaCorrenteComboDataTable dtbContaCorrente;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbContaCorrente = new DsDataSource.DtbContaCorrenteComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" conta," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.CONTA_CORRENTE_ID) + " + ' - ' + conta + '  ' + descricao AS idContaDescricao," +
						" conta + ' - ' + descricao AS contaComDescricao" +
					" FROM t_FIN_CONTA_CORRENTE" +
					strWhere;

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" '' AS conta," +
						" '' AS descricao," +
						" '' AS idContaDescricao," +
						" '' AS contaComDescricao";
			}

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" '' AS conta," +
						" 'TODAS' AS descricao," +
						" 'TODAS' AS idContaDescricao," +
						" 'TODAS' AS contaComDescricao";
			}

			strSql +=
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbContaCorrente);
			return dtbContaCorrente;
		}
		#endregion

		#region [ criaDtbPlanoContasEmpresaCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasEmpresaComboDataTable criaDtbPlanoContasEmpresaCombo(eFiltraStAtivo stAtivo)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasEmpresaComboDataTable dtbPlanoContasEmpresa;
			DataView dvPlanoContasEmpresa;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbPlanoContasEmpresaComboBuffer != null)
			{
				dvPlanoContasEmpresa = new DataView();
				dvPlanoContasEmpresa.Table = _dtbPlanoContasEmpresaComboBuffer;
				dvPlanoContasEmpresa.Sort = strOrderBy;
				if (strWhere.Length > 0) dvPlanoContasEmpresa.RowFilter = strWhere;
				dtbPlanoContasEmpresa = new DsDataSource.DtbPlanoContasEmpresaComboDataTable();
				dtbPlanoContasEmpresa.Merge(dvPlanoContasEmpresa.ToTable());
				return dtbPlanoContasEmpresa;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasEmpresa = new DsDataSource.DtbPlanoContasEmpresaComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_EMPRESA) + " + ' - ' + descricao AS idComDescricao" +
					" FROM t_FIN_PLANO_CONTAS_EMPRESA" +
					strWhere +
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasEmpresa);
			return dtbPlanoContasEmpresa;
		}
		#endregion

		#region [ criaDtbPlanoContasEmpresaCombo (opção para incluir item BRANCO/TODOS) ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasEmpresaComboDataTable criaDtbPlanoContasEmpresaCombo(eFiltraStAtivo stAtivo, Global.eOpcaoIncluirItemTodos opcaoIncluir)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasEmpresaComboDataTable dtbPlanoContasEmpresa;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasEmpresa = new DsDataSource.DtbPlanoContasEmpresaComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_EMPRESA) + " + ' - ' + descricao AS idComDescricao" +
					" FROM t_FIN_PLANO_CONTAS_EMPRESA" +
					strWhere;

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" '' AS descricao," +
						" '' AS idComDescricao";
			}

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" 'TODAS' AS descricao," +
						" 'TODAS' AS idComDescricao";
			}

			strSql +=
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasEmpresa);
			return dtbPlanoContasEmpresa;
		}
		#endregion

		#region [ criaDtbPlanoContasGrupoCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasGrupoComboDataTable criaDtbPlanoContasGrupoCombo(eFiltraStAtivo stAtivo)
		{
			#region [ Declarações ]
			String strTamanho;
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasGrupoComboDataTable dtbPlanoContasGrupo;
			DataView dvPlanoContasGrupo;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbPlanoContasGrupoComboBuffer != null)
			{
				dvPlanoContasGrupo = new DataView();
				dvPlanoContasGrupo.Table = _dtbPlanoContasGrupoComboBuffer;
				dvPlanoContasGrupo.Sort = strOrderBy;
				if (strWhere.Length > 0) dvPlanoContasGrupo.RowFilter = strWhere;
				dtbPlanoContasGrupo = new DsDataSource.DtbPlanoContasGrupoComboDataTable();
				dtbPlanoContasGrupo.Merge(dvPlanoContasGrupo.ToTable());
				return dtbPlanoContasGrupo;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasGrupo = new DsDataSource.DtbPlanoContasGrupoComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strTamanho = Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO.ToString();
			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO) + " + ' - ' + descricao AS idComDescricao" +
					" FROM t_FIN_PLANO_CONTAS_GRUPO" +
					strWhere +
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasGrupo);
			return dtbPlanoContasGrupo;
		}
		#endregion

		#region [ criaDtbPlanoContasGrupoCombo (opção para incluir item BRANCO/TODOS) ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasGrupoComboDataTable criaDtbPlanoContasGrupoCombo(eFiltraStAtivo stAtivo, Global.eOpcaoIncluirItemTodos opcaoIncluir)
		{
			#region [ Declarações ]
			String strTamanho;
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasGrupoComboDataTable dtbPlanoContasGrupo;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasGrupo = new DsDataSource.DtbPlanoContasGrupoComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strTamanho = Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO.ToString();
			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" descricao," +
						Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO) + " + ' - ' + descricao AS idComDescricao" +
					" FROM t_FIN_PLANO_CONTAS_GRUPO" +
					strWhere;

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" '' AS descricao," +
						" '' AS idComDescricao";
			}

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" 0 AS st_ativo," +
						" 'TODAS' AS descricao," +
						" 'TODAS' AS idComDescricao";
			}

			strSql +=
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasGrupo);
			return dtbPlanoContasGrupo;
		}
		#endregion

		#region [ criaDtbPlanoContasContaCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que forem da natureza especificada e estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasContaComboDataTable criaDtbPlanoContasContaCombo(eFiltraNatureza opcaoNatureza, eFiltraStAtivo stAtivo, eFiltraStSistema stSistema)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			String strCampoIdComDescricao;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasContaComboDataTable dtbPlanoContasConta;
			DataView dvPlanoContasConta;
			#endregion

			#region [ Monta restrições ]

			#region [ Natureza ]
			if (opcaoNatureza == eFiltraNatureza.SOMENTE_CREDITO)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (natureza = '" + Global.Cte.FIN.Natureza.CREDITO + "')";
			}
			else if (opcaoNatureza == eFiltraNatureza.SOMENTE_DEBITO)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (natureza = '" + Global.Cte.FIN.Natureza.DEBITO + "')";
			}
			#endregion

			#region [ Status Ativo ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			}
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			}
			#endregion

			#region [ Status Sistema ]
			if (stSistema == eFiltraStSistema.SOMENTE_CONTAS_NORMAIS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_sistema = " + Global.Cte.FIN.StSistema.NAO.ToString() + ")";
			}
			else if (stSistema == eFiltraStSistema.SOMENTE_CONTAS_SISTEMA)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_sistema = " + Global.Cte.FIN.StSistema.SIM.ToString() + ")";
			}
			#endregion

			#endregion

			#region [ Order by ]
			strOrderBy = " id," +
						 " natureza";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbPlanoContasContaComboBuffer != null)
			{
				dvPlanoContasConta = new DataView();
				dvPlanoContasConta.Table = _dtbPlanoContasContaComboBuffer;
				dvPlanoContasConta.Sort = strOrderBy;
				if (strWhere.Length > 0) dvPlanoContasConta.RowFilter = strWhere;
				dtbPlanoContasConta = new DsDataSource.DtbPlanoContasContaComboDataTable();
				dtbPlanoContasConta.Merge(dvPlanoContasConta.ToTable());
				return dtbPlanoContasConta;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasConta = new DsDataSource.DtbPlanoContasContaComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			if (opcaoNatureza == eFiltraNatureza.TODOS)
				strCampoIdComDescricao = Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA) + " + ' (' + natureza + ') - ' + descricao AS idComDescricao";
			else
				strCampoIdComDescricao = Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA) + " + ' - ' + descricao AS idComDescricao";

			strSql = "SELECT" +
						" id," +
						" id_plano_contas_grupo," +
						" natureza," +
						" st_sistema," +
						" st_ativo," +
						" descricao," +
						strCampoIdComDescricao +
					" FROM t_FIN_PLANO_CONTAS_CONTA" +
					strWhere +
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasConta);
			return dtbPlanoContasConta;
		}
		#endregion

		#region [ criaDtbPlanoContasContaCombo (opção para incluir item BRANCO/TODOS) ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que forem da natureza especificada e estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbPlanoContasContaComboDataTable criaDtbPlanoContasContaCombo(eFiltraNatureza opcaoNatureza, eFiltraStAtivo stAtivo, eFiltraStSistema stSistema, Global.eOpcaoIncluirItemTodos opcaoIncluir)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			String strCampoIdComDescricao;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbPlanoContasContaComboDataTable dtbPlanoContasConta;
			#endregion

			#region [ Monta restrições ]

			#region [ Natureza ]
			if (opcaoNatureza == eFiltraNatureza.SOMENTE_CREDITO)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (natureza = '" + Global.Cte.FIN.Natureza.CREDITO + "')";
			}
			else if (opcaoNatureza == eFiltraNatureza.SOMENTE_DEBITO)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (natureza = '" + Global.Cte.FIN.Natureza.DEBITO + "')";
			}
			#endregion

			#region [ Status Ativo ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			}
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			}
			#endregion

			#region [ Status Sistema ]
			if (stSistema == eFiltraStSistema.SOMENTE_CONTAS_NORMAIS)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_sistema = " + Global.Cte.FIN.StSistema.NAO.ToString() + ")";
			}
			else if (stSistema == eFiltraStSistema.SOMENTE_CONTAS_SISTEMA)
			{
				if (strWhere.Length > 0) strWhere += " AND";
				strWhere += " (st_sistema = " + Global.Cte.FIN.StSistema.SIM.ToString() + ")";
			}
			#endregion

			#endregion

			#region [ Order by ]
			strOrderBy = " id," +
						 " natureza";
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbPlanoContasConta = new DsDataSource.DtbPlanoContasContaComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			if (opcaoNatureza == eFiltraNatureza.TODOS)
				strCampoIdComDescricao = Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA) + " + ' (' + natureza + ') - ' + descricao AS idComDescricao";
			else
				strCampoIdComDescricao = Global.sqlMontaPadLeftCampoNumerico("id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA) + " + ' - ' + descricao AS idComDescricao";

			strSql = "SELECT" +
						" id," +
						" id_plano_contas_grupo," +
						" natureza," +
						" st_sistema," +
						" st_ativo," +
						" descricao," +
						strCampoIdComDescricao +
					" FROM t_FIN_PLANO_CONTAS_CONTA" +
					strWhere;

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" NULL AS id_plano_contas_grupo," +
						" ' ' AS natureza," +
						" 0 AS st_sistema," +
						" 0 AS st_ativo," +
						" '' AS descricao," +
						" '' AS idComDescricao";
			}

			if (opcaoIncluir == Global.eOpcaoIncluirItemTodos.INCLUIR)
			{
				strSql +=
					" UNION ALL " +
					"SELECT" +
						" NULL AS id," +
						" NULL AS id_plano_contas_grupo," +
						" ' ' AS natureza," +
						" 0 AS st_sistema," +
						" 0 AS st_ativo," +
						" 'TODAS' AS descricao," +
						" 'TODAS' AS idComDescricao";
			}

			strSql +=
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbPlanoContasConta);
			return dtbPlanoContasConta;
		}
		#endregion

		#region [ criaDtbBoletoCedenteCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbBoletoCedenteComboDataTable criaDtbBoletoCedenteCombo(eFiltraStAtivo stAtivo)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbBoletoCedenteComboDataTable dtbBoletoCedente;
			DataView dvBoletoCedente;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " id";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbBoletoCedenteComboBuffer != null)
			{
				dvBoletoCedente = new DataView();
				dvBoletoCedente.Table = _dtbBoletoCedenteComboBuffer;
				dvBoletoCedente.Sort = strOrderBy;
				if (strWhere.Length > 0) dvBoletoCedente.RowFilter = strWhere;
				dtbBoletoCedente = new DsDataSource.DtbBoletoCedenteComboDataTable();
				dtbBoletoCedente.Merge(dvBoletoCedente.ToTable());
				return dtbBoletoCedente;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbBoletoCedente = new DsDataSource.DtbBoletoCedenteComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" 'Bco:' + num_banco + '  Ag:' + agencia + '  Cta:' + conta + '      ' + nome_empresa AS descricao_formatada" +
					" FROM t_FIN_BOLETO_CEDENTE" +
					strWhere +
					" ORDER BY" +
						strOrderBy;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbBoletoCedente);
			return dtbBoletoCedente;
		}
		#endregion

		#region [ criaDtbEquipeVendasCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbEquipeVendasComboDataTable criaDtbEquipeVendasCombo()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbEquipeVendasComboDataTable dtbEquipeVendas;
			#endregion

			#region [ Há dados no "buffer" local? ]
			if (_dtbEquipeVendasComboBuffer != null) return (DsDataSource.DtbEquipeVendasComboDataTable)_dtbEquipeVendasComboBuffer.Copy();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbEquipeVendas = new DsDataSource.DtbEquipeVendasComboDataTable();

			strSql = "SELECT" +
						" id," +
						" apelido," +
						" descricao," +
						" apelido + ' - ' + descricao AS apelidoComDescricao" +
					" FROM t_EQUIPE_VENDAS" +
					" ORDER BY" +
						" apelido";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbEquipeVendas);
			return dtbEquipeVendas;
		}
		#endregion

		#region [ criaDtbVendedor ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbVendedorComboDataTable criaDtbVendedor()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbVendedorComboDataTable dtbVendedor;
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbVendedorComboBuffer != null) return (DsDataSource.DtbVendedorComboDataTable)_dtbVendedorComboBuffer.Copy();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbVendedor = new DsDataSource.DtbVendedorComboDataTable();

			strSql = "SELECT DISTINCT" +
						" usuario," +
						" nome," +
						" usuario + ' - ' + nome AS usuarioComNome" +
					" FROM" +
						"(" +
							"SELECT" +
								" usuario," +
								" nome" +
							" FROM t_USUARIO" +
							" WHERE" +
								" (vendedor_loja <> 0)" +
							" UNION" +
							" SELECT" +
								" t_USUARIO.usuario AS usuario," +
								" t_USUARIO.nome AS nome" +
							" FROM t_USUARIO" +
								" INNER JOIN t_PERFIL_X_USUARIO ON (t_USUARIO.usuario=t_PERFIL_X_USUARIO.usuario)" +
								" INNER JOIN t_PERFIL ON (t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id)" +
								" INNER JOIN t_PERFIL_ITEM ON (t_PERFIL.id=t_PERFIL_ITEM.id_perfil)" +
							" WHERE" +
								" (t_PERFIL_ITEM.id_operacao = " + Global.Acesso.OP_CEN_ACESSO_TODAS_LOJAS + ")" +
						") AS t" +
					" ORDER BY" +
						" usuario";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbVendedor);
			return dtbVendedor;
		}
		#endregion

		#region [ criaDtbIndicador ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbIndicadorComboDataTable criaDtbIndicador()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbIndicadorComboDataTable dtbIndicador;
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbIndicadorComboBuffer != null) return (DsDataSource.DtbIndicadorComboDataTable)_dtbIndicadorComboBuffer.Copy();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbIndicador = new DsDataSource.DtbIndicadorComboDataTable();

			strSql = "SELECT" +
						" apelido," +
						" razao_social_nome," +
						" apelido + ' - ' + razao_social_nome AS apelidoComRazaoSocialNome" +
					" FROM t_ORCAMENTISTA_E_INDICADOR" +
					" ORDER BY" +
						" apelido";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbIndicador);
			return dtbIndicador;
		}
		#endregion

		#region [ criaDtbUnidadeNegocioCombo ]
		/// <summary>
		/// Cria um DataTable apenas com os dados necessários para uso em combo box, retornando apenas os registros que estiverem com o status ativo/inativo especificado
		/// </summary>
		/// <returns>
		/// Retorna um DataTable apenas com os campos para uso em combo box
		/// </returns>
		public static DsDataSource.DtbUnidadeNegocioComboDataTable criaDtbUnidadeNegocioCombo(eFiltraStAtivo stAtivo)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			String strOrderBy;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbUnidadeNegocioComboDataTable dtbUnidadeNegocio;
			DataView dvUnidadeNegocio;
			#endregion

			#region [ Monta Restrições ]
			if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.ATIVO.ToString() + ")";
			else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
				strWhere = " (st_ativo = " + Global.Cte.FIN.StAtivo.INATIVO.ToString() + ")";
			#endregion

			#region [ Order by ]
			strOrderBy = " descricao";
			#endregion

			#region [ Há dados no buffer local? ]
			if (_dtbUnidadeNegocioComboBuffer != null)
			{
				dvUnidadeNegocio = new DataView();
				dvUnidadeNegocio.Table = _dtbUnidadeNegocioComboBuffer;
				dvUnidadeNegocio.Sort = strOrderBy;
				if (strWhere.Length > 0) dvUnidadeNegocio.RowFilter = strWhere;
				dtbUnidadeNegocio = new DsDataSource.DtbUnidadeNegocioComboDataTable();
				dtbUnidadeNegocio.Merge(dvUnidadeNegocio.ToTable());
				return dtbUnidadeNegocio;
			}
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			dtbUnidadeNegocio = new DsDataSource.DtbUnidadeNegocioComboDataTable();

			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT" +
						" id," +
						" st_ativo," +
						" apelido," +
						" descricao" +
					" FROM t_FIN_UNIDADE_NEGOCIO" +
					strWhere +
					" ORDER BY" +
						strOrderBy;

			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbUnidadeNegocio);
			return dtbUnidadeNegocio;
		}
		#endregion

		#endregion
	}
}
