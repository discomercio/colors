#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient; 
#endregion

namespace ConsolidadorXlsEC
{
    public class ComboDAO
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

        #endregion

        #region [ Atributos ]
        private BancoDados _bd;
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

        #region [ Construtor ]
        public ComboDAO(ref BancoDados bd)
        {
            _bd = bd;
        }
        #endregion

        #region [ Métodos ]

        #region [ criaDtbTransportadoraCombo ]
        public DataTable criaDtbTransportadoraCombo()
        {
            #region [ Declarações ]
            string strSql;
            SqlCommand cmCommand;
            SqlDataAdapter daDataAdapter;
            DataTable dtbTransportadora;
            #endregion

            cmCommand = _bd.criaSqlCommand();
            daDataAdapter = _bd.criaSqlDataAdapter();
            dtbTransportadora = new DataTable();

            strSql = "SELECT id, (id + ' - ' + razao_social) AS id_razao_social FROM t_TRANSPORTADORA ORDER BY id";
            cmCommand.CommandText = strSql;
            daDataAdapter.SelectCommand = cmCommand;
            daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
            daDataAdapter.Fill(dtbTransportadora);
            return dtbTransportadora;
        }
        #endregion

        #region [ criaDtbOrigemPedidoGrupoCombo ]
        public DataTable criaDtbOrigemPedidoGrupoCombo(eFiltraStAtivo stAtivo)
        {
            #region [ Declarações ]
            string strSql;
            string strWhere = "";
            SqlCommand cmCommand;
            SqlDataAdapter daDataAdapter;
            DataTable dtbOrigemPedidoGrupo;
            #endregion

            #region [ Monta Restrições ]
            strWhere = " (grupo='PedidoECommerce_Origem_Grupo')";

            if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
                strWhere = strWhere + " AND (st_inativo = 0)";
            else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
                strWhere = strWhere + " AND (st_inativo = 1)";
            #endregion

            if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

            cmCommand = _bd.criaSqlCommand();
            daDataAdapter = _bd.criaSqlDataAdapter();
            dtbOrigemPedidoGrupo = new DataTable();

            strSql = "SELECT codigo, descricao FROM t_CODIGO_DESCRICAO" +
                strWhere +
                    "ORDER BY descricao";
            cmCommand.CommandText = strSql;
            daDataAdapter.SelectCommand = cmCommand;
            daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
            daDataAdapter.Fill(dtbOrigemPedidoGrupo);
            return dtbOrigemPedidoGrupo;
        }
        #endregion

        #region [ criaDtbOrigemPedidoCombo ]
        public DataTable criaDtbOrigemPedidoCombo(eFiltraStAtivo stAtivo)
        {
            #region [ Declarações ]
            string strSql;
            string strWhere = "";
            SqlCommand cmCommand;
            SqlDataAdapter daDataAdapter;
            DataTable dtbOrigemPedido;
            #endregion

            #region [ Monta Restrições ]
            strWhere = " (grupo='PedidoECommerce_Origem')";

            if (stAtivo == eFiltraStAtivo.SOMENTE_ATIVOS)
                strWhere = strWhere + " AND (st_inativo = 0)";
            else if (stAtivo == eFiltraStAtivo.SOMENTE_INATIVOS)
                strWhere = strWhere + " AND (st_inativo = 1)";
            #endregion

            if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

            cmCommand = _bd.criaSqlCommand();
            daDataAdapter = _bd.criaSqlDataAdapter();
            dtbOrigemPedido = new DataTable();

            strSql = "SELECT codigo, descricao FROM t_CODIGO_DESCRICAO" +
                strWhere +
                    "ORDER BY descricao";
            cmCommand.CommandText = strSql;
            daDataAdapter.SelectCommand = cmCommand;
            daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
            daDataAdapter.Fill(dtbOrigemPedido);
            return dtbOrigemPedido;
        }
		#endregion

		#region [ criaDtbPlataforma ]
		public DataTable criaDtbPlataforma()
		{
			#region [ Declarações ]
			DataTable dtbPlataforma;
			DataRow rowPlataforma;
			#endregion

			dtbPlataforma = new DataTable();
			dtbPlataforma.Columns.Add("codigo", typeof(System.Int32));
			dtbPlataforma.Columns.Add("descricao", typeof(System.String));
			rowPlataforma = dtbPlataforma.NewRow();
			rowPlataforma["codigo"] = Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML;
			rowPlataforma["descricao"] = "Magento v1";
			dtbPlataforma.Rows.Add(rowPlataforma);
			rowPlataforma = dtbPlataforma.NewRow();
			rowPlataforma["codigo"] = Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON;
			rowPlataforma["descricao"] = "Magento v2";
			dtbPlataforma.Rows.Add(rowPlataforma);

			return dtbPlataforma;
		}
		#endregion

		#endregion
	}
}
