using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class LojaDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
		private SqlCommand cmSelectLoja;
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
		public LojaDAO(ref BancoDados bd)
		{
			_bd = bd;
			inicializaObjetos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetos ]
		public void inicializaObjetos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmSelectLoja ]
			strSql = "SELECT " +
						"*" +
					" FROM t_LOJA" +
					" WHERE" +
						" (loja = @loja)";
			cmSelectLoja = _bd.criaSqlCommand();
			cmSelectLoja.CommandText = strSql;
			cmSelectLoja.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmSelectLoja.Prepare();
			#endregion
		}
		#endregion

		#region [ GetLoja ]
		public Loja GetLoja(string numeroLoja, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "LojaDAO.GetLoja()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			Loja loja = new Loja();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(numeroLoja))
				{
					msg_erro = "Número da loja não informado";
					return null;
				}

				numeroLoja = numeroLoja.Trim();
				numeroLoja = Global.normalizaNumeroLoja(numeroLoja);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				cmSelectLoja.Parameters["@loja"].Value = numeroLoja;
				daDataAdapter.SelectCommand = cmSelectLoja;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

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
				loja.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
				loja.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
				loja.comissao_indicacao = BD.readToSingle(rowResultado["comissao_indicacao"]);
				loja.PercMaxSenhaDesconto = BD.readToSingle(rowResultado["PercMaxSenhaDesconto"]);
				loja.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
				loja.id_plano_contas_grupo = BD.readToInt(rowResultado["id_plano_contas_grupo"]);
				loja.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
				loja.natureza = BD.readToString(rowResultado["natureza"]);
				loja.PercMaxDescSemZerarRT = BD.readToSingle(rowResultado["PercMaxDescSemZerarRT"]);
				loja.perc_max_comissao = BD.readToSingle(rowResultado["perc_max_comissao"]);
				loja.perc_max_comissao_e_desconto = BD.readToSingle(rowResultado["perc_max_comissao_e_desconto"]);
				loja.perc_max_comissao_e_desconto_nivel2 = BD.readToSingle(rowResultado["perc_max_comissao_e_desconto_nivel2"]);
				loja.perc_max_comissao_e_desconto_nivel2_pj = BD.readToSingle(rowResultado["perc_max_comissao_e_desconto_nivel2_pj"]);
				loja.perc_max_comissao_e_desconto_pj = BD.readToSingle(rowResultado["perc_max_comissao_e_desconto_pj"]);
				loja.unidade_negocio = BD.readToString(rowResultado["unidade_negocio"]);
				loja.magento_api_versao = BD.readToInt(rowResultado["magento_api_versao"]);
				loja.magento_api_urlWebService = BD.readToString(rowResultado["magento_api_urlWebService"]);
				loja.magento_api_username = BD.readToString(rowResultado["magento_api_username"]);
				loja.magento_api_password = Criptografia.Descriptografa(BD.readToString(rowResultado["magento_api_password"]));
				loja.magento_api_rest_endpoint = BD.readToString(rowResultado["magento_api_rest_endpoint"]);
				loja.magento_api_rest_access_token = BD.readToString(rowResultado["magento_api_rest_access_token"]);
				loja.magento_api_rest_force_get_sales_order_by_entity_id = BD.readToByte(rowResultado["magento_api_rest_force_get_sales_order_by_entity_id"]);
				loja.magento_api_rest_prefixo_num_magento = BD.readToString(rowResultado["magento_api_rest_prefixo_num_magento"]);

				return loja;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#endregion
	}
}
