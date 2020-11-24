using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace FinanceiroService
{
	class ClienteDAO
	{
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
		static ClienteDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{

		}
		#endregion

		#region [ carregaDadosClienteFromDataRow ]
		public static Cliente carregaDadosClienteFromDataRow(DataRow rowDadosCliente)
		{
			#region [ Declarações ]
			Cliente cliente = new Cliente();
			#endregion

			cliente.id = rowDadosCliente["id"].ToString();
			cliente.cnpj_cpf = !Convert.IsDBNull(rowDadosCliente["cnpj_cpf"]) ? rowDadosCliente["cnpj_cpf"].ToString() : "";
			cliente.tipo = !Convert.IsDBNull(rowDadosCliente["tipo"]) ? rowDadosCliente["tipo"].ToString() : "";
			cliente.ie = !Convert.IsDBNull(rowDadosCliente["ie"]) ? rowDadosCliente["ie"].ToString() : "";
			cliente.rg = !Convert.IsDBNull(rowDadosCliente["rg"]) ? rowDadosCliente["rg"].ToString() : "";
			cliente.nome = !Convert.IsDBNull(rowDadosCliente["nome"]) ? rowDadosCliente["nome"].ToString() : "";
			cliente.sexo = !Convert.IsDBNull(rowDadosCliente["sexo"]) ? rowDadosCliente["sexo"].ToString() : "";
			cliente.endereco = !Convert.IsDBNull(rowDadosCliente["endereco"]) ? rowDadosCliente["endereco"].ToString() : "";
			cliente.endereco_numero = !Convert.IsDBNull(rowDadosCliente["endereco_numero"]) ? rowDadosCliente["endereco_numero"].ToString() : "";
			cliente.endereco_complemento = !Convert.IsDBNull(rowDadosCliente["endereco_complemento"]) ? rowDadosCliente["endereco_complemento"].ToString() : "";
			cliente.bairro = !Convert.IsDBNull(rowDadosCliente["bairro"]) ? rowDadosCliente["bairro"].ToString() : "";
			cliente.cidade = !Convert.IsDBNull(rowDadosCliente["cidade"]) ? rowDadosCliente["cidade"].ToString() : "";
			cliente.uf = !Convert.IsDBNull(rowDadosCliente["uf"]) ? rowDadosCliente["uf"].ToString() : "";
			cliente.cep = !Convert.IsDBNull(rowDadosCliente["cep"]) ? rowDadosCliente["cep"].ToString() : "";
			cliente.ddd_res = !Convert.IsDBNull(rowDadosCliente["ddd_res"]) ? rowDadosCliente["ddd_res"].ToString() : "";
			cliente.tel_res = !Convert.IsDBNull(rowDadosCliente["tel_res"]) ? rowDadosCliente["tel_res"].ToString() : "";
			cliente.ddd_com = !Convert.IsDBNull(rowDadosCliente["ddd_com"]) ? rowDadosCliente["ddd_com"].ToString() : "";
			cliente.tel_com = !Convert.IsDBNull(rowDadosCliente["tel_com"]) ? rowDadosCliente["tel_com"].ToString() : "";
			cliente.ramal_com = !Convert.IsDBNull(rowDadosCliente["ramal_com"]) ? rowDadosCliente["ramal_com"].ToString() : "";
			cliente.contato = !Convert.IsDBNull(rowDadosCliente["contato"]) ? rowDadosCliente["contato"].ToString() : "";
			cliente.dt_nasc = !Convert.IsDBNull(rowDadosCliente["dt_nasc"]) ? (DateTime)rowDadosCliente["dt_nasc"] : DateTime.MinValue;
			cliente.filiacao = !Convert.IsDBNull(rowDadosCliente["filiacao"]) ? rowDadosCliente["filiacao"].ToString() : "";
			cliente.obs_crediticias = !Convert.IsDBNull(rowDadosCliente["obs_crediticias"]) ? rowDadosCliente["obs_crediticias"].ToString() : "";
			cliente.midia = !Convert.IsDBNull(rowDadosCliente["midia"]) ? rowDadosCliente["midia"].ToString() : "";
			cliente.email = !Convert.IsDBNull(rowDadosCliente["email"]) ? rowDadosCliente["email"].ToString() : "";
			cliente.email_opcoes = !Convert.IsDBNull(rowDadosCliente["email_opcoes"]) ? rowDadosCliente["email_opcoes"].ToString() : "";
			cliente.dt_cadastro = !Convert.IsDBNull(rowDadosCliente["dt_cadastro"]) ? (DateTime)rowDadosCliente["dt_cadastro"] : DateTime.MinValue;
			cliente.dt_ult_atualizacao = !Convert.IsDBNull(rowDadosCliente["dt_ult_atualizacao"]) ? (DateTime)rowDadosCliente["dt_ult_atualizacao"] : DateTime.MinValue;
			cliente.socMaj_Nome = !Convert.IsDBNull(rowDadosCliente["SocMaj_Nome"]) ? rowDadosCliente["SocMaj_Nome"].ToString() : "";
			cliente.socMaj_CPF = !Convert.IsDBNull(rowDadosCliente["SocMaj_CPF"]) ? rowDadosCliente["SocMaj_CPF"].ToString() : "";
			cliente.socMaj_banco = !Convert.IsDBNull(rowDadosCliente["SocMaj_banco"]) ? rowDadosCliente["SocMaj_banco"].ToString() : "";
			cliente.socMaj_agencia = !Convert.IsDBNull(rowDadosCliente["SocMaj_agencia"]) ? rowDadosCliente["SocMaj_agencia"].ToString() : "";
			cliente.socMaj_conta = !Convert.IsDBNull(rowDadosCliente["SocMaj_conta"]) ? rowDadosCliente["SocMaj_conta"].ToString() : "";
			cliente.socMaj_ddd = !Convert.IsDBNull(rowDadosCliente["SocMaj_ddd"]) ? rowDadosCliente["SocMaj_ddd"].ToString() : "";
			cliente.socMaj_telefone = !Convert.IsDBNull(rowDadosCliente["SocMaj_telefone"]) ? rowDadosCliente["SocMaj_telefone"].ToString() : "";
			cliente.socMaj_contato = !Convert.IsDBNull(rowDadosCliente["SocMaj_contato"]) ? rowDadosCliente["SocMaj_contato"].ToString() : "";
			cliente.usuario_cadastro = !Convert.IsDBNull(rowDadosCliente["usuario_cadastro"]) ? rowDadosCliente["usuario_cadastro"].ToString() : "";
			cliente.usuario_ult_atualizacao = !Convert.IsDBNull(rowDadosCliente["usuario_ult_atualizacao"]) ? rowDadosCliente["usuario_ult_atualizacao"].ToString() : "";
			cliente.indicador = !Convert.IsDBNull(rowDadosCliente["indicador"]) ? rowDadosCliente["indicador"].ToString() : "";
			cliente.spc_negativado_status = BD.readToByte(rowDadosCliente["spc_negativado_status"]);
			cliente.spc_negativado_data_negativacao = !Convert.IsDBNull(rowDadosCliente["spc_negativado_data_negativacao"]) ? (DateTime)rowDadosCliente["spc_negativado_data_negativacao"] : DateTime.MinValue;
			cliente.spc_negativado_data = !Convert.IsDBNull(rowDadosCliente["spc_negativado_data"]) ? (DateTime)rowDadosCliente["spc_negativado_data"] : DateTime.MinValue;
			cliente.spc_negativado_data_hora = !Convert.IsDBNull(rowDadosCliente["spc_negativado_data_hora"]) ? (DateTime)rowDadosCliente["spc_negativado_data_hora"] : DateTime.MinValue;
			cliente.spc_negativado_usuario = BD.readToString(rowDadosCliente["spc_negativado_usuario"]);
			cliente.contribuinte_icms_status = BD.readToByte(rowDadosCliente["contribuinte_icms_status"]);
			cliente.produtor_rural_status = BD.readToByte(rowDadosCliente["produtor_rural_status"]);
			cliente.email_xml = BD.readToString(rowDadosCliente["email_xml"]);
			cliente.ddd_cel = BD.readToString(rowDadosCliente["ddd_cel"]);
			cliente.tel_cel = BD.readToString(rowDadosCliente["tel_cel"]);
			cliente.ddd_com_2 = BD.readToString(rowDadosCliente["ddd_com_2"]);
			cliente.tel_com_2 = BD.readToString(rowDadosCliente["tel_com_2"]);
			cliente.ramal_com_2 = BD.readToString(rowDadosCliente["ramal_com_2"]);

			return cliente;
		}
		#endregion

		#region [ getCliente ]
		/// <summary>
		/// Retorna um objeto Cliente contendo os dados lidos do BD
		/// </summary>
		/// <param name="id">
		/// Identificação do registro
		/// </param>
		/// <returns>
		/// Retorna um objeto Cliente contendo os dados lidos do BD
		/// </returns>
		public static Cliente getCliente(String id)
		{
			#region [ Declarações ]
			String strSql;
			Cliente cliente;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			#region [ Consistências ]
			if (id == null) throw new Exception("O identificador do registro não foi fornecido!!");
			if (id.Trim().Length == 0) throw new Exception("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cliente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_CLIENTE" +
					" WHERE" +
						" (id = '" + id.Trim() + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new Exception("Registro id=" + id.ToString() + " não localizado na tabela t_CLIENTE!!");

			cliente = carregaDadosClienteFromDataRow(dtbResultado.Rows[0]);
			#endregion

			return cliente;
		}
		#endregion

		#region [ getClienteByCnpjCpf ]
		/// <summary>
		/// Retorna um objeto Cliente contendo os dados lidos do BD
		/// </summary>
		/// <param name="cnpjCpf">
		/// Número do CNPJ/CPF
		/// </param>
		/// <returns>
		/// Retorna um objeto Cliente contendo os dados lidos do BD
		/// </returns>
		public static Cliente getClienteCnpjCpf(String cnpjCpf)
		{
			#region [ Declarações ]
			String strSql;
			Cliente cliente;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			#region [ Consistências ]
			if (cnpjCpf == null) return null;
			if (cnpjCpf.Trim().Length == 0) return null;
			if (!Global.isCnpjCpfOk(cnpjCpf)) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cliente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_CLIENTE" +
					" WHERE" +
						" (cnpj_cpf = '" + Global.digitos(cnpjCpf.Trim()) + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) return null;

			cliente = carregaDadosClienteFromDataRow(dtbResultado.Rows[0]);
			#endregion

			return cliente;
		}
		#endregion
	}
}
