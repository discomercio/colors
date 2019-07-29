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
			Cliente cliente = new Cliente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
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

			rowResultado = dtbResultado.Rows[0];

			cliente.id = rowResultado["id"].ToString();
			cliente.cnpj_cpf = !Convert.IsDBNull(rowResultado["cnpj_cpf"]) ? rowResultado["cnpj_cpf"].ToString() : "";
			cliente.tipo = !Convert.IsDBNull(rowResultado["tipo"]) ? rowResultado["tipo"].ToString() : "";
			cliente.ie = !Convert.IsDBNull(rowResultado["ie"]) ? rowResultado["ie"].ToString() : "";
			cliente.rg = !Convert.IsDBNull(rowResultado["rg"]) ? rowResultado["rg"].ToString() : "";
			cliente.nome = !Convert.IsDBNull(rowResultado["nome"]) ? rowResultado["nome"].ToString() : "";
			cliente.sexo = !Convert.IsDBNull(rowResultado["sexo"]) ? rowResultado["sexo"].ToString() : "";
			cliente.endereco = !Convert.IsDBNull(rowResultado["endereco"]) ? rowResultado["endereco"].ToString() : "";
			cliente.endereco_numero = !Convert.IsDBNull(rowResultado["endereco_numero"]) ? rowResultado["endereco_numero"].ToString() : "";
			cliente.endereco_complemento = !Convert.IsDBNull(rowResultado["endereco_complemento"]) ? rowResultado["endereco_complemento"].ToString() : "";
			cliente.bairro = !Convert.IsDBNull(rowResultado["bairro"]) ? rowResultado["bairro"].ToString() : "";
			cliente.cidade = !Convert.IsDBNull(rowResultado["cidade"]) ? rowResultado["cidade"].ToString() : "";
			cliente.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
			cliente.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
			cliente.ddd_res = !Convert.IsDBNull(rowResultado["ddd_res"]) ? rowResultado["ddd_res"].ToString() : "";
			cliente.tel_res = !Convert.IsDBNull(rowResultado["tel_res"]) ? rowResultado["tel_res"].ToString() : "";
			cliente.ddd_com = !Convert.IsDBNull(rowResultado["ddd_com"]) ? rowResultado["ddd_com"].ToString() : "";
			cliente.tel_com = !Convert.IsDBNull(rowResultado["tel_com"]) ? rowResultado["tel_com"].ToString() : "";
			cliente.ramal_com = !Convert.IsDBNull(rowResultado["ramal_com"]) ? rowResultado["ramal_com"].ToString() : "";
			cliente.contato = !Convert.IsDBNull(rowResultado["contato"]) ? rowResultado["contato"].ToString() : "";
			cliente.dt_nasc = !Convert.IsDBNull(rowResultado["dt_nasc"]) ? (DateTime)rowResultado["dt_nasc"] : DateTime.MinValue;
			cliente.filiacao = !Convert.IsDBNull(rowResultado["filiacao"]) ? rowResultado["filiacao"].ToString() : "";
			cliente.obs_crediticias = !Convert.IsDBNull(rowResultado["obs_crediticias"]) ? rowResultado["obs_crediticias"].ToString() : "";
			cliente.midia = !Convert.IsDBNull(rowResultado["midia"]) ? rowResultado["midia"].ToString() : "";
			cliente.email = !Convert.IsDBNull(rowResultado["email"]) ? rowResultado["email"].ToString() : "";
			cliente.email_opcoes = !Convert.IsDBNull(rowResultado["email_opcoes"]) ? rowResultado["email_opcoes"].ToString() : "";
			cliente.dt_cadastro = !Convert.IsDBNull(rowResultado["dt_cadastro"]) ? (DateTime)rowResultado["dt_cadastro"] : DateTime.MinValue;
			cliente.dt_ult_atualizacao = !Convert.IsDBNull(rowResultado["dt_ult_atualizacao"]) ? (DateTime)rowResultado["dt_ult_atualizacao"] : DateTime.MinValue;
			cliente.socMaj_Nome = !Convert.IsDBNull(rowResultado["SocMaj_Nome"]) ? rowResultado["SocMaj_Nome"].ToString() : "";
			cliente.socMaj_CPF = !Convert.IsDBNull(rowResultado["SocMaj_CPF"]) ? rowResultado["SocMaj_CPF"].ToString() : "";
			cliente.socMaj_banco = !Convert.IsDBNull(rowResultado["SocMaj_banco"]) ? rowResultado["SocMaj_banco"].ToString() : "";
			cliente.socMaj_agencia = !Convert.IsDBNull(rowResultado["SocMaj_agencia"]) ? rowResultado["SocMaj_agencia"].ToString() : "";
			cliente.socMaj_conta = !Convert.IsDBNull(rowResultado["SocMaj_conta"]) ? rowResultado["SocMaj_conta"].ToString() : "";
			cliente.socMaj_ddd = !Convert.IsDBNull(rowResultado["SocMaj_ddd"]) ? rowResultado["SocMaj_ddd"].ToString() : "";
			cliente.socMaj_telefone = !Convert.IsDBNull(rowResultado["SocMaj_telefone"]) ? rowResultado["SocMaj_telefone"].ToString() : "";
			cliente.socMaj_contato = !Convert.IsDBNull(rowResultado["SocMaj_contato"]) ? rowResultado["SocMaj_contato"].ToString() : "";
			cliente.usuario_cadastro = !Convert.IsDBNull(rowResultado["usuario_cadastro"]) ? rowResultado["usuario_cadastro"].ToString() : "";
			cliente.usuario_ult_atualizacao = !Convert.IsDBNull(rowResultado["usuario_ult_atualizacao"]) ? rowResultado["usuario_ult_atualizacao"].ToString() : "";
			cliente.indicador = !Convert.IsDBNull(rowResultado["indicador"]) ? rowResultado["indicador"].ToString() : "";
			cliente.spc_negativado_status = BD.readToByte(rowResultado["spc_negativado_status"]);
			cliente.spc_negativado_data_negativacao = !Convert.IsDBNull(rowResultado["spc_negativado_data_negativacao"]) ? (DateTime)rowResultado["spc_negativado_data_negativacao"] : DateTime.MinValue;
			cliente.spc_negativado_data = !Convert.IsDBNull(rowResultado["spc_negativado_data"]) ? (DateTime)rowResultado["spc_negativado_data"] : DateTime.MinValue;
			cliente.spc_negativado_data_hora = !Convert.IsDBNull(rowResultado["spc_negativado_data_hora"]) ? (DateTime)rowResultado["spc_negativado_data_hora"] : DateTime.MinValue;
			cliente.spc_negativado_usuario = BD.readToString(rowResultado["spc_negativado_usuario"]);
			cliente.contribuinte_icms_status = BD.readToByte(rowResultado["contribuinte_icms_status"]);
			cliente.produtor_rural_status = BD.readToByte(rowResultado["produtor_rural_status"]);
			cliente.email_xml = BD.readToString(rowResultado["email_xml"]);
			cliente.ddd_cel = BD.readToString(rowResultado["ddd_cel"]);
			cliente.tel_cel = BD.readToString(rowResultado["tel_cel"]);
			cliente.ddd_com_2 = BD.readToString(rowResultado["ddd_com_2"]);
			cliente.tel_com_2 = BD.readToString(rowResultado["tel_com_2"]);
			cliente.ramal_com_2 = BD.readToString(rowResultado["ramal_com_2"]);
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
			Cliente cliente = new Cliente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
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

			rowResultado = dtbResultado.Rows[0];

			cliente.id = rowResultado["id"].ToString();
			cliente.cnpj_cpf = !Convert.IsDBNull(rowResultado["cnpj_cpf"]) ? rowResultado["cnpj_cpf"].ToString() : "";
			cliente.tipo = !Convert.IsDBNull(rowResultado["tipo"]) ? rowResultado["tipo"].ToString() : "";
			cliente.ie = !Convert.IsDBNull(rowResultado["ie"]) ? rowResultado["ie"].ToString() : "";
			cliente.rg = !Convert.IsDBNull(rowResultado["rg"]) ? rowResultado["rg"].ToString() : "";
			cliente.nome = !Convert.IsDBNull(rowResultado["nome"]) ? rowResultado["nome"].ToString() : "";
			cliente.sexo = !Convert.IsDBNull(rowResultado["sexo"]) ? rowResultado["sexo"].ToString() : "";
			cliente.endereco = !Convert.IsDBNull(rowResultado["endereco"]) ? rowResultado["endereco"].ToString() : "";
			cliente.endereco_numero = !Convert.IsDBNull(rowResultado["endereco_numero"]) ? rowResultado["endereco_numero"].ToString() : "";
			cliente.endereco_complemento = !Convert.IsDBNull(rowResultado["endereco_complemento"]) ? rowResultado["endereco_complemento"].ToString() : "";
			cliente.bairro = !Convert.IsDBNull(rowResultado["bairro"]) ? rowResultado["bairro"].ToString() : "";
			cliente.cidade = !Convert.IsDBNull(rowResultado["cidade"]) ? rowResultado["cidade"].ToString() : "";
			cliente.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
			cliente.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
			cliente.ddd_res = !Convert.IsDBNull(rowResultado["ddd_res"]) ? rowResultado["ddd_res"].ToString() : "";
			cliente.tel_res = !Convert.IsDBNull(rowResultado["tel_res"]) ? rowResultado["tel_res"].ToString() : "";
			cliente.ddd_com = !Convert.IsDBNull(rowResultado["ddd_com"]) ? rowResultado["ddd_com"].ToString() : "";
			cliente.tel_com = !Convert.IsDBNull(rowResultado["tel_com"]) ? rowResultado["tel_com"].ToString() : "";
			cliente.ramal_com = !Convert.IsDBNull(rowResultado["ramal_com"]) ? rowResultado["ramal_com"].ToString() : "";
			cliente.contato = !Convert.IsDBNull(rowResultado["contato"]) ? rowResultado["contato"].ToString() : "";
			cliente.dt_nasc = !Convert.IsDBNull(rowResultado["dt_nasc"]) ? (DateTime)rowResultado["dt_nasc"] : DateTime.MinValue;
			cliente.filiacao = !Convert.IsDBNull(rowResultado["filiacao"]) ? rowResultado["filiacao"].ToString() : "";
			cliente.obs_crediticias = !Convert.IsDBNull(rowResultado["obs_crediticias"]) ? rowResultado["obs_crediticias"].ToString() : "";
			cliente.midia = !Convert.IsDBNull(rowResultado["midia"]) ? rowResultado["midia"].ToString() : "";
			cliente.email = !Convert.IsDBNull(rowResultado["email"]) ? rowResultado["email"].ToString() : "";
			cliente.email_opcoes = !Convert.IsDBNull(rowResultado["email_opcoes"]) ? rowResultado["email_opcoes"].ToString() : "";
			cliente.dt_cadastro = !Convert.IsDBNull(rowResultado["dt_cadastro"]) ? (DateTime)rowResultado["dt_cadastro"] : DateTime.MinValue;
			cliente.dt_ult_atualizacao = !Convert.IsDBNull(rowResultado["dt_ult_atualizacao"]) ? (DateTime)rowResultado["dt_ult_atualizacao"] : DateTime.MinValue;
			cliente.socMaj_Nome = !Convert.IsDBNull(rowResultado["SocMaj_Nome"]) ? rowResultado["SocMaj_Nome"].ToString() : "";
			cliente.socMaj_CPF = !Convert.IsDBNull(rowResultado["SocMaj_CPF"]) ? rowResultado["SocMaj_CPF"].ToString() : "";
			cliente.socMaj_banco = !Convert.IsDBNull(rowResultado["SocMaj_banco"]) ? rowResultado["SocMaj_banco"].ToString() : "";
			cliente.socMaj_agencia = !Convert.IsDBNull(rowResultado["SocMaj_agencia"]) ? rowResultado["SocMaj_agencia"].ToString() : "";
			cliente.socMaj_conta = !Convert.IsDBNull(rowResultado["SocMaj_conta"]) ? rowResultado["SocMaj_conta"].ToString() : "";
			cliente.socMaj_ddd = !Convert.IsDBNull(rowResultado["SocMaj_ddd"]) ? rowResultado["SocMaj_ddd"].ToString() : "";
			cliente.socMaj_telefone = !Convert.IsDBNull(rowResultado["SocMaj_telefone"]) ? rowResultado["SocMaj_telefone"].ToString() : "";
			cliente.socMaj_contato = !Convert.IsDBNull(rowResultado["SocMaj_contato"]) ? rowResultado["SocMaj_contato"].ToString() : "";
			cliente.usuario_cadastro = !Convert.IsDBNull(rowResultado["usuario_cadastro"]) ? rowResultado["usuario_cadastro"].ToString() : "";
			cliente.usuario_ult_atualizacao = !Convert.IsDBNull(rowResultado["usuario_ult_atualizacao"]) ? rowResultado["usuario_ult_atualizacao"].ToString() : "";
			cliente.indicador = !Convert.IsDBNull(rowResultado["indicador"]) ? rowResultado["indicador"].ToString() : "";
			#endregion

			return cliente;
		}
		#endregion
	}
}
