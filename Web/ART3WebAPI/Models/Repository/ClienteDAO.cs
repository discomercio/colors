using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data.SqlClient;
using System.Data;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Models.Repository
{
	public class ClienteDAO
	{
		#region [ getClienteByCpfCnpj ]
		public static Cliente getClienteByCpfCnpj(string cpfCnpj)
		{
			#region [ Declarações ]
			string strSql;
			string id_cliente;
			Cliente cliente = null;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if ((cpfCnpj ?? "").Trim().Length == 0) throw new Exception("CPF/CNPJ do cliente não foi informado!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try
			{
				#region [ Monta Select ]
				strSql = "SELECT " +
							"id" +
						" FROM t_CLIENTE" +
						" WHERE" +
							" (cnpj_cpf = '" + Global.digitos(cpfCnpj) + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				rowResultado = dtbResultado.Rows[0];
				id_cliente = BD.readToString(rowResultado["id"]);
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			cliente = getClienteById(id_cliente);
			return cliente;
		}
		#endregion

		#region [ getClienteById ]
		public static Cliente getClienteById(string id)
		{
			#region [ Declarações ]
			String strSql;
			Cliente cliente = new Cliente();
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if ((id ?? "").Trim().Length == 0) throw new Exception("ID do cliente não foi informado!");
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
				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_CLIENTE" +
						" WHERE" +
							" (id = '" + id + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Cliente com ID " + id + " não foi encontrado!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				cliente.id = BD.readToString(rowResultado["id"]);
				cliente.cnpj_cpf = BD.readToString(rowResultado["cnpj_cpf"]);
				cliente.tipo = BD.readToString(rowResultado["tipo"]);
				cliente.ie = BD.readToString(rowResultado["ie"]);
				cliente.rg = BD.readToString(rowResultado["rg"]);
				cliente.nome = BD.readToString(rowResultado["nome"]);
				cliente.sexo = BD.readToString(rowResultado["sexo"]);
				cliente.endereco = BD.readToString(rowResultado["endereco"]);
				cliente.bairro = BD.readToString(rowResultado["bairro"]);
				cliente.cidade = BD.readToString(rowResultado["cidade"]);
				cliente.uf = BD.readToString(rowResultado["uf"]);
				cliente.cep = BD.readToString(rowResultado["cep"]);
				cliente.ddd_res = BD.readToString(rowResultado["ddd_res"]);
				cliente.tel_res = BD.readToString(rowResultado["tel_res"]);
				cliente.ddd_com = BD.readToString(rowResultado["ddd_com"]);
				cliente.tel_com = BD.readToString(rowResultado["tel_com"]);
				cliente.ramal_com = BD.readToString(rowResultado["ramal_com"]);
				cliente.contato = BD.readToString(rowResultado["contato"]);
				cliente.dt_nasc = BD.readToDateTime(rowResultado["dt_nasc"]);
				cliente.filiacao = BD.readToString(rowResultado["filiacao"]);
				cliente.obs_crediticias = BD.readToString(rowResultado["obs_crediticias"]);
				cliente.midia = BD.readToString(rowResultado["midia"]);
				cliente.email = BD.readToString(rowResultado["email"]);
				cliente.email_opcoes = BD.readToString(rowResultado["email_opcoes"]);
				cliente.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
				cliente.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
				cliente.SocMaj_Nome = BD.readToString(rowResultado["SocMaj_Nome"]);
				cliente.SocMaj_CPF = BD.readToString(rowResultado["SocMaj_CPF"]);
				cliente.SocMaj_banco = BD.readToString(rowResultado["SocMaj_banco"]);
				cliente.SocMaj_agencia = BD.readToString(rowResultado["SocMaj_agencia"]);
				cliente.SocMaj_conta = BD.readToString(rowResultado["SocMaj_conta"]);
				cliente.SocMaj_ddd = BD.readToString(rowResultado["SocMaj_ddd"]);
				cliente.SocMaj_telefone = BD.readToString(rowResultado["SocMaj_telefone"]);
				cliente.SocMaj_contato = BD.readToString(rowResultado["SocMaj_contato"]);
				cliente.usuario_cadastro = BD.readToString(rowResultado["usuario_cadastro"]);
				cliente.usuario_ult_atualizacao = BD.readToString(rowResultado["usuario_ult_atualizacao"]);
				cliente.indicador = BD.readToString(rowResultado["indicador"]);
				cliente.endereco_numero = BD.readToString(rowResultado["endereco_numero"]);
				cliente.endereco_complemento = BD.readToString(rowResultado["endereco_complemento"]);
				cliente.nome_iniciais_em_maiusculas = BD.readToString(rowResultado["nome_iniciais_em_maiusculas"]);
				cliente.spc_negativado_status = BD.readToByte(rowResultado["spc_negativado_status"]);
				cliente.spc_negativado_data_negativacao = BD.readToDateTime(rowResultado["spc_negativado_data_negativacao"]);
				cliente.spc_negativado_data = BD.readToDateTime(rowResultado["spc_negativado_data"]);
				cliente.spc_negativado_data_hora = BD.readToDateTime(rowResultado["spc_negativado_data_hora"]);
				cliente.spc_negativado_usuario = BD.readToString(rowResultado["spc_negativado_usuario"]);
				cliente.email_anterior = BD.readToString(rowResultado["email_anterior"]);
				cliente.email_atualizacao_data = BD.readToDateTime(rowResultado["email_atualizacao_data"]);
				cliente.email_atualizacao_data_hora = BD.readToDateTime(rowResultado["email_atualizacao_data_hora"]);
				cliente.email_atualizacao_usuario = BD.readToString(rowResultado["email_atualizacao_usuario"]);
				cliente.contribuinte_icms_status = BD.readToByte(rowResultado["contribuinte_icms_status"]);
				cliente.contribuinte_icms_data = BD.readToDateTime(rowResultado["contribuinte_icms_data"]);
				cliente.contribuinte_icms_data_hora = BD.readToDateTime(rowResultado["contribuinte_icms_data_hora"]);
				cliente.contribuinte_icms_usuario = BD.readToString(rowResultado["contribuinte_icms_usuario"]);
				cliente.produtor_rural_status = BD.readToByte(rowResultado["produtor_rural_status"]);
				cliente.produtor_rural_data = BD.readToDateTime(rowResultado["produtor_rural_data"]);
				cliente.produtor_rural_data_hora = BD.readToDateTime(rowResultado["produtor_rural_data_hora"]);
				cliente.produtor_rural_usuario = BD.readToString(rowResultado["produtor_rural_usuario"]);
				cliente.email_xml = BD.readToString(rowResultado["email_xml"]);
				cliente.ddd_cel = BD.readToString(rowResultado["ddd_cel"]);
				cliente.tel_cel = BD.readToString(rowResultado["tel_cel"]);
				cliente.ddd_com_2 = BD.readToString(rowResultado["ddd_com_2"]);
				cliente.tel_com_2 = BD.readToString(rowResultado["tel_com_2"]);
				cliente.ramal_com_2 = BD.readToString(rowResultado["ramal_com_2"]);
				cliente.sistema_responsavel_cadastro = BD.readToInt(rowResultado["sistema_responsavel_cadastro"]);
				cliente.sistema_responsavel_atualizacao = BD.readToInt(rowResultado["sistema_responsavel_atualizacao"]);
				#endregion
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return cliente;
		}
		#endregion
	}
}