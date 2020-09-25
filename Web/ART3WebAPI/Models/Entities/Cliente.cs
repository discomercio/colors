using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class Cliente
	{
		public string id { get; set; }
		public string cnpj_cpf { get; set; }
		public string tipo { get; set; }
		public string ie { get; set; }
		public string rg { get; set; }
		public string nome { get; set; }
		public string sexo { get; set; }
		public string endereco { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string bairro { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
		public string cep { get; set; }
		public string ddd_res { get; set; }
		public string tel_res { get; set; }
		public string ddd_com { get; set; }
		public string tel_com { get; set; }
		public string ramal_com { get; set; }
		public string contato { get; set; }
		public DateTime dt_nasc { get; set; }
		public string filiacao { get; set; }
		public string obs_crediticias { get; set; }
		public string midia { get; set; }
		public string email { get; set; }
		public string email_opcoes { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public string SocMaj_Nome { get; set; }
		public string SocMaj_CPF { get; set; }
		public string SocMaj_banco { get; set; }
		public string SocMaj_agencia { get; set; }
		public string SocMaj_conta { get; set; }
		public string SocMaj_ddd { get; set; }
		public string SocMaj_telefone { get; set; }
		public string SocMaj_contato { get; set; }
		public string usuario_cadastro { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public string indicador { get; set; }
		public string nome_iniciais_em_maiusculas { get; set; }
		public byte spc_negativado_status { get; set; }
		public DateTime spc_negativado_data_negativacao { get; set; }
		public DateTime spc_negativado_data { get; set; }
		public DateTime spc_negativado_data_hora { get; set; }
		public string spc_negativado_usuario { get; set; }
		public string email_anterior { get; set; }
		public DateTime email_atualizacao_data { get; set; }
		public DateTime email_atualizacao_data_hora { get; set; }
		public string email_atualizacao_usuario { get; set; }
		public byte contribuinte_icms_status { get; set; }
		public DateTime contribuinte_icms_data { get; set; }
		public DateTime contribuinte_icms_data_hora { get; set; }
		public string contribuinte_icms_usuario { get; set; }
		public byte produtor_rural_status { get; set; }
		public DateTime produtor_rural_data { get; set; }
		public DateTime produtor_rural_data_hora { get; set; }
		public string produtor_rural_usuario { get; set; }
		public string email_xml { get; set; }
		public string ddd_cel { get; set; }
		public string tel_cel { get; set; }
		public string ddd_com_2 { get; set; }
		public string tel_com_2 { get; set; }
		public string ramal_com_2 { get; set; }
		public int sistema_responsavel_cadastro { get; set; }
		public int sistema_responsavel_atualizacao { get; set; }
	}
}