using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ OrcamentistaIndicadorCompleto ]
	public class OrcamentistaIndicadorCompleto
	{
		public string apelido { get; set; }
		public string cnpj_cpf { get; set; }
		public string tipo { get; set; }
		public string ie_rg { get; set; }
		public string razao_social_nome { get; set; }
		public string endereco { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string bairro { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
		public string cep { get; set; }
		public string ddd { get; set; }
		public string telefone { get; set; }
		public string fax { get; set; }
		public string ddd_cel { get; set; }
		public string tel_cel { get; set; }
		public string contato { get; set; }
		public string banco { get; set; }
		public string agencia { get; set; }
		public string conta { get; set; }
		public string favorecido { get; set; }
		public string loja { get; set; }
		public string vendedor { get; set; }
		public int hab_acesso_sistema { get; set; }
		public string status { get; set; }
		public string senha { get; set; }
		public string datastamp { get; set; }
		public DateTime dt_ult_alteracao_senha { get; set; }
		public DateTime dt_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public DateTime dt_ult_acesso { get; set; }
		public string desempenho_nota { get; set; }
		public DateTime desempenho_nota_data { get; set; }
		public string desempenho_nota_usuario { get; set; }
		public double perc_desagio_RA { get; set; }
		public decimal vl_limite_mensal { get; set; }
		public string email { get; set; }
		public string captador { get; set; }
		public int checado_status { get; set; }
		public DateTime checado_data { get; set; }
		public string checado_usuario { get; set; }
		public string obs { get; set; }
		public decimal vl_meta { get; set; }
		public string UsuarioUltAtualizVlMeta { get; set; }
		public DateTime DtHrUltAtualizVlMeta { get; set; }
		public int permite_RA_status { get; set; }
		public string permite_RA_usuario { get; set; }
		public DateTime permite_RA_data_hora { get; set; }
		public string forma_como_conheceu_codigo { get; set; }
		public string forma_como_conheceu_usuario { get; set; }
		public DateTime forma_como_conheceu_data { get; set; }
		public DateTime forma_como_conheceu_data_hora { get; set; }
		public string forma_como_conheceu_codigo_anterior { get; set; }
		public string nome_fantasia { get; set; }
		public int tipo_estabelecimento { get; set; }
		public string nextel { get; set; }
		public string email2 { get; set; }
		public string email3 { get; set; }
		public string razao_social_nome_iniciais_em_maiusculas { get; set; }
		public byte st_reg_copiado_automaticamente { get; set; }
		public DateTime dt_hr_reg_atualizado_automaticamente { get; set; }
		public string etq_endereco { get; set; }
		public string etq_endereco_numero { get; set; }
		public string etq_endereco_complemento { get; set; }
		public string etq_bairro { get; set; }
		public string etq_cidade { get; set; }
		public string etq_uf { get; set; }
		public string etq_cep { get; set; }
		public string etq_email { get; set; }
		public string etq_ddd_1 { get; set; }
		public string etq_tel_1 { get; set; }
		public string etq_ddd_2 { get; set; }
		public string etq_tel_2 { get; set; }
		public string favorecido_cnpj_cpf { get; set; }
		public string agencia_dv { get; set; }
		public string conta_operacao { get; set; }
		public string conta_dv { get; set; }
		public string tipo_conta { get; set; }
		public DateTime vendedor_dt_ult_atualizacao { get; set; }
		public DateTime vendedor_dt_hr_ult_atualizacao { get; set; }
		public string vendedor_usuario_ult_atualizacao { get; set; }
	}
	#endregion

	#region [ OrcamentistaIndicadorBasico ]
	public class OrcamentistaIndicadorBasico
	{
		public string apelido { get; set; }
		public string cnpj_cpf { get; set; }
		public string tipo { get; set; }
		public string ie_rg { get; set; }
		public string razao_social_nome { get; set; }
		public string endereco { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string bairro { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
		public string cep { get; set; }
		public string ddd { get; set; }
		public string telefone { get; set; }
		public string fax { get; set; }
		public string ddd_cel { get; set; }
		public string tel_cel { get; set; }
		public string contato { get; set; }
		public string loja { get; set; }
		public string vendedor { get; set; }
		public string status { get; set; }
		public DateTime dt_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public string email { get; set; }
		public string captador { get; set; }
		public int permite_RA_status { get; set; }
		public string nome_fantasia { get; set; }
		public int tipo_estabelecimento { get; set; }
		public string nextel { get; set; }
		public string email2 { get; set; }
		public string email3 { get; set; }
		public string razao_social_nome_iniciais_em_maiusculas { get; set; }
		public string etq_endereco { get; set; }
		public string etq_endereco_numero { get; set; }
		public string etq_endereco_complemento { get; set; }
		public string etq_bairro { get; set; }
		public string etq_cidade { get; set; }
		public string etq_uf { get; set; }
		public string etq_cep { get; set; }
		public string etq_email { get; set; }
		public string etq_ddd_1 { get; set; }
		public string etq_tel_1 { get; set; }
		public string etq_ddd_2 { get; set; }
		public string etq_tel_2 { get; set; }
	}
	#endregion

	#region [ OrcamentistaIndicadorResumoPesquisa ]
	public class OrcamentistaIndicadorResumoPesquisa
	{
		public string apelido { get; set; }
		public string cnpj_cpf { get; set; }
		public string razao_social_nome { get; set; }
		public string razao_social_nome_iniciais_em_maiusculas { get; set; }
		public string loja { get; set; }
		public string vendedor { get; set; }
		public string captador { get; set; }
		public string status { get; set; }
		public int permite_RA_status { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
	}
	#endregion
}