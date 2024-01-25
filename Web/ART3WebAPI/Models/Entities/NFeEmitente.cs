using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class NFeEmitente
	{
		public int id { get; set; }
		public int id_boleto_cedente { get; set; }
		public byte st_ativo { get; set; }
		public string apelido { get; set; }
		public string cnpj { get; set; }
		public string razao_social { get; set; }
		public string endereco { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string bairro { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
		public string cep { get; set; }
		public byte NFe_st_emitente_padrao { get; set; }
		public int NFe_serie_NF { get; set; }
		public int NFe_numero_NF { get; set; }
		public string NFe_T1_servidor_BD { get; set; }
		public string NFe_T1_nome_BD { get; set; }
		public string NFe_T1_usuario_BD { get; set; }
		public string NFe_T1_senha_BD { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public DateTime dt_hr_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public byte st_habilitado_ctrl_estoque { get; set; }
		public int ordem { get; set; }
		public string texto_fixo_especifico { get; set; }
		public List<NFeEmitenteCfgDanfe> listaCfgDanfe { get; set; } = new List<NFeEmitenteCfgDanfe>();
	}

	public class NFeEmitenteCfgDanfe
	{
		public int id { get; set; }
		public int id_nfe_emitente { get; set; }
		public byte min_tamanho_serie_NFe { get; set; }
		public byte min_tamanho_numero_NFe { get; set; }
		public string convencao_nome_arq_pdf_danfe { get; set; }
		public string diretorio_pdf_danfe { get; set; }
		public string convencao_nome_arq_xml_nfe { get; set; }
		public string diretorio_xml_nfe { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public int ordenacao { get; set; }
		public string observacao { get; set; }
	}
}