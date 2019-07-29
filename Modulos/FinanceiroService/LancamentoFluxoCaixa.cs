using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceiroService
{
	#region [ LancamentoFluxoCaixa ]
	class LancamentoFluxoCaixa
	{
		public int id { get; set; }
		public byte id_conta_corrente { get; set; }
		public byte id_plano_contas_empresa { get; set; }
		public int id_plano_contas_grupo { get; set; }
		public int id_plano_contas_conta { get; set; }
		public char natureza { get; set; }
		public byte st_sem_efeito { get; set; }
		public DateTime dt_competencia { get; set; }
		public decimal valor { get; set; }
		public string descricao { get; set; }
		public int ctrl_pagto_id_parcela { get; set; }
		public byte ctrl_pagto_modulo { get; set; }
		public byte ctrl_pagto_status { get; set; }
		public string id_cliente { get; set; }
		public string cnpj_cpf { get; set; }
		public char tipo_cadastro { get; set; }
		public char editado_manual { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public DateTime dt_hr_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public byte st_confirmacao_pendente { get; set; }
		public byte st_boleto_pago_cheque { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_pago_cheque { get; set; }
		public byte st_boleto_ocorrencia_17 { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_17 { get; set; }
		public byte st_boleto_ocorrencia_15 { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_15 { get; set; }
		public byte st_boleto_ocorrencia_23 { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_23 { get; set; }
		public byte st_boleto_ocorrencia_34 { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_34 { get; set; }
		public byte st_boleto_baixado { get; set; }
		public DateTime dt_ocorrencia_banco_boleto_baixado { get; set; }
		public int id_boleto_cedente { get; set; }
		public byte tamanho_cnpj_cpf { get; set; }
		public byte ctrl_pagto_id_ambiente_origem { get; set; }
	}
	#endregion

	#region [ LancamentoFluxoCaixaInsertDevidoBoletoEC ]
	public class LancamentoFluxoCaixaInsertDevidoBoletoEC
	{
		public int id { get; set; }
		public byte id_conta_corrente { get; set; }
		public byte id_plano_contas_empresa { get; set; }
		public int id_plano_contas_grupo { get; set; }
		public int id_plano_contas_conta { get; set; }
		public char natureza { get; set; }
		public DateTime dt_competencia { get; set; }
		public decimal valor { get; set; }
		public string descricao { get; set; }
		public int ctrl_pagto_id_parcela { get; set; }
		public byte ctrl_pagto_modulo { get; set; }
		public string id_cliente { get; set; }
		public string cnpj_cpf { get; set; }
		public char tipo_cadastro { get; set; }
		public char editado_manual { get; set; }
		public string usuario_cadastro { get; set; }
		public string usuario_ult_atualizacao { get; set; }
	}
	#endregion
}
