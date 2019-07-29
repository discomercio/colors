#region [ MyRegion ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class LancamentoFluxoCaixa
	{
		#region [ Getters/Setters ]

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private byte _id_conta_corrente;
		public byte id_conta_corrente
		{
			get { return _id_conta_corrente; }
			set { _id_conta_corrente = value; }
		}

		private byte _id_plano_contas_empresa;
		public byte id_plano_contas_empresa
		{
			get { return _id_plano_contas_empresa; }
			set { _id_plano_contas_empresa = value; }
		}

		private int _id_plano_contas_grupo;
		public int id_plano_contas_grupo
		{
			get { return _id_plano_contas_grupo; }
			set { _id_plano_contas_grupo = value; }
		}

		private int _id_plano_contas_conta;
		public int id_plano_contas_conta
		{
			get { return _id_plano_contas_conta; }
			set { _id_plano_contas_conta = value; }
		}

		private char _natureza;
		public char natureza
		{
			get { return _natureza; }
			set { _natureza = value; }
		}

		private byte _st_sem_efeito;
		public byte st_sem_efeito
		{
			get { return _st_sem_efeito; }
			set { _st_sem_efeito = value; }
		}

		private DateTime _dt_competencia;
		public DateTime dt_competencia
		{
			get { return _dt_competencia; }
			set { _dt_competencia = value; }
		}

        private DateTime _dt_mes_competencia;
        public DateTime dt_mes_competencia
        {
            get { return _dt_mes_competencia; }
            set { _dt_mes_competencia = value; }
        }

        private decimal _valor;
		public decimal valor
		{
			get { return _valor; }
			set { _valor = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private int _ctrl_pagto_id_parcela;
		public int ctrl_pagto_id_parcela
		{
			get { return _ctrl_pagto_id_parcela; }
			set { _ctrl_pagto_id_parcela = value; }
		}

		private byte _ctrl_pagto_modulo;
		public byte ctrl_pagto_modulo
		{
			get { return _ctrl_pagto_modulo; }
			set { _ctrl_pagto_modulo = value; }
		}

		private byte _ctrl_pagto_status;
		public byte ctrl_pagto_status
		{
			get { return _ctrl_pagto_status; }
			set { _ctrl_pagto_status = value; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private String _cnpj_cpf;
		public String cnpj_cpf
		{
			get { return _cnpj_cpf; }
			set { _cnpj_cpf = value; }
		}

		private int _numero_NF;
		public int numero_NF
		{
			get { return _numero_NF; }
			set { _numero_NF = value; }
		}

		private char _tipo_cadastro;
		public char tipo_cadastro
		{
			get { return _tipo_cadastro; }
			set { _tipo_cadastro = value; }
		}

		private char _editado_manual;
		public char editado_manual
		{
			get { return _editado_manual; }
			set { _editado_manual = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
		}

		private DateTime _dt_hr_cadastro;
		public DateTime dt_hr_cadastro
		{
			get { return _dt_hr_cadastro; }
			set { _dt_hr_cadastro = value; }
		}

		private String _usuario_cadastro;
		public String usuario_cadastro
		{
			get { return _usuario_cadastro; }
			set { _usuario_cadastro = value; }
		}

		private DateTime _dt_ult_atualizacao;
		public DateTime dt_ult_atualizacao
		{
			get { return _dt_ult_atualizacao; }
			set { _dt_ult_atualizacao = value; }
		}

		private DateTime _dt_hr_ult_atualizacao;
		public DateTime dt_hr_ult_atualizacao
		{
			get { return _dt_hr_ult_atualizacao; }
			set { _dt_hr_ult_atualizacao = value; }
		}

		private String _usuario_ult_atualizacao;
		public String usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}

		private byte _st_confirmacao_pendente;
		public byte st_confirmacao_pendente
		{
			get { return _st_confirmacao_pendente; }
			set { _st_confirmacao_pendente = value; }
		}

		private byte _st_boleto_pago_cheque;
		public byte st_boleto_pago_cheque
		{
			get { return _st_boleto_pago_cheque; }
			set { _st_boleto_pago_cheque = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_pago_cheque;
		public DateTime dt_ocorrencia_banco_boleto_pago_cheque
		{
			get { return _dt_ocorrencia_banco_boleto_pago_cheque; }
			set { _dt_ocorrencia_banco_boleto_pago_cheque = value; }
		}

		private byte _st_boleto_ocorrencia_17;
		public byte st_boleto_ocorrencia_17
		{
			get { return _st_boleto_ocorrencia_17; }
			set { _st_boleto_ocorrencia_17 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_17;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_17
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_17; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_17 = value; }
		}

		private byte _st_boleto_ocorrencia_15;
		public byte st_boleto_ocorrencia_15
		{
			get { return _st_boleto_ocorrencia_15; }
			set { _st_boleto_ocorrencia_15 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_15;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_15
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_15; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_15 = value; }
		}

		private byte _st_boleto_ocorrencia_23;
		public byte st_boleto_ocorrencia_23
		{
			get { return _st_boleto_ocorrencia_23; }
			set { _st_boleto_ocorrencia_23 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_23;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_23
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_23; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_23 = value; }
		}

		private byte _st_boleto_ocorrencia_34;
		public byte st_boleto_ocorrencia_34
		{
			get { return _st_boleto_ocorrencia_34; }
			set { _st_boleto_ocorrencia_34 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_34;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_34
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_34; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_34 = value; }
		}

		private byte _st_boleto_baixado;
		public byte st_boleto_baixado
		{
			get { return _st_boleto_baixado; }
			set { _st_boleto_baixado = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_baixado;
		public DateTime dt_ocorrencia_banco_boleto_baixado
		{
			get { return _dt_ocorrencia_banco_boleto_baixado; }
			set { _dt_ocorrencia_banco_boleto_baixado = value; }
		}

		private int _id_boleto_cedente;
		public int id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
		}
		#endregion

	}
}
