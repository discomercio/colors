#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class BoletoCedente
	{
		#region [ Getters / Setters ]

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

		private byte _st_ativo;
		public byte st_ativo
		{
			get { return _st_ativo; }
			set { _st_ativo = value; }
		}

		private int _nsu_arq_remessa;
		public int nsu_arq_remessa
		{
			get { return _nsu_arq_remessa; }
			set { _nsu_arq_remessa = value; }
		}

		private String _codigo_empresa;
		public String codigo_empresa
		{
			get { return _codigo_empresa; }
			set { _codigo_empresa = value; }
		}

		private String _nome_empresa;
		public String nome_empresa
		{
			get { return _nome_empresa; }
			set { _nome_empresa = value; }
		}

		private String _num_banco;
		public String num_banco
		{
			get { return _num_banco; }
			set { _num_banco = value; }
		}

		private String _nome_banco;
		public String nome_banco
		{
			get { return _nome_banco; }
			set { _nome_banco = value; }
		}

		private String _agencia;
		public String agencia
		{
			get { return _agencia; }
			set { _agencia = value; }
		}

		private String _digito_agencia;
		public String digito_agencia
		{
			get { return _digito_agencia; }
			set { _digito_agencia = value; }
		}

		private String _conta;
		public String conta
		{
			get { return _conta; }
			set { _conta = value; }
		}

		private String _digito_conta;
		public String digito_conta
		{
			get { return _digito_conta; }
			set { _digito_conta = value; }
		}

		private String _carteira;
		public String carteira
		{
			get { return _carteira; }
			set { _carteira = value; }
		}

		private double _juros_mora;
		public double juros_mora
		{
			get { return _juros_mora; }
			set { _juros_mora = value; }
		}

		private double _perc_multa;
		public double perc_multa
		{
			get { return _perc_multa; }
			set { _perc_multa = value; }
		}

		private byte _qtde_dias_protestar_apos_padrao;
		public byte qtde_dias_protestar_apos_padrao
		{
			get { return _qtde_dias_protestar_apos_padrao; }
			set { _qtde_dias_protestar_apos_padrao = value; }
		}

		private String _segunda_mensagem_padrao;
		public String segunda_mensagem_padrao
		{
			get { return _segunda_mensagem_padrao; }
			set { _segunda_mensagem_padrao = value; }
		}

		private String _mensagem_1_padrao;
		public String mensagem_1_padrao
		{
			get { return _mensagem_1_padrao; }
			set { _mensagem_1_padrao = value; }
		}

		private String _mensagem_2_padrao;
		public String mensagem_2_padrao
		{
			get { return _mensagem_2_padrao; }
			set { _mensagem_2_padrao = value; }
		}

		private String _mensagem_3_padrao;
		public String mensagem_3_padrao
		{
			get { return _mensagem_3_padrao; }
			set { _mensagem_3_padrao = value; }
		}

		private String _mensagem_4_padrao;
		public String mensagem_4_padrao
		{
			get { return _mensagem_4_padrao; }
			set { _mensagem_4_padrao = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
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

		private String _usuario_ult_atualizacao;
		public String usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}

		private String _endereco;
		public String endereco
		{
			get { return _endereco; }
			set { _endereco = value; }
		}

		private String _endereco_numero;
		public String endereco_numero
		{
			get { return _endereco_numero; }
			set { _endereco_numero = value; }
		}

		private String _endereco_complemento;
		public String endereco_complemento
		{
			get { return _endereco_complemento; }
			set { _endereco_complemento = value; }
		}

		private String _bairro;
		public String bairro
		{
			get { return _bairro; }
			set { _bairro = value; }
		}

		private String _cidade;
		public String cidade
		{
			get { return _cidade; }
			set { _cidade = value; }
		}

		private String _uf;
		public String uf
		{
			get { return _uf; }
			set { _uf = value; }
		}

		private String _cep;
		public String cep
		{
			get { return _cep; }
			set { _cep = value; }
		}

		private byte _st_boleto_cedente_padrao;
		public byte st_boleto_cedente_padrao
		{
			get { return _st_boleto_cedente_padrao; }
			set { _st_boleto_cedente_padrao = value; }
		}

		private String _apelido;
		public String apelido
		{
			get { return _apelido; }
			set { _apelido = value; }
		}

		private String _loja_default_boleto_plano_contas;
		public String loja_default_boleto_plano_contas
		{
			get { return _loja_default_boleto_plano_contas; }
			set { _loja_default_boleto_plano_contas = value; }
		}

        private string _cnpj;
        public string cnpj
        {
            get { return _cnpj; }
            set { _cnpj = value; }
        }

        private byte _st_participante_serasa_reciprocidade;
		public byte st_participante_serasa_reciprocidade
		{
			get { return _st_participante_serasa_reciprocidade; }
			set { _st_participante_serasa_reciprocidade = value; }
		}
		#endregion
	}
}
