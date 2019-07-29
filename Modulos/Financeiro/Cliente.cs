#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class Cliente
	{
		#region [ Getters / Setters ]

		private String _id;
		public String id
		{
			get { return _id; }
			set { _id = value; }
		}

		private String _cnpj_cpf;
		public String cnpj_cpf
		{
			get { return _cnpj_cpf; }
			set { _cnpj_cpf = value; }
		}

		private String _tipo;
		public String tipo
		{
			get { return _tipo; }
			set { _tipo = value; }
		}

		private String _ie;
		public String ie
		{
			get { return _ie; }
			set { _ie = value; }
		}

		private String _rg;
		public String rg
		{
			get { return _rg; }
			set { _rg = value; }
		}

		private String _nome;
		public String nome
		{
			get { return _nome; }
			set { _nome = value; }
		}

		private String _sexo;
		public String sexo
		{
			get { return _sexo; }
			set { _sexo = value; }
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

		private String _ddd_res;
		public String ddd_res
		{
			get { return _ddd_res; }
			set { _ddd_res = value; }
		}

		private String _tel_res;
		public String tel_res
		{
			get { return _tel_res; }
			set { _tel_res = value; }
		}

		private String _ddd_com;
		public String ddd_com
		{
			get { return _ddd_com; }
			set { _ddd_com = value; }
		}

		private String _tel_com;
		public String tel_com
		{
			get { return _tel_com; }
			set { _tel_com = value; }
		}

		private String _ramal_com;
		public String ramal_com
		{
			get { return _ramal_com; }
			set { _ramal_com = value; }
		}

		private String _contato;
		public String contato
		{
			get { return _contato; }
			set { _contato = value; }
		}

		private DateTime _dt_nasc;
		public DateTime dt_nasc
		{
			get { return _dt_nasc; }
			set { _dt_nasc = value; }
		}

		private String _filiacao;
		public String filiacao
		{
			get { return _filiacao; }
			set { _filiacao = value; }
		}

		private String _obs_crediticias;
		public String obs_crediticias
		{
			get { return _obs_crediticias; }
			set { _obs_crediticias = value; }
		}

		private String _midia;
		public String midia
		{
			get { return _midia; }
			set { _midia = value; }
		}

		private String _email;
		public String email
		{
			get { return _email; }
			set { _email = value; }
		}

		private String _email_opcoes;
		public String email_opcoes
		{
			get { return _email_opcoes; }
			set { _email_opcoes = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
		}

		private DateTime _dt_ult_atualizacao;
		public DateTime dt_ult_atualizacao
		{
			get { return _dt_ult_atualizacao; }
			set { _dt_ult_atualizacao = value; }
		}

		private String _SocMaj_Nome;
		public String socMaj_Nome
		{
			get { return _SocMaj_Nome; }
			set { _SocMaj_Nome = value; }
		}

		private String _SocMaj_CPF;
		public String socMaj_CPF
		{
			get { return _SocMaj_CPF; }
			set { _SocMaj_CPF = value; }
		}

		private String _SocMaj_banco;
		public String socMaj_banco
		{
			get { return _SocMaj_banco; }
			set { _SocMaj_banco = value; }
		}

		private String _SocMaj_agencia;
		public String socMaj_agencia
		{
			get { return _SocMaj_agencia; }
			set { _SocMaj_agencia = value; }
		}

		private String _SocMaj_conta;
		public String socMaj_conta
		{
			get { return _SocMaj_conta; }
			set { _SocMaj_conta = value; }
		}

		private String _SocMaj_ddd;
		public String socMaj_ddd
		{
			get { return _SocMaj_ddd; }
			set { _SocMaj_ddd = value; }
		}

		private String _SocMaj_telefone;
		public String socMaj_telefone
		{
			get { return _SocMaj_telefone; }
			set { _SocMaj_telefone = value; }
		}

		private String _SocMaj_contato;
		public String socMaj_contato
		{
			get { return _SocMaj_contato; }
			set { _SocMaj_contato = value; }
		}

		private String _usuario_cadastro;
		public String usuario_cadastro
		{
			get { return _usuario_cadastro; }
			set { _usuario_cadastro = value; }
		}

		private String _usuario_ult_atualizacao;
		public String usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}

		private String _indicador;
		public String indicador
		{
			get { return _indicador; }
			set { _indicador = value; }
		}

		#endregion
	}
}
