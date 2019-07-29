#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class Loja
	{
		#region [ Getters / Setters ]

		private String _loja;
		public String loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private String _cnpj;
		public String cnpj
		{
			get { return _cnpj; }
			set { _cnpj = value; }
		}

		private String _ie;
		public String ie
		{
			get { return _ie; }
			set { _ie = value; }
		}

		private String _nome;
		public String nome
		{
			get { return _nome; }
			set { _nome = value; }
		}

		private String _razao_social;
		public String razao_social
		{
			get { return _razao_social; }
			set { _razao_social = value; }
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

		private String _ddd;
		public String ddd
		{
			get { return _ddd; }
			set { _ddd = value; }
		}

		private String _telefone;
		public String telefone
		{
			get { return _telefone; }
			set { _telefone = value; }
		}

		private String _fax;
		public String fax
		{
			get { return _fax; }
			set { _fax = value; }
		}

		private double _comissao_indicacao;
		public double comissao_indicacao
		{
			get { return _comissao_indicacao; }
			set { _comissao_indicacao = value; }
		}

		private double _PercMaxSenhaDesconto;
		public double percMaxSenhaDesconto
		{
			get { return _PercMaxSenhaDesconto; }
			set { _PercMaxSenhaDesconto = value; }
		}

		private byte _id_plano_contas_empresa;
		public byte id_plano_contas_empresa
		{
			get { return _id_plano_contas_empresa; }
			set { _id_plano_contas_empresa = value; }
		}

		private short _id_plano_contas_grupo;
		public short id_plano_contas_grupo
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

		#endregion
	}
}
