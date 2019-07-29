using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Financeiro
{
	class NFeEmitente
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_boleto_cedente;
		public int id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
		}

		private byte _st_ativo;
		public byte st_ativo
		{
			get { return _st_ativo; }
			set { _st_ativo = value; }
		}

		private string _apelido;
		public string apelido
		{
			get { return _apelido; }
			set { _apelido = value; }
		}

		private string _cnpj;
		public string cnpj
		{
			get { return _cnpj; }
			set { _cnpj = value; }
		}

		private string _razao_social;
		public string razao_social
		{
			get { return _razao_social; }
			set { _razao_social = value; }
		}

		private string _endereco;
		public string endereco
		{
			get { return _endereco; }
			set { _endereco = value; }
		}

		private string _endereco_numero;
		public string endereco_numero
		{
			get { return _endereco_numero; }
			set { _endereco_numero = value; }
		}

		private string _endereco_complemento;
		public string endereco_complemento
		{
			get { return _endereco_complemento; }
			set { _endereco_complemento = value; }
		}

		private string _bairro;
		public string bairro
		{
			get { return _bairro; }
			set { _bairro = value; }
		}

		private string _cidade;
		public string cidade
		{
			get { return _cidade; }
			set { _cidade = value; }
		}

		private string _uf;
		public string uf
		{
			get { return _uf; }
			set { _uf = value; }
		}

		private string _cep;
		public string cep
		{
			get { return _cep; }
			set { _cep = value; }
		}

		private byte _NFe_st_emitente_padrao;
		public byte NFe_st_emitente_padrao
		{
			get { return _NFe_st_emitente_padrao; }
			set { _NFe_st_emitente_padrao = value; }
		}

		private int _NFe_serie_NF;
		public int NFe_serie_NF
		{
			get { return _NFe_serie_NF; }
			set { _NFe_serie_NF = value; }
		}

		private int _NFe_numero_NF;
		public int NFe_numero_NF
		{
			get { return _NFe_numero_NF; }
			set { _NFe_numero_NF = value; }
		}

		private string _NFe_T1_servidor_BD;
		public string NFe_T1_servidor_BD
		{
			get { return _NFe_T1_servidor_BD; }
			set { _NFe_T1_servidor_BD = value; }
		}

		private string _NFe_T1_nome_BD;
		public string NFe_T1_nome_BD
		{
			get { return _NFe_T1_nome_BD; }
			set { _NFe_T1_nome_BD = value; }
		}

		private string _NFe_T1_usuario_BD;
		public string NFe_T1_usuario_BD
		{
			get { return _NFe_T1_usuario_BD; }
			set { _NFe_T1_usuario_BD = value; }
		}

		private string _NFe_T1_senha_BD;
		public string NFe_T1_senha_BD
		{
			get { return _NFe_T1_senha_BD; }
			set { _NFe_T1_senha_BD = value; }
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

		private string _usuario_cadastro;
		public string usuario_cadastro
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

		private string _usuario_ult_atualizacao;
		public string usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}
	}
}
