#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class FinLog
	{
		#region [ Getters/Setters ]
		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private String _usuario;
		public String usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private String _operacao;
		public String operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
		}

		private byte _st_sem_efeito;
		public byte st_sem_efeito
		{
			get { return _st_sem_efeito; }
			set { _st_sem_efeito = value; }
		}

		private byte _ctrl_pagto_status;
		public byte ctrl_pagto_status
		{
			get { return _ctrl_pagto_status; }
			set { _ctrl_pagto_status = value; }
		}

		private char _natureza;
		public char natureza
		{
			get { return _natureza; }
			set { _natureza = value; }
		}

		private char _tipo_cadastro;
		public char tipo_cadastro
		{
			get { return _tipo_cadastro; }
			set { _tipo_cadastro = value; }
		}

		private String _fin_modulo;
		public String fin_modulo
		{
			get { return _fin_modulo; }
			set { _fin_modulo = value; }
		}

		private byte _cod_tabela_origem;
		public byte cod_tabela_origem
		{
			get { return _cod_tabela_origem; }
			set { _cod_tabela_origem = value; }
		}

		private int _id_registro_origem;
		public int id_registro_origem
		{
			get { return _id_registro_origem; }
			set { _id_registro_origem = value; }
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

		private int _id_boleto_cedente;
		public int id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
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

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}
		#endregion
	}
}
