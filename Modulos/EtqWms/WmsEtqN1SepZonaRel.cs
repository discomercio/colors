#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace EtqWms
{
	#region [ WmsEtqN1SepZonaRel ]
	class WmsEtqN1SepZonaRel
	{
		#region [ Getters / Setters ]
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
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

		private DateTime _dt_emissao;
		public DateTime dt_emissao
		{
			get { return _dt_emissao; }
			set { _dt_emissao = value; }
		}

		private DateTime _dt_hr_emissao;
		public DateTime dt_hr_emissao
		{
			get { return _dt_hr_emissao; }
			set { _dt_hr_emissao = value; }
		}

		private string _usuario;
		public string usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private string _filtro_dt_inicio;
		public string filtro_dt_inicio
		{
			get { return _filtro_dt_inicio; }
			set { _filtro_dt_inicio = value; }
		}

		private string _filtro_dt_termino;
		public string filtro_dt_termino
		{
			get { return _filtro_dt_termino; }
			set { _filtro_dt_termino = value; }
		}

		private string _filtro_NFe_emitida;
		public string filtro_NFe_emitida
		{
			get { return _filtro_NFe_emitida; }
			set { _filtro_NFe_emitida = value; }
		}

		private string _filtro_transportadora;
		public string filtro_transportadora
		{
			get { return _filtro_transportadora; }
			set { _filtro_transportadora = value; }
		}

		private string _filtro_qtde_max_pedidos;
		public string filtro_qtde_max_pedidos
		{
			get { return _filtro_qtde_max_pedidos; }
			set { _filtro_qtde_max_pedidos = value; }
		}

		private string _filtro_qtde_disponivel_pedidos;
		public string filtro_qtde_disponivel_pedidos
		{
			get { return _filtro_qtde_disponivel_pedidos; }
			set { _filtro_qtde_disponivel_pedidos = value; }
		}

		private string _lista_zonas_cadastradas;
		public string lista_zonas_cadastradas
		{
			get { return _lista_zonas_cadastradas; }
			set { _lista_zonas_cadastradas = value; }
		}

		private byte _etiqueta_impressao_status;
		public byte etiqueta_impressao_status
		{
			get { return _etiqueta_impressao_status; }
			set { _etiqueta_impressao_status = value; }
		}

		private int _etiqueta_impressao_qtde_impressoes;
		public int etiqueta_impressao_qtde_impressoes
		{
			get { return _etiqueta_impressao_qtde_impressoes; }
			set { _etiqueta_impressao_qtde_impressoes = value; }
		}

		private DateTime _etiqueta_impressao_primeira_vez_data;
		public DateTime etiqueta_impressao_primeira_vez_data
		{
			get { return _etiqueta_impressao_primeira_vez_data; }
			set { _etiqueta_impressao_primeira_vez_data = value; }
		}

		private DateTime _etiqueta_impressao_primeira_vez_data_hora;
		public DateTime etiqueta_impressao_primeira_vez_data_hora
		{
			get { return _etiqueta_impressao_primeira_vez_data_hora; }
			set { _etiqueta_impressao_primeira_vez_data_hora = value; }
		}

		private string _etiqueta_impressao_primeira_vez_usuario;
		public string etiqueta_impressao_primeira_vez_usuario
		{
			get { return _etiqueta_impressao_primeira_vez_usuario; }
			set { _etiqueta_impressao_primeira_vez_usuario = value; }
		}

		private DateTime _etiqueta_impressao_ultima_vez_data;
		public DateTime etiqueta_impressao_ultima_vez_data
		{
			get { return _etiqueta_impressao_ultima_vez_data; }
			set { _etiqueta_impressao_ultima_vez_data = value; }
		}

		private DateTime _etiqueta_impressao_ultima_vez_data_hora;
		public DateTime etiqueta_impressao_ultima_vez_data_hora
		{
			get { return _etiqueta_impressao_ultima_vez_data_hora; }
			set { _etiqueta_impressao_ultima_vez_data_hora = value; }
		}

		private string _etiqueta_impressao_ultima_vez_usuario;
		public string etiqueta_impressao_ultima_vez_usuario
		{
			get { return _etiqueta_impressao_ultima_vez_usuario; }
			set { _etiqueta_impressao_ultima_vez_usuario = value; }
		}
		#endregion
	}
	#endregion
}
