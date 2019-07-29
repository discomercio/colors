#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace ADM2
{
	public class Parametro
	{
		#region [ Getters/Setters ]
		private String _id;
		public String id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _campo_inteiro;
		public int campo_inteiro
		{
			get { return _campo_inteiro; }
			set { _campo_inteiro = value; }
		}

		private decimal _campo_monetario;
		public decimal campo_monetario
		{
			get { return _campo_monetario; }
			set { _campo_monetario = value; }
		}

		private double _campo_real;
		public double campo_real
		{
			get { return _campo_real; }
			set { _campo_real = value; }
		}

		private DateTime _campo_data;
		public DateTime campo_data
		{
			get { return _campo_data; }
			set { _campo_data = value; }
		}

		private String _campo_texto;
		public String campo_texto
		{
			get { return _campo_texto; }
			set { _campo_texto = value; }
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
		#endregion
	}
}
