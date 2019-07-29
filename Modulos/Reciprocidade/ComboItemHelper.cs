#region [ using ]
using System;
#endregion

namespace Reciprocidade
{
	class ComboItemHelper
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}
		private DateTime _dataHora;
		public DateTime dataHora
		{
			get { return _dataHora; }
			set { _dataHora = value; }
		}

		public ComboItemHelper(int id, DateTime dataHora)
		{
			_id = id;
			_dataHora = dataHora;
		}

		public override string ToString()
		{
			return _dataHora.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador) + "  " + _dataHora.ToString(Global.Cte.DataHora.FmtHhMmComSeparador) + "  (id: " + Global.formataInteiro(_id) + ")";
		}
	}
}
