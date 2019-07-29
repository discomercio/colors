#region [ using ]
using System;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	partial class PainelSobre : Form
	{
		#region [ Construtor ]
		public PainelSobre()
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			#region [ Calcula data/hora do build ]
			Version versao = Assembly.GetExecutingAssembly().GetName().Version;
			DateTime dtBuildVersao = new DateTime(2000, 1, 1).AddDays(versao.Build).AddSeconds(versao.MinorRevision * 2);
			DaylightTime dlt = TimeZone.CurrentTimeZone.GetDaylightChanges(dtBuildVersao.Year);
			if ((dlt.Start <= dtBuildVersao) && (dtBuildVersao < dlt.End)) dtBuildVersao += dlt.Delta;
			#endregion

			#region [ Preenche os dados para exibição ]
			this.Text = String.Format("Informações sobre: {0}", AssemblyTitle);
			this.labelProductName.Text = Global.Cte.Aplicativo.NOME_SISTEMA;
			this.labelVersion.Text = "Versão: " + Global.Cte.Aplicativo.VERSAO;
			this.labelCopyright.Text = "Build: " + String.Format("{0}", versao.ToString()) + " (" + dtBuildVersao.ToString() + ")";
			this.labelCompanyName.Text = AssemblyCopyright;
			this.textBoxDescription.Text = Global.Cte.Aplicativo.M_DESCRICAO;
			#endregion
		}
		#endregion

		#region Assembly Attribute Accessors

		public string AssemblyTitle
		{
			get
			{
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
				if (attributes.Length > 0)
				{
					AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
					if (titleAttribute.Title != "")
					{
						return titleAttribute.Title;
					}
				}
				return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
			}
		}

		public string AssemblyVersion
		{
			get
			{
				return Assembly.GetExecutingAssembly().GetName().Version.ToString();
			}
		}

		public string AssemblyDescription
		{
			get
			{
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
				if (attributes.Length == 0)
				{
					return "";
				}
				return ((AssemblyDescriptionAttribute)attributes[0]).Description;
			}
		}

		public string AssemblyProduct
		{
			get
			{
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
				if (attributes.Length == 0)
				{
					return "";
				}
				return ((AssemblyProductAttribute)attributes[0]).Product;
			}
		}

		public string AssemblyCopyright
		{
			get
			{
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
				if (attributes.Length == 0)
				{
					return "";
				}
				return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
			}
		}

		public string AssemblyCompany
		{
			get
			{
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
				if (attributes.Length == 0)
				{
					return "";
				}
				return ((AssemblyCompanyAttribute)attributes[0]).Company;
			}
		}
		#endregion

		#region [ Eventos ]
		private void okButton_Click(object sender, EventArgs e)
		{
			Close();
		}
		#endregion
	}
}
