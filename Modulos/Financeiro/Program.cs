using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace Financeiro
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			// Verifica se já existe uma instância em execução 
			if (Global.haOutraInstanciaEmExecucao())
			{
				Console.Beep();
				Application.Exit();
				return;
			}

			Color? backColor = Global.getBackColorFromAppConfig();
			if (backColor != null) Global.BackColorPainelPadrao = (Color)backColor;

			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new FMain());
		}
	}
}
