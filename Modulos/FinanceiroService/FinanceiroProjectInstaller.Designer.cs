namespace FinanceiroService
{
	partial class FinanceiroProjectInstaller
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.FinanceiroServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
			this.FinanceiroServiceInstaller = new System.ServiceProcess.ServiceInstaller();
			// 
			// FinanceiroServiceProcessInstaller
			// 
			this.FinanceiroServiceProcessInstaller.Password = null;
			this.FinanceiroServiceProcessInstaller.Username = null;
			// 
			// FinanceiroServiceInstaller
			// 
			this.FinanceiroServiceInstaller.Description = "Serviço para execução de rotinas automáticas do Financeiro";
			this.FinanceiroServiceInstaller.DisplayName = "Financeiro";
			this.FinanceiroServiceInstaller.ServiceName = "FinanceiroService";
			// 
			// FinanceiroProjectInstaller
			// 
			this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.FinanceiroServiceProcessInstaller,
            this.FinanceiroServiceInstaller});

		}

		#endregion

		private System.ServiceProcess.ServiceProcessInstaller FinanceiroServiceProcessInstaller;
		private System.ServiceProcess.ServiceInstaller FinanceiroServiceInstaller;
	}
}