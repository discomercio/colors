namespace Financeiro
{
	partial class FCobrancaMain
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

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FCobrancaMain));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.gboxModuloCobranca = new System.Windows.Forms.GroupBox();
			this.btnBoletoConsulta = new System.Windows.Forms.Button();
			this.btnAdministracaoCarteiraEmAtraso = new System.Windows.Forms.Button();
			this.btnFluxoCaixaConsulta = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxModuloCobranca.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxModuloCobranca);
			this.pnCampos.Controls.Add(this.lblTitulo);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// lblTitulo
			// 
			this.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTitulo.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitulo.ForeColor = System.Drawing.Color.Black;
			this.lblTitulo.Image = ((System.Drawing.Image)(resources.GetObject("lblTitulo.Image")));
			this.lblTitulo.Location = new System.Drawing.Point(0, 0);
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.Size = new System.Drawing.Size(1014, 40);
			this.lblTitulo.TabIndex = 2;
			this.lblTitulo.Text = "Módulo de Cobrança";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// gboxModuloCobranca
			// 
			this.gboxModuloCobranca.Controls.Add(this.btnFluxoCaixaConsulta);
			this.gboxModuloCobranca.Controls.Add(this.btnBoletoConsulta);
			this.gboxModuloCobranca.Controls.Add(this.btnAdministracaoCarteiraEmAtraso);
			this.gboxModuloCobranca.Location = new System.Drawing.Point(39, 46);
			this.gboxModuloCobranca.Name = "gboxModuloCobranca";
			this.gboxModuloCobranca.Size = new System.Drawing.Size(428, 179);
			this.gboxModuloCobranca.TabIndex = 3;
			this.gboxModuloCobranca.TabStop = false;
			// 
			// btnBoletoConsulta
			// 
			this.btnBoletoConsulta.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoConsulta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoConsulta.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoConsulta.Image")));
			this.btnBoletoConsulta.Location = new System.Drawing.Point(14, 72);
			this.btnBoletoConsulta.Name = "btnBoletoConsulta";
			this.btnBoletoConsulta.Size = new System.Drawing.Size(400, 38);
			this.btnBoletoConsulta.TabIndex = 1;
			this.btnBoletoConsulta.Text = "Boleto: Consulta";
			this.btnBoletoConsulta.UseVisualStyleBackColor = true;
			this.btnBoletoConsulta.Click += new System.EventHandler(this.btnBoletoConsulta_Click);
			// 
			// btnAdministracaoCarteiraEmAtraso
			// 
			this.btnAdministracaoCarteiraEmAtraso.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnAdministracaoCarteiraEmAtraso.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAdministracaoCarteiraEmAtraso.ForeColor = System.Drawing.Color.Black;
			this.btnAdministracaoCarteiraEmAtraso.Image = ((System.Drawing.Image)(resources.GetObject("btnAdministracaoCarteiraEmAtraso.Image")));
			this.btnAdministracaoCarteiraEmAtraso.Location = new System.Drawing.Point(14, 16);
			this.btnAdministracaoCarteiraEmAtraso.Name = "btnAdministracaoCarteiraEmAtraso";
			this.btnAdministracaoCarteiraEmAtraso.Size = new System.Drawing.Size(400, 38);
			this.btnAdministracaoCarteiraEmAtraso.TabIndex = 0;
			this.btnAdministracaoCarteiraEmAtraso.Text = "Administração da Carteira em Atraso";
			this.btnAdministracaoCarteiraEmAtraso.UseVisualStyleBackColor = true;
			this.btnAdministracaoCarteiraEmAtraso.Click += new System.EventHandler(this.btnAdministracaoCarteiraEmAtraso_Click);
			// 
			// btnFluxoCaixaConsulta
			// 
			this.btnFluxoCaixaConsulta.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaConsulta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaConsulta.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaConsulta.Image")));
			this.btnFluxoCaixaConsulta.Location = new System.Drawing.Point(14, 128);
			this.btnFluxoCaixaConsulta.Name = "btnFluxoCaixaConsulta";
			this.btnFluxoCaixaConsulta.Size = new System.Drawing.Size(400, 38);
			this.btnFluxoCaixaConsulta.TabIndex = 2;
			this.btnFluxoCaixaConsulta.Text = "Fluxo de Caixa: Consulta";
			this.btnFluxoCaixaConsulta.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaConsulta.Click += new System.EventHandler(this.btnFluxoCaixaConsulta_Click);
			// 
			// FCobrancaMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FCobrancaMain";
			this.Text = "Artven - Financeiro  -  1.11 - XX.XXX.2009";
			this.Load += new System.EventHandler(this.FCobrancaMain_Load);
			this.Shown += new System.EventHandler(this.FCobrancaMain_Shown);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FCobrancaMain_FormClosing);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxModuloCobranca.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.GroupBox gboxModuloCobranca;
		private System.Windows.Forms.Button btnAdministracaoCarteiraEmAtraso;
		private System.Windows.Forms.Button btnBoletoConsulta;
		private System.Windows.Forms.Button btnFluxoCaixaConsulta;
	}
}
