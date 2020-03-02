namespace ADM2
{
	partial class FModelo
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FModelo));
            this.pnStatus = new System.Windows.Forms.Panel();
            this.pnMensagem = new System.Windows.Forms.Panel();
            this.lblMensagem = new System.Windows.Forms.Label();
            this.pnHora = new System.Windows.Forms.Panel();
            this.lblHora = new System.Windows.Forms.Label();
            this.pnData = new System.Windows.Forms.Panel();
            this.lblData = new System.Windows.Forms.Label();
            this.menuPrincipal = new System.Windows.Forms.MenuStrip();
            this.menuArquivo = new System.Windows.Forms.ToolStripMenuItem();
            this.menuArquivoFechar = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAjuda = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAjudaSobre = new System.Windows.Forms.ToolStripMenuItem();
            this.pnBotoes = new System.Windows.Forms.Panel();
            this.btnSobre = new System.Windows.Forms.Button();
            this.pbLogo = new System.Windows.Forms.PictureBox();
            this.btnDummy = new System.Windows.Forms.Button();
            this.btnFechar = new System.Windows.Forms.Button();
            this.pnCampos = new System.Windows.Forms.Panel();
            this.tmrRelogio = new System.Windows.Forms.Timer(this.components);
            this.pnStatus.SuspendLayout();
            this.pnMensagem.SuspendLayout();
            this.pnHora.SuspendLayout();
            this.pnData.SuspendLayout();
            this.menuPrincipal.SuspendLayout();
            this.pnBotoes.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // pnStatus
            // 
            this.pnStatus.Controls.Add(this.pnMensagem);
            this.pnStatus.Controls.Add(this.pnHora);
            this.pnStatus.Controls.Add(this.pnData);
            this.pnStatus.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnStatus.Location = new System.Drawing.Point(0, 544);
            this.pnStatus.Name = "pnStatus";
            this.pnStatus.Size = new System.Drawing.Size(1008, 18);
            this.pnStatus.TabIndex = 6;
            // 
            // pnMensagem
            // 
            this.pnMensagem.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnMensagem.Controls.Add(this.lblMensagem);
            this.pnMensagem.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnMensagem.Location = new System.Drawing.Point(115, 0);
            this.pnMensagem.Name = "pnMensagem";
            this.pnMensagem.Size = new System.Drawing.Size(893, 18);
            this.pnMensagem.TabIndex = 6;
            // 
            // lblMensagem
            // 
            this.lblMensagem.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblMensagem.Location = new System.Drawing.Point(0, 0);
            this.lblMensagem.Name = "lblMensagem";
            this.lblMensagem.Size = new System.Drawing.Size(889, 14);
            this.lblMensagem.TabIndex = 0;
            this.lblMensagem.Text = "mensagem";
            this.lblMensagem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnHora
            // 
            this.pnHora.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnHora.Controls.Add(this.lblHora);
            this.pnHora.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnHora.Location = new System.Drawing.Point(64, 0);
            this.pnHora.Name = "pnHora";
            this.pnHora.Size = new System.Drawing.Size(51, 18);
            this.pnHora.TabIndex = 5;
            // 
            // lblHora
            // 
            this.lblHora.AutoSize = true;
            this.lblHora.Location = new System.Drawing.Point(0, 1);
            this.lblHora.Name = "lblHora";
            this.lblHora.Size = new System.Drawing.Size(49, 13);
            this.lblHora.TabIndex = 0;
            this.lblHora.Text = "00:00:00";
            this.lblHora.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnData
            // 
            this.pnData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnData.Controls.Add(this.lblData);
            this.pnData.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnData.Location = new System.Drawing.Point(0, 0);
            this.pnData.Name = "pnData";
            this.pnData.Size = new System.Drawing.Size(64, 18);
            this.pnData.TabIndex = 4;
            // 
            // lblData
            // 
            this.lblData.AutoSize = true;
            this.lblData.Location = new System.Drawing.Point(0, 1);
            this.lblData.Name = "lblData";
            this.lblData.Size = new System.Drawing.Size(61, 13);
            this.lblData.TabIndex = 0;
            this.lblData.Text = "99.99.9999";
            this.lblData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // menuPrincipal
            // 
            this.menuPrincipal.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuArquivo,
            this.menuAjuda});
            this.menuPrincipal.Location = new System.Drawing.Point(0, 0);
            this.menuPrincipal.Name = "menuPrincipal";
            this.menuPrincipal.Size = new System.Drawing.Size(1008, 24);
            this.menuPrincipal.TabIndex = 7;
            this.menuPrincipal.Text = "menuStrip1";
            // 
            // menuArquivo
            // 
            this.menuArquivo.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuArquivoFechar});
            this.menuArquivo.Name = "menuArquivo";
            this.menuArquivo.Size = new System.Drawing.Size(61, 20);
            this.menuArquivo.Text = "&Arquivo";
            // 
            // menuArquivoFechar
            // 
            this.menuArquivoFechar.Name = "menuArquivoFechar";
            this.menuArquivoFechar.Size = new System.Drawing.Size(109, 22);
            this.menuArquivoFechar.Text = "&Fechar";
            this.menuArquivoFechar.Click += new System.EventHandler(this.menuArquivoFechar_Click);
            // 
            // menuAjuda
            // 
            this.menuAjuda.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuAjudaSobre});
            this.menuAjuda.Name = "menuAjuda";
            this.menuAjuda.Size = new System.Drawing.Size(50, 20);
            this.menuAjuda.Text = "A&juda";
            // 
            // menuAjudaSobre
            // 
            this.menuAjudaSobre.Name = "menuAjudaSobre";
            this.menuAjudaSobre.Size = new System.Drawing.Size(113, 22);
            this.menuAjudaSobre.Text = "So&bre...";
            this.menuAjudaSobre.Click += new System.EventHandler(this.menuAjudaSobre_Click);
            // 
            // pnBotoes
            // 
            this.pnBotoes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnBotoes.Controls.Add(this.btnSobre);
            this.pnBotoes.Controls.Add(this.pbLogo);
            this.pnBotoes.Controls.Add(this.btnDummy);
            this.pnBotoes.Controls.Add(this.btnFechar);
            this.pnBotoes.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnBotoes.Location = new System.Drawing.Point(0, 24);
            this.pnBotoes.Name = "pnBotoes";
            this.pnBotoes.Size = new System.Drawing.Size(1008, 55);
            this.pnBotoes.TabIndex = 8;
            // 
            // btnSobre
            // 
            this.btnSobre.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnSobre.Image = ((System.Drawing.Image)(resources.GetObject("btnSobre.Image")));
            this.btnSobre.Location = new System.Drawing.Point(914, 4);
            this.btnSobre.Name = "btnSobre";
            this.btnSobre.Size = new System.Drawing.Size(40, 44);
            this.btnSobre.TabIndex = 7;
            this.btnSobre.TabStop = false;
            this.btnSobre.UseVisualStyleBackColor = true;
            this.btnSobre.Click += new System.EventHandler(this.btnSobre_Click);
            // 
            // pbLogo
            // 
            this.pbLogo.BackColor = System.Drawing.Color.Transparent;
            this.pbLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pbLogo.Image = ((System.Drawing.Image)(resources.GetObject("pbLogo.Image")));
            this.pbLogo.Location = new System.Drawing.Point(0, 0);
            this.pbLogo.Name = "pbLogo";
            this.pbLogo.Size = new System.Drawing.Size(200, 52);
            this.pbLogo.TabIndex = 6;
            this.pbLogo.TabStop = false;
            // 
            // btnDummy
            // 
            this.btnDummy.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnDummy.Location = new System.Drawing.Point(375, 22);
            this.btnDummy.Name = "btnDummy";
            this.btnDummy.Size = new System.Drawing.Size(73, 25);
            this.btnDummy.TabIndex = 0;
            this.btnDummy.Text = "btnDummy";
            this.btnDummy.UseVisualStyleBackColor = true;
            // 
            // btnFechar
            // 
            this.btnFechar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnFechar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnFechar.Image = ((System.Drawing.Image)(resources.GetObject("btnFechar.Image")));
            this.btnFechar.Location = new System.Drawing.Point(959, 4);
            this.btnFechar.Name = "btnFechar";
            this.btnFechar.Size = new System.Drawing.Size(40, 44);
            this.btnFechar.TabIndex = 5;
            this.btnFechar.TabStop = false;
            this.btnFechar.UseVisualStyleBackColor = true;
            this.btnFechar.Click += new System.EventHandler(this.btnFechar_Click);
            // 
            // pnCampos
            // 
            this.pnCampos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnCampos.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnCampos.Location = new System.Drawing.Point(0, 79);
            this.pnCampos.Name = "pnCampos";
            this.pnCampos.Size = new System.Drawing.Size(1008, 465);
            this.pnCampos.TabIndex = 9;
            // 
            // tmrRelogio
            // 
            this.tmrRelogio.Enabled = true;
            this.tmrRelogio.Interval = 200;
            this.tmrRelogio.Tick += new System.EventHandler(this.tmrRelogio_Tick);
            // 
            // FModelo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnFechar;
            this.ClientSize = new System.Drawing.Size(1008, 562);
            this.Controls.Add(this.pnCampos);
            this.Controls.Add(this.pnBotoes);
            this.Controls.Add(this.pnStatus);
            this.Controls.Add(this.menuPrincipal);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuPrincipal;
            this.MaximizeBox = false;
            this.Name = "FModelo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ADM - Módulo Administrativo  -  1.00 - XX.XXX.20XX";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FModelo_FormClosing);
            this.Load += new System.EventHandler(this.FModelo_Load);
            this.pnStatus.ResumeLayout(false);
            this.pnMensagem.ResumeLayout(false);
            this.pnHora.ResumeLayout(false);
            this.pnHora.PerformLayout();
            this.pnData.ResumeLayout(false);
            this.pnData.PerformLayout();
            this.menuPrincipal.ResumeLayout(false);
            this.menuPrincipal.PerformLayout();
            this.pnBotoes.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Panel pnStatus;
		private System.Windows.Forms.Panel pnMensagem;
		private System.Windows.Forms.Label lblMensagem;
		private System.Windows.Forms.Panel pnHora;
		private System.Windows.Forms.Label lblHora;
		private System.Windows.Forms.Panel pnData;
		private System.Windows.Forms.Label lblData;
		private System.Windows.Forms.MenuStrip menuPrincipal;
		protected System.Windows.Forms.Panel pnBotoes;
		protected System.Windows.Forms.Button btnSobre;
		private System.Windows.Forms.PictureBox pbLogo;
		protected System.Windows.Forms.Button btnDummy;
		protected System.Windows.Forms.Button btnFechar;
		private System.Windows.Forms.ToolStripMenuItem menuArquivo;
		private System.Windows.Forms.ToolStripMenuItem menuArquivoFechar;
		private System.Windows.Forms.ToolStripMenuItem menuAjuda;
		private System.Windows.Forms.ToolStripMenuItem menuAjudaSobre;
		private System.Windows.Forms.Timer tmrRelogio;
		protected System.Windows.Forms.Panel pnCampos;
	}
}