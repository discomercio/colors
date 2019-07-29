namespace ConsolidadorXlsEC
{
	partial class FAtualizaPrecosSistema
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FAtualizaPrecosSistema));
			this.btnAbrePlanilhaControle = new System.Windows.Forms.Button();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.btnSelecionaPlanilhaControle = new System.Windows.Forms.Button();
			this.txtPlanilhaControle = new System.Windows.Forms.TextBox();
			this.lblPlanilhaControle = new System.Windows.Forms.Label();
			this.lblTituloPainel = new System.Windows.Forms.Label();
			this.btnAtualizaPrecos = new System.Windows.Forms.Button();
			this.openFileDialogCtrl = new System.Windows.Forms.OpenFileDialog();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.gboxMensagensInformativas.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnAtualizaPrecos);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnAtualizaPrecos, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.btnAbrePlanilhaControle);
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.btnSelecionaPlanilhaControle);
			this.pnCampos.Controls.Add(this.txtPlanilhaControle);
			this.pnCampos.Controls.Add(this.lblPlanilhaControle);
			this.pnCampos.Controls.Add(this.lblTituloPainel);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.TabIndex = 2;
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 1;
			// 
			// btnAbrePlanilhaControle
			// 
			this.btnAbrePlanilhaControle.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrePlanilhaControle.Image")));
			this.btnAbrePlanilhaControle.Location = new System.Drawing.Point(953, 47);
			this.btnAbrePlanilhaControle.Name = "btnAbrePlanilhaControle";
			this.btnAbrePlanilhaControle.Size = new System.Drawing.Size(39, 25);
			this.btnAbrePlanilhaControle.TabIndex = 2;
			this.btnAbrePlanilhaControle.UseVisualStyleBackColor = true;
			this.btnAbrePlanilhaControle.Click += new System.EventHandler(this.btnAbrePlanilhaControle_Click);
			// 
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(12, 343);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(987, 254);
			this.gboxMsgErro.TabIndex = 18;
			this.gboxMsgErro.TabStop = false;
			this.gboxMsgErro.Text = "Mensagens de Erro";
			// 
			// lbErro
			// 
			this.lbErro.ForeColor = System.Drawing.Color.Red;
			this.lbErro.FormattingEnabled = true;
			this.lbErro.Location = new System.Drawing.Point(15, 19);
			this.lbErro.Name = "lbErro";
			this.lbErro.ScrollAlwaysVisible = true;
			this.lbErro.Size = new System.Drawing.Size(965, 225);
			this.lbErro.TabIndex = 0;
			this.lbErro.DoubleClick += new System.EventHandler(this.lbErro_DoubleClick);
			// 
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(12, 88);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(987, 239);
			this.gboxMensagensInformativas.TabIndex = 19;
			this.gboxMensagensInformativas.TabStop = false;
			this.gboxMensagensInformativas.Text = "Mensagens Informativas";
			// 
			// lbMensagem
			// 
			this.lbMensagem.FormattingEnabled = true;
			this.lbMensagem.Location = new System.Drawing.Point(15, 19);
			this.lbMensagem.Name = "lbMensagem";
			this.lbMensagem.ScrollAlwaysVisible = true;
			this.lbMensagem.Size = new System.Drawing.Size(965, 212);
			this.lbMensagem.TabIndex = 0;
			this.lbMensagem.DoubleClick += new System.EventHandler(this.lbMensagem_DoubleClick);
			// 
			// btnSelecionaPlanilhaControle
			// 
			this.btnSelecionaPlanilhaControle.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaPlanilhaControle.Image")));
			this.btnSelecionaPlanilhaControle.Location = new System.Drawing.Point(906, 47);
			this.btnSelecionaPlanilhaControle.Name = "btnSelecionaPlanilhaControle";
			this.btnSelecionaPlanilhaControle.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaPlanilhaControle.TabIndex = 1;
			this.btnSelecionaPlanilhaControle.UseVisualStyleBackColor = true;
			this.btnSelecionaPlanilhaControle.Click += new System.EventHandler(this.btnSelecionaPlanilhaControle_Click);
			// 
			// txtPlanilhaControle
			// 
			this.txtPlanilhaControle.BackColor = System.Drawing.Color.White;
			this.txtPlanilhaControle.Location = new System.Drawing.Point(113, 50);
			this.txtPlanilhaControle.Name = "txtPlanilhaControle";
			this.txtPlanilhaControle.ReadOnly = true;
			this.txtPlanilhaControle.Size = new System.Drawing.Size(787, 20);
			this.txtPlanilhaControle.TabIndex = 0;
			this.txtPlanilhaControle.DoubleClick += new System.EventHandler(this.txtPlanilhaControle_DoubleClick);
			this.txtPlanilhaControle.Enter += new System.EventHandler(this.txtPlanilhaControle_Enter);
			// 
			// lblPlanilhaControle
			// 
			this.lblPlanilhaControle.AutoSize = true;
			this.lblPlanilhaControle.Location = new System.Drawing.Point(21, 53);
			this.lblPlanilhaControle.Name = "lblPlanilhaControle";
			this.lblPlanilhaControle.Size = new System.Drawing.Size(86, 13);
			this.lblPlanilhaControle.TabIndex = 17;
			this.lblPlanilhaControle.Text = "Planilha Controle";
			// 
			// lblTituloPainel
			// 
			this.lblTituloPainel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTituloPainel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTituloPainel.Image = ((System.Drawing.Image)(resources.GetObject("lblTituloPainel.Image")));
			this.lblTituloPainel.Location = new System.Drawing.Point(-2, 1);
			this.lblTituloPainel.Name = "lblTituloPainel";
			this.lblTituloPainel.Size = new System.Drawing.Size(1018, 40);
			this.lblTituloPainel.TabIndex = 15;
			this.lblTituloPainel.Text = "Atualiza Preços no Sistema";
			this.lblTituloPainel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAtualizaPrecos
			// 
			this.btnAtualizaPrecos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnAtualizaPrecos.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizaPrecos.Image")));
			this.btnAtualizaPrecos.Location = new System.Drawing.Point(879, 4);
			this.btnAtualizaPrecos.Name = "btnAtualizaPrecos";
			this.btnAtualizaPrecos.Size = new System.Drawing.Size(40, 44);
			this.btnAtualizaPrecos.TabIndex = 0;
			this.btnAtualizaPrecos.TabStop = false;
			this.btnAtualizaPrecos.UseVisualStyleBackColor = true;
			this.btnAtualizaPrecos.Click += new System.EventHandler(this.btnAtualizaPrecos_Click);
			// 
			// openFileDialogCtrl
			// 
			this.openFileDialogCtrl.AddExtension = false;
			this.openFileDialogCtrl.Filter = "Planilha Excel|*.xls;*.xlsx;*.xlsm|Todos os arquivos|*.*";
			this.openFileDialogCtrl.InitialDirectory = "\\";
			// 
			// FAtualizaPrecosSistema
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FAtualizaPrecosSistema";
			this.Text = "Atualiza Preços no Sistema";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FAtualizaPrecosSistema_FormClosing);
			this.Load += new System.EventHandler(this.FAtualizaPrecosSistema_Load);
			this.Shown += new System.EventHandler(this.FAtualizaPrecosSistema_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxMsgErro.ResumeLayout(false);
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnAbrePlanilhaControle;
		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.Button btnSelecionaPlanilhaControle;
		private System.Windows.Forms.TextBox txtPlanilhaControle;
		private System.Windows.Forms.Label lblPlanilhaControle;
		private System.Windows.Forms.Label lblTituloPainel;
		private System.Windows.Forms.Button btnAtualizaPrecos;
		private System.Windows.Forms.OpenFileDialog openFileDialogCtrl;
	}
}