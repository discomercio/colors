namespace ConsolidadorXlsEC
{
	partial class FConsolidaDadosPlanilha
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FConsolidaDadosPlanilha));
			this.lblTituloPainel = new System.Windows.Forms.Label();
			this.lblPlanilhaControle = new System.Windows.Forms.Label();
			this.txtPlanilhaControle = new System.Windows.Forms.TextBox();
			this.btnSelecionaPlanilhaControle = new System.Windows.Forms.Button();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.btnConsolidaPlanilha = new System.Windows.Forms.Button();
			this.openFileDialogCtrl = new System.Windows.Forms.OpenFileDialog();
			this.openFileDialogPrecos = new System.Windows.Forms.OpenFileDialog();
			this.btnSelecionaPlanilhaFerramentaPrecos = new System.Windows.Forms.Button();
			this.txtPlanilhaFerramentaPrecos = new System.Windows.Forms.TextBox();
			this.lblPlanilhaPrecos = new System.Windows.Forms.Label();
			this.btnAbrePlanilhaControle = new System.Windows.Forms.Button();
			this.btnAbrePlanilhaPrecos = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMensagensInformativas.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnConsolidaPlanilha);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnConsolidaPlanilha, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.btnAbrePlanilhaPrecos);
			this.pnCampos.Controls.Add(this.btnAbrePlanilhaControle);
			this.pnCampos.Controls.Add(this.btnSelecionaPlanilhaFerramentaPrecos);
			this.pnCampos.Controls.Add(this.txtPlanilhaFerramentaPrecos);
			this.pnCampos.Controls.Add(this.lblPlanilhaPrecos);
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
			// lblTituloPainel
			// 
			this.lblTituloPainel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTituloPainel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTituloPainel.Image = ((System.Drawing.Image)(resources.GetObject("lblTituloPainel.Image")));
			this.lblTituloPainel.Location = new System.Drawing.Point(-2, 1);
			this.lblTituloPainel.Name = "lblTituloPainel";
			this.lblTituloPainel.Size = new System.Drawing.Size(1018, 40);
			this.lblTituloPainel.TabIndex = 1;
			this.lblTituloPainel.Text = "Consolidação de Dados da Planilha";
			this.lblTituloPainel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblPlanilhaControle
			// 
			this.lblPlanilhaControle.AutoSize = true;
			this.lblPlanilhaControle.Location = new System.Drawing.Point(21, 53);
			this.lblPlanilhaControle.Name = "lblPlanilhaControle";
			this.lblPlanilhaControle.Size = new System.Drawing.Size(86, 13);
			this.lblPlanilhaControle.TabIndex = 2;
			this.lblPlanilhaControle.Text = "Planilha Controle";
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
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(12, 141);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(987, 216);
			this.gboxMensagensInformativas.TabIndex = 12;
			this.gboxMensagensInformativas.TabStop = false;
			this.gboxMensagensInformativas.Text = "Mensagens Informativas";
			// 
			// lbMensagem
			// 
			this.lbMensagem.FormattingEnabled = true;
			this.lbMensagem.Location = new System.Drawing.Point(15, 19);
			this.lbMensagem.Name = "lbMensagem";
			this.lbMensagem.ScrollAlwaysVisible = true;
			this.lbMensagem.Size = new System.Drawing.Size(965, 186);
			this.lbMensagem.TabIndex = 0;
			this.lbMensagem.DoubleClick += new System.EventHandler(this.lbMensagem_DoubleClick);
			// 
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(12, 373);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(987, 218);
			this.gboxMsgErro.TabIndex = 10;
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
			this.lbErro.Size = new System.Drawing.Size(965, 186);
			this.lbErro.TabIndex = 0;
			this.lbErro.DoubleClick += new System.EventHandler(this.lbErro_DoubleClick);
			// 
			// btnConsolidaPlanilha
			// 
			this.btnConsolidaPlanilha.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnConsolidaPlanilha.Image = ((System.Drawing.Image)(resources.GetObject("btnConsolidaPlanilha.Image")));
			this.btnConsolidaPlanilha.Location = new System.Drawing.Point(879, 4);
			this.btnConsolidaPlanilha.Name = "btnConsolidaPlanilha";
			this.btnConsolidaPlanilha.Size = new System.Drawing.Size(40, 44);
			this.btnConsolidaPlanilha.TabIndex = 0;
			this.btnConsolidaPlanilha.TabStop = false;
			this.btnConsolidaPlanilha.UseVisualStyleBackColor = true;
			this.btnConsolidaPlanilha.Click += new System.EventHandler(this.btnConsolidaPlanilha_Click);
			// 
			// openFileDialogCtrl
			// 
			this.openFileDialogCtrl.AddExtension = false;
			this.openFileDialogCtrl.Filter = "Planilha Excel|*.xls;*.xlsx;*.xlsm|Todos os arquivos|*.*";
			this.openFileDialogCtrl.InitialDirectory = "\\";
			// 
			// openFileDialogPrecos
			// 
			this.openFileDialogPrecos.AddExtension = false;
			this.openFileDialogPrecos.Filter = "Planilha Excel|*.xls;*.xlsx|Todos os arquivos|*.*";
			this.openFileDialogPrecos.InitialDirectory = "\\";
			// 
			// btnSelecionaPlanilhaFerramentaPrecos
			// 
			this.btnSelecionaPlanilhaFerramentaPrecos.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaPlanilhaFerramentaPrecos.Image")));
			this.btnSelecionaPlanilhaFerramentaPrecos.Location = new System.Drawing.Point(906, 83);
			this.btnSelecionaPlanilhaFerramentaPrecos.Name = "btnSelecionaPlanilhaFerramentaPrecos";
			this.btnSelecionaPlanilhaFerramentaPrecos.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaPlanilhaFerramentaPrecos.TabIndex = 4;
			this.btnSelecionaPlanilhaFerramentaPrecos.UseVisualStyleBackColor = true;
			this.btnSelecionaPlanilhaFerramentaPrecos.Click += new System.EventHandler(this.btnSelecionaPlanilhaFerramentaPrecos_Click);
			// 
			// txtPlanilhaFerramentaPrecos
			// 
			this.txtPlanilhaFerramentaPrecos.BackColor = System.Drawing.Color.White;
			this.txtPlanilhaFerramentaPrecos.Location = new System.Drawing.Point(113, 86);
			this.txtPlanilhaFerramentaPrecos.Name = "txtPlanilhaFerramentaPrecos";
			this.txtPlanilhaFerramentaPrecos.ReadOnly = true;
			this.txtPlanilhaFerramentaPrecos.Size = new System.Drawing.Size(787, 20);
			this.txtPlanilhaFerramentaPrecos.TabIndex = 3;
			this.txtPlanilhaFerramentaPrecos.DoubleClick += new System.EventHandler(this.txtPlanilhaFerramentaPrecos_DoubleClick);
			this.txtPlanilhaFerramentaPrecos.Enter += new System.EventHandler(this.txtPlanilhaFerramentaPrecos_Enter);
			// 
			// lblPlanilhaPrecos
			// 
			this.lblPlanilhaPrecos.AutoSize = true;
			this.lblPlanilhaPrecos.Location = new System.Drawing.Point(27, 89);
			this.lblPlanilhaPrecos.Name = "lblPlanilhaPrecos";
			this.lblPlanilhaPrecos.Size = new System.Drawing.Size(80, 13);
			this.lblPlanilhaPrecos.TabIndex = 13;
			this.lblPlanilhaPrecos.Text = "Planilha Preços";
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
			// btnAbrePlanilhaPrecos
			// 
			this.btnAbrePlanilhaPrecos.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrePlanilhaPrecos.Image")));
			this.btnAbrePlanilhaPrecos.Location = new System.Drawing.Point(953, 83);
			this.btnAbrePlanilhaPrecos.Name = "btnAbrePlanilhaPrecos";
			this.btnAbrePlanilhaPrecos.Size = new System.Drawing.Size(39, 25);
			this.btnAbrePlanilhaPrecos.TabIndex = 5;
			this.btnAbrePlanilhaPrecos.UseVisualStyleBackColor = true;
			this.btnAbrePlanilhaPrecos.Click += new System.EventHandler(this.btnAbrePlanilhaPrecos_Click);
			// 
			// FConsolidaDadosPlanilha
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FConsolidaDadosPlanilha";
			this.Text = "Consolida Dados da Planilha";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FConsolidaDadosPlanilha_FormClosing);
			this.Load += new System.EventHandler(this.FConsolidaDadosPlanilha_Load);
			this.Shown += new System.EventHandler(this.FConsolidaDadosPlanilha_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.gboxMsgErro.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTituloPainel;
		private System.Windows.Forms.TextBox txtPlanilhaControle;
		private System.Windows.Forms.Label lblPlanilhaControle;
		private System.Windows.Forms.Button btnSelecionaPlanilhaControle;
		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.Button btnConsolidaPlanilha;
		private System.Windows.Forms.OpenFileDialog openFileDialogCtrl;
		private System.Windows.Forms.OpenFileDialog openFileDialogPrecos;
		private System.Windows.Forms.Button btnSelecionaPlanilhaFerramentaPrecos;
		private System.Windows.Forms.TextBox txtPlanilhaFerramentaPrecos;
		private System.Windows.Forms.Label lblPlanilhaPrecos;
		private System.Windows.Forms.Button btnAbrePlanilhaPrecos;
		private System.Windows.Forms.Button btnAbrePlanilhaControle;
	}
}