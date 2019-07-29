namespace ADM2
{
	partial class FAtualizaPlanilhaEstoque
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FAtualizaPlanilhaEstoque));
			this.btnAbrePlanilhaEstoque = new System.Windows.Forms.Button();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.btnSelecionaPlanilhaEstoque = new System.Windows.Forms.Button();
			this.txtPlanilhaEstoque = new System.Windows.Forms.TextBox();
			this.lblPlanilhaEstoque = new System.Windows.Forms.Label();
			this.lblTituloPainel = new System.Windows.Forms.Label();
			this.openFileDialogCtrl = new System.Windows.Forms.OpenFileDialog();
			this.btnAtualizaPlanilhaEstoque = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.gboxMensagensInformativas.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnAtualizaPlanilhaEstoque);
			this.pnBotoes.Size = new System.Drawing.Size(1018, 55);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnAtualizaPlanilhaEstoque, 0);
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(924, 4);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(969, 4);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.btnAbrePlanilhaEstoque);
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.btnSelecionaPlanilhaEstoque);
			this.pnCampos.Controls.Add(this.txtPlanilhaEstoque);
			this.pnCampos.Controls.Add(this.lblPlanilhaEstoque);
			this.pnCampos.Controls.Add(this.lblTituloPainel);
			this.pnCampos.Size = new System.Drawing.Size(1018, 559);
			// 
			// btnAbrePlanilhaEstoque
			// 
			this.btnAbrePlanilhaEstoque.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrePlanilhaEstoque.Image")));
			this.btnAbrePlanilhaEstoque.Location = new System.Drawing.Point(953, 47);
			this.btnAbrePlanilhaEstoque.Name = "btnAbrePlanilhaEstoque";
			this.btnAbrePlanilhaEstoque.Size = new System.Drawing.Size(39, 25);
			this.btnAbrePlanilhaEstoque.TabIndex = 17;
			this.btnAbrePlanilhaEstoque.UseVisualStyleBackColor = true;
			this.btnAbrePlanilhaEstoque.Click += new System.EventHandler(this.btnAbrePlanilhaEstoque_Click);
			// 
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(12, 323);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(987, 218);
			this.gboxMsgErro.TabIndex = 22;
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
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(12, 91);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(987, 216);
			this.gboxMensagensInformativas.TabIndex = 23;
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
			// btnSelecionaPlanilhaEstoque
			// 
			this.btnSelecionaPlanilhaEstoque.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaPlanilhaEstoque.Image")));
			this.btnSelecionaPlanilhaEstoque.Location = new System.Drawing.Point(906, 47);
			this.btnSelecionaPlanilhaEstoque.Name = "btnSelecionaPlanilhaEstoque";
			this.btnSelecionaPlanilhaEstoque.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaPlanilhaEstoque.TabIndex = 15;
			this.btnSelecionaPlanilhaEstoque.UseVisualStyleBackColor = true;
			this.btnSelecionaPlanilhaEstoque.Click += new System.EventHandler(this.btnSelecionaPlanilhaEstoque_Click);
			// 
			// txtPlanilhaEstoque
			// 
			this.txtPlanilhaEstoque.BackColor = System.Drawing.Color.White;
			this.txtPlanilhaEstoque.Location = new System.Drawing.Point(113, 50);
			this.txtPlanilhaEstoque.Name = "txtPlanilhaEstoque";
			this.txtPlanilhaEstoque.ReadOnly = true;
			this.txtPlanilhaEstoque.Size = new System.Drawing.Size(787, 20);
			this.txtPlanilhaEstoque.TabIndex = 14;
			this.txtPlanilhaEstoque.DoubleClick += new System.EventHandler(this.txtPlanilhaEstoque_DoubleClick);
			this.txtPlanilhaEstoque.Enter += new System.EventHandler(this.txtPlanilhaEstoque_Enter);
			// 
			// lblPlanilhaEstoque
			// 
			this.lblPlanilhaEstoque.AutoSize = true;
			this.lblPlanilhaEstoque.Location = new System.Drawing.Point(21, 53);
			this.lblPlanilhaEstoque.Name = "lblPlanilhaEstoque";
			this.lblPlanilhaEstoque.Size = new System.Drawing.Size(86, 13);
			this.lblPlanilhaEstoque.TabIndex = 18;
			this.lblPlanilhaEstoque.Text = "Planilha Estoque";
			// 
			// lblTituloPainel
			// 
			this.lblTituloPainel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTituloPainel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTituloPainel.Image = ((System.Drawing.Image)(resources.GetObject("lblTituloPainel.Image")));
			this.lblTituloPainel.Location = new System.Drawing.Point(-2, 1);
			this.lblTituloPainel.Name = "lblTituloPainel";
			this.lblTituloPainel.Size = new System.Drawing.Size(1018, 40);
			this.lblTituloPainel.TabIndex = 16;
			this.lblTituloPainel.Text = "Atualização da Planilha do Estoque";
			this.lblTituloPainel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// openFileDialogCtrl
			// 
			this.openFileDialogCtrl.AddExtension = false;
			this.openFileDialogCtrl.Filter = "Planilha Excel|*.xls;*.xlsx;*.xlsm|Todos os arquivos|*.*";
			this.openFileDialogCtrl.InitialDirectory = "\\";
			// 
			// btnAtualizaPlanilhaEstoque
			// 
			this.btnAtualizaPlanilhaEstoque.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnAtualizaPlanilhaEstoque.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizaPlanilhaEstoque.Image")));
			this.btnAtualizaPlanilhaEstoque.Location = new System.Drawing.Point(879, 4);
			this.btnAtualizaPlanilhaEstoque.Name = "btnAtualizaPlanilhaEstoque";
			this.btnAtualizaPlanilhaEstoque.Size = new System.Drawing.Size(40, 44);
			this.btnAtualizaPlanilhaEstoque.TabIndex = 8;
			this.btnAtualizaPlanilhaEstoque.TabStop = false;
			this.btnAtualizaPlanilhaEstoque.UseVisualStyleBackColor = true;
			this.btnAtualizaPlanilhaEstoque.Click += new System.EventHandler(this.btnAtualizaPlanilhaEstoque_Click);
			// 
			// FAtualizaPlanilhaEstoque
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1018, 656);
			this.Name = "FAtualizaPlanilhaEstoque";
			this.Text = "FAtualizaPlanilhaEstoque";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FAtualizaPlanilhaEstoque_FormClosing);
			this.Load += new System.EventHandler(this.FAtualizaPlanilhaEstoque_Load);
			this.Shown += new System.EventHandler(this.FAtualizaPlanilhaEstoque_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxMsgErro.ResumeLayout(false);
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.Button btnAbrePlanilhaEstoque;
		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.Button btnSelecionaPlanilhaEstoque;
		private System.Windows.Forms.TextBox txtPlanilhaEstoque;
		private System.Windows.Forms.Label lblPlanilhaEstoque;
		private System.Windows.Forms.Label lblTituloPainel;
		private System.Windows.Forms.OpenFileDialog openFileDialogCtrl;
		private System.Windows.Forms.Button btnAtualizaPlanilhaEstoque;
	}
}