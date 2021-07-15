namespace ConsolidadorXlsEC
{
	partial class FConferenciaPreco
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FConferenciaPreco));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.btnSelecionaArquivo = new System.Windows.Forms.Button();
			this.txtArquivo = new System.Windows.Forms.TextBox();
			this.lblArquivo = new System.Windows.Forms.Label();
			this.lblTituloPainel = new System.Windows.Forms.Label();
			this.btnAbreArquivo = new System.Windows.Forms.Button();
			this.openFileDialogCtrl = new System.Windows.Forms.OpenFileDialog();
			this.btnIniciaProcessamento = new System.Windows.Forms.Button();
			this.grid = new System.Windows.Forms.DataGridView();
			this.ColVisibleOrdenacaoPadrao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.SKU = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Descricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.PrecoMagento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.PrecoCSV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.DiferencaValor = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.DiferencaPerc = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.ColHiddenValorOrdenacaoPadrao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cbPlataforma = new System.Windows.Forms.ComboBox();
			this.lblPlataforma = new System.Windows.Forms.Label();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.gboxMensagensInformativas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnIniciaProcessamento);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnIniciaProcessamento, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.cbPlataforma);
			this.pnCampos.Controls.Add(this.lblPlataforma);
			this.pnCampos.Controls.Add(this.grid);
			this.pnCampos.Controls.Add(this.btnAbreArquivo);
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.btnSelecionaArquivo);
			this.pnCampos.Controls.Add(this.txtArquivo);
			this.pnCampos.Controls.Add(this.lblArquivo);
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
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(12, 506);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(987, 95);
			this.gboxMsgErro.TabIndex = 8;
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
			this.lbErro.Size = new System.Drawing.Size(965, 69);
			this.lbErro.TabIndex = 0;
			this.lbErro.DoubleClick += new System.EventHandler(this.lbErro_DoubleClick);
			// 
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(12, 400);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(987, 95);
			this.gboxMensagensInformativas.TabIndex = 7;
			this.gboxMensagensInformativas.TabStop = false;
			this.gboxMensagensInformativas.Text = "Mensagens Informativas";
			// 
			// lbMensagem
			// 
			this.lbMensagem.FormattingEnabled = true;
			this.lbMensagem.Location = new System.Drawing.Point(15, 19);
			this.lbMensagem.Name = "lbMensagem";
			this.lbMensagem.ScrollAlwaysVisible = true;
			this.lbMensagem.Size = new System.Drawing.Size(965, 69);
			this.lbMensagem.TabIndex = 0;
			this.lbMensagem.DoubleClick += new System.EventHandler(this.lbMensagem_DoubleClick);
			// 
			// btnSelecionaArquivo
			// 
			this.btnSelecionaArquivo.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArquivo.Image")));
			this.btnSelecionaArquivo.Location = new System.Drawing.Point(906, 47);
			this.btnSelecionaArquivo.Name = "btnSelecionaArquivo";
			this.btnSelecionaArquivo.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaArquivo.TabIndex = 2;
			this.btnSelecionaArquivo.UseVisualStyleBackColor = true;
			this.btnSelecionaArquivo.Click += new System.EventHandler(this.btnSelecionaArquivo_Click);
			// 
			// txtArquivo
			// 
			this.txtArquivo.BackColor = System.Drawing.Color.White;
			this.txtArquivo.Location = new System.Drawing.Point(94, 50);
			this.txtArquivo.Name = "txtArquivo";
			this.txtArquivo.ReadOnly = true;
			this.txtArquivo.Size = new System.Drawing.Size(806, 20);
			this.txtArquivo.TabIndex = 1;
			this.txtArquivo.DoubleClick += new System.EventHandler(this.txtArquivo_DoubleClick);
			this.txtArquivo.Enter += new System.EventHandler(this.txtArquivo_Enter);
			// 
			// lblArquivo
			// 
			this.lblArquivo.AutoSize = true;
			this.lblArquivo.Location = new System.Drawing.Point(21, 53);
			this.lblArquivo.Name = "lblArquivo";
			this.lblArquivo.Size = new System.Drawing.Size(67, 13);
			this.lblArquivo.TabIndex = 0;
			this.lblArquivo.Text = "Arquivo CSV";
			// 
			// lblTituloPainel
			// 
			this.lblTituloPainel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTituloPainel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTituloPainel.Image = ((System.Drawing.Image)(resources.GetObject("lblTituloPainel.Image")));
			this.lblTituloPainel.Location = new System.Drawing.Point(-2, 1);
			this.lblTituloPainel.Name = "lblTituloPainel";
			this.lblTituloPainel.Size = new System.Drawing.Size(1018, 40);
			this.lblTituloPainel.TabIndex = 22;
			this.lblTituloPainel.Text = "Conferência de Preços";
			this.lblTituloPainel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAbreArquivo
			// 
			this.btnAbreArquivo.Image = ((System.Drawing.Image)(resources.GetObject("btnAbreArquivo.Image")));
			this.btnAbreArquivo.Location = new System.Drawing.Point(953, 47);
			this.btnAbreArquivo.Name = "btnAbreArquivo";
			this.btnAbreArquivo.Size = new System.Drawing.Size(39, 25);
			this.btnAbreArquivo.TabIndex = 3;
			this.btnAbreArquivo.UseVisualStyleBackColor = true;
			this.btnAbreArquivo.Click += new System.EventHandler(this.btnAbreArquivo_Click);
			// 
			// openFileDialogCtrl
			// 
			this.openFileDialogCtrl.AddExtension = false;
			this.openFileDialogCtrl.Filter = "Arquivo CSV|*.csv|Todos os arquivos|*.*";
			this.openFileDialogCtrl.InitialDirectory = "\\";
			// 
			// btnIniciaProcessamento
			// 
			this.btnIniciaProcessamento.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnIniciaProcessamento.Image = ((System.Drawing.Image)(resources.GetObject("btnIniciaProcessamento.Image")));
			this.btnIniciaProcessamento.Location = new System.Drawing.Point(879, 4);
			this.btnIniciaProcessamento.Name = "btnIniciaProcessamento";
			this.btnIniciaProcessamento.Size = new System.Drawing.Size(40, 44);
			this.btnIniciaProcessamento.TabIndex = 0;
			this.btnIniciaProcessamento.TabStop = false;
			this.btnIniciaProcessamento.UseVisualStyleBackColor = true;
			this.btnIniciaProcessamento.Click += new System.EventHandler(this.btnIniciaProcessamento_Click);
			// 
			// grid
			// 
			this.grid.AllowUserToAddRows = false;
			this.grid.AllowUserToDeleteRows = false;
			this.grid.AllowUserToResizeRows = false;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColVisibleOrdenacaoPadrao,
            this.SKU,
            this.Descricao,
            this.PrecoMagento,
            this.PrecoCSV,
            this.DiferencaValor,
            this.DiferencaPerc,
            this.ColHiddenValorOrdenacaoPadrao});
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grid.DefaultCellStyle = dataGridViewCellStyle9;
			this.grid.Location = new System.Drawing.Point(12, 108);
			this.grid.MultiSelect = false;
			this.grid.Name = "grid";
			this.grid.ReadOnly = true;
			this.grid.RowHeadersVisible = false;
			this.grid.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grid.ShowEditingIcon = false;
			this.grid.Size = new System.Drawing.Size(987, 282);
			this.grid.TabIndex = 6;
			this.grid.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.grid_SortCompare);
			// 
			// ColVisibleOrdenacaoPadrao
			// 
			this.ColVisibleOrdenacaoPadrao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.ColVisibleOrdenacaoPadrao.DefaultCellStyle = dataGridViewCellStyle2;
			this.ColVisibleOrdenacaoPadrao.HeaderText = "";
			this.ColVisibleOrdenacaoPadrao.MinimumWidth = 30;
			this.ColVisibleOrdenacaoPadrao.Name = "ColVisibleOrdenacaoPadrao";
			this.ColVisibleOrdenacaoPadrao.ReadOnly = true;
			this.ColVisibleOrdenacaoPadrao.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.ColVisibleOrdenacaoPadrao.Width = 30;
			// 
			// SKU
			// 
			this.SKU.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.SKU.DefaultCellStyle = dataGridViewCellStyle3;
			this.SKU.HeaderText = "SKU";
			this.SKU.MinimumWidth = 70;
			this.SKU.Name = "SKU";
			this.SKU.ReadOnly = true;
			this.SKU.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.SKU.Width = 70;
			// 
			// Descricao
			// 
			this.Descricao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Descricao.DefaultCellStyle = dataGridViewCellStyle4;
			this.Descricao.HeaderText = "Descrição";
			this.Descricao.Name = "Descricao";
			this.Descricao.ReadOnly = true;
			// 
			// PrecoMagento
			// 
			this.PrecoMagento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.PrecoMagento.DefaultCellStyle = dataGridViewCellStyle5;
			this.PrecoMagento.HeaderText = "Preço Magento";
			this.PrecoMagento.MinimumWidth = 100;
			this.PrecoMagento.Name = "PrecoMagento";
			this.PrecoMagento.ReadOnly = true;
			this.PrecoMagento.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			// 
			// PrecoCSV
			// 
			this.PrecoCSV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.PrecoCSV.DefaultCellStyle = dataGridViewCellStyle6;
			this.PrecoCSV.HeaderText = "Preço Novo";
			this.PrecoCSV.MinimumWidth = 100;
			this.PrecoCSV.Name = "PrecoCSV";
			this.PrecoCSV.ReadOnly = true;
			this.PrecoCSV.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			// 
			// DiferencaValor
			// 
			this.DiferencaValor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.DiferencaValor.DefaultCellStyle = dataGridViewCellStyle7;
			this.DiferencaValor.HeaderText = "Dif (R$)";
			this.DiferencaValor.MinimumWidth = 80;
			this.DiferencaValor.Name = "DiferencaValor";
			this.DiferencaValor.ReadOnly = true;
			this.DiferencaValor.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.DiferencaValor.Width = 80;
			// 
			// DiferencaPerc
			// 
			this.DiferencaPerc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.DiferencaPerc.DefaultCellStyle = dataGridViewCellStyle8;
			this.DiferencaPerc.HeaderText = "Dif (%)";
			this.DiferencaPerc.MinimumWidth = 80;
			this.DiferencaPerc.Name = "DiferencaPerc";
			this.DiferencaPerc.ReadOnly = true;
			this.DiferencaPerc.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.DiferencaPerc.Width = 80;
			// 
			// ColHiddenValorOrdenacaoPadrao
			// 
			this.ColHiddenValorOrdenacaoPadrao.HeaderText = "Campo Ordenação Padrão";
			this.ColHiddenValorOrdenacaoPadrao.Name = "ColHiddenValorOrdenacaoPadrao";
			this.ColHiddenValorOrdenacaoPadrao.ReadOnly = true;
			this.ColHiddenValorOrdenacaoPadrao.Visible = false;
			// 
			// cbPlataforma
			// 
			this.cbPlataforma.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlataforma.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlataforma.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlataforma.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlataforma.FormattingEnabled = true;
			this.cbPlataforma.Location = new System.Drawing.Point(94, 78);
			this.cbPlataforma.Name = "cbPlataforma";
			this.cbPlataforma.Size = new System.Drawing.Size(150, 24);
			this.cbPlataforma.TabIndex = 5;
			// 
			// lblPlataforma
			// 
			this.lblPlataforma.AutoSize = true;
			this.lblPlataforma.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblPlataforma.Location = new System.Drawing.Point(31, 83);
			this.lblPlataforma.Name = "lblPlataforma";
			this.lblPlataforma.Size = new System.Drawing.Size(57, 13);
			this.lblPlataforma.TabIndex = 4;
			this.lblPlataforma.Text = "Plataforma";
			// 
			// FConferenciaPreco
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FConferenciaPreco";
			this.Text = "FConferenciaPreco";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FConferenciaPreco_FormClosing);
			this.Load += new System.EventHandler(this.FConferenciaPreco_Load);
			this.Shown += new System.EventHandler(this.FConferenciaPreco_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxMsgErro.ResumeLayout(false);
			this.gboxMensagensInformativas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.Button btnSelecionaArquivo;
		private System.Windows.Forms.TextBox txtArquivo;
		private System.Windows.Forms.Label lblArquivo;
		private System.Windows.Forms.Label lblTituloPainel;
		private System.Windows.Forms.Button btnAbreArquivo;
		private System.Windows.Forms.OpenFileDialog openFileDialogCtrl;
		private System.Windows.Forms.Button btnIniciaProcessamento;
		private System.Windows.Forms.DataGridView grid;
		private System.Windows.Forms.DataGridViewTextBoxColumn ColVisibleOrdenacaoPadrao;
		private System.Windows.Forms.DataGridViewTextBoxColumn SKU;
		private System.Windows.Forms.DataGridViewTextBoxColumn Descricao;
		private System.Windows.Forms.DataGridViewTextBoxColumn PrecoMagento;
		private System.Windows.Forms.DataGridViewTextBoxColumn PrecoCSV;
		private System.Windows.Forms.DataGridViewTextBoxColumn DiferencaValor;
		private System.Windows.Forms.DataGridViewTextBoxColumn DiferencaPerc;
		private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenValorOrdenacaoPadrao;
		private System.Windows.Forms.ComboBox cbPlataforma;
		private System.Windows.Forms.Label lblPlataforma;
	}
}