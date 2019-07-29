namespace Financeiro
{
	partial class FCepPesquisa
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FCepPesquisa));
			this.gboxPesquisaPorCep = new System.Windows.Forms.GroupBox();
			this.btnPesquisaPorCep = new System.Windows.Forms.Button();
			this.txtCep = new System.Windows.Forms.TextBox();
			this.lblTitCep = new System.Windows.Forms.Label();
			this.gboxPesquisaPorEndereco = new System.Windows.Forms.GroupBox();
			this.btnPesquisarPorEndereco = new System.Windows.Forms.Button();
			this.txtEndereco = new System.Windows.Forms.TextBox();
			this.lblTitEndereco = new System.Windows.Forms.Label();
			this.lblTitLocalidade = new System.Windows.Forms.Label();
			this.cbLocalidade = new System.Windows.Forms.ComboBox();
			this.lblTitUF = new System.Windows.Forms.Label();
			this.cbUF = new System.Windows.Forms.ComboBox();
			this.gboxResultado = new System.Windows.Forms.GroupBox();
			this.txtNumeroOuComplemento = new System.Windows.Forms.TextBox();
			this.lblTitNumeroOuComplemento = new System.Windows.Forms.Label();
			this.grdResultado = new System.Windows.Forms.DataGridView();
			this.check = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.cep = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.uf = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cidade = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.bairro = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.logradouro = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.complemento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnResultadoBotoes = new System.Windows.Forms.Panel();
			this.btnConfirma = new System.Windows.Forms.Button();
			this.btnCancela = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxPesquisaPorCep.SuspendLayout();
			this.gboxPesquisaPorEndereco.SuspendLayout();
			this.gboxResultado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdResultado)).BeginInit();
			this.pnResultadoBotoes.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxResultado);
			this.pnCampos.Controls.Add(this.gboxPesquisaPorEndereco);
			this.pnCampos.Controls.Add(this.gboxPesquisaPorCep);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// gboxPesquisaPorCep
			// 
			this.gboxPesquisaPorCep.Controls.Add(this.btnPesquisaPorCep);
			this.gboxPesquisaPorCep.Controls.Add(this.txtCep);
			this.gboxPesquisaPorCep.Controls.Add(this.lblTitCep);
			this.gboxPesquisaPorCep.Location = new System.Drawing.Point(22, 19);
			this.gboxPesquisaPorCep.Name = "gboxPesquisaPorCep";
			this.gboxPesquisaPorCep.Size = new System.Drawing.Size(284, 64);
			this.gboxPesquisaPorCep.TabIndex = 0;
			this.gboxPesquisaPorCep.TabStop = false;
			this.gboxPesquisaPorCep.Text = "Pesquisa por CEP";
			// 
			// btnPesquisaPorCep
			// 
			this.btnPesquisaPorCep.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisaPorCep.Image")));
			this.btnPesquisaPorCep.Location = new System.Drawing.Point(177, 25);
			this.btnPesquisaPorCep.Name = "btnPesquisaPorCep";
			this.btnPesquisaPorCep.Size = new System.Drawing.Size(80, 25);
			this.btnPesquisaPorCep.TabIndex = 1;
			this.btnPesquisaPorCep.Text = "Pesquisar";
			this.btnPesquisaPorCep.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
			this.btnPesquisaPorCep.UseVisualStyleBackColor = true;
			this.btnPesquisaPorCep.Click += new System.EventHandler(this.btnPesquisaPorCep_Click);
			// 
			// txtCep
			// 
			this.txtCep.BackColor = System.Drawing.SystemColors.Window;
			this.txtCep.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCep.Location = new System.Drawing.Point(67, 28);
			this.txtCep.MaxLength = 9;
			this.txtCep.Name = "txtCep";
			this.txtCep.Size = new System.Drawing.Size(104, 20);
			this.txtCep.TabIndex = 0;
			this.txtCep.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCep.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCep_KeyDown);
			this.txtCep.Leave += new System.EventHandler(this.txtCep_Leave);
			this.txtCep.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCep_KeyPress);
			this.txtCep.Enter += new System.EventHandler(this.txtCep_Enter);
			// 
			// lblTitCep
			// 
			this.lblTitCep.AutoSize = true;
			this.lblTitCep.Location = new System.Drawing.Point(33, 31);
			this.lblTitCep.Name = "lblTitCep";
			this.lblTitCep.Size = new System.Drawing.Size(28, 13);
			this.lblTitCep.TabIndex = 27;
			this.lblTitCep.Text = "CEP";
			// 
			// gboxPesquisaPorEndereco
			// 
			this.gboxPesquisaPorEndereco.Controls.Add(this.btnPesquisarPorEndereco);
			this.gboxPesquisaPorEndereco.Controls.Add(this.txtEndereco);
			this.gboxPesquisaPorEndereco.Controls.Add(this.lblTitEndereco);
			this.gboxPesquisaPorEndereco.Controls.Add(this.lblTitLocalidade);
			this.gboxPesquisaPorEndereco.Controls.Add(this.cbLocalidade);
			this.gboxPesquisaPorEndereco.Controls.Add(this.lblTitUF);
			this.gboxPesquisaPorEndereco.Controls.Add(this.cbUF);
			this.gboxPesquisaPorEndereco.Location = new System.Drawing.Point(22, 104);
			this.gboxPesquisaPorEndereco.Name = "gboxPesquisaPorEndereco";
			this.gboxPesquisaPorEndereco.Size = new System.Drawing.Size(693, 105);
			this.gboxPesquisaPorEndereco.TabIndex = 1;
			this.gboxPesquisaPorEndereco.TabStop = false;
			this.gboxPesquisaPorEndereco.Text = "Pesquisa por Endereço";
			// 
			// btnPesquisarPorEndereco
			// 
			this.btnPesquisarPorEndereco.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisarPorEndereco.Image")));
			this.btnPesquisarPorEndereco.Location = new System.Drawing.Point(599, 67);
			this.btnPesquisarPorEndereco.Name = "btnPesquisarPorEndereco";
			this.btnPesquisarPorEndereco.Size = new System.Drawing.Size(80, 25);
			this.btnPesquisarPorEndereco.TabIndex = 3;
			this.btnPesquisarPorEndereco.Text = "Pesquisar";
			this.btnPesquisarPorEndereco.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
			this.btnPesquisarPorEndereco.UseVisualStyleBackColor = true;
			this.btnPesquisarPorEndereco.Click += new System.EventHandler(this.btnPesquisarPorEndereco_Click);
			// 
			// txtEndereco
			// 
			this.txtEndereco.Location = new System.Drawing.Point(67, 70);
			this.txtEndereco.Name = "txtEndereco";
			this.txtEndereco.Size = new System.Drawing.Size(526, 20);
			this.txtEndereco.TabIndex = 2;
			this.txtEndereco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEndereco_KeyDown);
			this.txtEndereco.Leave += new System.EventHandler(this.txtEndereco_Leave);
			this.txtEndereco.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEndereco_KeyPress);
			this.txtEndereco.Enter += new System.EventHandler(this.txtEndereco_Enter);
			// 
			// lblTitEndereco
			// 
			this.lblTitEndereco.AutoSize = true;
			this.lblTitEndereco.Location = new System.Drawing.Point(8, 73);
			this.lblTitEndereco.Name = "lblTitEndereco";
			this.lblTitEndereco.Size = new System.Drawing.Size(53, 13);
			this.lblTitEndereco.TabIndex = 31;
			this.lblTitEndereco.Text = "Endereço";
			// 
			// lblTitLocalidade
			// 
			this.lblTitLocalidade.AutoSize = true;
			this.lblTitLocalidade.Location = new System.Drawing.Point(175, 31);
			this.lblTitLocalidade.Name = "lblTitLocalidade";
			this.lblTitLocalidade.Size = new System.Drawing.Size(59, 13);
			this.lblTitLocalidade.TabIndex = 30;
			this.lblTitLocalidade.Text = "Localidade";
			// 
			// cbLocalidade
			// 
			this.cbLocalidade.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbLocalidade.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbLocalidade.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbLocalidade.FormattingEnabled = true;
			this.cbLocalidade.Location = new System.Drawing.Point(240, 28);
			this.cbLocalidade.Name = "cbLocalidade";
			this.cbLocalidade.Size = new System.Drawing.Size(353, 21);
			this.cbLocalidade.TabIndex = 1;
			this.cbLocalidade.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbLocalidade_KeyDown);
			// 
			// lblTitUF
			// 
			this.lblTitUF.AutoSize = true;
			this.lblTitUF.Location = new System.Drawing.Point(40, 31);
			this.lblTitUF.Name = "lblTitUF";
			this.lblTitUF.Size = new System.Drawing.Size(21, 13);
			this.lblTitUF.TabIndex = 28;
			this.lblTitUF.Text = "UF";
			// 
			// cbUF
			// 
			this.cbUF.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbUF.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbUF.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbUF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbUF.FormattingEnabled = true;
			this.cbUF.Location = new System.Drawing.Point(67, 28);
			this.cbUF.Name = "cbUF";
			this.cbUF.Size = new System.Drawing.Size(55, 21);
			this.cbUF.TabIndex = 0;
			this.cbUF.SelectionChangeCommitted += new System.EventHandler(this.cbUF_SelectionChangeCommitted);
			this.cbUF.Leave += new System.EventHandler(this.cbUF_Leave);
			this.cbUF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbUF_KeyDown);
			// 
			// gboxResultado
			// 
			this.gboxResultado.Controls.Add(this.txtNumeroOuComplemento);
			this.gboxResultado.Controls.Add(this.lblTitNumeroOuComplemento);
			this.gboxResultado.Controls.Add(this.grdResultado);
			this.gboxResultado.Controls.Add(this.pnResultadoBotoes);
			this.gboxResultado.Location = new System.Drawing.Point(22, 230);
			this.gboxResultado.Name = "gboxResultado";
			this.gboxResultado.Size = new System.Drawing.Size(971, 363);
			this.gboxResultado.TabIndex = 2;
			this.gboxResultado.TabStop = false;
			this.gboxResultado.Text = "Resultado";
			// 
			// txtNumeroOuComplemento
			// 
			this.txtNumeroOuComplemento.Location = new System.Drawing.Point(384, 291);
			this.txtNumeroOuComplemento.Name = "txtNumeroOuComplemento";
			this.txtNumeroOuComplemento.Size = new System.Drawing.Size(302, 20);
			this.txtNumeroOuComplemento.TabIndex = 1;
			this.txtNumeroOuComplemento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumeroOuComplemento_KeyDown);
			this.txtNumeroOuComplemento.Leave += new System.EventHandler(this.txtNumeroOuComplemento_Leave);
			this.txtNumeroOuComplemento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumeroOuComplemento_KeyPress);
			this.txtNumeroOuComplemento.Enter += new System.EventHandler(this.txtNumeroOuComplemento_Enter);
			// 
			// lblTitNumeroOuComplemento
			// 
			this.lblTitNumeroOuComplemento.AutoSize = true;
			this.lblTitNumeroOuComplemento.Location = new System.Drawing.Point(284, 294);
			this.lblTitNumeroOuComplemento.Name = "lblTitNumeroOuComplemento";
			this.lblTitNumeroOuComplemento.Size = new System.Drawing.Size(94, 13);
			this.lblTitNumeroOuComplemento.TabIndex = 33;
			this.lblTitNumeroOuComplemento.Text = "Nº / Complemento";
			// 
			// grdResultado
			// 
			this.grdResultado.AllowUserToAddRows = false;
			this.grdResultado.AllowUserToDeleteRows = false;
			this.grdResultado.AllowUserToResizeRows = false;
			this.grdResultado.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdResultado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdResultado.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdResultado.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdResultado.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.check,
            this.cep,
            this.uf,
            this.cidade,
            this.bairro,
            this.logradouro,
            this.complemento});
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdResultado.DefaultCellStyle = dataGridViewCellStyle8;
			this.grdResultado.Location = new System.Drawing.Point(11, 21);
			this.grdResultado.MultiSelect = false;
			this.grdResultado.Name = "grdResultado";
			this.grdResultado.RowHeadersVisible = false;
			this.grdResultado.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdResultado.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdResultado.Size = new System.Drawing.Size(949, 261);
			this.grdResultado.StandardTab = true;
			this.grdResultado.TabIndex = 0;
			// 
			// check
			// 
			this.check.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.check.HeaderText = "";
			this.check.MinimumWidth = 20;
			this.check.Name = "check";
			this.check.ReadOnly = true;
			this.check.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.check.Visible = false;
			this.check.Width = 20;
			// 
			// cep
			// 
			this.cep.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.cep.DefaultCellStyle = dataGridViewCellStyle2;
			this.cep.HeaderText = "CEP";
			this.cep.MinimumWidth = 80;
			this.cep.Name = "cep";
			this.cep.ReadOnly = true;
			this.cep.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.cep.Width = 80;
			// 
			// uf
			// 
			this.uf.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.uf.DefaultCellStyle = dataGridViewCellStyle3;
			this.uf.HeaderText = "UF";
			this.uf.MinimumWidth = 50;
			this.uf.Name = "uf";
			this.uf.ReadOnly = true;
			this.uf.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.uf.Width = 50;
			// 
			// cidade
			// 
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.cidade.DefaultCellStyle = dataGridViewCellStyle4;
			this.cidade.HeaderText = "Cidade";
			this.cidade.MinimumWidth = 60;
			this.cidade.Name = "cidade";
			this.cidade.ReadOnly = true;
			this.cidade.Width = 180;
			// 
			// bairro
			// 
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.bairro.DefaultCellStyle = dataGridViewCellStyle5;
			this.bairro.HeaderText = "Bairro";
			this.bairro.MinimumWidth = 60;
			this.bairro.Name = "bairro";
			this.bairro.ReadOnly = true;
			this.bairro.Width = 180;
			// 
			// logradouro
			// 
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.logradouro.DefaultCellStyle = dataGridViewCellStyle6;
			this.logradouro.HeaderText = "Logradouro";
			this.logradouro.MinimumWidth = 60;
			this.logradouro.Name = "logradouro";
			this.logradouro.ReadOnly = true;
			this.logradouro.Width = 220;
			// 
			// complemento
			// 
			this.complemento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.complemento.DefaultCellStyle = dataGridViewCellStyle7;
			this.complemento.HeaderText = "Complemento";
			this.complemento.MinimumWidth = 60;
			this.complemento.Name = "complemento";
			this.complemento.ReadOnly = true;
			// 
			// pnResultadoBotoes
			// 
			this.pnResultadoBotoes.Controls.Add(this.btnConfirma);
			this.pnResultadoBotoes.Controls.Add(this.btnCancela);
			this.pnResultadoBotoes.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.pnResultadoBotoes.Location = new System.Drawing.Point(3, 317);
			this.pnResultadoBotoes.Name = "pnResultadoBotoes";
			this.pnResultadoBotoes.Size = new System.Drawing.Size(965, 43);
			this.pnResultadoBotoes.TabIndex = 2;
			// 
			// btnConfirma
			// 
			this.btnConfirma.Image = ((System.Drawing.Image)(resources.GetObject("btnConfirma.Image")));
			this.btnConfirma.Location = new System.Drawing.Point(564, 6);
			this.btnConfirma.Name = "btnConfirma";
			this.btnConfirma.Size = new System.Drawing.Size(87, 31);
			this.btnConfirma.TabIndex = 1;
			this.btnConfirma.Text = "Co&nfirma";
			this.btnConfirma.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
			this.btnConfirma.UseVisualStyleBackColor = true;
			this.btnConfirma.Click += new System.EventHandler(this.btnConfirma_Click);
			// 
			// btnCancela
			// 
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(314, 6);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(87, 31);
			this.btnCancela.TabIndex = 0;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// FCepPesquisa
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FCepPesquisa";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FCepPesquisa_Load);
			this.Shown += new System.EventHandler(this.FCepPesquisa_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxPesquisaPorCep.ResumeLayout(false);
			this.gboxPesquisaPorCep.PerformLayout();
			this.gboxPesquisaPorEndereco.ResumeLayout(false);
			this.gboxPesquisaPorEndereco.PerformLayout();
			this.gboxResultado.ResumeLayout(false);
			this.gboxResultado.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdResultado)).EndInit();
			this.pnResultadoBotoes.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxPesquisaPorCep;
		private System.Windows.Forms.Button btnPesquisaPorCep;
		private System.Windows.Forms.TextBox txtCep;
		private System.Windows.Forms.Label lblTitCep;
		private System.Windows.Forms.GroupBox gboxPesquisaPorEndereco;
		private System.Windows.Forms.Label lblTitLocalidade;
		private System.Windows.Forms.ComboBox cbLocalidade;
		private System.Windows.Forms.Label lblTitUF;
		private System.Windows.Forms.ComboBox cbUF;
		private System.Windows.Forms.TextBox txtEndereco;
		private System.Windows.Forms.Label lblTitEndereco;
		private System.Windows.Forms.GroupBox gboxResultado;
		private System.Windows.Forms.Panel pnResultadoBotoes;
		private System.Windows.Forms.Button btnConfirma;
		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.DataGridView grdResultado;
		private System.Windows.Forms.Button btnPesquisarPorEndereco;
		private System.Windows.Forms.DataGridViewCheckBoxColumn check;
		private System.Windows.Forms.DataGridViewTextBoxColumn cep;
		private System.Windows.Forms.DataGridViewTextBoxColumn uf;
		private System.Windows.Forms.DataGridViewTextBoxColumn cidade;
		private System.Windows.Forms.DataGridViewTextBoxColumn bairro;
		private System.Windows.Forms.DataGridViewTextBoxColumn logradouro;
		private System.Windows.Forms.DataGridViewTextBoxColumn complemento;
		private System.Windows.Forms.TextBox txtNumeroOuComplemento;
		private System.Windows.Forms.Label lblTitNumeroOuComplemento;
	}
}
