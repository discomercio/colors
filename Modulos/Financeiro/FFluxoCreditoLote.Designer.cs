namespace Financeiro
{
	partial class FFluxoCreditoLote
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoCreditoLote));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblContaCorrente = new System.Windows.Forms.Label();
			this.cbContaCorrente = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasEmpresa = new System.Windows.Forms.Label();
			this.cbPlanoContasEmpresa = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasConta = new System.Windows.Forms.Label();
			this.cbPlanoContasConta = new System.Windows.Forms.ComboBox();
			this.lblDataCompetencia = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.lblDescricao = new System.Windows.Forms.Label();
			this.txtDescricao = new System.Windows.Forms.TextBox();
			this.txtDataCompetencia = new System.Windows.Forms.TextBox();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.btnGravar = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.gboxBotoes = new System.Windows.Forms.GroupBox();
			this.gboxDefault = new System.Windows.Forms.GroupBox();
			this.lblTitValorTotal = new System.Windows.Forms.Label();
			this.lblValorTotal = new System.Windows.Forms.Label();
			this.lblQtdeLancamentos = new System.Windows.Forms.Label();
			this.lblTitQtdeLancamentos = new System.Windows.Forms.Label();
			this.gboxLote = new System.Windows.Forms.GroupBox();
			this.grdLote = new Financeiro.DataGridViewEditavel();
			this.colPlanoContasConta = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.colDataCompetencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colValorLancto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colCnpjCpf = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colNF = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDescricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxBotoes.SuspendLayout();
			this.gboxDefault.SuspendLayout();
			this.gboxLote.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdLote)).BeginInit();
			this.SuspendLayout();
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.lblTitValorTotal);
			this.pnCampos.Controls.Add(this.lblValorTotal);
			this.pnCampos.Controls.Add(this.lblQtdeLancamentos);
			this.pnCampos.Controls.Add(this.lblTitQtdeLancamentos);
			this.pnCampos.Controls.Add(this.gboxLote);
			this.pnCampos.Controls.Add(this.gboxDefault);
			this.pnCampos.Controls.Add(this.gboxBotoes);
			this.pnCampos.Controls.Add(this.lblTitulo);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.TabIndex = 1;
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 0;
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
			this.lblTitulo.TabIndex = 0;
			this.lblTitulo.Text = "Lançamento de Crédito em Lote";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContaCorrente
			// 
			this.lblContaCorrente.AutoSize = true;
			this.lblContaCorrente.Location = new System.Drawing.Point(69, 27);
			this.lblContaCorrente.Name = "lblContaCorrente";
			this.lblContaCorrente.Size = new System.Drawing.Size(78, 13);
			this.lblContaCorrente.TabIndex = 1;
			this.lblContaCorrente.Text = "Conta Corrente";
			// 
			// cbContaCorrente
			// 
			this.cbContaCorrente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbContaCorrente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbContaCorrente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbContaCorrente.FormattingEnabled = true;
			this.cbContaCorrente.Location = new System.Drawing.Point(153, 22);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(518, 24);
			this.cbContaCorrente.TabIndex = 0;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(99, 58);
			this.lblPlanoContasEmpresa.Name = "lblPlanoContasEmpresa";
			this.lblPlanoContasEmpresa.Size = new System.Drawing.Size(48, 13);
			this.lblPlanoContasEmpresa.TabIndex = 3;
			this.lblPlanoContasEmpresa.Text = "Empresa";
			// 
			// cbPlanoContasEmpresa
			// 
			this.cbPlanoContasEmpresa.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasEmpresa.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasEmpresa.FormattingEnabled = true;
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(153, 53);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasEmpresa.TabIndex = 1;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasConta
			// 
			this.lblPlanoContasConta.AutoSize = true;
			this.lblPlanoContasConta.Location = new System.Drawing.Point(67, 89);
			this.lblPlanoContasConta.Name = "lblPlanoContasConta";
			this.lblPlanoContasConta.Size = new System.Drawing.Size(80, 13);
			this.lblPlanoContasConta.TabIndex = 5;
			this.lblPlanoContasConta.Text = "Plano de Conta";
			// 
			// cbPlanoContasConta
			// 
			this.cbPlanoContasConta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasConta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasConta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasConta.FormattingEnabled = true;
			this.cbPlanoContasConta.Location = new System.Drawing.Point(153, 84);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasConta.TabIndex = 2;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblDataCompetencia
			// 
			this.lblDataCompetencia.AutoSize = true;
			this.lblDataCompetencia.Location = new System.Drawing.Point(37, 120);
			this.lblDataCompetencia.Name = "lblDataCompetencia";
			this.lblDataCompetencia.Size = new System.Drawing.Size(110, 13);
			this.lblDataCompetencia.TabIndex = 7;
			this.lblDataCompetencia.Text = "Data de Competência";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Location = new System.Drawing.Point(266, 120);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(54, 13);
			this.lblValor.TabIndex = 9;
			this.lblValor.Text = "Valor (R$)";
			// 
			// lblDescricao
			// 
			this.lblDescricao.AutoSize = true;
			this.lblDescricao.Location = new System.Drawing.Point(92, 150);
			this.lblDescricao.Name = "lblDescricao";
			this.lblDescricao.Size = new System.Drawing.Size(55, 13);
			this.lblDescricao.TabIndex = 11;
			this.lblDescricao.Text = "Descrição";
			// 
			// txtDescricao
			// 
			this.txtDescricao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDescricao.Location = new System.Drawing.Point(153, 145);
			this.txtDescricao.MaxLength = 40;
			this.txtDescricao.Name = "txtDescricao";
			this.txtDescricao.Size = new System.Drawing.Size(518, 23);
			this.txtDescricao.TabIndex = 6;
			this.txtDescricao.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescricao_KeyDown);
			this.txtDescricao.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescricao_KeyPress);
			// 
			// txtDataCompetencia
			// 
			this.txtDataCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetencia.Location = new System.Drawing.Point(153, 115);
			this.txtDataCompetencia.MaxLength = 10;
			this.txtDataCompetencia.Name = "txtDataCompetencia";
			this.txtDataCompetencia.Size = new System.Drawing.Size(91, 23);
			this.txtDataCompetencia.TabIndex = 3;
			this.txtDataCompetencia.Text = "01/01/2000";
			this.txtDataCompetencia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCompetencia.Enter += new System.EventHandler(this.txtDataCompetencia_Enter);
			this.txtDataCompetencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetencia_KeyDown);
			this.txtDataCompetencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetencia_KeyPress);
			this.txtDataCompetencia.Leave += new System.EventHandler(this.txtDataCompetencia_Leave);
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(326, 115);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(111, 23);
			this.txtValor.TabIndex = 4;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			// 
			// btnGravar
			// 
			this.btnGravar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnGravar.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGravar.Image = ((System.Drawing.Image)(resources.GetObject("btnGravar.Image")));
			this.btnGravar.Location = new System.Drawing.Point(45, 75);
			this.btnGravar.Name = "btnGravar";
			this.btnGravar.Size = new System.Drawing.Size(129, 40);
			this.btnGravar.TabIndex = 0;
			this.btnGravar.Text = "&Gravar Lote";
			this.btnGravar.UseVisualStyleBackColor = true;
			this.btnGravar.Click += new System.EventHandler(this.btnGravar_Click);
			// 
			// btnLimpar
			// 
			this.btnLimpar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnLimpar.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnLimpar.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpar.Image")));
			this.btnLimpar.Location = new System.Drawing.Point(45, 128);
			this.btnLimpar.Name = "btnLimpar";
			this.btnLimpar.Size = new System.Drawing.Size(129, 40);
			this.btnLimpar.TabIndex = 1;
			this.btnLimpar.Text = "&Limpar";
			this.btnLimpar.UseVisualStyleBackColor = true;
			this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
			// 
			// lblCnpjCpf
			// 
			this.lblCnpjCpf.AutoSize = true;
			this.lblCnpjCpf.Location = new System.Drawing.Point(461, 120);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(59, 13);
			this.lblCnpjCpf.TabIndex = 12;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(526, 115);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(145, 23);
			this.txtCnpjCpf.TabIndex = 5;
			this.txtCnpjCpf.Text = "00.000.000/0000-00";
			this.txtCnpjCpf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCnpjCpf.Enter += new System.EventHandler(this.txtCnpjCpf_Enter);
			this.txtCnpjCpf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCnpjCpf_KeyDown);
			this.txtCnpjCpf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCnpjCpf_KeyPress);
			this.txtCnpjCpf.Leave += new System.EventHandler(this.txtCnpjCpf_Leave);
			// 
			// lblContador
			// 
			this.lblContador.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 22F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblContador.Image = ((System.Drawing.Image)(resources.GetObject("lblContador.Image")));
			this.lblContador.Location = new System.Drawing.Point(72, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(75, 40);
			this.lblContador.TabIndex = 13;
			this.lblContador.Text = "999";
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// gboxBotoes
			// 
			this.gboxBotoes.Controls.Add(this.lblContador);
			this.gboxBotoes.Controls.Add(this.btnLimpar);
			this.gboxBotoes.Controls.Add(this.btnGravar);
			this.gboxBotoes.Location = new System.Drawing.Point(786, 48);
			this.gboxBotoes.Name = "gboxBotoes";
			this.gboxBotoes.Size = new System.Drawing.Size(219, 181);
			this.gboxBotoes.TabIndex = 2;
			this.gboxBotoes.TabStop = false;
			// 
			// gboxDefault
			// 
			this.gboxDefault.Controls.Add(this.txtCnpjCpf);
			this.gboxDefault.Controls.Add(this.lblCnpjCpf);
			this.gboxDefault.Controls.Add(this.txtValor);
			this.gboxDefault.Controls.Add(this.txtDataCompetencia);
			this.gboxDefault.Controls.Add(this.txtDescricao);
			this.gboxDefault.Controls.Add(this.lblDescricao);
			this.gboxDefault.Controls.Add(this.lblValor);
			this.gboxDefault.Controls.Add(this.lblDataCompetencia);
			this.gboxDefault.Controls.Add(this.cbPlanoContasConta);
			this.gboxDefault.Controls.Add(this.lblPlanoContasConta);
			this.gboxDefault.Controls.Add(this.cbPlanoContasEmpresa);
			this.gboxDefault.Controls.Add(this.lblPlanoContasEmpresa);
			this.gboxDefault.Controls.Add(this.cbContaCorrente);
			this.gboxDefault.Controls.Add(this.lblContaCorrente);
			this.gboxDefault.Location = new System.Drawing.Point(10, 48);
			this.gboxDefault.Name = "gboxDefault";
			this.gboxDefault.Size = new System.Drawing.Size(731, 181);
			this.gboxDefault.TabIndex = 0;
			this.gboxDefault.TabStop = false;
			this.gboxDefault.Text = "Valores Padrão";
			// 
			// lblTitValorTotal
			// 
			this.lblTitValorTotal.AutoSize = true;
			this.lblTitValorTotal.Location = new System.Drawing.Point(339, 585);
			this.lblTitValorTotal.Name = "lblTitValorTotal";
			this.lblTitValorTotal.Size = new System.Drawing.Size(58, 13);
			this.lblTitValorTotal.TabIndex = 24;
			this.lblTitValorTotal.Text = "Valor Total";
			// 
			// lblValorTotal
			// 
			this.lblValorTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValorTotal.Location = new System.Drawing.Point(386, 585);
			this.lblValorTotal.Name = "lblValorTotal";
			this.lblValorTotal.Size = new System.Drawing.Size(112, 13);
			this.lblValorTotal.TabIndex = 25;
			this.lblValorTotal.Text = "999.999.999,00";
			this.lblValorTotal.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblQtdeLancamentos
			// 
			this.lblQtdeLancamentos.AutoSize = true;
			this.lblQtdeLancamentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeLancamentos.Location = new System.Drawing.Point(147, 585);
			this.lblQtdeLancamentos.Name = "lblQtdeLancamentos";
			this.lblQtdeLancamentos.Size = new System.Drawing.Size(21, 13);
			this.lblQtdeLancamentos.TabIndex = 23;
			this.lblQtdeLancamentos.Text = "00";
			// 
			// lblTitQtdeLancamentos
			// 
			this.lblTitQtdeLancamentos.AutoSize = true;
			this.lblTitQtdeLancamentos.Location = new System.Drawing.Point(47, 585);
			this.lblTitQtdeLancamentos.Name = "lblTitQtdeLancamentos";
			this.lblTitQtdeLancamentos.Size = new System.Drawing.Size(97, 13);
			this.lblTitQtdeLancamentos.TabIndex = 22;
			this.lblTitQtdeLancamentos.Text = "Qtde Lançamentos";
			// 
			// gboxLote
			// 
			this.gboxLote.Controls.Add(this.grdLote);
			this.gboxLote.Location = new System.Drawing.Point(10, 247);
			this.gboxLote.Name = "gboxLote";
			this.gboxLote.Size = new System.Drawing.Size(995, 332);
			this.gboxLote.TabIndex = 1;
			this.gboxLote.TabStop = false;
			this.gboxLote.Text = "Lançamentos";
			// 
			// grdLote
			// 
			this.grdLote.AllowUserToAddRows = false;
			this.grdLote.AllowUserToDeleteRows = false;
			this.grdLote.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdLote.BorderStyle = System.Windows.Forms.BorderStyle.None;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdLote.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdLote.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdLote.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colPlanoContasConta,
            this.colDataCompetencia,
            this.colValorLancto,
            this.colCnpjCpf,
            this.colNF,
            this.colDescricao});
			this.grdLote.Dock = System.Windows.Forms.DockStyle.Fill;
			this.grdLote.Location = new System.Drawing.Point(3, 16);
			this.grdLote.MultiSelect = false;
			this.grdLote.Name = "grdLote";
			this.grdLote.RowHeadersWidth = 35;
			this.grdLote.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdLote.Size = new System.Drawing.Size(989, 313);
			this.grdLote.TabIndex = 0;
			this.grdLote.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.grdLote_CellValidating);
			this.grdLote.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdLote_CellValueChanged);
			this.grdLote.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.grdLote_EditingControlShowing);
			// 
			// colPlanoContasConta
			// 
			dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colPlanoContasConta.DefaultCellStyle = dataGridViewCellStyle2;
			this.colPlanoContasConta.Frozen = true;
			this.colPlanoContasConta.HeaderText = "Plano de Conta";
			this.colPlanoContasConta.MinimumWidth = 250;
			this.colPlanoContasConta.Name = "colPlanoContasConta";
			this.colPlanoContasConta.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.colPlanoContasConta.Width = 250;
			// 
			// colDataCompetencia
			// 
			this.colDataCompetencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colDataCompetencia.DefaultCellStyle = dataGridViewCellStyle3;
			this.colDataCompetencia.HeaderText = "Competência";
			this.colDataCompetencia.MaxInputLength = 10;
			this.colDataCompetencia.MinimumWidth = 100;
			this.colDataCompetencia.Name = "colDataCompetencia";
			this.colDataCompetencia.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colDataCompetencia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// colValorLancto
			// 
			this.colValorLancto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colValorLancto.DefaultCellStyle = dataGridViewCellStyle4;
			this.colValorLancto.HeaderText = "Valor";
			this.colValorLancto.MaxInputLength = 20;
			this.colValorLancto.MinimumWidth = 100;
			this.colValorLancto.Name = "colValorLancto";
			this.colValorLancto.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colValorLancto.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// colCnpjCpf
			// 
			this.colCnpjCpf.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colCnpjCpf.DefaultCellStyle = dataGridViewCellStyle5;
			this.colCnpjCpf.HeaderText = "CNPJ/CPF";
			this.colCnpjCpf.MaxInputLength = 18;
			this.colCnpjCpf.MinimumWidth = 140;
			this.colCnpjCpf.Name = "colCnpjCpf";
			this.colCnpjCpf.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colCnpjCpf.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colCnpjCpf.Width = 140;
			// 
			// colNF
			// 
			this.colNF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colNF.DefaultCellStyle = dataGridViewCellStyle6;
			this.colNF.HeaderText = "NF";
			this.colNF.MaxInputLength = 13;
			this.colNF.MinimumWidth = 80;
			this.colNF.Name = "colNF";
			this.colNF.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colNF.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colNF.Width = 80;
			// 
			// colDescricao
			// 
			this.colDescricao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colDescricao.DefaultCellStyle = dataGridViewCellStyle7;
			this.colDescricao.HeaderText = "Descrição";
			this.colDescricao.MaxInputLength = 40;
			this.colDescricao.Name = "colDescricao";
			this.colDescricao.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// FFluxoCreditoLote
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FFluxoCreditoLote";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FFluxoCreditoLote_FormClosing);
			this.Load += new System.EventHandler(this.FFluxoCredito_Load);
			this.Shown += new System.EventHandler(this.FFluxoCredito_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FFluxoCreditoLote_KeyDown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FFluxoCreditoLote_KeyPress);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxBotoes.ResumeLayout(false);
			this.gboxDefault.ResumeLayout(false);
			this.gboxDefault.PerformLayout();
			this.gboxLote.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdLote)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.ComboBox cbContaCorrente;
		private System.Windows.Forms.Label lblContaCorrente;
		private System.Windows.Forms.ComboBox cbPlanoContasConta;
		private System.Windows.Forms.Label lblPlanoContasConta;
		private System.Windows.Forms.ComboBox cbPlanoContasEmpresa;
		private System.Windows.Forms.Label lblPlanoContasEmpresa;
		private System.Windows.Forms.Label lblDataCompetencia;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtDescricao;
		private System.Windows.Forms.Label lblDescricao;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.TextBox txtDataCompetencia;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.Button btnGravar;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox gboxBotoes;
		private System.Windows.Forms.GroupBox gboxDefault;
		private System.Windows.Forms.Label lblTitValorTotal;
		private System.Windows.Forms.Label lblValorTotal;
		private System.Windows.Forms.Label lblQtdeLancamentos;
		private System.Windows.Forms.Label lblTitQtdeLancamentos;
		private System.Windows.Forms.GroupBox gboxLote;
		private DataGridViewEditavel grdLote;
		private System.Windows.Forms.DataGridViewComboBoxColumn colPlanoContasConta;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDataCompetencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn colValorLancto;
		private System.Windows.Forms.DataGridViewTextBoxColumn colCnpjCpf;
		private System.Windows.Forms.DataGridViewTextBoxColumn colNF;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDescricao;
	}
}
