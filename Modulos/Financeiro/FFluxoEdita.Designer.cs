namespace Financeiro
{
	partial class FFluxoEdita
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoEdita));
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
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.btnAtualizar = new System.Windows.Forms.Button();
			this.btnExcluir = new System.Windows.Forms.Button();
			this.lblTitCadastradoEm = new System.Windows.Forms.Label();
			this.lblCadastradoEm = new System.Windows.Forms.Label();
			this.lblCadastradoPor = new System.Windows.Forms.Label();
			this.lblTitCadastradoPor = new System.Windows.Forms.Label();
			this.lblCadastradoModo = new System.Windows.Forms.Label();
			this.lblTitCadastradoModo = new System.Windows.Forms.Label();
			this.lblAlteradoPor = new System.Windows.Forms.Label();
			this.lblTitAlteradoPor = new System.Windows.Forms.Label();
			this.lblAlteradoEm = new System.Windows.Forms.Label();
			this.lblTitAlteradoEm = new System.Windows.Forms.Label();
			this.lblNatureza = new System.Windows.Forms.Label();
			this.lblTitNatureza = new System.Windows.Forms.Label();
			this.cbStSemEfeito = new System.Windows.Forms.ComboBox();
			this.lblTitStSemEfeito = new System.Windows.Forms.Label();
			this.cbCtrlPagtoStatus = new System.Windows.Forms.ComboBox();
			this.lblTitCtrlPagtoStatus = new System.Windows.Forms.Label();
			this.gboxCamposSistema = new System.Windows.Forms.GroupBox();
			this.gboxCamposLancamento = new System.Windows.Forms.GroupBox();
			this.txtComp2 = new System.Windows.Forms.TextBox();
			this.lblTitComp2 = new System.Windows.Forms.Label();
			this.cbStConfirmacaoPendente = new System.Windows.Forms.ComboBox();
			this.lblTitStConfirmacaoPendente = new System.Windows.Forms.Label();
			this.gboxDadosCliente = new System.Windows.Forms.GroupBox();
			this.lblNome = new System.Windows.Forms.Label();
			this.lblTitNome = new System.Windows.Forms.Label();
			this.txtNF = new System.Windows.Forms.TextBox();
			this.lblNF = new System.Windows.Forms.Label();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxCamposSistema.SuspendLayout();
			this.gboxCamposLancamento.SuspendLayout();
			this.gboxDadosCliente.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnExcluir);
			this.pnBotoes.Controls.Add(this.btnAtualizar);
			this.pnBotoes.Size = new System.Drawing.Size(675, 55);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnAtualizar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnExcluir, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxDadosCliente);
			this.pnCampos.Controls.Add(this.gboxCamposLancamento);
			this.pnCampos.Controls.Add(this.gboxCamposSistema);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Size = new System.Drawing.Size(675, 486);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(626, 4);
			this.btnFechar.TabIndex = 3;
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(581, 4);
			this.btnSobre.TabIndex = 2;
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
			this.lblTitulo.Size = new System.Drawing.Size(671, 40);
			this.lblTitulo.TabIndex = 0;
			this.lblTitulo.Text = "Edição de Lançamento";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContaCorrente
			// 
			this.lblContaCorrente.AutoSize = true;
			this.lblContaCorrente.Location = new System.Drawing.Point(38, 93);
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
			this.cbContaCorrente.Location = new System.Drawing.Point(122, 88);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(518, 24);
			this.cbContaCorrente.TabIndex = 3;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(68, 131);
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
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(122, 126);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasEmpresa.TabIndex = 4;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasConta
			// 
			this.lblPlanoContasConta.AutoSize = true;
			this.lblPlanoContasConta.Location = new System.Drawing.Point(36, 169);
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
			this.cbPlanoContasConta.Location = new System.Drawing.Point(122, 164);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasConta.TabIndex = 5;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblDataCompetencia
			// 
			this.lblDataCompetencia.AutoSize = true;
			this.lblDataCompetencia.Location = new System.Drawing.Point(21, 208);
			this.lblDataCompetencia.Name = "lblDataCompetencia";
			this.lblDataCompetencia.Size = new System.Drawing.Size(95, 13);
			this.lblDataCompetencia.TabIndex = 7;
			this.lblDataCompetencia.Text = "Data Competência";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Location = new System.Drawing.Point(62, 246);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(54, 13);
			this.lblValor.TabIndex = 9;
			this.lblValor.Text = "Valor (R$)";
			// 
			// lblDescricao
			// 
			this.lblDescricao.AutoSize = true;
			this.lblDescricao.Location = new System.Drawing.Point(61, 285);
			this.lblDescricao.Name = "lblDescricao";
			this.lblDescricao.Size = new System.Drawing.Size(55, 13);
			this.lblDescricao.TabIndex = 11;
			this.lblDescricao.Text = "Descrição";
			// 
			// txtDescricao
			// 
			this.txtDescricao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDescricao.Location = new System.Drawing.Point(122, 280);
			this.txtDescricao.MaxLength = 40;
			this.txtDescricao.Name = "txtDescricao";
			this.txtDescricao.Size = new System.Drawing.Size(518, 23);
			this.txtDescricao.TabIndex = 11;
			this.txtDescricao.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescricao_KeyDown);
			this.txtDescricao.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescricao_KeyPress);
			// 
			// txtDataCompetencia
			// 
			this.txtDataCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetencia.Location = new System.Drawing.Point(122, 203);
			this.txtDataCompetencia.MaxLength = 10;
			this.txtDataCompetencia.Name = "txtDataCompetencia";
			this.txtDataCompetencia.Size = new System.Drawing.Size(91, 23);
			this.txtDataCompetencia.TabIndex = 6;
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
			this.txtValor.Location = new System.Drawing.Point(122, 241);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(111, 23);
			this.txtValor.TabIndex = 8;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			// 
			// lblCnpjCpf
			// 
			this.lblCnpjCpf.AutoSize = true;
			this.lblCnpjCpf.Location = new System.Drawing.Point(257, 246);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(59, 13);
			this.lblCnpjCpf.TabIndex = 12;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(322, 241);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(145, 23);
			this.txtCnpjCpf.TabIndex = 9;
			this.txtCnpjCpf.Text = "00.000.000/0000-00";
			this.txtCnpjCpf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCnpjCpf.Enter += new System.EventHandler(this.txtCnpjCpf_Enter);
			this.txtCnpjCpf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCnpjCpf_KeyDown);
			this.txtCnpjCpf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCnpjCpf_KeyPress);
			this.txtCnpjCpf.Leave += new System.EventHandler(this.txtCnpjCpf_Leave);
			// 
			// btnAtualizar
			// 
			this.btnAtualizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnAtualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizar.Image")));
			this.btnAtualizar.Location = new System.Drawing.Point(491, 4);
			this.btnAtualizar.Name = "btnAtualizar";
			this.btnAtualizar.Size = new System.Drawing.Size(40, 44);
			this.btnAtualizar.TabIndex = 0;
			this.btnAtualizar.TabStop = false;
			this.btnAtualizar.UseVisualStyleBackColor = true;
			this.btnAtualizar.Click += new System.EventHandler(this.btnAtualizar_Click);
			// 
			// btnExcluir
			// 
			this.btnExcluir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnExcluir.Image = ((System.Drawing.Image)(resources.GetObject("btnExcluir.Image")));
			this.btnExcluir.Location = new System.Drawing.Point(536, 4);
			this.btnExcluir.Name = "btnExcluir";
			this.btnExcluir.Size = new System.Drawing.Size(40, 44);
			this.btnExcluir.TabIndex = 1;
			this.btnExcluir.TabStop = false;
			this.btnExcluir.UseVisualStyleBackColor = true;
			this.btnExcluir.Click += new System.EventHandler(this.btnExcluir_Click);
			// 
			// lblTitCadastradoEm
			// 
			this.lblTitCadastradoEm.AutoSize = true;
			this.lblTitCadastradoEm.Location = new System.Drawing.Point(38, 14);
			this.lblTitCadastradoEm.Name = "lblTitCadastradoEm";
			this.lblTitCadastradoEm.Size = new System.Drawing.Size(78, 13);
			this.lblTitCadastradoEm.TabIndex = 13;
			this.lblTitCadastradoEm.Text = "Cadastrado em";
			// 
			// lblCadastradoEm
			// 
			this.lblCadastradoEm.AutoSize = true;
			this.lblCadastradoEm.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCadastradoEm.Location = new System.Drawing.Point(119, 14);
			this.lblCadastradoEm.Name = "lblCadastradoEm";
			this.lblCadastradoEm.Size = new System.Drawing.Size(111, 13);
			this.lblCadastradoEm.TabIndex = 14;
			this.lblCadastradoEm.Text = "01/01/2000 12:00";
			// 
			// lblCadastradoPor
			// 
			this.lblCadastradoPor.AutoSize = true;
			this.lblCadastradoPor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCadastradoPor.Location = new System.Drawing.Point(352, 14);
			this.lblCadastradoPor.Name = "lblCadastradoPor";
			this.lblCadastradoPor.Size = new System.Drawing.Size(61, 13);
			this.lblCadastradoPor.TabIndex = 16;
			this.lblCadastradoPor.Text = "SISTEMA";
			// 
			// lblTitCadastradoPor
			// 
			this.lblTitCadastradoPor.AutoSize = true;
			this.lblTitCadastradoPor.Location = new System.Drawing.Point(270, 14);
			this.lblTitCadastradoPor.Name = "lblTitCadastradoPor";
			this.lblTitCadastradoPor.Size = new System.Drawing.Size(79, 13);
			this.lblTitCadastradoPor.TabIndex = 15;
			this.lblTitCadastradoPor.Text = "Cadastrado por";
			// 
			// lblCadastradoModo
			// 
			this.lblCadastradoModo.AutoSize = true;
			this.lblCadastradoModo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCadastradoModo.Location = new System.Drawing.Point(517, 14);
			this.lblCadastradoModo.Name = "lblCadastradoModo";
			this.lblCadastradoModo.Size = new System.Drawing.Size(61, 13);
			this.lblCadastradoModo.TabIndex = 18;
			this.lblCadastradoModo.Text = "SISTEMA";
			// 
			// lblTitCadastradoModo
			// 
			this.lblTitCadastradoModo.AutoSize = true;
			this.lblTitCadastradoModo.Location = new System.Drawing.Point(453, 14);
			this.lblTitCadastradoModo.Name = "lblTitCadastradoModo";
			this.lblTitCadastradoModo.Size = new System.Drawing.Size(61, 13);
			this.lblTitCadastradoModo.TabIndex = 17;
			this.lblTitCadastradoModo.Text = "Cadastrado";
			// 
			// lblAlteradoPor
			// 
			this.lblAlteradoPor.AutoSize = true;
			this.lblAlteradoPor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblAlteradoPor.Location = new System.Drawing.Point(352, 42);
			this.lblAlteradoPor.Name = "lblAlteradoPor";
			this.lblAlteradoPor.Size = new System.Drawing.Size(56, 13);
			this.lblAlteradoPor.TabIndex = 22;
			this.lblAlteradoPor.Text = "FULANO";
			// 
			// lblTitAlteradoPor
			// 
			this.lblTitAlteradoPor.AutoSize = true;
			this.lblTitAlteradoPor.Location = new System.Drawing.Point(285, 42);
			this.lblTitAlteradoPor.Name = "lblTitAlteradoPor";
			this.lblTitAlteradoPor.Size = new System.Drawing.Size(64, 13);
			this.lblTitAlteradoPor.TabIndex = 21;
			this.lblTitAlteradoPor.Text = "Alterado por";
			// 
			// lblAlteradoEm
			// 
			this.lblAlteradoEm.AutoSize = true;
			this.lblAlteradoEm.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblAlteradoEm.Location = new System.Drawing.Point(119, 42);
			this.lblAlteradoEm.Name = "lblAlteradoEm";
			this.lblAlteradoEm.Size = new System.Drawing.Size(111, 13);
			this.lblAlteradoEm.TabIndex = 20;
			this.lblAlteradoEm.Text = "02/02/2002 12:02";
			// 
			// lblTitAlteradoEm
			// 
			this.lblTitAlteradoEm.AutoSize = true;
			this.lblTitAlteradoEm.Location = new System.Drawing.Point(53, 42);
			this.lblTitAlteradoEm.Name = "lblTitAlteradoEm";
			this.lblTitAlteradoEm.Size = new System.Drawing.Size(63, 13);
			this.lblTitAlteradoEm.TabIndex = 19;
			this.lblTitAlteradoEm.Text = "Alterado em";
			// 
			// lblNatureza
			// 
			this.lblNatureza.AutoSize = true;
			this.lblNatureza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblNatureza.Location = new System.Drawing.Point(517, 42);
			this.lblNatureza.Name = "lblNatureza";
			this.lblNatureza.Size = new System.Drawing.Size(62, 13);
			this.lblNatureza.TabIndex = 24;
			this.lblNatureza.Text = "CRÉDITO";
			// 
			// lblTitNatureza
			// 
			this.lblTitNatureza.AutoSize = true;
			this.lblTitNatureza.Location = new System.Drawing.Point(464, 42);
			this.lblTitNatureza.Name = "lblTitNatureza";
			this.lblTitNatureza.Size = new System.Drawing.Size(50, 13);
			this.lblTitNatureza.TabIndex = 23;
			this.lblTitNatureza.Text = "Natureza";
			// 
			// cbStSemEfeito
			// 
			this.cbStSemEfeito.BackColor = System.Drawing.SystemColors.Window;
			this.cbStSemEfeito.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbStSemEfeito.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbStSemEfeito.FormattingEnabled = true;
			this.cbStSemEfeito.Location = new System.Drawing.Point(122, 13);
			this.cbStSemEfeito.Name = "cbStSemEfeito";
			this.cbStSemEfeito.Size = new System.Drawing.Size(155, 24);
			this.cbStSemEfeito.TabIndex = 0;
			this.cbStSemEfeito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbStSemEfeito_KeyDown);
			// 
			// lblTitStSemEfeito
			// 
			this.lblTitStSemEfeito.AutoSize = true;
			this.lblTitStSemEfeito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitStSemEfeito.Location = new System.Drawing.Point(82, 18);
			this.lblTitStSemEfeito.Name = "lblTitStSemEfeito";
			this.lblTitStSemEfeito.Size = new System.Drawing.Size(34, 13);
			this.lblTitStSemEfeito.TabIndex = 35;
			this.lblTitStSemEfeito.Text = "Efeito";
			// 
			// cbCtrlPagtoStatus
			// 
			this.cbCtrlPagtoStatus.BackColor = System.Drawing.SystemColors.Window;
			this.cbCtrlPagtoStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbCtrlPagtoStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbCtrlPagtoStatus.FormattingEnabled = true;
			this.cbCtrlPagtoStatus.Location = new System.Drawing.Point(122, 50);
			this.cbCtrlPagtoStatus.Name = "cbCtrlPagtoStatus";
			this.cbCtrlPagtoStatus.Size = new System.Drawing.Size(518, 24);
			this.cbCtrlPagtoStatus.TabIndex = 2;
			this.cbCtrlPagtoStatus.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbCtrlPagtoStatus_KeyDown);
			// 
			// lblTitCtrlPagtoStatus
			// 
			this.lblTitCtrlPagtoStatus.AutoSize = true;
			this.lblTitCtrlPagtoStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCtrlPagtoStatus.Location = new System.Drawing.Point(79, 55);
			this.lblTitCtrlPagtoStatus.Name = "lblTitCtrlPagtoStatus";
			this.lblTitCtrlPagtoStatus.Size = new System.Drawing.Size(37, 13);
			this.lblTitCtrlPagtoStatus.TabIndex = 37;
			this.lblTitCtrlPagtoStatus.Text = "Status";
			// 
			// gboxCamposSistema
			// 
			this.gboxCamposSistema.Controls.Add(this.lblNatureza);
			this.gboxCamposSistema.Controls.Add(this.lblTitNatureza);
			this.gboxCamposSistema.Controls.Add(this.lblAlteradoPor);
			this.gboxCamposSistema.Controls.Add(this.lblTitAlteradoPor);
			this.gboxCamposSistema.Controls.Add(this.lblAlteradoEm);
			this.gboxCamposSistema.Controls.Add(this.lblTitAlteradoEm);
			this.gboxCamposSistema.Controls.Add(this.lblCadastradoModo);
			this.gboxCamposSistema.Controls.Add(this.lblTitCadastradoModo);
			this.gboxCamposSistema.Controls.Add(this.lblCadastradoPor);
			this.gboxCamposSistema.Controls.Add(this.lblTitCadastradoPor);
			this.gboxCamposSistema.Controls.Add(this.lblCadastradoEm);
			this.gboxCamposSistema.Controls.Add(this.lblTitCadastradoEm);
			this.gboxCamposSistema.Location = new System.Drawing.Point(10, 45);
			this.gboxCamposSistema.Name = "gboxCamposSistema";
			this.gboxCamposSistema.Size = new System.Drawing.Size(652, 64);
			this.gboxCamposSistema.TabIndex = 38;
			this.gboxCamposSistema.TabStop = false;
			// 
			// gboxCamposLancamento
			// 
			this.gboxCamposLancamento.Controls.Add(this.txtNF);
			this.gboxCamposLancamento.Controls.Add(this.lblNF);
			this.gboxCamposLancamento.Controls.Add(this.txtComp2);
			this.gboxCamposLancamento.Controls.Add(this.lblTitComp2);
			this.gboxCamposLancamento.Controls.Add(this.cbStConfirmacaoPendente);
			this.gboxCamposLancamento.Controls.Add(this.lblTitStConfirmacaoPendente);
			this.gboxCamposLancamento.Controls.Add(this.cbCtrlPagtoStatus);
			this.gboxCamposLancamento.Controls.Add(this.lblTitCtrlPagtoStatus);
			this.gboxCamposLancamento.Controls.Add(this.cbStSemEfeito);
			this.gboxCamposLancamento.Controls.Add(this.lblTitStSemEfeito);
			this.gboxCamposLancamento.Controls.Add(this.txtCnpjCpf);
			this.gboxCamposLancamento.Controls.Add(this.lblCnpjCpf);
			this.gboxCamposLancamento.Controls.Add(this.txtValor);
			this.gboxCamposLancamento.Controls.Add(this.txtDataCompetencia);
			this.gboxCamposLancamento.Controls.Add(this.txtDescricao);
			this.gboxCamposLancamento.Controls.Add(this.lblDescricao);
			this.gboxCamposLancamento.Controls.Add(this.lblValor);
			this.gboxCamposLancamento.Controls.Add(this.lblDataCompetencia);
			this.gboxCamposLancamento.Controls.Add(this.cbPlanoContasConta);
			this.gboxCamposLancamento.Controls.Add(this.lblPlanoContasConta);
			this.gboxCamposLancamento.Controls.Add(this.cbPlanoContasEmpresa);
			this.gboxCamposLancamento.Controls.Add(this.lblPlanoContasEmpresa);
			this.gboxCamposLancamento.Controls.Add(this.cbContaCorrente);
			this.gboxCamposLancamento.Controls.Add(this.lblContaCorrente);
			this.gboxCamposLancamento.Location = new System.Drawing.Point(10, 158);
			this.gboxCamposLancamento.Name = "gboxCamposLancamento";
			this.gboxCamposLancamento.Size = new System.Drawing.Size(652, 313);
			this.gboxCamposLancamento.TabIndex = 39;
			this.gboxCamposLancamento.TabStop = false;
			// 
			// txtComp2
			// 
			this.txtComp2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtComp2.Location = new System.Drawing.Point(322, 203);
			this.txtComp2.MaxLength = 7;
			this.txtComp2.Name = "txtComp2";
			this.txtComp2.Size = new System.Drawing.Size(91, 23);
			this.txtComp2.TabIndex = 7;
			this.txtComp2.Text = "01/01/2000";
			this.txtComp2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtComp2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtComp2_KeyDown);
			this.txtComp2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComp2_KeyPress);
			this.txtComp2.Leave += new System.EventHandler(this.txtComp2_Leave);
			// 
			// lblTitComp2
			// 
			this.lblTitComp2.AutoSize = true;
			this.lblTitComp2.Location = new System.Drawing.Point(276, 208);
			this.lblTitComp2.Name = "lblTitComp2";
			this.lblTitComp2.Size = new System.Drawing.Size(40, 13);
			this.lblTitComp2.TabIndex = 8;
			this.lblTitComp2.Text = "Comp2";
			// 
			// cbStConfirmacaoPendente
			// 
			this.cbStConfirmacaoPendente.BackColor = System.Drawing.SystemColors.Window;
			this.cbStConfirmacaoPendente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbStConfirmacaoPendente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbStConfirmacaoPendente.FormattingEnabled = true;
			this.cbStConfirmacaoPendente.Location = new System.Drawing.Point(473, 13);
			this.cbStConfirmacaoPendente.Name = "cbStConfirmacaoPendente";
			this.cbStConfirmacaoPendente.Size = new System.Drawing.Size(167, 24);
			this.cbStConfirmacaoPendente.TabIndex = 1;
			this.cbStConfirmacaoPendente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbStConfirmacaoPendente_KeyDown);
			// 
			// lblTitStConfirmacaoPendente
			// 
			this.lblTitStConfirmacaoPendente.AutoSize = true;
			this.lblTitStConfirmacaoPendente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitStConfirmacaoPendente.Location = new System.Drawing.Point(352, 18);
			this.lblTitStConfirmacaoPendente.Name = "lblTitStConfirmacaoPendente";
			this.lblTitStConfirmacaoPendente.Size = new System.Drawing.Size(115, 13);
			this.lblTitStConfirmacaoPendente.TabIndex = 39;
			this.lblTitStConfirmacaoPendente.Text = "Confirmação Pendente";
			// 
			// gboxDadosCliente
			// 
			this.gboxDadosCliente.Controls.Add(this.lblNome);
			this.gboxDadosCliente.Controls.Add(this.lblTitNome);
			this.gboxDadosCliente.Location = new System.Drawing.Point(10, 116);
			this.gboxDadosCliente.Name = "gboxDadosCliente";
			this.gboxDadosCliente.Size = new System.Drawing.Size(652, 35);
			this.gboxDadosCliente.TabIndex = 40;
			this.gboxDadosCliente.TabStop = false;
			// 
			// lblNome
			// 
			this.lblNome.AutoSize = true;
			this.lblNome.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblNome.Location = new System.Drawing.Point(119, 14);
			this.lblNome.Name = "lblNome";
			this.lblNome.Size = new System.Drawing.Size(104, 13);
			this.lblNome.TabIndex = 15;
			this.lblNome.Text = "FULANO DE TAL";
			// 
			// lblTitNome
			// 
			this.lblTitNome.AutoSize = true;
			this.lblTitNome.Location = new System.Drawing.Point(81, 14);
			this.lblTitNome.Name = "lblTitNome";
			this.lblTitNome.Size = new System.Drawing.Size(35, 13);
			this.lblTitNome.TabIndex = 14;
			this.lblTitNome.Text = "Nome";
			// 
			// txtNF
			// 
			this.txtNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNF.Location = new System.Drawing.Point(547, 241);
			this.txtNF.MaxLength = 18;
			this.txtNF.Name = "txtNF";
			this.txtNF.Size = new System.Drawing.Size(93, 23);
			this.txtNF.TabIndex = 10;
			this.txtNF.Text = "999.999.999";
			this.txtNF.Enter += new System.EventHandler(this.txtNF_Enter);
			this.txtNF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNF_KeyDown);
			this.txtNF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNF_KeyPress);
			this.txtNF.Leave += new System.EventHandler(this.txtNF_Leave);
			// 
			// lblNF
			// 
			this.lblNF.AutoSize = true;
			this.lblNF.Location = new System.Drawing.Point(520, 246);
			this.lblNF.Name = "lblNF";
			this.lblNF.Size = new System.Drawing.Size(21, 13);
			this.lblNF.TabIndex = 41;
			this.lblNF.Text = "NF";
			// 
			// FFluxoEdita
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(675, 583);
			this.Name = "FFluxoEdita";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FFluxoEdita_FormClosing);
			this.Load += new System.EventHandler(this.FFluxoEdita_Load);
			this.Shown += new System.EventHandler(this.FFluxoEdita_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxCamposSistema.ResumeLayout(false);
			this.gboxCamposSistema.PerformLayout();
			this.gboxCamposLancamento.ResumeLayout(false);
			this.gboxCamposLancamento.PerformLayout();
			this.gboxDadosCliente.ResumeLayout(false);
			this.gboxDadosCliente.PerformLayout();
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
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.Button btnExcluir;
		private System.Windows.Forms.Button btnAtualizar;
		private System.Windows.Forms.Label lblCadastradoPor;
		private System.Windows.Forms.Label lblTitCadastradoPor;
		private System.Windows.Forms.Label lblCadastradoEm;
		private System.Windows.Forms.Label lblTitCadastradoEm;
		private System.Windows.Forms.Label lblCadastradoModo;
		private System.Windows.Forms.Label lblTitCadastradoModo;
		private System.Windows.Forms.Label lblAlteradoPor;
		private System.Windows.Forms.Label lblTitAlteradoPor;
		private System.Windows.Forms.Label lblAlteradoEm;
		private System.Windows.Forms.Label lblTitAlteradoEm;
		private System.Windows.Forms.Label lblNatureza;
		private System.Windows.Forms.Label lblTitNatureza;
		private System.Windows.Forms.ComboBox cbStSemEfeito;
		private System.Windows.Forms.Label lblTitStSemEfeito;
		private System.Windows.Forms.ComboBox cbCtrlPagtoStatus;
		private System.Windows.Forms.Label lblTitCtrlPagtoStatus;
		private System.Windows.Forms.GroupBox gboxCamposSistema;
		private System.Windows.Forms.GroupBox gboxCamposLancamento;
		private System.Windows.Forms.GroupBox gboxDadosCliente;
		private System.Windows.Forms.Label lblNome;
		private System.Windows.Forms.Label lblTitNome;
		private System.Windows.Forms.ComboBox cbStConfirmacaoPendente;
		private System.Windows.Forms.Label lblTitStConfirmacaoPendente;
        private System.Windows.Forms.TextBox txtComp2;
        private System.Windows.Forms.Label lblTitComp2;
		private System.Windows.Forms.TextBox txtNF;
		private System.Windows.Forms.Label lblNF;
	}
}
