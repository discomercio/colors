namespace Financeiro
{
	partial class FFluxoRelatorioMovimentoSintetico
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoRelatorioMovimentoSintetico));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.pnParametros = new System.Windows.Forms.Panel();
			this.txtMesCompetenciaFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodoComp2 = new System.Windows.Forms.Label();
			this.txtMesCompetenciaInicial = new System.Windows.Forms.TextBox();
			this.lblMesCompetenciaAte = new System.Windows.Forms.Label();
			this.chkCNPJ = new System.Windows.Forms.CheckBox();
			this.chkCPF = new System.Windows.Forms.CheckBox();
			this.cbPlanoContasGrupoFinal = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasGrupoFinal = new System.Windows.Forms.Label();
			this.cbPlanoContasGrupoInicial = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasGrupoInicial = new System.Windows.Forms.Label();
			this.chkIncluirAtrasados = new System.Windows.Forms.CheckBox();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.txtDescricao = new System.Windows.Forms.TextBox();
			this.lblDescricao = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.cbPlanoContasConta = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasConta = new System.Windows.Forms.Label();
			this.cbPlanoContasGrupo = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasGrupo = new System.Windows.Forms.Label();
			this.cbPlanoContasEmpresa = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasEmpresa = new System.Windows.Forms.Label();
			this.cbContaCorrente = new System.Windows.Forms.ComboBox();
			this.lblTitContaCorrente = new System.Windows.Forms.Label();
			this.cbNatureza = new System.Windows.Forms.ComboBox();
			this.lblTitNatureza = new System.Windows.Forms.Label();
			this.txtDataCadastroFinal = new System.Windows.Forms.TextBox();
			this.lblTitDataCadastro = new System.Windows.Forms.Label();
			this.txtDataCadastroInicial = new System.Windows.Forms.TextBox();
			this.lblTitDataCadastroA = new System.Windows.Forms.Label();
			this.txtDataCompetenciaFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodoCompetencia = new System.Windows.Forms.Label();
			this.txtDataCompetenciaInicial = new System.Windows.Forms.TextBox();
			this.lblDataCompetenciaAte = new System.Windows.Forms.Label();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
			this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.btnPrintPreview = new System.Windows.Forms.Button();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.pnParametros.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnPrinterDialog);
			this.pnBotoes.Controls.Add(this.btnPrintPreview);
			this.pnBotoes.Controls.Add(this.btnLimpar);
			this.pnBotoes.Controls.Add(this.btnImprimir);
			this.pnBotoes.Controls.SetChildIndex(this.btnImprimir, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnLimpar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrintPreview, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.pnParametros);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Size = new System.Drawing.Size(1018, 299);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 4;
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
			this.lblTitulo.TabIndex = 1;
			this.lblTitulo.Text = "Relatório Sintético de Movimentos";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pnParametros
			// 
			this.pnParametros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnParametros.Controls.Add(this.txtMesCompetenciaFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodoComp2);
			this.pnParametros.Controls.Add(this.txtMesCompetenciaInicial);
			this.pnParametros.Controls.Add(this.lblMesCompetenciaAte);
			this.pnParametros.Controls.Add(this.chkCNPJ);
			this.pnParametros.Controls.Add(this.chkCPF);
			this.pnParametros.Controls.Add(this.cbPlanoContasGrupoFinal);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasGrupoFinal);
			this.pnParametros.Controls.Add(this.cbPlanoContasGrupoInicial);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasGrupoInicial);
			this.pnParametros.Controls.Add(this.chkIncluirAtrasados);
			this.pnParametros.Controls.Add(this.txtCnpjCpf);
			this.pnParametros.Controls.Add(this.lblCnpjCpf);
			this.pnParametros.Controls.Add(this.txtValor);
			this.pnParametros.Controls.Add(this.txtDescricao);
			this.pnParametros.Controls.Add(this.lblDescricao);
			this.pnParametros.Controls.Add(this.lblValor);
			this.pnParametros.Controls.Add(this.cbPlanoContasConta);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasConta);
			this.pnParametros.Controls.Add(this.cbPlanoContasGrupo);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasGrupo);
			this.pnParametros.Controls.Add(this.cbPlanoContasEmpresa);
			this.pnParametros.Controls.Add(this.lblPlanoContasEmpresa);
			this.pnParametros.Controls.Add(this.cbContaCorrente);
			this.pnParametros.Controls.Add(this.lblTitContaCorrente);
			this.pnParametros.Controls.Add(this.cbNatureza);
			this.pnParametros.Controls.Add(this.lblTitNatureza);
			this.pnParametros.Controls.Add(this.txtDataCadastroFinal);
			this.pnParametros.Controls.Add(this.lblTitDataCadastro);
			this.pnParametros.Controls.Add(this.txtDataCadastroInicial);
			this.pnParametros.Controls.Add(this.lblTitDataCadastroA);
			this.pnParametros.Controls.Add(this.txtDataCompetenciaFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodoCompetencia);
			this.pnParametros.Controls.Add(this.txtDataCompetenciaInicial);
			this.pnParametros.Controls.Add(this.lblDataCompetenciaAte);
			this.pnParametros.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pnParametros.Location = new System.Drawing.Point(0, 40);
			this.pnParametros.Name = "pnParametros";
			this.pnParametros.Size = new System.Drawing.Size(1014, 255);
			this.pnParametros.TabIndex = 2;
			// 
			// txtMesCompetenciaFinal
			// 
			this.txtMesCompetenciaFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMesCompetenciaFinal.Location = new System.Drawing.Point(547, 18);
			this.txtMesCompetenciaFinal.MaxLength = 7;
			this.txtMesCompetenciaFinal.Name = "txtMesCompetenciaFinal";
			this.txtMesCompetenciaFinal.Size = new System.Drawing.Size(91, 23);
			this.txtMesCompetenciaFinal.TabIndex = 3;
			this.txtMesCompetenciaFinal.Text = "01/2000";
			this.txtMesCompetenciaFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtMesCompetenciaFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMesCompetenciaFinal_KeyDown);
			this.txtMesCompetenciaFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMesCompetenciaFinal_KeyPress);
			this.txtMesCompetenciaFinal.Leave += new System.EventHandler(this.txtMesCompetenciaFinal_Leave);
			// 
			// lblTitPeriodoComp2
			// 
			this.lblTitPeriodoComp2.AutoSize = true;
			this.lblTitPeriodoComp2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodoComp2.Location = new System.Drawing.Point(379, 23);
			this.lblTitPeriodoComp2.Name = "lblTitPeriodoComp2";
			this.lblTitPeriodoComp2.Size = new System.Drawing.Size(45, 13);
			this.lblTitPeriodoComp2.TabIndex = 51;
			this.lblTitPeriodoComp2.Text = "Comp2";
			// 
			// txtMesCompetenciaInicial
			// 
			this.txtMesCompetenciaInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMesCompetenciaInicial.Location = new System.Drawing.Point(430, 18);
			this.txtMesCompetenciaInicial.MaxLength = 7;
			this.txtMesCompetenciaInicial.Name = "txtMesCompetenciaInicial";
			this.txtMesCompetenciaInicial.Size = new System.Drawing.Size(91, 23);
			this.txtMesCompetenciaInicial.TabIndex = 2;
			this.txtMesCompetenciaInicial.Text = "01/2000";
			this.txtMesCompetenciaInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtMesCompetenciaInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMesCompetenciaInicial_KeyDown);
			this.txtMesCompetenciaInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMesCompetenciaInicial_KeyPress);
			this.txtMesCompetenciaInicial.Leave += new System.EventHandler(this.txtMesCompetenciaInicial_Leave);
			// 
			// lblMesCompetenciaAte
			// 
			this.lblMesCompetenciaAte.AutoSize = true;
			this.lblMesCompetenciaAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblMesCompetenciaAte.Location = new System.Drawing.Point(527, 23);
			this.lblMesCompetenciaAte.Name = "lblMesCompetenciaAte";
			this.lblMesCompetenciaAte.Size = new System.Drawing.Size(14, 13);
			this.lblMesCompetenciaAte.TabIndex = 50;
			this.lblMesCompetenciaAte.Text = "a";
			// 
			// chkCNPJ
			// 
			this.chkCNPJ.AutoSize = true;
			this.chkCNPJ.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.chkCNPJ.Location = new System.Drawing.Point(447, 216);
			this.chkCNPJ.Name = "chkCNPJ";
			this.chkCNPJ.Size = new System.Drawing.Size(57, 17);
			this.chkCNPJ.TabIndex = 18;
			this.chkCNPJ.Text = "CNPJ";
			this.chkCNPJ.UseVisualStyleBackColor = true;
			// 
			// chkCPF
			// 
			this.chkCPF.AutoSize = true;
			this.chkCPF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.chkCPF.Location = new System.Drawing.Point(377, 216);
			this.chkCPF.Name = "chkCPF";
			this.chkCPF.Size = new System.Drawing.Size(49, 17);
			this.chkCPF.TabIndex = 17;
			this.chkCPF.Text = "CPF";
			this.chkCPF.UseVisualStyleBackColor = true;
			// 
			// cbPlanoContasGrupoFinal
			// 
			this.cbPlanoContasGrupoFinal.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasGrupoFinal.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasGrupoFinal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasGrupoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasGrupoFinal.FormattingEnabled = true;
			this.cbPlanoContasGrupoFinal.Location = new System.Drawing.Point(605, 176);
			this.cbPlanoContasGrupoFinal.MaxDropDownItems = 12;
			this.cbPlanoContasGrupoFinal.Name = "cbPlanoContasGrupoFinal";
			this.cbPlanoContasGrupoFinal.Size = new System.Drawing.Size(400, 24);
			this.cbPlanoContasGrupoFinal.TabIndex = 15;
			this.cbPlanoContasGrupoFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasGrupoFinal_KeyDown);
			// 
			// lblTitPlanoContasGrupoFinal
			// 
			this.lblTitPlanoContasGrupoFinal.AutoSize = true;
			this.lblTitPlanoContasGrupoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasGrupoFinal.Location = new System.Drawing.Point(522, 181);
			this.lblTitPlanoContasGrupoFinal.Name = "lblTitPlanoContasGrupoFinal";
			this.lblTitPlanoContasGrupoFinal.Size = new System.Drawing.Size(77, 13);
			this.lblTitPlanoContasGrupoFinal.TabIndex = 39;
			this.lblTitPlanoContasGrupoFinal.Text = "Grupo (final)";
			// 
			// cbPlanoContasGrupoInicial
			// 
			this.cbPlanoContasGrupoInicial.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasGrupoInicial.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasGrupoInicial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasGrupoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasGrupoInicial.FormattingEnabled = true;
			this.cbPlanoContasGrupoInicial.Location = new System.Drawing.Point(104, 176);
			this.cbPlanoContasGrupoInicial.MaxDropDownItems = 12;
			this.cbPlanoContasGrupoInicial.Name = "cbPlanoContasGrupoInicial";
			this.cbPlanoContasGrupoInicial.Size = new System.Drawing.Size(400, 24);
			this.cbPlanoContasGrupoInicial.TabIndex = 14;
			this.cbPlanoContasGrupoInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasGrupoInicial_KeyDown);
			// 
			// lblTitPlanoContasGrupoInicial
			// 
			this.lblTitPlanoContasGrupoInicial.AutoSize = true;
			this.lblTitPlanoContasGrupoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasGrupoInicial.Location = new System.Drawing.Point(12, 181);
			this.lblTitPlanoContasGrupoInicial.Name = "lblTitPlanoContasGrupoInicial";
			this.lblTitPlanoContasGrupoInicial.Size = new System.Drawing.Size(86, 13);
			this.lblTitPlanoContasGrupoInicial.TabIndex = 38;
			this.lblTitPlanoContasGrupoInicial.Text = "Grupo (inicial)";
			// 
			// chkIncluirAtrasados
			// 
			this.chkIncluirAtrasados.AutoSize = true;
			this.chkIncluirAtrasados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.chkIncluirAtrasados.Location = new System.Drawing.Point(104, 216);
			this.chkIncluirAtrasados.Name = "chkIncluirAtrasados";
			this.chkIncluirAtrasados.Size = new System.Drawing.Size(121, 17);
			this.chkIncluirAtrasados.TabIndex = 16;
			this.chkIncluirAtrasados.Text = "Incluir Atrasados";
			this.chkIncluirAtrasados.UseVisualStyleBackColor = true;
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(359, 57);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(145, 23);
			this.txtCnpjCpf.TabIndex = 7;
			this.txtCnpjCpf.Text = "00.000.000/0000-00";
			this.txtCnpjCpf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCnpjCpf.Enter += new System.EventHandler(this.txtCnpjCpf_Enter);
			this.txtCnpjCpf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCnpjCpf_KeyDown);
			this.txtCnpjCpf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCnpjCpf_KeyPress);
			this.txtCnpjCpf.Leave += new System.EventHandler(this.txtCnpjCpf_Leave);
			// 
			// lblCnpjCpf
			// 
			this.lblCnpjCpf.AutoSize = true;
			this.lblCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCnpjCpf.Location = new System.Drawing.Point(286, 62);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(67, 13);
			this.lblCnpjCpf.TabIndex = 31;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(104, 57);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(111, 23);
			this.txtValor.TabIndex = 6;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			// 
			// txtDescricao
			// 
			this.txtDescricao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDescricao.Location = new System.Drawing.Point(605, 57);
			this.txtDescricao.MaxLength = 80;
			this.txtDescricao.Name = "txtDescricao";
			this.txtDescricao.Size = new System.Drawing.Size(218, 23);
			this.txtDescricao.TabIndex = 8;
			this.txtDescricao.Enter += new System.EventHandler(this.txtDescricao_Enter);
			this.txtDescricao.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescricao_KeyDown);
			this.txtDescricao.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescricao_KeyPress);
			this.txtDescricao.Leave += new System.EventHandler(this.txtDescricao_Leave);
			// 
			// lblDescricao
			// 
			this.lblDescricao.AutoSize = true;
			this.lblDescricao.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblDescricao.Location = new System.Drawing.Point(535, 62);
			this.lblDescricao.Name = "lblDescricao";
			this.lblDescricao.Size = new System.Drawing.Size(64, 13);
			this.lblDescricao.TabIndex = 29;
			this.lblDescricao.Text = "Descrição";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValor.Location = new System.Drawing.Point(34, 62);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(64, 13);
			this.lblValor.TabIndex = 28;
			this.lblValor.Text = "Valor (R$)";
			// 
			// cbPlanoContasConta
			// 
			this.cbPlanoContasConta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasConta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasConta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasConta.FormattingEnabled = true;
			this.cbPlanoContasConta.Location = new System.Drawing.Point(605, 136);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(400, 24);
			this.cbPlanoContasConta.TabIndex = 13;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblTitPlanoContasConta
			// 
			this.lblTitPlanoContasConta.AutoSize = true;
			this.lblTitPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasConta.Location = new System.Drawing.Point(559, 141);
			this.lblTitPlanoContasConta.Name = "lblTitPlanoContasConta";
			this.lblTitPlanoContasConta.Size = new System.Drawing.Size(40, 13);
			this.lblTitPlanoContasConta.TabIndex = 25;
			this.lblTitPlanoContasConta.Text = "Conta";
			// 
			// cbPlanoContasGrupo
			// 
			this.cbPlanoContasGrupo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasGrupo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasGrupo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasGrupo.FormattingEnabled = true;
			this.cbPlanoContasGrupo.Location = new System.Drawing.Point(104, 136);
			this.cbPlanoContasGrupo.MaxDropDownItems = 12;
			this.cbPlanoContasGrupo.Name = "cbPlanoContasGrupo";
			this.cbPlanoContasGrupo.Size = new System.Drawing.Size(400, 24);
			this.cbPlanoContasGrupo.TabIndex = 12;
			this.cbPlanoContasGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasGrupo_KeyDown);
			// 
			// lblTitPlanoContasGrupo
			// 
			this.lblTitPlanoContasGrupo.AutoSize = true;
			this.lblTitPlanoContasGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasGrupo.Location = new System.Drawing.Point(57, 141);
			this.lblTitPlanoContasGrupo.Name = "lblTitPlanoContasGrupo";
			this.lblTitPlanoContasGrupo.Size = new System.Drawing.Size(41, 13);
			this.lblTitPlanoContasGrupo.TabIndex = 23;
			this.lblTitPlanoContasGrupo.Text = "Grupo";
			// 
			// cbPlanoContasEmpresa
			// 
			this.cbPlanoContasEmpresa.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasEmpresa.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasEmpresa.FormattingEnabled = true;
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(605, 96);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(400, 24);
			this.cbPlanoContasEmpresa.TabIndex = 11;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(544, 101);
			this.lblPlanoContasEmpresa.Name = "lblPlanoContasEmpresa";
			this.lblPlanoContasEmpresa.Size = new System.Drawing.Size(55, 13);
			this.lblPlanoContasEmpresa.TabIndex = 21;
			this.lblPlanoContasEmpresa.Text = "Empresa";
			// 
			// cbContaCorrente
			// 
			this.cbContaCorrente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbContaCorrente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbContaCorrente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbContaCorrente.FormattingEnabled = true;
			this.cbContaCorrente.Location = new System.Drawing.Point(104, 96);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(400, 24);
			this.cbContaCorrente.TabIndex = 10;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblTitContaCorrente
			// 
			this.lblTitContaCorrente.AutoSize = true;
			this.lblTitContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitContaCorrente.Location = new System.Drawing.Point(6, 101);
			this.lblTitContaCorrente.Name = "lblTitContaCorrente";
			this.lblTitContaCorrente.Size = new System.Drawing.Size(92, 13);
			this.lblTitContaCorrente.TabIndex = 19;
			this.lblTitContaCorrente.Text = "Conta Corrente";
			// 
			// cbNatureza
			// 
			this.cbNatureza.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbNatureza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbNatureza.FormattingEnabled = true;
			this.cbNatureza.Location = new System.Drawing.Point(898, 59);
			this.cbNatureza.Name = "cbNatureza";
			this.cbNatureza.Size = new System.Drawing.Size(107, 21);
			this.cbNatureza.TabIndex = 9;
			this.cbNatureza.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbNatureza_KeyDown);
			// 
			// lblTitNatureza
			// 
			this.lblTitNatureza.AutoSize = true;
			this.lblTitNatureza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNatureza.Location = new System.Drawing.Point(834, 62);
			this.lblTitNatureza.Name = "lblTitNatureza";
			this.lblTitNatureza.Size = new System.Drawing.Size(58, 13);
			this.lblTitNatureza.TabIndex = 16;
			this.lblTitNatureza.Text = "Natureza";
			// 
			// txtDataCadastroFinal
			// 
			this.txtDataCadastroFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCadastroFinal.Location = new System.Drawing.Point(914, 18);
			this.txtDataCadastroFinal.MaxLength = 10;
			this.txtDataCadastroFinal.Name = "txtDataCadastroFinal";
			this.txtDataCadastroFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataCadastroFinal.TabIndex = 5;
			this.txtDataCadastroFinal.Text = "01/01/2000";
			this.txtDataCadastroFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCadastroFinal.Enter += new System.EventHandler(this.txtDataCadastroFinal_Enter);
			this.txtDataCadastroFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCadastroFinal_KeyDown);
			this.txtDataCadastroFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCadastroFinal_KeyPress);
			this.txtDataCadastroFinal.Leave += new System.EventHandler(this.txtDataCadastroFinal_Leave);
			// 
			// lblTitDataCadastro
			// 
			this.lblTitDataCadastro.AutoSize = true;
			this.lblTitDataCadastro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitDataCadastro.Location = new System.Drawing.Point(700, 23);
			this.lblTitDataCadastro.Name = "lblTitDataCadastro";
			this.lblTitDataCadastro.Size = new System.Drawing.Size(91, 13);
			this.lblTitDataCadastro.TabIndex = 14;
			this.lblTitDataCadastro.Text = "Cadastramento";
			// 
			// txtDataCadastroInicial
			// 
			this.txtDataCadastroInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCadastroInicial.Location = new System.Drawing.Point(797, 18);
			this.txtDataCadastroInicial.MaxLength = 10;
			this.txtDataCadastroInicial.Name = "txtDataCadastroInicial";
			this.txtDataCadastroInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataCadastroInicial.TabIndex = 4;
			this.txtDataCadastroInicial.Text = "01/01/2000";
			this.txtDataCadastroInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCadastroInicial.Enter += new System.EventHandler(this.txtDataCadastroInicial_Enter);
			this.txtDataCadastroInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCadastroInicial_KeyDown);
			this.txtDataCadastroInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCadastroInicial_KeyPress);
			this.txtDataCadastroInicial.Leave += new System.EventHandler(this.txtDataCadastroInicial_Leave);
			// 
			// lblTitDataCadastroA
			// 
			this.lblTitDataCadastroA.AutoSize = true;
			this.lblTitDataCadastroA.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitDataCadastroA.Location = new System.Drawing.Point(894, 23);
			this.lblTitDataCadastroA.Name = "lblTitDataCadastroA";
			this.lblTitDataCadastroA.Size = new System.Drawing.Size(14, 13);
			this.lblTitDataCadastroA.TabIndex = 13;
			this.lblTitDataCadastroA.Text = "a";
			// 
			// txtDataCompetenciaFinal
			// 
			this.txtDataCompetenciaFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetenciaFinal.Location = new System.Drawing.Point(221, 18);
			this.txtDataCompetenciaFinal.MaxLength = 10;
			this.txtDataCompetenciaFinal.Name = "txtDataCompetenciaFinal";
			this.txtDataCompetenciaFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataCompetenciaFinal.TabIndex = 1;
			this.txtDataCompetenciaFinal.Text = "01/01/2000";
			this.txtDataCompetenciaFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCompetenciaFinal.Enter += new System.EventHandler(this.txtDataCompetenciaFinal_Enter);
			this.txtDataCompetenciaFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetenciaFinal_KeyDown);
			this.txtDataCompetenciaFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetenciaFinal_KeyPress);
			this.txtDataCompetenciaFinal.Leave += new System.EventHandler(this.txtDataCompetenciaFinal_Leave);
			// 
			// lblTitPeriodoCompetencia
			// 
			this.lblTitPeriodoCompetencia.AutoSize = true;
			this.lblTitPeriodoCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodoCompetencia.Location = new System.Drawing.Point(18, 23);
			this.lblTitPeriodoCompetencia.Name = "lblTitPeriodoCompetencia";
			this.lblTitPeriodoCompetencia.Size = new System.Drawing.Size(80, 13);
			this.lblTitPeriodoCompetencia.TabIndex = 10;
			this.lblTitPeriodoCompetencia.Text = "Competência";
			// 
			// txtDataCompetenciaInicial
			// 
			this.txtDataCompetenciaInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetenciaInicial.Location = new System.Drawing.Point(104, 18);
			this.txtDataCompetenciaInicial.MaxLength = 10;
			this.txtDataCompetenciaInicial.Name = "txtDataCompetenciaInicial";
			this.txtDataCompetenciaInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataCompetenciaInicial.TabIndex = 0;
			this.txtDataCompetenciaInicial.Text = "01/01/2000";
			this.txtDataCompetenciaInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCompetenciaInicial.Enter += new System.EventHandler(this.txtDataCompetenciaInicial_Enter);
			this.txtDataCompetenciaInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetenciaInicial_KeyDown);
			this.txtDataCompetenciaInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetenciaInicial_KeyPress);
			this.txtDataCompetenciaInicial.Leave += new System.EventHandler(this.txtDataCompetenciaInicial_Leave);
			// 
			// lblDataCompetenciaAte
			// 
			this.lblDataCompetenciaAte.AutoSize = true;
			this.lblDataCompetenciaAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblDataCompetenciaAte.Location = new System.Drawing.Point(201, 23);
			this.lblDataCompetenciaAte.Name = "lblDataCompetenciaAte";
			this.lblDataCompetenciaAte.Size = new System.Drawing.Size(14, 13);
			this.lblDataCompetenciaAte.TabIndex = 9;
			this.lblDataCompetenciaAte.Text = "a";
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(789, 4);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(40, 44);
			this.btnImprimir.TabIndex = 1;
			this.btnImprimir.TabStop = false;
			this.btnImprimir.UseVisualStyleBackColor = true;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnLimpar
			// 
			this.btnLimpar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnLimpar.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpar.Image")));
			this.btnLimpar.Location = new System.Drawing.Point(744, 4);
			this.btnLimpar.Name = "btnLimpar";
			this.btnLimpar.Size = new System.Drawing.Size(40, 44);
			this.btnLimpar.TabIndex = 0;
			this.btnLimpar.TabStop = false;
			this.btnLimpar.UseVisualStyleBackColor = true;
			this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
			// 
			// prnPreviewConsulta
			// 
			this.prnPreviewConsulta.AutoScrollMargin = new System.Drawing.Size(0, 0);
			this.prnPreviewConsulta.AutoScrollMinSize = new System.Drawing.Size(0, 0);
			this.prnPreviewConsulta.ClientSize = new System.Drawing.Size(400, 300);
			this.prnPreviewConsulta.Document = this.prnDocConsulta;
			this.prnPreviewConsulta.Enabled = true;
			this.prnPreviewConsulta.Icon = ((System.Drawing.Icon)(resources.GetObject("prnPreviewConsulta.Icon")));
			this.prnPreviewConsulta.Name = "prnPreview";
			this.prnPreviewConsulta.UseAntiAlias = true;
			this.prnPreviewConsulta.Visible = false;
			// 
			// prnDocConsulta
			// 
			this.prnDocConsulta.DocumentName = "Relatório Sintético de Movimentos";
			this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
			this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
			// 
			// prnDialogConsulta
			// 
			this.prnDialogConsulta.Document = this.prnDocConsulta;
			this.prnDialogConsulta.UseEXDialog = true;
			// 
			// btnPrintPreview
			// 
			this.btnPrintPreview.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrintPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnPrintPreview.Image")));
			this.btnPrintPreview.Location = new System.Drawing.Point(834, 4);
			this.btnPrintPreview.Name = "btnPrintPreview";
			this.btnPrintPreview.Size = new System.Drawing.Size(40, 44);
			this.btnPrintPreview.TabIndex = 2;
			this.btnPrintPreview.TabStop = false;
			this.btnPrintPreview.UseVisualStyleBackColor = true;
			this.btnPrintPreview.Click += new System.EventHandler(this.btnPrintPreview_Click);
			// 
			// btnPrinterDialog
			// 
			this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
			this.btnPrinterDialog.Location = new System.Drawing.Point(879, 4);
			this.btnPrinterDialog.Name = "btnPrinterDialog";
			this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
			this.btnPrinterDialog.TabIndex = 3;
			this.btnPrinterDialog.TabStop = false;
			this.btnPrinterDialog.UseVisualStyleBackColor = true;
			this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
			// 
			// FFluxoRelatorioMovimentoSintetico
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 396);
			this.KeyPreview = true;
			this.Name = "FFluxoRelatorioMovimentoSintetico";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FFluxoRelatorio_FormClosing);
			this.Load += new System.EventHandler(this.FFluxoRelatorio_Load);
			this.Shown += new System.EventHandler(this.FFluxoRelatorio_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FFluxoRelatorioMovimento_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnParametros.ResumeLayout(false);
			this.pnParametros.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Panel pnParametros;
		private System.Windows.Forms.TextBox txtDataCompetenciaInicial;
		private System.Windows.Forms.Label lblDataCompetenciaAte;
		private System.Windows.Forms.Label lblTitPeriodoCompetencia;
		private System.Windows.Forms.TextBox txtDataCompetenciaFinal;
		private System.Windows.Forms.TextBox txtDataCadastroFinal;
		private System.Windows.Forms.Label lblTitDataCadastro;
		private System.Windows.Forms.TextBox txtDataCadastroInicial;
		private System.Windows.Forms.Label lblTitDataCadastroA;
		private System.Windows.Forms.Label lblTitNatureza;
		private System.Windows.Forms.ComboBox cbNatureza;
		private System.Windows.Forms.ComboBox cbContaCorrente;
		private System.Windows.Forms.Label lblTitContaCorrente;
		private System.Windows.Forms.ComboBox cbPlanoContasEmpresa;
		private System.Windows.Forms.Label lblPlanoContasEmpresa;
		private System.Windows.Forms.ComboBox cbPlanoContasGrupo;
		private System.Windows.Forms.Label lblTitPlanoContasGrupo;
		private System.Windows.Forms.ComboBox cbPlanoContasConta;
		private System.Windows.Forms.Label lblTitPlanoContasConta;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.TextBox txtDescricao;
		private System.Windows.Forms.Label lblDescricao;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnPrintPreview;
		private System.Windows.Forms.CheckBox chkIncluirAtrasados;
		private System.Windows.Forms.ComboBox cbPlanoContasGrupoFinal;
		private System.Windows.Forms.Label lblTitPlanoContasGrupoFinal;
		private System.Windows.Forms.ComboBox cbPlanoContasGrupoInicial;
		private System.Windows.Forms.Label lblTitPlanoContasGrupoInicial;
		private System.Windows.Forms.CheckBox chkCNPJ;
		private System.Windows.Forms.CheckBox chkCPF;
        private System.Windows.Forms.TextBox txtMesCompetenciaFinal;
        private System.Windows.Forms.Label lblTitPeriodoComp2;
        private System.Windows.Forms.TextBox txtMesCompetenciaInicial;
        private System.Windows.Forms.Label lblMesCompetenciaAte;
    }
}
