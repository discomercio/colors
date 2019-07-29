namespace Financeiro
{
	partial class FCobrancaAdministracao
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FCobrancaAdministracao));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.pnParametros = new System.Windows.Forms.Panel();
			this.ckb_somente_primeira_parcela_em_atraso = new System.Windows.Forms.CheckBox();
			this.cbGarantia = new System.Windows.Forms.ComboBox();
			this.lblTitPedidosComGarantia = new System.Windows.Forms.Label();
			this.cbIndicador = new System.Windows.Forms.ComboBox();
			this.lblTitIndicador = new System.Windows.Forms.Label();
			this.cbVendedor = new System.Windows.Forms.ComboBox();
			this.lblTitVendedor = new System.Windows.Forms.Label();
			this.cbEquipeVendas = new System.Windows.Forms.ComboBox();
			this.lblTitEquipeVendas = new System.Windows.Forms.Label();
			this.lblTitAtrasadoEntreDias = new System.Windows.Forms.Label();
			this.lblTitAtrasadoEntreE = new System.Windows.Forms.Label();
			this.txtQtdeDiasAtrasoFinal = new System.Windows.Forms.TextBox();
			this.txtQtdeDiasAtrasoInicial = new System.Windows.Forms.TextBox();
			this.lblTitAtrasadoEntre = new System.Windows.Forms.Label();
			this.cbSituacao = new System.Windows.Forms.ComboBox();
			this.lblTitSituacao = new System.Windows.Forms.Label();
			this.txtNomeCliente = new System.Windows.Forms.TextBox();
			this.lblTitNomeCliente = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.cbPlanoContasConta = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasConta = new System.Windows.Forms.Label();
			this.cbPlanoContasGrupo = new System.Windows.Forms.ComboBox();
			this.lblTitPlanoContasGrupo = new System.Windows.Forms.Label();
			this.cbPlanoContasEmpresa = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasEmpresa = new System.Windows.Forms.Label();
			this.cbContaCorrente = new System.Windows.Forms.ComboBox();
			this.lblTitContaCorrente = new System.Windows.Forms.Label();
			this.pnTotalizacao = new System.Windows.Forms.Panel();
			this.btnDesmarcarTodos = new System.Windows.Forms.Button();
			this.btnMarcarTodos = new System.Windows.Forms.Button();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.pnResultado = new System.Windows.Forms.Panel();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.colCheckBox = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.colGridNomeCnpjCpf = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridQtdeParcelasEmAtraso = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridMaxDiasEmAtraso = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridValorTotalEmAtraso = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridDescricaoParcelas = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridNumParcelaMaiorAtraso = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridVendedor = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridIndicador = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridUF = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colGridIdCliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.btnPrintPreview = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
			this.btnExcel = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.pnParametros.SuspendLayout();
			this.pnTotalizacao.SuspendLayout();
			this.pnResultado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnExcel);
			this.pnBotoes.Controls.Add(this.btnPrinterDialog);
			this.pnBotoes.Controls.Add(this.btnPrintPreview);
			this.pnBotoes.Controls.Add(this.btnImprimir);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnImprimir, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrintPreview, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnExcel, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.pnResultado);
			this.pnCampos.Controls.Add(this.pnTotalizacao);
			this.pnCampos.Controls.Add(this.pnParametros);
			this.pnCampos.Controls.Add(this.lblTitulo);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.TabIndex = 6;
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 5;
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
			this.lblTitulo.TabIndex = 3;
			this.lblTitulo.Text = "Cobrança: Administração da Carteira em Atraso";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pnParametros
			// 
			this.pnParametros.Controls.Add(this.ckb_somente_primeira_parcela_em_atraso);
			this.pnParametros.Controls.Add(this.cbGarantia);
			this.pnParametros.Controls.Add(this.lblTitPedidosComGarantia);
			this.pnParametros.Controls.Add(this.cbIndicador);
			this.pnParametros.Controls.Add(this.lblTitIndicador);
			this.pnParametros.Controls.Add(this.cbVendedor);
			this.pnParametros.Controls.Add(this.lblTitVendedor);
			this.pnParametros.Controls.Add(this.cbEquipeVendas);
			this.pnParametros.Controls.Add(this.lblTitEquipeVendas);
			this.pnParametros.Controls.Add(this.lblTitAtrasadoEntreDias);
			this.pnParametros.Controls.Add(this.lblTitAtrasadoEntreE);
			this.pnParametros.Controls.Add(this.txtQtdeDiasAtrasoFinal);
			this.pnParametros.Controls.Add(this.txtQtdeDiasAtrasoInicial);
			this.pnParametros.Controls.Add(this.lblTitAtrasadoEntre);
			this.pnParametros.Controls.Add(this.cbSituacao);
			this.pnParametros.Controls.Add(this.lblTitSituacao);
			this.pnParametros.Controls.Add(this.txtNomeCliente);
			this.pnParametros.Controls.Add(this.lblTitNomeCliente);
			this.pnParametros.Controls.Add(this.txtCnpjCpf);
			this.pnParametros.Controls.Add(this.lblCnpjCpf);
			this.pnParametros.Controls.Add(this.cbPlanoContasConta);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasConta);
			this.pnParametros.Controls.Add(this.cbPlanoContasGrupo);
			this.pnParametros.Controls.Add(this.lblTitPlanoContasGrupo);
			this.pnParametros.Controls.Add(this.cbPlanoContasEmpresa);
			this.pnParametros.Controls.Add(this.lblPlanoContasEmpresa);
			this.pnParametros.Controls.Add(this.cbContaCorrente);
			this.pnParametros.Controls.Add(this.lblTitContaCorrente);
			this.pnParametros.Dock = System.Windows.Forms.DockStyle.Top;
			this.pnParametros.Location = new System.Drawing.Point(0, 40);
			this.pnParametros.Name = "pnParametros";
			this.pnParametros.Size = new System.Drawing.Size(1014, 153);
			this.pnParametros.TabIndex = 0;
			// 
			// ckb_somente_primeira_parcela_em_atraso
			// 
			this.ckb_somente_primeira_parcela_em_atraso.AutoSize = true;
			this.ckb_somente_primeira_parcela_em_atraso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ckb_somente_primeira_parcela_em_atraso.Location = new System.Drawing.Point(811, 9);
			this.ckb_somente_primeira_parcela_em_atraso.Name = "ckb_somente_primeira_parcela_em_atraso";
			this.ckb_somente_primeira_parcela_em_atraso.Size = new System.Drawing.Size(196, 17);
			this.ckb_somente_primeira_parcela_em_atraso.TabIndex = 3;
			this.ckb_somente_primeira_parcela_em_atraso.Text = "Somente 1ª parcela em atraso";
			this.ckb_somente_primeira_parcela_em_atraso.UseVisualStyleBackColor = true;
			// 
			// cbGarantia
			// 
			this.cbGarantia.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbGarantia.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbGarantia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbGarantia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbGarantia.FormattingEnabled = true;
			this.cbGarantia.Location = new System.Drawing.Point(615, 104);
			this.cbGarantia.MaxDropDownItems = 12;
			this.cbGarantia.Name = "cbGarantia";
			this.cbGarantia.Size = new System.Drawing.Size(392, 21);
			this.cbGarantia.TabIndex = 11;
			this.cbGarantia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbGarantia_KeyDown);
			// 
			// lblTitPedidosComGarantia
			// 
			this.lblTitPedidosComGarantia.AutoSize = true;
			this.lblTitPedidosComGarantia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPedidosComGarantia.Location = new System.Drawing.Point(554, 109);
			this.lblTitPedidosComGarantia.Name = "lblTitPedidosComGarantia";
			this.lblTitPedidosComGarantia.Size = new System.Drawing.Size(55, 13);
			this.lblTitPedidosComGarantia.TabIndex = 67;
			this.lblTitPedidosComGarantia.Text = "Garantia";
			// 
			// cbIndicador
			// 
			this.cbIndicador.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbIndicador.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbIndicador.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbIndicador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbIndicador.FormattingEnabled = true;
			this.cbIndicador.Location = new System.Drawing.Point(106, 104);
			this.cbIndicador.MaxDropDownItems = 12;
			this.cbIndicador.Name = "cbIndicador";
			this.cbIndicador.Size = new System.Drawing.Size(392, 21);
			this.cbIndicador.TabIndex = 10;
			this.cbIndicador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbIndicador_KeyDown);
			// 
			// lblTitIndicador
			// 
			this.lblTitIndicador.AutoSize = true;
			this.lblTitIndicador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitIndicador.Location = new System.Drawing.Point(39, 109);
			this.lblTitIndicador.Name = "lblTitIndicador";
			this.lblTitIndicador.Size = new System.Drawing.Size(60, 13);
			this.lblTitIndicador.TabIndex = 66;
			this.lblTitIndicador.Text = "Indicador";
			// 
			// cbVendedor
			// 
			this.cbVendedor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbVendedor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbVendedor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbVendedor.FormattingEnabled = true;
			this.cbVendedor.Location = new System.Drawing.Point(615, 79);
			this.cbVendedor.MaxDropDownItems = 12;
			this.cbVendedor.Name = "cbVendedor";
			this.cbVendedor.Size = new System.Drawing.Size(392, 21);
			this.cbVendedor.TabIndex = 9;
			this.cbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbVendedor_KeyDown);
			// 
			// lblTitVendedor
			// 
			this.lblTitVendedor.AutoSize = true;
			this.lblTitVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitVendedor.Location = new System.Drawing.Point(548, 84);
			this.lblTitVendedor.Name = "lblTitVendedor";
			this.lblTitVendedor.Size = new System.Drawing.Size(61, 13);
			this.lblTitVendedor.TabIndex = 64;
			this.lblTitVendedor.Text = "Vendedor";
			// 
			// cbEquipeVendas
			// 
			this.cbEquipeVendas.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbEquipeVendas.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbEquipeVendas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbEquipeVendas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbEquipeVendas.FormattingEnabled = true;
			this.cbEquipeVendas.Location = new System.Drawing.Point(106, 79);
			this.cbEquipeVendas.MaxDropDownItems = 12;
			this.cbEquipeVendas.Name = "cbEquipeVendas";
			this.cbEquipeVendas.Size = new System.Drawing.Size(392, 21);
			this.cbEquipeVendas.TabIndex = 8;
			this.cbEquipeVendas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbEquipeVendas_KeyDown);
			// 
			// lblTitEquipeVendas
			// 
			this.lblTitEquipeVendas.AutoSize = true;
			this.lblTitEquipeVendas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitEquipeVendas.Location = new System.Drawing.Point(8, 84);
			this.lblTitEquipeVendas.Name = "lblTitEquipeVendas";
			this.lblTitEquipeVendas.Size = new System.Drawing.Size(92, 13);
			this.lblTitEquipeVendas.TabIndex = 62;
			this.lblTitEquipeVendas.Text = "Equipe Vendas";
			// 
			// lblTitAtrasadoEntreDias
			// 
			this.lblTitAtrasadoEntreDias.AutoSize = true;
			this.lblTitAtrasadoEntreDias.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitAtrasadoEntreDias.Location = new System.Drawing.Point(751, 10);
			this.lblTitAtrasadoEntreDias.Name = "lblTitAtrasadoEntreDias";
			this.lblTitAtrasadoEntreDias.Size = new System.Drawing.Size(30, 13);
			this.lblTitAtrasadoEntreDias.TabIndex = 60;
			this.lblTitAtrasadoEntreDias.Text = "dias";
			// 
			// lblTitAtrasadoEntreE
			// 
			this.lblTitAtrasadoEntreE.AutoSize = true;
			this.lblTitAtrasadoEntreE.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitAtrasadoEntreE.Location = new System.Drawing.Point(673, 10);
			this.lblTitAtrasadoEntreE.Name = "lblTitAtrasadoEntreE";
			this.lblTitAtrasadoEntreE.Size = new System.Drawing.Size(14, 13);
			this.lblTitAtrasadoEntreE.TabIndex = 59;
			this.lblTitAtrasadoEntreE.Text = "e";
			// 
			// txtQtdeDiasAtrasoFinal
			// 
			this.txtQtdeDiasAtrasoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtQtdeDiasAtrasoFinal.Location = new System.Drawing.Point(693, 5);
			this.txtQtdeDiasAtrasoFinal.MaxLength = 4;
			this.txtQtdeDiasAtrasoFinal.Name = "txtQtdeDiasAtrasoFinal";
			this.txtQtdeDiasAtrasoFinal.Size = new System.Drawing.Size(52, 20);
			this.txtQtdeDiasAtrasoFinal.TabIndex = 2;
			this.txtQtdeDiasAtrasoFinal.Text = "9999";
			this.txtQtdeDiasAtrasoFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtQtdeDiasAtrasoFinal.Enter += new System.EventHandler(this.txtQtdeDiasAtrasoFinal_Enter);
			this.txtQtdeDiasAtrasoFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtQtdeDiasAtrasoFinal_KeyDown);
			this.txtQtdeDiasAtrasoFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtQtdeDiasAtrasoFinal_KeyPress);
			this.txtQtdeDiasAtrasoFinal.Leave += new System.EventHandler(this.txtQtdeDiasAtrasoFinal_Leave);
			// 
			// txtQtdeDiasAtrasoInicial
			// 
			this.txtQtdeDiasAtrasoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtQtdeDiasAtrasoInicial.Location = new System.Drawing.Point(615, 5);
			this.txtQtdeDiasAtrasoInicial.MaxLength = 4;
			this.txtQtdeDiasAtrasoInicial.Name = "txtQtdeDiasAtrasoInicial";
			this.txtQtdeDiasAtrasoInicial.Size = new System.Drawing.Size(52, 20);
			this.txtQtdeDiasAtrasoInicial.TabIndex = 1;
			this.txtQtdeDiasAtrasoInicial.Text = "0000";
			this.txtQtdeDiasAtrasoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtQtdeDiasAtrasoInicial.Enter += new System.EventHandler(this.txtQtdeDiasAtrasoInicial_Enter);
			this.txtQtdeDiasAtrasoInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtQtdeDiasAtrasoInicial_KeyDown);
			this.txtQtdeDiasAtrasoInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtQtdeDiasAtrasoInicial_KeyPress);
			this.txtQtdeDiasAtrasoInicial.Leave += new System.EventHandler(this.txtQtdeDiasAtrasoInicial_Leave);
			// 
			// lblTitAtrasadoEntre
			// 
			this.lblTitAtrasadoEntre.AutoSize = true;
			this.lblTitAtrasadoEntre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitAtrasadoEntre.Location = new System.Drawing.Point(519, 10);
			this.lblTitAtrasadoEntre.Name = "lblTitAtrasadoEntre";
			this.lblTitAtrasadoEntre.Size = new System.Drawing.Size(90, 13);
			this.lblTitAtrasadoEntre.TabIndex = 57;
			this.lblTitAtrasadoEntre.Text = "Atrasado entre";
			// 
			// cbSituacao
			// 
			this.cbSituacao.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbSituacao.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbSituacao.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbSituacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbSituacao.FormattingEnabled = true;
			this.cbSituacao.Location = new System.Drawing.Point(106, 4);
			this.cbSituacao.MaxDropDownItems = 12;
			this.cbSituacao.Name = "cbSituacao";
			this.cbSituacao.Size = new System.Drawing.Size(392, 21);
			this.cbSituacao.TabIndex = 0;
			this.cbSituacao.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbSituacao_KeyDown);
			// 
			// lblTitSituacao
			// 
			this.lblTitSituacao.AutoSize = true;
			this.lblTitSituacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitSituacao.Location = new System.Drawing.Point(43, 9);
			this.lblTitSituacao.Name = "lblTitSituacao";
			this.lblTitSituacao.Size = new System.Drawing.Size(57, 13);
			this.lblTitSituacao.TabIndex = 55;
			this.lblTitSituacao.Text = "Situação";
			// 
			// txtNomeCliente
			// 
			this.txtNomeCliente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.txtNomeCliente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
			this.txtNomeCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNomeCliente.Location = new System.Drawing.Point(106, 129);
			this.txtNomeCliente.MaxLength = 60;
			this.txtNomeCliente.Name = "txtNomeCliente";
			this.txtNomeCliente.Size = new System.Drawing.Size(392, 20);
			this.txtNomeCliente.TabIndex = 12;
			this.txtNomeCliente.Enter += new System.EventHandler(this.txtNomeCliente_Enter);
			this.txtNomeCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNomeCliente_KeyDown);
			this.txtNomeCliente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNomeCliente_KeyPress);
			this.txtNomeCliente.Leave += new System.EventHandler(this.txtNomeCliente_Leave);
			// 
			// lblTitNomeCliente
			// 
			this.lblTitNomeCliente.AutoSize = true;
			this.lblTitNomeCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNomeCliente.Location = new System.Drawing.Point(18, 134);
			this.lblTitNomeCliente.Name = "lblTitNomeCliente";
			this.lblTitNomeCliente.Size = new System.Drawing.Size(82, 13);
			this.lblTitNomeCliente.TabIndex = 53;
			this.lblTitNomeCliente.Text = "Nome Cliente";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(615, 129);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(171, 20);
			this.txtCnpjCpf.TabIndex = 13;
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
			this.lblCnpjCpf.Location = new System.Drawing.Point(541, 134);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(67, 13);
			this.lblCnpjCpf.TabIndex = 52;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// cbPlanoContasConta
			// 
			this.cbPlanoContasConta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasConta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasConta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasConta.FormattingEnabled = true;
			this.cbPlanoContasConta.Location = new System.Drawing.Point(615, 54);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(392, 21);
			this.cbPlanoContasConta.TabIndex = 7;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblTitPlanoContasConta
			// 
			this.lblTitPlanoContasConta.AutoSize = true;
			this.lblTitPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasConta.Location = new System.Drawing.Point(569, 59);
			this.lblTitPlanoContasConta.Name = "lblTitPlanoContasConta";
			this.lblTitPlanoContasConta.Size = new System.Drawing.Size(40, 13);
			this.lblTitPlanoContasConta.TabIndex = 51;
			this.lblTitPlanoContasConta.Text = "Conta";
			// 
			// cbPlanoContasGrupo
			// 
			this.cbPlanoContasGrupo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasGrupo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasGrupo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasGrupo.FormattingEnabled = true;
			this.cbPlanoContasGrupo.Location = new System.Drawing.Point(106, 54);
			this.cbPlanoContasGrupo.MaxDropDownItems = 12;
			this.cbPlanoContasGrupo.Name = "cbPlanoContasGrupo";
			this.cbPlanoContasGrupo.Size = new System.Drawing.Size(392, 21);
			this.cbPlanoContasGrupo.TabIndex = 6;
			this.cbPlanoContasGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasGrupo_KeyDown);
			// 
			// lblTitPlanoContasGrupo
			// 
			this.lblTitPlanoContasGrupo.AutoSize = true;
			this.lblTitPlanoContasGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPlanoContasGrupo.Location = new System.Drawing.Point(59, 59);
			this.lblTitPlanoContasGrupo.Name = "lblTitPlanoContasGrupo";
			this.lblTitPlanoContasGrupo.Size = new System.Drawing.Size(41, 13);
			this.lblTitPlanoContasGrupo.TabIndex = 50;
			this.lblTitPlanoContasGrupo.Text = "Grupo";
			// 
			// cbPlanoContasEmpresa
			// 
			this.cbPlanoContasEmpresa.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasEmpresa.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasEmpresa.FormattingEnabled = true;
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(615, 29);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(392, 21);
			this.cbPlanoContasEmpresa.TabIndex = 5;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(554, 34);
			this.lblPlanoContasEmpresa.Name = "lblPlanoContasEmpresa";
			this.lblPlanoContasEmpresa.Size = new System.Drawing.Size(55, 13);
			this.lblPlanoContasEmpresa.TabIndex = 49;
			this.lblPlanoContasEmpresa.Text = "Empresa";
			// 
			// cbContaCorrente
			// 
			this.cbContaCorrente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbContaCorrente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbContaCorrente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbContaCorrente.FormattingEnabled = true;
			this.cbContaCorrente.Location = new System.Drawing.Point(106, 29);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(392, 21);
			this.cbContaCorrente.TabIndex = 4;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblTitContaCorrente
			// 
			this.lblTitContaCorrente.AutoSize = true;
			this.lblTitContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitContaCorrente.Location = new System.Drawing.Point(8, 34);
			this.lblTitContaCorrente.Name = "lblTitContaCorrente";
			this.lblTitContaCorrente.Size = new System.Drawing.Size(92, 13);
			this.lblTitContaCorrente.TabIndex = 48;
			this.lblTitContaCorrente.Text = "Conta Corrente";
			// 
			// pnTotalizacao
			// 
			this.pnTotalizacao.Controls.Add(this.btnDesmarcarTodos);
			this.pnTotalizacao.Controls.Add(this.btnMarcarTodos);
			this.pnTotalizacao.Controls.Add(this.lblTitTotalizacaoRegistros);
			this.pnTotalizacao.Controls.Add(this.lblTitTotalizacaoValor);
			this.pnTotalizacao.Controls.Add(this.lblTotalizacaoValor);
			this.pnTotalizacao.Controls.Add(this.lblTotalizacaoRegistros);
			this.pnTotalizacao.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.pnTotalizacao.Location = new System.Drawing.Point(0, 576);
			this.pnTotalizacao.Name = "pnTotalizacao";
			this.pnTotalizacao.Size = new System.Drawing.Size(1014, 29);
			this.pnTotalizacao.TabIndex = 2;
			// 
			// btnDesmarcarTodos
			// 
			this.btnDesmarcarTodos.Location = new System.Drawing.Point(118, 3);
			this.btnDesmarcarTodos.Name = "btnDesmarcarTodos";
			this.btnDesmarcarTodos.Size = new System.Drawing.Size(110, 23);
			this.btnDesmarcarTodos.TabIndex = 1;
			this.btnDesmarcarTodos.Text = "Desmarcar Todos";
			this.btnDesmarcarTodos.UseVisualStyleBackColor = true;
			this.btnDesmarcarTodos.Click += new System.EventHandler(this.btnDesmarcarTodos_Click);
			// 
			// btnMarcarTodos
			// 
			this.btnMarcarTodos.Location = new System.Drawing.Point(3, 3);
			this.btnMarcarTodos.Name = "btnMarcarTodos";
			this.btnMarcarTodos.Size = new System.Drawing.Size(110, 23);
			this.btnMarcarTodos.TabIndex = 0;
			this.btnMarcarTodos.Text = "Marcar Todos";
			this.btnMarcarTodos.UseVisualStyleBackColor = true;
			this.btnMarcarTodos.Click += new System.EventHandler(this.btnMarcarTodos_Click);
			// 
			// lblTitTotalizacaoRegistros
			// 
			this.lblTitTotalizacaoRegistros.AutoSize = true;
			this.lblTitTotalizacaoRegistros.Location = new System.Drawing.Point(697, 8);
			this.lblTitTotalizacaoRegistros.Name = "lblTitTotalizacaoRegistros";
			this.lblTitTotalizacaoRegistros.Size = new System.Drawing.Size(54, 13);
			this.lblTitTotalizacaoRegistros.TabIndex = 5;
			this.lblTitTotalizacaoRegistros.Text = "Registros:";
			// 
			// lblTitTotalizacaoValor
			// 
			this.lblTitTotalizacaoValor.AutoSize = true;
			this.lblTitTotalizacaoValor.Location = new System.Drawing.Point(822, 8);
			this.lblTitTotalizacaoValor.Name = "lblTitTotalizacaoValor";
			this.lblTitTotalizacaoValor.Size = new System.Drawing.Size(61, 13);
			this.lblTitTotalizacaoValor.TabIndex = 4;
			this.lblTitTotalizacaoValor.Text = "Valor Total:";
			// 
			// lblTotalizacaoValor
			// 
			this.lblTotalizacaoValor.AutoSize = true;
			this.lblTotalizacaoValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoValor.Location = new System.Drawing.Point(889, 8);
			this.lblTotalizacaoValor.Name = "lblTotalizacaoValor";
			this.lblTotalizacaoValor.Size = new System.Drawing.Size(96, 13);
			this.lblTotalizacaoValor.TabIndex = 7;
			this.lblTotalizacaoValor.Text = "999.999.999,99";
			// 
			// lblTotalizacaoRegistros
			// 
			this.lblTotalizacaoRegistros.AutoSize = true;
			this.lblTotalizacaoRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoRegistros.Location = new System.Drawing.Point(757, 8);
			this.lblTotalizacaoRegistros.Name = "lblTotalizacaoRegistros";
			this.lblTotalizacaoRegistros.Size = new System.Drawing.Size(21, 13);
			this.lblTotalizacaoRegistros.TabIndex = 6;
			this.lblTotalizacaoRegistros.Text = "99";
			// 
			// pnResultado
			// 
			this.pnResultado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnResultado.Controls.Add(this.gridDados);
			this.pnResultado.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pnResultado.Location = new System.Drawing.Point(0, 193);
			this.pnResultado.Name = "pnResultado";
			this.pnResultado.Size = new System.Drawing.Size(1014, 383);
			this.pnResultado.TabIndex = 1;
			// 
			// gridDados
			// 
			this.gridDados.AllowUserToAddRows = false;
			this.gridDados.AllowUserToDeleteRows = false;
			this.gridDados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.gridDados.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.gridDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.gridDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colCheckBox,
            this.colGridNomeCnpjCpf,
            this.colGridQtdeParcelasEmAtraso,
            this.colGridMaxDiasEmAtraso,
            this.colGridValorTotalEmAtraso,
            this.colGridDescricaoParcelas,
            this.colGridNumParcelaMaiorAtraso,
            this.colGridVendedor,
            this.colGridIndicador,
            this.colGridUF,
            this.colGridIdCliente});
			dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.gridDados.DefaultCellStyle = dataGridViewCellStyle11;
			this.gridDados.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridDados.Location = new System.Drawing.Point(0, 0);
			this.gridDados.MultiSelect = false;
			this.gridDados.Name = "gridDados";
			dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.RowHeadersDefaultCellStyle = dataGridViewCellStyle12;
			this.gridDados.RowHeadersVisible = false;
			this.gridDados.RowHeadersWidth = 15;
			this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridDados.Size = new System.Drawing.Size(1010, 379);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 0;
			this.gridDados.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridDados_CellContentClick);
			this.gridDados.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.gridDados_SortCompare);
			// 
			// colCheckBox
			// 
			this.colCheckBox.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.colCheckBox.Frozen = true;
			this.colCheckBox.HeaderText = "";
			this.colCheckBox.MinimumWidth = 20;
			this.colCheckBox.Name = "colCheckBox";
			this.colCheckBox.Width = 20;
			// 
			// colGridNomeCnpjCpf
			// 
			this.colGridNomeCnpjCpf.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colGridNomeCnpjCpf.DefaultCellStyle = dataGridViewCellStyle2;
			this.colGridNomeCnpjCpf.HeaderText = "Nome/CNPJ/CPF";
			this.colGridNomeCnpjCpf.MinimumWidth = 150;
			this.colGridNomeCnpjCpf.Name = "colGridNomeCnpjCpf";
			this.colGridNomeCnpjCpf.ReadOnly = true;
			// 
			// colGridQtdeParcelasEmAtraso
			// 
			this.colGridQtdeParcelasEmAtraso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridQtdeParcelasEmAtraso.DefaultCellStyle = dataGridViewCellStyle3;
			this.colGridQtdeParcelasEmAtraso.HeaderText = "Parcelas em atraso";
			this.colGridQtdeParcelasEmAtraso.MinimumWidth = 90;
			this.colGridQtdeParcelasEmAtraso.Name = "colGridQtdeParcelasEmAtraso";
			this.colGridQtdeParcelasEmAtraso.ReadOnly = true;
			this.colGridQtdeParcelasEmAtraso.Width = 90;
			// 
			// colGridMaxDiasEmAtraso
			// 
			this.colGridMaxDiasEmAtraso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridMaxDiasEmAtraso.DefaultCellStyle = dataGridViewCellStyle4;
			this.colGridMaxDiasEmAtraso.HeaderText = "Maior atraso";
			this.colGridMaxDiasEmAtraso.MinimumWidth = 70;
			this.colGridMaxDiasEmAtraso.Name = "colGridMaxDiasEmAtraso";
			this.colGridMaxDiasEmAtraso.ReadOnly = true;
			this.colGridMaxDiasEmAtraso.Width = 70;
			// 
			// colGridValorTotalEmAtraso
			// 
			this.colGridValorTotalEmAtraso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.colGridValorTotalEmAtraso.DefaultCellStyle = dataGridViewCellStyle5;
			this.colGridValorTotalEmAtraso.HeaderText = "Valor total em atraso";
			this.colGridValorTotalEmAtraso.MinimumWidth = 110;
			this.colGridValorTotalEmAtraso.Name = "colGridValorTotalEmAtraso";
			this.colGridValorTotalEmAtraso.ReadOnly = true;
			this.colGridValorTotalEmAtraso.Width = 110;
			// 
			// colGridDescricaoParcelas
			// 
			this.colGridDescricaoParcelas.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle6.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colGridDescricaoParcelas.DefaultCellStyle = dataGridViewCellStyle6;
			this.colGridDescricaoParcelas.HeaderText = "Descrição das parcelas em atraso";
			this.colGridDescricaoParcelas.MinimumWidth = 180;
			this.colGridDescricaoParcelas.Name = "colGridDescricaoParcelas";
			this.colGridDescricaoParcelas.ReadOnly = true;
			this.colGridDescricaoParcelas.Width = 180;
			// 
			// colGridNumParcelaMaiorAtraso
			// 
			this.colGridNumParcelaMaiorAtraso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridNumParcelaMaiorAtraso.DefaultCellStyle = dataGridViewCellStyle7;
			this.colGridNumParcelaMaiorAtraso.HeaderText = "Parc maior atraso";
			this.colGridNumParcelaMaiorAtraso.MinimumWidth = 50;
			this.colGridNumParcelaMaiorAtraso.Name = "colGridNumParcelaMaiorAtraso";
			this.colGridNumParcelaMaiorAtraso.ReadOnly = true;
			this.colGridNumParcelaMaiorAtraso.Width = 50;
			// 
			// colGridVendedor
			// 
			this.colGridVendedor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridVendedor.DefaultCellStyle = dataGridViewCellStyle8;
			this.colGridVendedor.HeaderText = "Vendedor";
			this.colGridVendedor.MinimumWidth = 75;
			this.colGridVendedor.Name = "colGridVendedor";
			this.colGridVendedor.ReadOnly = true;
			this.colGridVendedor.Width = 75;
			// 
			// colGridIndicador
			// 
			this.colGridIndicador.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridIndicador.DefaultCellStyle = dataGridViewCellStyle9;
			this.colGridIndicador.HeaderText = "Parceiro";
			this.colGridIndicador.MinimumWidth = 75;
			this.colGridIndicador.Name = "colGridIndicador";
			this.colGridIndicador.ReadOnly = true;
			this.colGridIndicador.Width = 75;
			// 
			// colGridUF
			// 
			this.colGridUF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colGridUF.DefaultCellStyle = dataGridViewCellStyle10;
			this.colGridUF.HeaderText = "UF";
			this.colGridUF.MinimumWidth = 30;
			this.colGridUF.Name = "colGridUF";
			this.colGridUF.ReadOnly = true;
			this.colGridUF.Width = 30;
			// 
			// colGridIdCliente
			// 
			this.colGridIdCliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.colGridIdCliente.HeaderText = "id_cliente";
			this.colGridIdCliente.Name = "colGridIdCliente";
			this.colGridIdCliente.ReadOnly = true;
			this.colGridIdCliente.Visible = false;
			// 
			// btnPesquisar
			// 
			this.btnPesquisar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPesquisar.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisar.Image")));
			this.btnPesquisar.Location = new System.Drawing.Point(699, 4);
			this.btnPesquisar.Name = "btnPesquisar";
			this.btnPesquisar.Size = new System.Drawing.Size(40, 44);
			this.btnPesquisar.TabIndex = 0;
			this.btnPesquisar.TabStop = false;
			this.btnPesquisar.UseVisualStyleBackColor = true;
			this.btnPesquisar.Click += new System.EventHandler(this.btnPesquisar_Click);
			// 
			// btnPrinterDialog
			// 
			this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
			this.btnPrinterDialog.Location = new System.Drawing.Point(879, 4);
			this.btnPrinterDialog.Name = "btnPrinterDialog";
			this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
			this.btnPrinterDialog.TabIndex = 4;
			this.btnPrinterDialog.TabStop = false;
			this.btnPrinterDialog.UseVisualStyleBackColor = true;
			this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
			// 
			// btnPrintPreview
			// 
			this.btnPrintPreview.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrintPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnPrintPreview.Image")));
			this.btnPrintPreview.Location = new System.Drawing.Point(834, 4);
			this.btnPrintPreview.Name = "btnPrintPreview";
			this.btnPrintPreview.Size = new System.Drawing.Size(40, 44);
			this.btnPrintPreview.TabIndex = 3;
			this.btnPrintPreview.TabStop = false;
			this.btnPrintPreview.UseVisualStyleBackColor = true;
			this.btnPrintPreview.Click += new System.EventHandler(this.btnPrintPreview_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(789, 4);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(40, 44);
			this.btnImprimir.TabIndex = 2;
			this.btnImprimir.TabStop = false;
			this.btnImprimir.UseVisualStyleBackColor = true;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// prnDocConsulta
			// 
			this.prnDocConsulta.DocumentName = "Carteira em Atraso";
			this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
			this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
			// 
			// prnDialogConsulta
			// 
			this.prnDialogConsulta.Document = this.prnDocConsulta;
			this.prnDialogConsulta.UseEXDialog = true;
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
			// btnExcel
			// 
			this.btnExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.Location = new System.Drawing.Point(744, 4);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(40, 44);
			this.btnExcel.TabIndex = 1;
			this.btnExcel.TabStop = false;
			this.btnExcel.UseVisualStyleBackColor = true;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// FCobrancaAdministracao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FCobrancaAdministracao";
			this.Text = "Artven - Financeiro  -  1.11 - XX.XXX.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FCobrancaAdministracao_FormClosing);
			this.Load += new System.EventHandler(this.FCobrancaAdministracao_Load);
			this.Shown += new System.EventHandler(this.FCobrancaAdministracao_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FCobrancaAdministracao_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnParametros.ResumeLayout(false);
			this.pnParametros.PerformLayout();
			this.pnTotalizacao.ResumeLayout(false);
			this.pnTotalizacao.PerformLayout();
			this.pnResultado.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Panel pnParametros;
		private System.Windows.Forms.Panel pnTotalizacao;
		private System.Windows.Forms.Button btnDesmarcarTodos;
		private System.Windows.Forms.Button btnMarcarTodos;
		private System.Windows.Forms.Label lblTitTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTitTotalizacaoValor;
		private System.Windows.Forms.Label lblTotalizacaoValor;
		private System.Windows.Forms.Label lblTotalizacaoRegistros;
		private System.Windows.Forms.Panel pnResultado;
		private System.Windows.Forms.DataGridView gridDados;
		private System.Windows.Forms.TextBox txtNomeCliente;
		private System.Windows.Forms.Label lblTitNomeCliente;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.ComboBox cbPlanoContasConta;
		private System.Windows.Forms.Label lblTitPlanoContasConta;
		private System.Windows.Forms.ComboBox cbPlanoContasGrupo;
		private System.Windows.Forms.Label lblTitPlanoContasGrupo;
		private System.Windows.Forms.ComboBox cbPlanoContasEmpresa;
		private System.Windows.Forms.Label lblPlanoContasEmpresa;
		private System.Windows.Forms.ComboBox cbContaCorrente;
		private System.Windows.Forms.Label lblTitContaCorrente;
		private System.Windows.Forms.TextBox txtQtdeDiasAtrasoInicial;
		private System.Windows.Forms.Label lblTitAtrasadoEntre;
		private System.Windows.Forms.ComboBox cbSituacao;
		private System.Windows.Forms.Label lblTitSituacao;
		private System.Windows.Forms.Label lblTitAtrasadoEntreDias;
		private System.Windows.Forms.Label lblTitAtrasadoEntreE;
		private System.Windows.Forms.TextBox txtQtdeDiasAtrasoFinal;
		private System.Windows.Forms.ComboBox cbEquipeVendas;
		private System.Windows.Forms.Label lblTitEquipeVendas;
		private System.Windows.Forms.ComboBox cbIndicador;
		private System.Windows.Forms.Label lblTitIndicador;
		private System.Windows.Forms.ComboBox cbVendedor;
		private System.Windows.Forms.Label lblTitVendedor;
		private System.Windows.Forms.Label lblTitPedidosComGarantia;
		private System.Windows.Forms.ComboBox cbGarantia;
		private System.Windows.Forms.Button btnPesquisar;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnPrintPreview;
		private System.Windows.Forms.Button btnImprimir;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Windows.Forms.CheckBox ckb_somente_primeira_parcela_em_atraso;
		private System.Windows.Forms.DataGridViewCheckBoxColumn colCheckBox;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridNomeCnpjCpf;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridQtdeParcelasEmAtraso;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridMaxDiasEmAtraso;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridValorTotalEmAtraso;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridDescricaoParcelas;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridNumParcelaMaiorAtraso;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridVendedor;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridIndicador;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridUF;
		private System.Windows.Forms.DataGridViewTextBoxColumn colGridIdCliente;
		private System.Windows.Forms.Button btnExcel;
	}
}
