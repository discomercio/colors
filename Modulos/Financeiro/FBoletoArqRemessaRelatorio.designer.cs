namespace Financeiro
{
	partial class FBoletoArqRemessaRelatorio
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoArqRemessaRelatorio));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblTitArqRemessa = new System.Windows.Forms.Label();
			this.txtArqRemessa = new System.Windows.Forms.TextBox();
			this.btnSelecionaArqRemessa = new System.Windows.Forms.Button();
			this.gboxBoletos = new System.Windows.Forms.GroupBox();
			this.lblTitTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalRegistros = new System.Windows.Forms.Label();
			this.grdBoletos = new System.Windows.Forms.DataGridView();
			this.colSacado = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colEndereco = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colLoja = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colNumeroDocumento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDataVencto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colValorTitulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colIdBoletoItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.gboxRelatorio = new System.Windows.Forms.GroupBox();
			this.gboxOpcaoSaida = new System.Windows.Forms.GroupBox();
			this.rbSaidaVisualizacao = new System.Windows.Forms.RadioButton();
			this.rbSaidaImpressora = new System.Windows.Forms.RadioButton();
			this.btnListagemArqRemessa = new System.Windows.Forms.Button();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
			this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxBoletos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).BeginInit();
			this.gboxRelatorio.SuspendLayout();
			this.gboxOpcaoSaida.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnPrinterDialog);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxRelatorio);
			this.pnCampos.Controls.Add(this.gboxBoletos);
			this.pnCampos.Controls.Add(this.btnSelecionaArqRemessa);
			this.pnCampos.Controls.Add(this.txtArqRemessa);
			this.pnCampos.Controls.Add(this.lblTitArqRemessa);
			this.pnCampos.Controls.Add(this.lblTitulo);
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
			this.lblTitulo.Text = "Boleto: Relatório do Arquivo de Remessa";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblTitArqRemessa
			// 
			this.lblTitArqRemessa.AutoSize = true;
			this.lblTitArqRemessa.Location = new System.Drawing.Point(12, 50);
			this.lblTitArqRemessa.Name = "lblTitArqRemessa";
			this.lblTitArqRemessa.Size = new System.Drawing.Size(105, 13);
			this.lblTitArqRemessa.TabIndex = 1;
			this.lblTitArqRemessa.Text = "Arquivo de Remessa";
			// 
			// txtArqRemessa
			// 
			this.txtArqRemessa.BackColor = System.Drawing.SystemColors.Window;
			this.txtArqRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtArqRemessa.Location = new System.Drawing.Point(123, 47);
			this.txtArqRemessa.Name = "txtArqRemessa";
			this.txtArqRemessa.ReadOnly = true;
			this.txtArqRemessa.Size = new System.Drawing.Size(695, 20);
			this.txtArqRemessa.TabIndex = 0;
			this.txtArqRemessa.DoubleClick += new System.EventHandler(this.txtArqRemessa_DoubleClick);
			this.txtArqRemessa.Enter += new System.EventHandler(this.txtArqRemessa_Enter);
			// 
			// btnSelecionaArqRemessa
			// 
			this.btnSelecionaArqRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSelecionaArqRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArqRemessa.Image")));
			this.btnSelecionaArqRemessa.Location = new System.Drawing.Point(825, 44);
			this.btnSelecionaArqRemessa.Name = "btnSelecionaArqRemessa";
			this.btnSelecionaArqRemessa.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaArqRemessa.TabIndex = 1;
			this.btnSelecionaArqRemessa.UseVisualStyleBackColor = true;
			this.btnSelecionaArqRemessa.Click += new System.EventHandler(this.btnSelecionaArqRemessa_Click);
			// 
			// gboxBoletos
			// 
			this.gboxBoletos.Controls.Add(this.lblTitTotalizacaoValor);
			this.gboxBoletos.Controls.Add(this.lblTotalizacaoValor);
			this.gboxBoletos.Controls.Add(this.lblTotalRegistros);
			this.gboxBoletos.Controls.Add(this.lblTitTotalRegistros);
			this.gboxBoletos.Controls.Add(this.grdBoletos);
			this.gboxBoletos.Location = new System.Drawing.Point(10, 71);
			this.gboxBoletos.Name = "gboxBoletos";
			this.gboxBoletos.Size = new System.Drawing.Size(995, 385);
			this.gboxBoletos.TabIndex = 4;
			this.gboxBoletos.TabStop = false;
			this.gboxBoletos.Text = "Dados do Arquivo de Remessa";
			// 
			// lblTitTotalizacaoValor
			// 
			this.lblTitTotalizacaoValor.AutoSize = true;
			this.lblTitTotalizacaoValor.Location = new System.Drawing.Point(807, 367);
			this.lblTitTotalizacaoValor.Name = "lblTitTotalizacaoValor";
			this.lblTitTotalizacaoValor.Size = new System.Drawing.Size(61, 13);
			this.lblTitTotalizacaoValor.TabIndex = 8;
			this.lblTitTotalizacaoValor.Text = "Valor Total:";
			// 
			// lblTotalizacaoValor
			// 
			this.lblTotalizacaoValor.AutoSize = true;
			this.lblTotalizacaoValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoValor.Location = new System.Drawing.Point(874, 367);
			this.lblTotalizacaoValor.Name = "lblTotalizacaoValor";
			this.lblTotalizacaoValor.Size = new System.Drawing.Size(96, 13);
			this.lblTotalizacaoValor.TabIndex = 9;
			this.lblTotalizacaoValor.Text = "999.999.999,99";
			// 
			// lblTotalRegistros
			// 
			this.lblTotalRegistros.AutoSize = true;
			this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalRegistros.Location = new System.Drawing.Point(100, 367);
			this.lblTotalRegistros.Name = "lblTotalRegistros";
			this.lblTotalRegistros.Size = new System.Drawing.Size(28, 13);
			this.lblTotalRegistros.TabIndex = 6;
			this.lblTotalRegistros.Text = "999";
			// 
			// lblTitTotalRegistros
			// 
			this.lblTitTotalRegistros.AutoSize = true;
			this.lblTitTotalRegistros.Location = new System.Drawing.Point(12, 367);
			this.lblTitTotalRegistros.Name = "lblTitTotalRegistros";
			this.lblTitTotalRegistros.Size = new System.Drawing.Size(88, 13);
			this.lblTitTotalRegistros.TabIndex = 5;
			this.lblTitTotalRegistros.Text = "Total de registros";
			// 
			// grdBoletos
			// 
			this.grdBoletos.AllowUserToAddRows = false;
			this.grdBoletos.AllowUserToDeleteRows = false;
			this.grdBoletos.AllowUserToResizeColumns = false;
			this.grdBoletos.AllowUserToResizeRows = false;
			this.grdBoletos.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdBoletos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdBoletos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdBoletos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdBoletos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSacado,
            this.colEndereco,
            this.colLoja,
            this.colPedido,
            this.colNumeroDocumento,
            this.colDataVencto,
            this.colValorTitulo,
            this.colIdBoletoItem});
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdBoletos.DefaultCellStyle = dataGridViewCellStyle9;
			this.grdBoletos.Location = new System.Drawing.Point(15, 19);
			this.grdBoletos.MultiSelect = false;
			this.grdBoletos.Name = "grdBoletos";
			this.grdBoletos.ReadOnly = true;
			this.grdBoletos.RowHeadersVisible = false;
			this.grdBoletos.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdBoletos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdBoletos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdBoletos.ShowEditingIcon = false;
			this.grdBoletos.Size = new System.Drawing.Size(965, 346);
			this.grdBoletos.StandardTab = true;
			this.grdBoletos.TabIndex = 0;
			// 
			// colSacado
			// 
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colSacado.DefaultCellStyle = dataGridViewCellStyle2;
			this.colSacado.HeaderText = "Sacado";
			this.colSacado.MinimumWidth = 230;
			this.colSacado.Name = "colSacado";
			this.colSacado.ReadOnly = true;
			this.colSacado.Width = 230;
			// 
			// colEndereco
			// 
			this.colEndereco.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colEndereco.DefaultCellStyle = dataGridViewCellStyle3;
			this.colEndereco.HeaderText = "Endereço";
			this.colEndereco.MinimumWidth = 120;
			this.colEndereco.Name = "colEndereco";
			this.colEndereco.ReadOnly = true;
			// 
			// colLoja
			// 
			this.colLoja.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colLoja.DefaultCellStyle = dataGridViewCellStyle4;
			this.colLoja.HeaderText = "Loja";
			this.colLoja.MinimumWidth = 58;
			this.colLoja.Name = "colLoja";
			this.colLoja.ReadOnly = true;
			this.colLoja.Width = 58;
			// 
			// colPedido
			// 
			this.colPedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colPedido.DefaultCellStyle = dataGridViewCellStyle5;
			this.colPedido.HeaderText = "Pedido";
			this.colPedido.MinimumWidth = 75;
			this.colPedido.Name = "colPedido";
			this.colPedido.ReadOnly = true;
			this.colPedido.Width = 75;
			// 
			// colNumeroDocumento
			// 
			this.colNumeroDocumento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colNumeroDocumento.DefaultCellStyle = dataGridViewCellStyle6;
			this.colNumeroDocumento.HeaderText = "Nº Documento";
			this.colNumeroDocumento.MinimumWidth = 120;
			this.colNumeroDocumento.Name = "colNumeroDocumento";
			this.colNumeroDocumento.ReadOnly = true;
			this.colNumeroDocumento.Width = 120;
			// 
			// colDataVencto
			// 
			this.colDataVencto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colDataVencto.DefaultCellStyle = dataGridViewCellStyle7;
			this.colDataVencto.HeaderText = "Vencto";
			this.colDataVencto.MinimumWidth = 80;
			this.colDataVencto.Name = "colDataVencto";
			this.colDataVencto.ReadOnly = true;
			this.colDataVencto.Width = 80;
			// 
			// colValorTitulo
			// 
			this.colValorTitulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			this.colValorTitulo.DefaultCellStyle = dataGridViewCellStyle8;
			this.colValorTitulo.HeaderText = "Valor";
			this.colValorTitulo.MinimumWidth = 140;
			this.colValorTitulo.Name = "colValorTitulo";
			this.colValorTitulo.ReadOnly = true;
			this.colValorTitulo.Width = 140;
			// 
			// colIdBoletoItem
			// 
			this.colIdBoletoItem.HeaderText = "id_boleto_item";
			this.colIdBoletoItem.Name = "colIdBoletoItem";
			this.colIdBoletoItem.ReadOnly = true;
			this.colIdBoletoItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colIdBoletoItem.Visible = false;
			// 
			// openFileDialog
			// 
			this.openFileDialog.AddExtension = false;
			this.openFileDialog.Filter = "Arquivo de remessa|*.REM|Todos os arquivos|*.*";
			this.openFileDialog.InitialDirectory = "\\";
			this.openFileDialog.Title = "Selecionar arquivo de remessa";
			// 
			// gboxRelatorio
			// 
			this.gboxRelatorio.Controls.Add(this.gboxOpcaoSaida);
			this.gboxRelatorio.Controls.Add(this.btnListagemArqRemessa);
			this.gboxRelatorio.Location = new System.Drawing.Point(10, 468);
			this.gboxRelatorio.Name = "gboxRelatorio";
			this.gboxRelatorio.Size = new System.Drawing.Size(994, 128);
			this.gboxRelatorio.TabIndex = 6;
			this.gboxRelatorio.TabStop = false;
			this.gboxRelatorio.Text = "Relatório";
			// 
			// gboxOpcaoSaida
			// 
			this.gboxOpcaoSaida.Controls.Add(this.rbSaidaVisualizacao);
			this.gboxOpcaoSaida.Controls.Add(this.rbSaidaImpressora);
			this.gboxOpcaoSaida.Location = new System.Drawing.Point(18, 28);
			this.gboxOpcaoSaida.Name = "gboxOpcaoSaida";
			this.gboxOpcaoSaida.Size = new System.Drawing.Size(132, 73);
			this.gboxOpcaoSaida.TabIndex = 0;
			this.gboxOpcaoSaida.TabStop = false;
			this.gboxOpcaoSaida.Text = "Opção de saída";
			// 
			// rbSaidaVisualizacao
			// 
			this.rbSaidaVisualizacao.AutoSize = true;
			this.rbSaidaVisualizacao.Location = new System.Drawing.Point(21, 46);
			this.rbSaidaVisualizacao.Name = "rbSaidaVisualizacao";
			this.rbSaidaVisualizacao.Size = new System.Drawing.Size(87, 17);
			this.rbSaidaVisualizacao.TabIndex = 1;
			this.rbSaidaVisualizacao.TabStop = true;
			this.rbSaidaVisualizacao.Text = "Print Preview";
			this.rbSaidaVisualizacao.UseVisualStyleBackColor = true;
			// 
			// rbSaidaImpressora
			// 
			this.rbSaidaImpressora.AutoSize = true;
			this.rbSaidaImpressora.Location = new System.Drawing.Point(21, 23);
			this.rbSaidaImpressora.Name = "rbSaidaImpressora";
			this.rbSaidaImpressora.Size = new System.Drawing.Size(76, 17);
			this.rbSaidaImpressora.TabIndex = 0;
			this.rbSaidaImpressora.TabStop = true;
			this.rbSaidaImpressora.Text = "Impressora";
			this.rbSaidaImpressora.UseVisualStyleBackColor = true;
			// 
			// btnListagemArqRemessa
			// 
			this.btnListagemArqRemessa.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnListagemArqRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnListagemArqRemessa.ForeColor = System.Drawing.Color.Black;
			this.btnListagemArqRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnListagemArqRemessa.Image")));
			this.btnListagemArqRemessa.Location = new System.Drawing.Point(297, 45);
			this.btnListagemArqRemessa.Name = "btnListagemArqRemessa";
			this.btnListagemArqRemessa.Size = new System.Drawing.Size(400, 38);
			this.btnListagemArqRemessa.TabIndex = 1;
			this.btnListagemArqRemessa.Text = "Listagem do Arquivo de Remessa";
			this.btnListagemArqRemessa.UseVisualStyleBackColor = true;
			this.btnListagemArqRemessa.Click += new System.EventHandler(this.btnListagemArqRemessa_Click);
			// 
			// btnPrinterDialog
			// 
			this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
			this.btnPrinterDialog.Location = new System.Drawing.Point(879, 4);
			this.btnPrinterDialog.Name = "btnPrinterDialog";
			this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
			this.btnPrinterDialog.TabIndex = 0;
			this.btnPrinterDialog.TabStop = false;
			this.btnPrinterDialog.UseVisualStyleBackColor = true;
			this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
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
			this.prnDocConsulta.DocumentName = "Relatório do Arquivo de Remessa";
			this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
			this.prnDocConsulta.QueryPageSettings += new System.Drawing.Printing.QueryPageSettingsEventHandler(this.prnDocConsulta_QueryPageSettings);
			this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
			// 
			// prnDialogConsulta
			// 
			this.prnDialogConsulta.Document = this.prnDocConsulta;
			this.prnDialogConsulta.UseEXDialog = true;
			// 
			// FBoletoArqRemessaRelatorio
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoArqRemessaRelatorio";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FBoletoArqRemessaRelatorio_Load);
			this.Shown += new System.EventHandler(this.FBoletoArqRemessaRelatorio_Shown);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoArqRemessaRelatorio_FormClosing);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxBoletos.ResumeLayout(false);
			this.gboxBoletos.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).EndInit();
			this.gboxRelatorio.ResumeLayout(false);
			this.gboxOpcaoSaida.ResumeLayout(false);
			this.gboxOpcaoSaida.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.TextBox txtArqRemessa;
		private System.Windows.Forms.Label lblTitArqRemessa;
		private System.Windows.Forms.Button btnSelecionaArqRemessa;
		private System.Windows.Forms.GroupBox gboxBoletos;
		private System.Windows.Forms.DataGridView grdBoletos;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.GroupBox gboxRelatorio;
		private System.Windows.Forms.Label lblTotalRegistros;
		private System.Windows.Forms.Label lblTitTotalRegistros;
		private System.Windows.Forms.Button btnListagemArqRemessa;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.GroupBox gboxOpcaoSaida;
		private System.Windows.Forms.RadioButton rbSaidaVisualizacao;
		private System.Windows.Forms.RadioButton rbSaidaImpressora;
		private System.Windows.Forms.Label lblTitTotalizacaoValor;
		private System.Windows.Forms.Label lblTotalizacaoValor;
		private System.Windows.Forms.DataGridViewTextBoxColumn colSacado;
		private System.Windows.Forms.DataGridViewTextBoxColumn colEndereco;
		private System.Windows.Forms.DataGridViewTextBoxColumn colLoja;
		private System.Windows.Forms.DataGridViewTextBoxColumn colPedido;
		private System.Windows.Forms.DataGridViewTextBoxColumn colNumeroDocumento;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDataVencto;
		private System.Windows.Forms.DataGridViewTextBoxColumn colValorTitulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn colIdBoletoItem;

	}
}
