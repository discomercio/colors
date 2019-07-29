namespace PrnDANFE
{
	partial class FDANFEImprime
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FDANFEImprime));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnPrinterDialog = new System.Windows.Forms.Button();
            this.gboxPesquisaPorDataEmissao = new System.Windows.Forms.GroupBox();
            this.btnMarcarTodos = new System.Windows.Forms.Button();
            this.btnLimparFiltro = new System.Windows.Forms.Button();
            this.cbTransportadora = new System.Windows.Forms.ComboBox();
            this.lblTotalRegistros = new System.Windows.Forms.Label();
            this.lblTitTotalRegistros = new System.Windows.Forms.Label();
            this.txtNFe = new System.Windows.Forms.TextBox();
            this.lblTitNsu = new System.Windows.Forms.Label();
            this.lblTitPedido = new System.Windows.Forms.Label();
            this.btnPesquisar = new System.Windows.Forms.Button();
            this.grdPesquisa = new System.Windows.Forms.DataGridView();
            this.colGrdPesqCheckBox = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colGrdPesqPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqCidade = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqUF = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqDataEntrega = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqTransportadora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqNfe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdPesqSerie = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblTitNumeroNFe = new System.Windows.Forms.Label();
            this.dtpDataEntrega = new System.Windows.Forms.DateTimePicker();
            this.lblTitDataEntrega = new System.Windows.Forms.Label();
            this.txtPedido = new System.Windows.Forms.TextBox();
            this.gboxOpcoes = new System.Windows.Forms.GroupBox();
            this.btnPastaPDFInd = new System.Windows.Forms.Button();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.btnPastaPDFAgrup = new System.Windows.Forms.Button();
            this.printDialog = new System.Windows.Forms.PrintDialog();
            this.lblEmit = new System.Windows.Forms.Label();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxPesquisaPorDataEmissao.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdPesquisa)).BeginInit();
            this.gboxOpcoes.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.lblEmit);
            this.pnBotoes.Controls.Add(this.btnPrinterDialog);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
            this.pnBotoes.Controls.SetChildIndex(this.lblEmit, 0);
            // 
            // btnSobre
            // 
            this.btnSobre.TabIndex = 1;
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // btnFechar
            // 
            this.btnFechar.TabIndex = 2;
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.gboxOpcoes);
            this.pnCampos.Controls.Add(this.gboxPesquisaPorDataEmissao);
            this.pnCampos.Size = new System.Drawing.Size(1008, 599);
            // 
            // btnPrinterDialog
            // 
            this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
            this.btnPrinterDialog.Location = new System.Drawing.Point(869, 4);
            this.btnPrinterDialog.Name = "btnPrinterDialog";
            this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
            this.btnPrinterDialog.TabIndex = 0;
            this.btnPrinterDialog.TabStop = false;
            this.btnPrinterDialog.UseVisualStyleBackColor = true;
            this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
            // 
            // gboxPesquisaPorDataEmissao
            // 
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.btnMarcarTodos);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.btnLimparFiltro);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.cbTransportadora);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTotalRegistros);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTitTotalRegistros);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.txtNFe);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTitNsu);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTitPedido);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.btnPesquisar);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.grdPesquisa);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTitNumeroNFe);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.dtpDataEntrega);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.lblTitDataEntrega);
            this.gboxPesquisaPorDataEmissao.Controls.Add(this.txtPedido);
            this.gboxPesquisaPorDataEmissao.Location = new System.Drawing.Point(5, 4);
            this.gboxPesquisaPorDataEmissao.Name = "gboxPesquisaPorDataEmissao";
            this.gboxPesquisaPorDataEmissao.Size = new System.Drawing.Size(994, 522);
            this.gboxPesquisaPorDataEmissao.TabIndex = 7;
            this.gboxPesquisaPorDataEmissao.TabStop = false;
            this.gboxPesquisaPorDataEmissao.Text = "Pesquisa de DANFEs";
            // 
            // btnMarcarTodos
            // 
            this.btnMarcarTodos.Location = new System.Drawing.Point(343, 486);
            this.btnMarcarTodos.Name = "btnMarcarTodos";
            this.btnMarcarTodos.Size = new System.Drawing.Size(288, 26);
            this.btnMarcarTodos.TabIndex = 13;
            this.btnMarcarTodos.Text = "Marcar/Desmarcar TODAS as DANFEs listadas";
            this.btnMarcarTodos.UseVisualStyleBackColor = true;
            this.btnMarcarTodos.Click += new System.EventHandler(this.btnMarcarTodos_Click);
            // 
            // btnLimparFiltro
            // 
            this.btnLimparFiltro.Image = ((System.Drawing.Image)(resources.GetObject("btnLimparFiltro.Image")));
            this.btnLimparFiltro.Location = new System.Drawing.Point(952, 14);
            this.btnLimparFiltro.Name = "btnLimparFiltro";
            this.btnLimparFiltro.Size = new System.Drawing.Size(32, 26);
            this.btnLimparFiltro.TabIndex = 6;
            this.btnLimparFiltro.UseVisualStyleBackColor = true;
            this.btnLimparFiltro.Click += new System.EventHandler(this.btnLimparFiltro_Click);
            // 
            // cbTransportadora
            // 
            this.cbTransportadora.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTransportadora.FormattingEnabled = true;
            this.cbTransportadora.Location = new System.Drawing.Point(314, 18);
            this.cbTransportadora.Name = "cbTransportadora";
            this.cbTransportadora.Size = new System.Drawing.Size(226, 21);
            this.cbTransportadora.TabIndex = 1;
            // 
            // lblTotalRegistros
            // 
            this.lblTotalRegistros.AutoSize = true;
            this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalRegistros.Location = new System.Drawing.Point(94, 506);
            this.lblTotalRegistros.Name = "lblTotalRegistros";
            this.lblTotalRegistros.Size = new System.Drawing.Size(42, 13);
            this.lblTotalRegistros.TabIndex = 11;
            this.lblTotalRegistros.Text = "99999";
            // 
            // lblTitTotalRegistros
            // 
            this.lblTitTotalRegistros.AutoSize = true;
            this.lblTitTotalRegistros.Location = new System.Drawing.Point(6, 506);
            this.lblTitTotalRegistros.Name = "lblTitTotalRegistros";
            this.lblTitTotalRegistros.Size = new System.Drawing.Size(88, 13);
            this.lblTitTotalRegistros.TabIndex = 10;
            this.lblTitTotalRegistros.Text = "Total de registros";
            // 
            // txtNFe
            // 
            this.txtNFe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNFe.Location = new System.Drawing.Point(595, 18);
            this.txtNFe.MaxLength = 9;
            this.txtNFe.Name = "txtNFe";
            this.txtNFe.Size = new System.Drawing.Size(90, 20);
            this.txtNFe.TabIndex = 2;
            this.txtNFe.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtNFe.Enter += new System.EventHandler(this.txtNFe_Enter);
            this.txtNFe.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNFe_KeyPress);
            this.txtNFe.Leave += new System.EventHandler(this.txtNFe_Leave);
            // 
            // lblTitNsu
            // 
            this.lblTitNsu.AutoSize = true;
            this.lblTitNsu.Location = new System.Drawing.Point(229, 21);
            this.lblTitNsu.Name = "lblTitNsu";
            this.lblTitNsu.Size = new System.Drawing.Size(79, 13);
            this.lblTitNsu.TabIndex = 9;
            this.lblTitNsu.Text = "Transportadora";
            // 
            // lblTitPedido
            // 
            this.lblTitPedido.AutoSize = true;
            this.lblTitPedido.Location = new System.Drawing.Point(703, 21);
            this.lblTitPedido.Name = "lblTitPedido";
            this.lblTitPedido.Size = new System.Drawing.Size(40, 13);
            this.lblTitPedido.TabIndex = 12;
            this.lblTitPedido.Text = "Pedido";
            // 
            // btnPesquisar
            // 
            this.btnPesquisar.Location = new System.Drawing.Point(859, 14);
            this.btnPesquisar.Name = "btnPesquisar";
            this.btnPesquisar.Size = new System.Drawing.Size(87, 26);
            this.btnPesquisar.TabIndex = 4;
            this.btnPesquisar.Text = "Pesquisar";
            this.btnPesquisar.UseVisualStyleBackColor = true;
            this.btnPesquisar.Click += new System.EventHandler(this.btnPesquisar_Click);
            // 
            // grdPesquisa
            // 
            this.grdPesquisa.AllowUserToAddRows = false;
            this.grdPesquisa.AllowUserToDeleteRows = false;
            this.grdPesquisa.AllowUserToResizeColumns = false;
            this.grdPesquisa.AllowUserToResizeRows = false;
            this.grdPesquisa.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.grdPesquisa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdPesquisa.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.grdPesquisa.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdPesquisa.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGrdPesqCheckBox,
            this.colGrdPesqPedido,
            this.colGrdPesqCidade,
            this.colGrdPesqUF,
            this.colGrdPesqDataEntrega,
            this.colGrdPesqTransportadora,
            this.colGrdPesqNfe,
            this.colGrdPesqSerie});
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdPesquisa.DefaultCellStyle = dataGridViewCellStyle7;
            this.grdPesquisa.Location = new System.Drawing.Point(11, 54);
            this.grdPesquisa.MultiSelect = false;
            this.grdPesquisa.Name = "grdPesquisa";
            this.grdPesquisa.RowHeadersVisible = false;
            this.grdPesquisa.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.grdPesquisa.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grdPesquisa.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdPesquisa.ShowEditingIcon = false;
            this.grdPesquisa.Size = new System.Drawing.Size(973, 423);
            this.grdPesquisa.StandardTab = true;
            this.grdPesquisa.TabIndex = 4;
            this.grdPesquisa.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdPesquisa_CellContentClick);
            // 
            // colGrdPesqCheckBox
            // 
            this.colGrdPesqCheckBox.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colGrdPesqCheckBox.HeaderText = "";
            this.colGrdPesqCheckBox.MinimumWidth = 25;
            this.colGrdPesqCheckBox.Name = "colGrdPesqCheckBox";
            this.colGrdPesqCheckBox.Width = 25;
            // 
            // colGrdPesqPedido
            // 
            this.colGrdPesqPedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colGrdPesqPedido.DefaultCellStyle = dataGridViewCellStyle2;
            this.colGrdPesqPedido.HeaderText = "Pedido";
            this.colGrdPesqPedido.MinimumWidth = 80;
            this.colGrdPesqPedido.Name = "colGrdPesqPedido";
            this.colGrdPesqPedido.Width = 80;
            // 
            // colGrdPesqCidade
            // 
            this.colGrdPesqCidade.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colGrdPesqCidade.DefaultCellStyle = dataGridViewCellStyle3;
            this.colGrdPesqCidade.HeaderText = "Cidade";
            this.colGrdPesqCidade.MinimumWidth = 200;
            this.colGrdPesqCidade.Name = "colGrdPesqCidade";
            this.colGrdPesqCidade.Width = 200;
            // 
            // colGrdPesqUF
            // 
            this.colGrdPesqUF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colGrdPesqUF.DefaultCellStyle = dataGridViewCellStyle4;
            this.colGrdPesqUF.HeaderText = "UF";
            this.colGrdPesqUF.MinimumWidth = 50;
            this.colGrdPesqUF.Name = "colGrdPesqUF";
            this.colGrdPesqUF.Width = 50;
            // 
            // colGrdPesqDataEntrega
            // 
            this.colGrdPesqDataEntrega.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colGrdPesqDataEntrega.DefaultCellStyle = dataGridViewCellStyle5;
            this.colGrdPesqDataEntrega.HeaderText = "Data Entrega";
            this.colGrdPesqDataEntrega.MinimumWidth = 80;
            this.colGrdPesqDataEntrega.Name = "colGrdPesqDataEntrega";
            this.colGrdPesqDataEntrega.Width = 120;
            // 
            // colGrdPesqTransportadora
            // 
            this.colGrdPesqTransportadora.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colGrdPesqTransportadora.DefaultCellStyle = dataGridViewCellStyle6;
            this.colGrdPesqTransportadora.HeaderText = "Transportadora";
            this.colGrdPesqTransportadora.MinimumWidth = 120;
            this.colGrdPesqTransportadora.Name = "colGrdPesqTransportadora";
            // 
            // colGrdPesqNfe
            // 
            this.colGrdPesqNfe.HeaderText = "Nº NFE";
            this.colGrdPesqNfe.Name = "colGrdPesqNfe";
            this.colGrdPesqNfe.Width = 80;
            // 
            // colGrdPesqSerie
            // 
            this.colGrdPesqSerie.HeaderText = "Série";
            this.colGrdPesqSerie.Name = "colGrdPesqSerie";
            // 
            // lblTitNumeroNFe
            // 
            this.lblTitNumeroNFe.AutoSize = true;
            this.lblTitNumeroNFe.Location = new System.Drawing.Point(546, 21);
            this.lblTitNumeroNFe.Name = "lblTitNumeroNFe";
            this.lblTitNumeroNFe.Size = new System.Drawing.Size(43, 13);
            this.lblTitNumeroNFe.TabIndex = 9;
            this.lblTitNumeroNFe.Text = "Nº NFE";
            // 
            // dtpDataEntrega
            // 
            this.dtpDataEntrega.Checked = false;
            this.dtpDataEntrega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDataEntrega.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataEntrega.Location = new System.Drawing.Point(97, 18);
            this.dtpDataEntrega.MaxDate = new System.DateTime(2099, 12, 31, 0, 0, 0, 0);
            this.dtpDataEntrega.MinDate = new System.DateTime(2013, 1, 1, 0, 0, 0, 0);
            this.dtpDataEntrega.Name = "dtpDataEntrega";
            this.dtpDataEntrega.ShowCheckBox = true;
            this.dtpDataEntrega.Size = new System.Drawing.Size(126, 20);
            this.dtpDataEntrega.TabIndex = 0;
            // 
            // lblTitDataEntrega
            // 
            this.lblTitDataEntrega.AutoSize = true;
            this.lblTitDataEntrega.Location = new System.Drawing.Point(8, 21);
            this.lblTitDataEntrega.Name = "lblTitDataEntrega";
            this.lblTitDataEntrega.Size = new System.Drawing.Size(85, 13);
            this.lblTitDataEntrega.TabIndex = 2;
            this.lblTitDataEntrega.Text = "Data de Entrega";
            // 
            // txtPedido
            // 
            this.txtPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPedido.Location = new System.Drawing.Point(749, 18);
            this.txtPedido.MaxLength = 9;
            this.txtPedido.Name = "txtPedido";
            this.txtPedido.Size = new System.Drawing.Size(92, 20);
            this.txtPedido.TabIndex = 3;
            this.txtPedido.Text = "000000K-A";
            this.txtPedido.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtPedido.Enter += new System.EventHandler(this.txtPedido_Enter);
            this.txtPedido.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPedido_KeyDown);
            this.txtPedido.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPedido_KeyPress);
            this.txtPedido.Leave += new System.EventHandler(this.txtPedido_Leave);
            // 
            // gboxOpcoes
            // 
            this.gboxOpcoes.Controls.Add(this.btnPastaPDFInd);
            this.gboxOpcoes.Controls.Add(this.btnImprimir);
            this.gboxOpcoes.Controls.Add(this.btnPastaPDFAgrup);
            this.gboxOpcoes.Location = new System.Drawing.Point(5, 541);
            this.gboxOpcoes.Name = "gboxOpcoes";
            this.gboxOpcoes.Size = new System.Drawing.Size(994, 49);
            this.gboxOpcoes.TabIndex = 9;
            this.gboxOpcoes.TabStop = false;
            this.gboxOpcoes.Text = "Opções";
            // 
            // btnPastaPDFInd
            // 
            this.btnPastaPDFInd.Location = new System.Drawing.Point(343, 17);
            this.btnPastaPDFInd.Name = "btnPastaPDFInd";
            this.btnPastaPDFInd.Size = new System.Drawing.Size(288, 26);
            this.btnPastaPDFInd.TabIndex = 2;
            this.btnPastaPDFInd.Text = "Abrir a pasta de PDFs individuais";
            this.btnPastaPDFInd.UseVisualStyleBackColor = true;
            this.btnPastaPDFInd.Click += new System.EventHandler(this.btnPastaPDFInd_Click);
            // 
            // btnImprimir
            // 
            this.btnImprimir.Location = new System.Drawing.Point(680, 17);
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.Size = new System.Drawing.Size(288, 26);
            this.btnImprimir.TabIndex = 1;
            this.btnImprimir.Text = "Imprimir DANFEs selecionados";
            this.btnImprimir.UseVisualStyleBackColor = true;
            this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
            // 
            // btnPastaPDFAgrup
            // 
            this.btnPastaPDFAgrup.Location = new System.Drawing.Point(20, 17);
            this.btnPastaPDFAgrup.Name = "btnPastaPDFAgrup";
            this.btnPastaPDFAgrup.Size = new System.Drawing.Size(288, 26);
            this.btnPastaPDFAgrup.TabIndex = 0;
            this.btnPastaPDFAgrup.Text = "Abrir a pasta de PDFs agrupados";
            this.btnPastaPDFAgrup.UseVisualStyleBackColor = true;
            this.btnPastaPDFAgrup.Click += new System.EventHandler(this.btnPastaPDFAgrup_Click);
            // 
            // printDialog
            // 
            this.printDialog.UseEXDialog = true;
            // 
            // lblEmit
            // 
            this.lblEmit.AutoSize = true;
            this.lblEmit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEmit.Location = new System.Drawing.Point(267, 18);
            this.lblEmit.Name = "lblEmit";
            this.lblEmit.Size = new System.Drawing.Size(55, 16);
            this.lblEmit.TabIndex = 7;
            this.lblEmit.Text = "lblEmit";
            this.lblEmit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FDANFEImprime
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1008, 696);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FDANFEImprime";
            this.Text = "PrnDANFE  -  1.03 - 15.ABR.2014";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FDANFEImprime_FormClosing);
            this.Load += new System.EventHandler(this.FDANFEImprime_Load);
            this.Shown += new System.EventHandler(this.FDANFEImprime_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnBotoes.PerformLayout();
            this.pnCampos.ResumeLayout(false);
            this.gboxPesquisaPorDataEmissao.ResumeLayout(false);
            this.gboxPesquisaPorDataEmissao.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdPesquisa)).EndInit();
            this.gboxOpcoes.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

        private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.GroupBox gboxPesquisaPorDataEmissao;
		private System.Windows.Forms.Label lblTitDataEntrega;
        private System.Windows.Forms.DateTimePicker dtpDataEntrega;
		private System.Windows.Forms.DataGridView grdPesquisa;
        private System.Windows.Forms.Button btnPesquisar;
        private System.Windows.Forms.Label lblTitNsu;
		private System.Windows.Forms.GroupBox gboxOpcoes;
		private System.Windows.Forms.Button btnImprimir;
        private System.Windows.Forms.Button btnPastaPDFAgrup;
        private System.Windows.Forms.Button btnLimparFiltro;
        private System.Windows.Forms.PrintDialog printDialog;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colGrdPesqCheckBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqPedido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqCidade;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqUF;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqDataEntrega;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqTransportadora;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqNfe;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdPesqSerie;
        private System.Windows.Forms.Label lblTotalRegistros;
        private System.Windows.Forms.Label lblTitTotalRegistros;
        private System.Windows.Forms.ComboBox cbTransportadora;
        private System.Windows.Forms.TextBox txtNFe;
        private System.Windows.Forms.Label lblTitPedido;
        private System.Windows.Forms.Label lblTitNumeroNFe;
        private System.Windows.Forms.TextBox txtPedido;
        private System.Windows.Forms.Button btnPastaPDFInd;
        private System.Windows.Forms.Button btnMarcarTodos;
        private System.Windows.Forms.Label lblEmit;
	}
}
