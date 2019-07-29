namespace Financeiro
{
	partial class FBoletoOcorrencias
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoOcorrencias));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.pnParametros = new System.Windows.Forms.Panel();
			this.ckb_somente_divergencia_valor = new System.Windows.Forms.CheckBox();
			this.lblTitCedente = new System.Windows.Forms.Label();
			this.cbBoletoCedente = new System.Windows.Forms.ComboBox();
			this.cbOcorrencia = new System.Windows.Forms.ComboBox();
			this.lblTitCtrlPagtoStatus = new System.Windows.Forms.Label();
			this.txtNumDocumento = new System.Windows.Forms.TextBox();
			this.lblTitNumDocumento = new System.Windows.Forms.Label();
			this.txtNossoNumero = new System.Windows.Forms.TextBox();
			this.lblTitNossoNumero = new System.Windows.Forms.Label();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.lblValor = new System.Windows.Forms.Label();
			this.txtDataFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodo = new System.Windows.Forms.Label();
			this.txtDataInicial = new System.Windows.Forms.TextBox();
			this.lblDataCompetenciaAte = new System.Windows.Forms.Label();
			this.pnResultado = new System.Windows.Forms.Panel();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.pnTotalizacao = new System.Windows.Forms.Panel();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.btnOcorrenciaTratar = new System.Windows.Forms.Button();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.btnMarcarComoJaTratada = new System.Windows.Forms.Button();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.btnPrintPreview = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
			this.dt_cadastro = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.numero_documento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_vencto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.vl_titulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.ocorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.obs = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto_ocorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto_item = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto_cedente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.identificacao_ocorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.registro_arq_retorno = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.pnParametros.SuspendLayout();
			this.pnResultado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
			this.pnTotalizacao.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnPrinterDialog);
			this.pnBotoes.Controls.Add(this.btnPrintPreview);
			this.pnBotoes.Controls.Add(this.btnImprimir);
			this.pnBotoes.Controls.Add(this.btnMarcarComoJaTratada);
			this.pnBotoes.Controls.Add(this.btnLimpar);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.Add(this.btnOcorrenciaTratar);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnOcorrenciaTratar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnLimpar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnMarcarComoJaTratada, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnImprimir, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrintPreview, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
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
			this.btnFechar.TabIndex = 8;
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
			this.lblTitulo.Text = "Boleto: Ocorrências";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pnParametros
			// 
			this.pnParametros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnParametros.Controls.Add(this.ckb_somente_divergencia_valor);
			this.pnParametros.Controls.Add(this.lblTitCedente);
			this.pnParametros.Controls.Add(this.cbBoletoCedente);
			this.pnParametros.Controls.Add(this.cbOcorrencia);
			this.pnParametros.Controls.Add(this.lblTitCtrlPagtoStatus);
			this.pnParametros.Controls.Add(this.txtNumDocumento);
			this.pnParametros.Controls.Add(this.lblTitNumDocumento);
			this.pnParametros.Controls.Add(this.txtNossoNumero);
			this.pnParametros.Controls.Add(this.lblTitNossoNumero);
			this.pnParametros.Controls.Add(this.txtValor);
			this.pnParametros.Controls.Add(this.lblValor);
			this.pnParametros.Controls.Add(this.txtDataFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodo);
			this.pnParametros.Controls.Add(this.txtDataInicial);
			this.pnParametros.Controls.Add(this.lblDataCompetenciaAte);
			this.pnParametros.Dock = System.Windows.Forms.DockStyle.Top;
			this.pnParametros.Location = new System.Drawing.Point(0, 40);
			this.pnParametros.Name = "pnParametros";
			this.pnParametros.Size = new System.Drawing.Size(1014, 97);
			this.pnParametros.TabIndex = 2;
			// 
			// ckb_somente_divergencia_valor
			// 
			this.ckb_somente_divergencia_valor.AutoSize = true;
			this.ckb_somente_divergencia_valor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ckb_somente_divergencia_valor.Location = new System.Drawing.Point(334, 10);
			this.ckb_somente_divergencia_valor.Name = "ckb_somente_divergencia_valor";
			this.ckb_somente_divergencia_valor.Size = new System.Drawing.Size(222, 17);
			this.ckb_somente_divergencia_valor.TabIndex = 2;
			this.ckb_somente_divergencia_valor.Text = "Somente com divergência de valor";
			this.ckb_somente_divergencia_valor.UseVisualStyleBackColor = true;
			// 
			// lblTitCedente
			// 
			this.lblTitCedente.AutoSize = true;
			this.lblTitCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCedente.Location = new System.Drawing.Point(44, 40);
			this.lblTitCedente.Name = "lblTitCedente";
			this.lblTitCedente.Size = new System.Drawing.Size(54, 13);
			this.lblTitCedente.TabIndex = 53;
			this.lblTitCedente.Text = "Cedente";
			// 
			// cbBoletoCedente
			// 
			this.cbBoletoCedente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbBoletoCedente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbBoletoCedente.BackColor = System.Drawing.SystemColors.Window;
			this.cbBoletoCedente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbBoletoCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbBoletoCedente.FormattingEnabled = true;
			this.cbBoletoCedente.Location = new System.Drawing.Point(104, 37);
			this.cbBoletoCedente.MaxDropDownItems = 12;
			this.cbBoletoCedente.Name = "cbBoletoCedente";
			this.cbBoletoCedente.Size = new System.Drawing.Size(452, 21);
			this.cbBoletoCedente.TabIndex = 4;
			this.cbBoletoCedente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbBoletoCedente_KeyDown);
			// 
			// cbOcorrencia
			// 
			this.cbOcorrencia.BackColor = System.Drawing.SystemColors.Window;
			this.cbOcorrencia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbOcorrencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbOcorrencia.FormattingEnabled = true;
			this.cbOcorrencia.Location = new System.Drawing.Point(104, 66);
			this.cbOcorrencia.Name = "cbOcorrencia";
			this.cbOcorrencia.Size = new System.Drawing.Size(452, 21);
			this.cbOcorrencia.TabIndex = 6;
			this.cbOcorrencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbOcorrencia_KeyDown);
			// 
			// lblTitCtrlPagtoStatus
			// 
			this.lblTitCtrlPagtoStatus.AutoSize = true;
			this.lblTitCtrlPagtoStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCtrlPagtoStatus.Location = new System.Drawing.Point(29, 69);
			this.lblTitCtrlPagtoStatus.Name = "lblTitCtrlPagtoStatus";
			this.lblTitCtrlPagtoStatus.Size = new System.Drawing.Size(69, 13);
			this.lblTitCtrlPagtoStatus.TabIndex = 47;
			this.lblTitCtrlPagtoStatus.Text = "Ocorrência";
			// 
			// txtNumDocumento
			// 
			this.txtNumDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumDocumento.Location = new System.Drawing.Point(876, 35);
			this.txtNumDocumento.MaxLength = 10;
			this.txtNumDocumento.Name = "txtNumDocumento";
			this.txtNumDocumento.Size = new System.Drawing.Size(129, 23);
			this.txtNumDocumento.TabIndex = 5;
			this.txtNumDocumento.Text = "1234567/10";
			this.txtNumDocumento.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumDocumento_KeyDown);
			this.txtNumDocumento.Leave += new System.EventHandler(this.txtNumDocumento_Leave);
			this.txtNumDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumDocumento_KeyPress);
			this.txtNumDocumento.Enter += new System.EventHandler(this.txtNumDocumento_Enter);
			// 
			// lblTitNumDocumento
			// 
			this.lblTitNumDocumento.AutoSize = true;
			this.lblTitNumDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNumDocumento.Location = new System.Drawing.Point(781, 40);
			this.lblTitNumDocumento.Name = "lblTitNumDocumento";
			this.lblTitNumDocumento.Size = new System.Drawing.Size(89, 13);
			this.lblTitNumDocumento.TabIndex = 41;
			this.lblTitNumDocumento.Text = "Nº Documento";
			// 
			// txtNossoNumero
			// 
			this.txtNossoNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNossoNumero.Location = new System.Drawing.Point(876, 64);
			this.txtNossoNumero.MaxLength = 11;
			this.txtNossoNumero.Name = "txtNossoNumero";
			this.txtNossoNumero.Size = new System.Drawing.Size(129, 23);
			this.txtNossoNumero.TabIndex = 7;
			this.txtNossoNumero.Text = "12345678901-8";
			this.txtNossoNumero.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNossoNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNossoNumero_KeyDown);
			this.txtNossoNumero.Leave += new System.EventHandler(this.txtNossoNumero_Leave);
			this.txtNossoNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNossoNumero_KeyPress);
			this.txtNossoNumero.Enter += new System.EventHandler(this.txtNossoNumero_Enter);
			// 
			// lblTitNossoNumero
			// 
			this.lblTitNossoNumero.AutoSize = true;
			this.lblTitNossoNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNossoNumero.Location = new System.Drawing.Point(781, 69);
			this.lblTitNossoNumero.Name = "lblTitNossoNumero";
			this.lblTitNossoNumero.Size = new System.Drawing.Size(89, 13);
			this.lblTitNossoNumero.TabIndex = 31;
			this.lblTitNossoNumero.Text = "Nosso Número";
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(876, 6);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(129, 23);
			this.txtValor.TabIndex = 3;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValor.Location = new System.Drawing.Point(806, 11);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(64, 13);
			this.lblValor.TabIndex = 28;
			this.lblValor.Text = "Valor (R$)";
			// 
			// txtDataFinal
			// 
			this.txtDataFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataFinal.Location = new System.Drawing.Point(221, 6);
			this.txtDataFinal.MaxLength = 10;
			this.txtDataFinal.Name = "txtDataFinal";
			this.txtDataFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataFinal.TabIndex = 1;
			this.txtDataFinal.Text = "01/01/2000";
			this.txtDataFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataFinal_KeyDown);
			this.txtDataFinal.Leave += new System.EventHandler(this.txtDataFinal_Leave);
			this.txtDataFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataFinal_KeyPress);
			this.txtDataFinal.Enter += new System.EventHandler(this.txtDataFinal_Enter);
			// 
			// lblTitPeriodo
			// 
			this.lblTitPeriodo.AutoSize = true;
			this.lblTitPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodo.Location = new System.Drawing.Point(64, 11);
			this.lblTitPeriodo.Name = "lblTitPeriodo";
			this.lblTitPeriodo.Size = new System.Drawing.Size(34, 13);
			this.lblTitPeriodo.TabIndex = 10;
			this.lblTitPeriodo.Text = "Data";
			// 
			// txtDataInicial
			// 
			this.txtDataInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataInicial.Location = new System.Drawing.Point(104, 6);
			this.txtDataInicial.MaxLength = 10;
			this.txtDataInicial.Name = "txtDataInicial";
			this.txtDataInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataInicial.TabIndex = 0;
			this.txtDataInicial.Text = "01/01/2000";
			this.txtDataInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataInicial_KeyDown);
			this.txtDataInicial.Leave += new System.EventHandler(this.txtDataInicial_Leave);
			this.txtDataInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataInicial_KeyPress);
			this.txtDataInicial.Enter += new System.EventHandler(this.txtDataInicial_Enter);
			// 
			// lblDataCompetenciaAte
			// 
			this.lblDataCompetenciaAte.AutoSize = true;
			this.lblDataCompetenciaAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblDataCompetenciaAte.Location = new System.Drawing.Point(201, 11);
			this.lblDataCompetenciaAte.Name = "lblDataCompetenciaAte";
			this.lblDataCompetenciaAte.Size = new System.Drawing.Size(14, 13);
			this.lblDataCompetenciaAte.TabIndex = 9;
			this.lblDataCompetenciaAte.Text = "a";
			// 
			// pnResultado
			// 
			this.pnResultado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnResultado.Controls.Add(this.gridDados);
			this.pnResultado.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pnResultado.Location = new System.Drawing.Point(0, 137);
			this.pnResultado.Name = "pnResultado";
			this.pnResultado.Size = new System.Drawing.Size(1014, 447);
			this.pnResultado.TabIndex = 4;
			// 
			// gridDados
			// 
			this.gridDados.AllowUserToAddRows = false;
			this.gridDados.AllowUserToDeleteRows = false;
			this.gridDados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.gridDados.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.gridDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.gridDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dt_cadastro,
            this.cliente,
            this.numero_documento,
            this.dt_vencto,
            this.vl_titulo,
            this.ocorrencia,
            this.obs,
            this.id_boleto_ocorrencia,
            this.id_boleto_item,
            this.id_boleto,
            this.id_boleto_cedente,
            this.identificacao_ocorrencia,
            this.registro_arq_retorno});
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.gridDados.DefaultCellStyle = dataGridViewCellStyle9;
			this.gridDados.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridDados.Location = new System.Drawing.Point(0, 0);
			this.gridDados.MultiSelect = false;
			this.gridDados.Name = "gridDados";
			this.gridDados.ReadOnly = true;
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.RowHeadersDefaultCellStyle = dataGridViewCellStyle10;
			this.gridDados.RowHeadersVisible = false;
			this.gridDados.RowHeadersWidth = 15;
			this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridDados.Size = new System.Drawing.Size(1010, 443);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 0;
			this.gridDados.DoubleClick += new System.EventHandler(this.gridDados_DoubleClick);
			this.gridDados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridDados_KeyDown);
			// 
			// pnTotalizacao
			// 
			this.pnTotalizacao.Controls.Add(this.lblTitTotalizacaoRegistros);
			this.pnTotalizacao.Controls.Add(this.lblTotalizacaoRegistros);
			this.pnTotalizacao.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.pnTotalizacao.Location = new System.Drawing.Point(0, 584);
			this.pnTotalizacao.Name = "pnTotalizacao";
			this.pnTotalizacao.Size = new System.Drawing.Size(1014, 21);
			this.pnTotalizacao.TabIndex = 3;
			// 
			// lblTitTotalizacaoRegistros
			// 
			this.lblTitTotalizacaoRegistros.AutoSize = true;
			this.lblTitTotalizacaoRegistros.Location = new System.Drawing.Point(697, 4);
			this.lblTitTotalizacaoRegistros.Name = "lblTitTotalizacaoRegistros";
			this.lblTitTotalizacaoRegistros.Size = new System.Drawing.Size(54, 13);
			this.lblTitTotalizacaoRegistros.TabIndex = 5;
			this.lblTitTotalizacaoRegistros.Text = "Registros:";
			// 
			// lblTotalizacaoRegistros
			// 
			this.lblTotalizacaoRegistros.AutoSize = true;
			this.lblTotalizacaoRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoRegistros.Location = new System.Drawing.Point(757, 4);
			this.lblTotalizacaoRegistros.Name = "lblTotalizacaoRegistros";
			this.lblTotalizacaoRegistros.Size = new System.Drawing.Size(21, 13);
			this.lblTotalizacaoRegistros.TabIndex = 6;
			this.lblTotalizacaoRegistros.Text = "99";
			// 
			// btnOcorrenciaTratar
			// 
			this.btnOcorrenciaTratar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnOcorrenciaTratar.Image = ((System.Drawing.Image)(resources.GetObject("btnOcorrenciaTratar.Image")));
			this.btnOcorrenciaTratar.Location = new System.Drawing.Point(699, 4);
			this.btnOcorrenciaTratar.Name = "btnOcorrenciaTratar";
			this.btnOcorrenciaTratar.Size = new System.Drawing.Size(40, 44);
			this.btnOcorrenciaTratar.TabIndex = 2;
			this.btnOcorrenciaTratar.TabStop = false;
			this.btnOcorrenciaTratar.UseVisualStyleBackColor = true;
			this.btnOcorrenciaTratar.Click += new System.EventHandler(this.btnOcorrenciaTratar_Click);
			// 
			// btnPesquisar
			// 
			this.btnPesquisar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPesquisar.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisar.Image")));
			this.btnPesquisar.Location = new System.Drawing.Point(609, 4);
			this.btnPesquisar.Name = "btnPesquisar";
			this.btnPesquisar.Size = new System.Drawing.Size(40, 44);
			this.btnPesquisar.TabIndex = 0;
			this.btnPesquisar.TabStop = false;
			this.btnPesquisar.UseVisualStyleBackColor = true;
			this.btnPesquisar.Click += new System.EventHandler(this.btnPesquisar_Click);
			// 
			// btnLimpar
			// 
			this.btnLimpar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnLimpar.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpar.Image")));
			this.btnLimpar.Location = new System.Drawing.Point(654, 4);
			this.btnLimpar.Name = "btnLimpar";
			this.btnLimpar.Size = new System.Drawing.Size(40, 44);
			this.btnLimpar.TabIndex = 1;
			this.btnLimpar.TabStop = false;
			this.btnLimpar.UseVisualStyleBackColor = true;
			this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
			// 
			// btnMarcarComoJaTratada
			// 
			this.btnMarcarComoJaTratada.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnMarcarComoJaTratada.Image = ((System.Drawing.Image)(resources.GetObject("btnMarcarComoJaTratada.Image")));
			this.btnMarcarComoJaTratada.Location = new System.Drawing.Point(744, 4);
			this.btnMarcarComoJaTratada.Name = "btnMarcarComoJaTratada";
			this.btnMarcarComoJaTratada.Size = new System.Drawing.Size(40, 44);
			this.btnMarcarComoJaTratada.TabIndex = 3;
			this.btnMarcarComoJaTratada.TabStop = false;
			this.btnMarcarComoJaTratada.UseVisualStyleBackColor = true;
			this.btnMarcarComoJaTratada.Click += new System.EventHandler(this.btnMarcarComoJaTratada_Click);
			// 
			// btnPrinterDialog
			// 
			this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
			this.btnPrinterDialog.Location = new System.Drawing.Point(879, 4);
			this.btnPrinterDialog.Name = "btnPrinterDialog";
			this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
			this.btnPrinterDialog.TabIndex = 6;
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
			this.btnPrintPreview.TabIndex = 5;
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
			this.btnImprimir.TabIndex = 4;
			this.btnImprimir.TabStop = false;
			this.btnImprimir.UseVisualStyleBackColor = true;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// prnDocConsulta
			// 
			this.prnDocConsulta.DocumentName = "Boleto: Ocorrências";
			this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
			this.prnDocConsulta.QueryPageSettings += new System.Drawing.Printing.QueryPageSettingsEventHandler(this.prnDocConsulta_QueryPageSettings);
			this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
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
			// dt_cadastro
			// 
			this.dt_cadastro.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.dt_cadastro.DefaultCellStyle = dataGridViewCellStyle2;
			this.dt_cadastro.HeaderText = "Data";
			this.dt_cadastro.MinimumWidth = 75;
			this.dt_cadastro.Name = "dt_cadastro";
			this.dt_cadastro.ReadOnly = true;
			this.dt_cadastro.Width = 75;
			// 
			// cliente
			// 
			this.cliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.cliente.DefaultCellStyle = dataGridViewCellStyle3;
			this.cliente.HeaderText = "Cliente";
			this.cliente.MinimumWidth = 180;
			this.cliente.Name = "cliente";
			this.cliente.ReadOnly = true;
			// 
			// numero_documento
			// 
			this.numero_documento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.numero_documento.DefaultCellStyle = dataGridViewCellStyle4;
			this.numero_documento.HeaderText = "Nº Doc";
			this.numero_documento.MinimumWidth = 85;
			this.numero_documento.Name = "numero_documento";
			this.numero_documento.ReadOnly = true;
			this.numero_documento.Width = 85;
			// 
			// dt_vencto
			// 
			this.dt_vencto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.dt_vencto.DefaultCellStyle = dataGridViewCellStyle5;
			this.dt_vencto.HeaderText = "Vencto";
			this.dt_vencto.MinimumWidth = 75;
			this.dt_vencto.Name = "dt_vencto";
			this.dt_vencto.ReadOnly = true;
			this.dt_vencto.Width = 75;
			// 
			// vl_titulo
			// 
			this.vl_titulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			this.vl_titulo.DefaultCellStyle = dataGridViewCellStyle6;
			this.vl_titulo.HeaderText = "Valor";
			this.vl_titulo.MinimumWidth = 90;
			this.vl_titulo.Name = "vl_titulo";
			this.vl_titulo.ReadOnly = true;
			this.vl_titulo.Width = 90;
			// 
			// ocorrencia
			// 
			this.ocorrencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.ocorrencia.DefaultCellStyle = dataGridViewCellStyle7;
			this.ocorrencia.HeaderText = "Ocorrência";
			this.ocorrencia.MinimumWidth = 220;
			this.ocorrencia.Name = "ocorrencia";
			this.ocorrencia.ReadOnly = true;
			this.ocorrencia.Width = 220;
			// 
			// obs
			// 
			this.obs.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.obs.DefaultCellStyle = dataGridViewCellStyle8;
			this.obs.HeaderText = "Obs";
			this.obs.MinimumWidth = 240;
			this.obs.Name = "obs";
			this.obs.ReadOnly = true;
			this.obs.Width = 240;
			// 
			// id_boleto_ocorrencia
			// 
			this.id_boleto_ocorrencia.HeaderText = "id_boleto_ocorrencia";
			this.id_boleto_ocorrencia.Name = "id_boleto_ocorrencia";
			this.id_boleto_ocorrencia.ReadOnly = true;
			this.id_boleto_ocorrencia.Visible = false;
			// 
			// id_boleto_item
			// 
			this.id_boleto_item.HeaderText = "id_boleto_item";
			this.id_boleto_item.Name = "id_boleto_item";
			this.id_boleto_item.ReadOnly = true;
			this.id_boleto_item.Visible = false;
			// 
			// id_boleto
			// 
			this.id_boleto.HeaderText = "id_boleto";
			this.id_boleto.Name = "id_boleto";
			this.id_boleto.ReadOnly = true;
			this.id_boleto.Visible = false;
			// 
			// id_boleto_cedente
			// 
			this.id_boleto_cedente.HeaderText = "id_boleto_cedente";
			this.id_boleto_cedente.Name = "id_boleto_cedente";
			this.id_boleto_cedente.ReadOnly = true;
			this.id_boleto_cedente.Visible = false;
			// 
			// identificacao_ocorrencia
			// 
			this.identificacao_ocorrencia.HeaderText = "identificacao_ocorrencia";
			this.identificacao_ocorrencia.Name = "identificacao_ocorrencia";
			this.identificacao_ocorrencia.ReadOnly = true;
			this.identificacao_ocorrencia.Visible = false;
			// 
			// registro_arq_retorno
			// 
			this.registro_arq_retorno.HeaderText = "registro_arq_retorno";
			this.registro_arq_retorno.Name = "registro_arq_retorno";
			this.registro_arq_retorno.ReadOnly = true;
			this.registro_arq_retorno.Visible = false;
			// 
			// FBoletoOcorrencias
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoOcorrencias";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FBoletoOcorrencias_Load);
			this.Shown += new System.EventHandler(this.FBoletoOcorrencias_Shown);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoOcorrencias_FormClosing);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FBoletoOcorrencias_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnParametros.ResumeLayout(false);
			this.pnParametros.PerformLayout();
			this.pnResultado.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).EndInit();
			this.pnTotalizacao.ResumeLayout(false);
			this.pnTotalizacao.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Panel pnParametros;
		private System.Windows.Forms.TextBox txtDataInicial;
		private System.Windows.Forms.Label lblDataCompetenciaAte;
		private System.Windows.Forms.Label lblTitPeriodo;
		private System.Windows.Forms.TextBox txtDataFinal;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtNossoNumero;
		private System.Windows.Forms.Label lblTitNossoNumero;
		private System.Windows.Forms.Panel pnResultado;
		private System.Windows.Forms.DataGridView gridDados;
		private System.Windows.Forms.Panel pnTotalizacao;
		private System.Windows.Forms.Label lblTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTitTotalizacaoRegistros;
		private System.Windows.Forms.Button btnOcorrenciaTratar;
		private System.Windows.Forms.Button btnPesquisar;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.TextBox txtNumDocumento;
		private System.Windows.Forms.Label lblTitNumDocumento;
		private System.Windows.Forms.ComboBox cbOcorrencia;
		private System.Windows.Forms.Label lblTitCtrlPagtoStatus;
		private System.Windows.Forms.Button btnMarcarComoJaTratada;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnPrintPreview;
		private System.Windows.Forms.Button btnImprimir;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Windows.Forms.Label lblTitCedente;
		private System.Windows.Forms.ComboBox cbBoletoCedente;
		private System.Windows.Forms.CheckBox ckb_somente_divergencia_valor;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_cadastro;
		private System.Windows.Forms.DataGridViewTextBoxColumn cliente;
		private System.Windows.Forms.DataGridViewTextBoxColumn numero_documento;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_vencto;
		private System.Windows.Forms.DataGridViewTextBoxColumn vl_titulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn ocorrencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn obs;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto_ocorrencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto_item;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto_cedente;
		private System.Windows.Forms.DataGridViewTextBoxColumn identificacao_ocorrencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn registro_arq_retorno;
	}
}
