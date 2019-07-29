namespace Financeiro
{
	partial class FBoletoConsulta
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoConsulta));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.pnParametros = new System.Windows.Forms.Panel();
			this.lblTitCedente = new System.Windows.Forms.Label();
			this.cbBoletoCedente = new System.Windows.Forms.ComboBox();
			this.txtDataCargaRetornoFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodoCargaRetorno = new System.Windows.Forms.Label();
			this.txtDataCargaRetornoInicial = new System.Windows.Forms.TextBox();
			this.lblDataCargaRetornoAte = new System.Windows.Forms.Label();
			this.cbOcorrencia = new System.Windows.Forms.ComboBox();
			this.lblTitCtrlPagtoStatus = new System.Windows.Forms.Label();
			this.txtNumPedido = new System.Windows.Forms.TextBox();
			this.lblTitNumPedido = new System.Windows.Forms.Label();
			this.txtNumNF = new System.Windows.Forms.TextBox();
			this.lblTitNumNF = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.txtNomeCliente = new System.Windows.Forms.TextBox();
			this.lblTitNomeCliente = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.txtDataVenctoFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodoVencto = new System.Windows.Forms.Label();
			this.txtDataVenctoInicial = new System.Windows.Forms.TextBox();
			this.lblDataVenctoAte = new System.Windows.Forms.Label();
			this.pnResultado = new System.Windows.Forms.Panel();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.colCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cnpj_cpf_formatado = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.num_documento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.num_parcela = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.situacao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_vencto_formatada = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valor_formatado = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto_item = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnTotalizacao = new System.Windows.Forms.Panel();
			this.btnDesmarcarTodos = new System.Windows.Forms.Button();
			this.btnMarcarTodos = new System.Windows.Forms.Button();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.btnDetalhe = new System.Windows.Forms.Button();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.btnBoletoEmail = new System.Windows.Forms.Button();
			this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.btnPrintPreview = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
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
			this.pnBotoes.Controls.Add(this.btnBoletoEmail);
			this.pnBotoes.Controls.Add(this.btnLimpar);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.Add(this.btnDetalhe);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDetalhe, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnLimpar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnBoletoEmail, 0);
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
			this.lblTitulo.Text = "Consulta de Boletos";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pnParametros
			// 
			this.pnParametros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnParametros.Controls.Add(this.lblTitCedente);
			this.pnParametros.Controls.Add(this.cbBoletoCedente);
			this.pnParametros.Controls.Add(this.txtDataCargaRetornoFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodoCargaRetorno);
			this.pnParametros.Controls.Add(this.txtDataCargaRetornoInicial);
			this.pnParametros.Controls.Add(this.lblDataCargaRetornoAte);
			this.pnParametros.Controls.Add(this.cbOcorrencia);
			this.pnParametros.Controls.Add(this.lblTitCtrlPagtoStatus);
			this.pnParametros.Controls.Add(this.txtNumPedido);
			this.pnParametros.Controls.Add(this.lblTitNumPedido);
			this.pnParametros.Controls.Add(this.txtNumNF);
			this.pnParametros.Controls.Add(this.lblTitNumNF);
			this.pnParametros.Controls.Add(this.txtCnpjCpf);
			this.pnParametros.Controls.Add(this.lblCnpjCpf);
			this.pnParametros.Controls.Add(this.txtValor);
			this.pnParametros.Controls.Add(this.txtNomeCliente);
			this.pnParametros.Controls.Add(this.lblTitNomeCliente);
			this.pnParametros.Controls.Add(this.lblValor);
			this.pnParametros.Controls.Add(this.txtDataVenctoFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodoVencto);
			this.pnParametros.Controls.Add(this.txtDataVenctoInicial);
			this.pnParametros.Controls.Add(this.lblDataVenctoAte);
			this.pnParametros.Dock = System.Windows.Forms.DockStyle.Top;
			this.pnParametros.Location = new System.Drawing.Point(0, 40);
			this.pnParametros.Name = "pnParametros";
			this.pnParametros.Size = new System.Drawing.Size(1014, 132);
			this.pnParametros.TabIndex = 2;
			// 
			// lblTitCedente
			// 
			this.lblTitCedente.AutoSize = true;
			this.lblTitCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCedente.Location = new System.Drawing.Point(59, 73);
			this.lblTitCedente.Name = "lblTitCedente";
			this.lblTitCedente.Size = new System.Drawing.Size(54, 13);
			this.lblTitCedente.TabIndex = 51;
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
			this.cbBoletoCedente.Location = new System.Drawing.Point(119, 70);
			this.cbBoletoCedente.MaxDropDownItems = 12;
			this.cbBoletoCedente.Name = "cbBoletoCedente";
			this.cbBoletoCedente.Size = new System.Drawing.Size(452, 21);
			this.cbBoletoCedente.TabIndex = 7;
			this.cbBoletoCedente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbBoletoCedente_KeyDown);
			// 
			// txtDataCargaRetornoFinal
			// 
			this.txtDataCargaRetornoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCargaRetornoFinal.Location = new System.Drawing.Point(236, 6);
			this.txtDataCargaRetornoFinal.MaxLength = 10;
			this.txtDataCargaRetornoFinal.Name = "txtDataCargaRetornoFinal";
			this.txtDataCargaRetornoFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataCargaRetornoFinal.TabIndex = 1;
			this.txtDataCargaRetornoFinal.Text = "01/01/2000";
			this.txtDataCargaRetornoFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCargaRetornoFinal.Enter += new System.EventHandler(this.txtDataCargaRetornoFinal_Enter);
			this.txtDataCargaRetornoFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCargaRetornoFinal_KeyDown);
			this.txtDataCargaRetornoFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCargaRetornoFinal_KeyPress);
			this.txtDataCargaRetornoFinal.Leave += new System.EventHandler(this.txtDataCargaRetornoFinal_Leave);
			// 
			// lblTitPeriodoCargaRetorno
			// 
			this.lblTitPeriodoCargaRetorno.AutoSize = true;
			this.lblTitPeriodoCargaRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodoCargaRetorno.Location = new System.Drawing.Point(6, 11);
			this.lblTitPeriodoCargaRetorno.Name = "lblTitPeriodoCargaRetorno";
			this.lblTitPeriodoCargaRetorno.Size = new System.Drawing.Size(107, 13);
			this.lblTitPeriodoCargaRetorno.TabIndex = 49;
			this.lblTitPeriodoCargaRetorno.Text = "Carga do Retorno";
			// 
			// txtDataCargaRetornoInicial
			// 
			this.txtDataCargaRetornoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCargaRetornoInicial.Location = new System.Drawing.Point(119, 6);
			this.txtDataCargaRetornoInicial.MaxLength = 10;
			this.txtDataCargaRetornoInicial.Name = "txtDataCargaRetornoInicial";
			this.txtDataCargaRetornoInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataCargaRetornoInicial.TabIndex = 0;
			this.txtDataCargaRetornoInicial.Text = "01/01/2000";
			this.txtDataCargaRetornoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCargaRetornoInicial.Enter += new System.EventHandler(this.txtDataCargaRetornoInicial_Enter);
			this.txtDataCargaRetornoInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCargaRetornoInicial_KeyDown);
			this.txtDataCargaRetornoInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCargaRetornoInicial_KeyPress);
			this.txtDataCargaRetornoInicial.Leave += new System.EventHandler(this.txtDataCargaRetornoInicial_Leave);
			// 
			// lblDataCargaRetornoAte
			// 
			this.lblDataCargaRetornoAte.AutoSize = true;
			this.lblDataCargaRetornoAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblDataCargaRetornoAte.Location = new System.Drawing.Point(216, 11);
			this.lblDataCargaRetornoAte.Name = "lblDataCargaRetornoAte";
			this.lblDataCargaRetornoAte.Size = new System.Drawing.Size(14, 13);
			this.lblDataCargaRetornoAte.TabIndex = 48;
			this.lblDataCargaRetornoAte.Text = "a";
			// 
			// cbOcorrencia
			// 
			this.cbOcorrencia.BackColor = System.Drawing.SystemColors.Window;
			this.cbOcorrencia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbOcorrencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbOcorrencia.FormattingEnabled = true;
			this.cbOcorrencia.Location = new System.Drawing.Point(119, 39);
			this.cbOcorrencia.Name = "cbOcorrencia";
			this.cbOcorrencia.Size = new System.Drawing.Size(452, 21);
			this.cbOcorrencia.TabIndex = 5;
			this.cbOcorrencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbOcorrencia_KeyDown);
			// 
			// lblTitCtrlPagtoStatus
			// 
			this.lblTitCtrlPagtoStatus.AutoSize = true;
			this.lblTitCtrlPagtoStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCtrlPagtoStatus.Location = new System.Drawing.Point(70, 42);
			this.lblTitCtrlPagtoStatus.Name = "lblTitCtrlPagtoStatus";
			this.lblTitCtrlPagtoStatus.Size = new System.Drawing.Size(43, 13);
			this.lblTitCtrlPagtoStatus.TabIndex = 45;
			this.lblTitCtrlPagtoStatus.Text = "Status";
			// 
			// txtNumPedido
			// 
			this.txtNumPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumPedido.Location = new System.Drawing.Point(883, 37);
			this.txtNumPedido.MaxLength = 9;
			this.txtNumPedido.Name = "txtNumPedido";
			this.txtNumPedido.Size = new System.Drawing.Size(122, 23);
			this.txtNumPedido.TabIndex = 6;
			this.txtNumPedido.Text = "123456789";
			this.txtNumPedido.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumPedido.Enter += new System.EventHandler(this.txtNumPedido_Enter);
			this.txtNumPedido.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumPedido_KeyDown);
			this.txtNumPedido.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumPedido_KeyPress);
			this.txtNumPedido.Leave += new System.EventHandler(this.txtNumPedido_Leave);
			// 
			// lblTitNumPedido
			// 
			this.lblTitNumPedido.AutoSize = true;
			this.lblTitNumPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNumPedido.Location = new System.Drawing.Point(813, 42);
			this.lblTitNumPedido.Name = "lblTitNumPedido";
			this.lblTitNumPedido.Size = new System.Drawing.Size(64, 13);
			this.lblTitNumPedido.TabIndex = 43;
			this.lblTitNumPedido.Text = "Nº Pedido";
			// 
			// txtNumNF
			// 
			this.txtNumNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumNF.Location = new System.Drawing.Point(883, 6);
			this.txtNumNF.MaxLength = 10;
			this.txtNumNF.Name = "txtNumNF";
			this.txtNumNF.Size = new System.Drawing.Size(122, 23);
			this.txtNumNF.TabIndex = 4;
			this.txtNumNF.Text = "1234567890";
			this.txtNumNF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumNF.Enter += new System.EventHandler(this.txtNumNF_Enter);
			this.txtNumNF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumNF_KeyDown);
			this.txtNumNF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumNF_KeyPress);
			this.txtNumNF.Leave += new System.EventHandler(this.txtNumNF_Leave);
			// 
			// lblTitNumNF
			// 
			this.lblTitNumNF.AutoSize = true;
			this.lblTitNumNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNumNF.Location = new System.Drawing.Point(836, 11);
			this.lblTitNumNF.Name = "lblTitNumNF";
			this.lblTitNumNF.Size = new System.Drawing.Size(41, 13);
			this.lblTitNumNF.TabIndex = 41;
			this.lblTitNumNF.Text = "Nº NF";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(860, 99);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(145, 23);
			this.txtCnpjCpf.TabIndex = 10;
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
			this.lblCnpjCpf.Location = new System.Drawing.Point(787, 104);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(67, 13);
			this.lblCnpjCpf.TabIndex = 31;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(883, 68);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(122, 23);
			this.txtValor.TabIndex = 8;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			// 
			// txtNomeCliente
			// 
			this.txtNomeCliente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.txtNomeCliente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
			this.txtNomeCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNomeCliente.Location = new System.Drawing.Point(119, 99);
			this.txtNomeCliente.MaxLength = 40;
			this.txtNomeCliente.Name = "txtNomeCliente";
			this.txtNomeCliente.Size = new System.Drawing.Size(618, 23);
			this.txtNomeCliente.TabIndex = 9;
			this.txtNomeCliente.Enter += new System.EventHandler(this.txtNomeCliente_Enter);
			this.txtNomeCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNomeCliente_KeyDown);
			this.txtNomeCliente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNomeCliente_KeyPress);
			this.txtNomeCliente.Leave += new System.EventHandler(this.txtNomeCliente_Leave);
			// 
			// lblTitNomeCliente
			// 
			this.lblTitNomeCliente.AutoSize = true;
			this.lblTitNomeCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNomeCliente.Location = new System.Drawing.Point(31, 104);
			this.lblTitNomeCliente.Name = "lblTitNomeCliente";
			this.lblTitNomeCliente.Size = new System.Drawing.Size(82, 13);
			this.lblTitNomeCliente.TabIndex = 29;
			this.lblTitNomeCliente.Text = "Nome Cliente";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValor.Location = new System.Drawing.Point(813, 73);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(64, 13);
			this.lblValor.TabIndex = 28;
			this.lblValor.Text = "Valor (R$)";
			// 
			// txtDataVenctoFinal
			// 
			this.txtDataVenctoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataVenctoFinal.Location = new System.Drawing.Point(646, 6);
			this.txtDataVenctoFinal.MaxLength = 10;
			this.txtDataVenctoFinal.Name = "txtDataVenctoFinal";
			this.txtDataVenctoFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataVenctoFinal.TabIndex = 3;
			this.txtDataVenctoFinal.Text = "01/01/2000";
			this.txtDataVenctoFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataVenctoFinal.Enter += new System.EventHandler(this.txtDataVenctoFinal_Enter);
			this.txtDataVenctoFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataVenctoFinal_KeyDown);
			this.txtDataVenctoFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataVenctoFinal_KeyPress);
			this.txtDataVenctoFinal.Leave += new System.EventHandler(this.txtDataVenctoFinal_Leave);
			// 
			// lblTitPeriodoVencto
			// 
			this.lblTitPeriodoVencto.AutoSize = true;
			this.lblTitPeriodoVencto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodoVencto.Location = new System.Drawing.Point(450, 11);
			this.lblTitPeriodoVencto.Name = "lblTitPeriodoVencto";
			this.lblTitPeriodoVencto.Size = new System.Drawing.Size(73, 13);
			this.lblTitPeriodoVencto.TabIndex = 10;
			this.lblTitPeriodoVencto.Text = "Vencimento";
			// 
			// txtDataVenctoInicial
			// 
			this.txtDataVenctoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataVenctoInicial.Location = new System.Drawing.Point(529, 6);
			this.txtDataVenctoInicial.MaxLength = 10;
			this.txtDataVenctoInicial.Name = "txtDataVenctoInicial";
			this.txtDataVenctoInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataVenctoInicial.TabIndex = 2;
			this.txtDataVenctoInicial.Text = "01/01/2000";
			this.txtDataVenctoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataVenctoInicial.Enter += new System.EventHandler(this.txtDataVenctoInicial_Enter);
			this.txtDataVenctoInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataVenctoInicial_KeyDown);
			this.txtDataVenctoInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataVenctoInicial_KeyPress);
			this.txtDataVenctoInicial.Leave += new System.EventHandler(this.txtDataVenctoInicial_Leave);
			// 
			// lblDataVenctoAte
			// 
			this.lblDataVenctoAte.AutoSize = true;
			this.lblDataVenctoAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblDataVenctoAte.Location = new System.Drawing.Point(626, 11);
			this.lblDataVenctoAte.Name = "lblDataVenctoAte";
			this.lblDataVenctoAte.Size = new System.Drawing.Size(14, 13);
			this.lblDataVenctoAte.TabIndex = 9;
			this.lblDataVenctoAte.Text = "a";
			// 
			// pnResultado
			// 
			this.pnResultado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnResultado.Controls.Add(this.gridDados);
			this.pnResultado.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pnResultado.Location = new System.Drawing.Point(0, 172);
			this.pnResultado.Name = "pnResultado";
			this.pnResultado.Size = new System.Drawing.Size(1014, 404);
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
            this.colCheck,
            this.cliente,
            this.cnpj_cpf_formatado,
            this.pedido,
            this.num_documento,
            this.num_parcela,
            this.situacao,
            this.dt_vencto_formatada,
            this.valor_formatado,
            this.id_boleto_item});
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.gridDados.DefaultCellStyle = dataGridViewCellStyle10;
			this.gridDados.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridDados.Location = new System.Drawing.Point(0, 0);
			this.gridDados.MultiSelect = false;
			this.gridDados.Name = "gridDados";
			dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.RowHeadersDefaultCellStyle = dataGridViewCellStyle11;
			this.gridDados.RowHeadersVisible = false;
			this.gridDados.RowHeadersWidth = 15;
			this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridDados.Size = new System.Drawing.Size(1010, 400);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 0;
			this.gridDados.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridDados_CellContentClick);
			this.gridDados.DoubleClick += new System.EventHandler(this.gridDados_DoubleClick);
			this.gridDados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridDados_KeyDown);
			// 
			// colCheck
			// 
			this.colCheck.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.colCheck.HeaderText = "";
			this.colCheck.MinimumWidth = 20;
			this.colCheck.Name = "colCheck";
			this.colCheck.Width = 20;
			// 
			// cliente
			// 
			this.cliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.cliente.DefaultCellStyle = dataGridViewCellStyle2;
			this.cliente.HeaderText = "Cliente";
			this.cliente.MinimumWidth = 180;
			this.cliente.Name = "cliente";
			this.cliente.ReadOnly = true;
			// 
			// cnpj_cpf_formatado
			// 
			this.cnpj_cpf_formatado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.cnpj_cpf_formatado.DefaultCellStyle = dataGridViewCellStyle3;
			this.cnpj_cpf_formatado.HeaderText = "CNPJ / CPF";
			this.cnpj_cpf_formatado.MinimumWidth = 130;
			this.cnpj_cpf_formatado.Name = "cnpj_cpf_formatado";
			this.cnpj_cpf_formatado.ReadOnly = true;
			this.cnpj_cpf_formatado.Width = 130;
			// 
			// pedido
			// 
			this.pedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.pedido.DefaultCellStyle = dataGridViewCellStyle4;
			this.pedido.HeaderText = "Pedido";
			this.pedido.MinimumWidth = 110;
			this.pedido.Name = "pedido";
			this.pedido.ReadOnly = true;
			this.pedido.Width = 110;
			// 
			// num_documento
			// 
			this.num_documento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.num_documento.DefaultCellStyle = dataGridViewCellStyle5;
			this.num_documento.HeaderText = "Nº Doc";
			this.num_documento.MinimumWidth = 85;
			this.num_documento.Name = "num_documento";
			this.num_documento.ReadOnly = true;
			this.num_documento.Width = 85;
			// 
			// num_parcela
			// 
			this.num_parcela.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.num_parcela.DefaultCellStyle = dataGridViewCellStyle6;
			this.num_parcela.HeaderText = "Parcela";
			this.num_parcela.MinimumWidth = 75;
			this.num_parcela.Name = "num_parcela";
			this.num_parcela.ReadOnly = true;
			this.num_parcela.Width = 75;
			// 
			// situacao
			// 
			this.situacao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.situacao.DefaultCellStyle = dataGridViewCellStyle7;
			this.situacao.HeaderText = "Situação";
			this.situacao.MinimumWidth = 160;
			this.situacao.Name = "situacao";
			this.situacao.ReadOnly = true;
			this.situacao.Width = 160;
			// 
			// dt_vencto_formatada
			// 
			this.dt_vencto_formatada.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.dt_vencto_formatada.DefaultCellStyle = dataGridViewCellStyle8;
			this.dt_vencto_formatada.HeaderText = "Vencto";
			this.dt_vencto_formatada.MinimumWidth = 80;
			this.dt_vencto_formatada.Name = "dt_vencto_formatada";
			this.dt_vencto_formatada.ReadOnly = true;
			this.dt_vencto_formatada.Width = 80;
			// 
			// valor_formatado
			// 
			this.valor_formatado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			this.valor_formatado.DefaultCellStyle = dataGridViewCellStyle9;
			this.valor_formatado.HeaderText = "Valor";
			this.valor_formatado.MinimumWidth = 100;
			this.valor_formatado.Name = "valor_formatado";
			this.valor_formatado.ReadOnly = true;
			// 
			// id_boleto_item
			// 
			this.id_boleto_item.HeaderText = "id_boleto_item";
			this.id_boleto_item.Name = "id_boleto_item";
			this.id_boleto_item.ReadOnly = true;
			this.id_boleto_item.Visible = false;
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
			this.pnTotalizacao.TabIndex = 3;
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
			// btnDetalhe
			// 
			this.btnDetalhe.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnDetalhe.Image = ((System.Drawing.Image)(resources.GetObject("btnDetalhe.Image")));
			this.btnDetalhe.Location = new System.Drawing.Point(699, 4);
			this.btnDetalhe.Name = "btnDetalhe";
			this.btnDetalhe.Size = new System.Drawing.Size(40, 44);
			this.btnDetalhe.TabIndex = 2;
			this.btnDetalhe.TabStop = false;
			this.btnDetalhe.UseVisualStyleBackColor = true;
			this.btnDetalhe.Click += new System.EventHandler(this.btnDetalhe_Click);
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
			// btnBoletoEmail
			// 
			this.btnBoletoEmail.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnBoletoEmail.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoEmail.Image")));
			this.btnBoletoEmail.Location = new System.Drawing.Point(744, 4);
			this.btnBoletoEmail.Name = "btnBoletoEmail";
			this.btnBoletoEmail.Size = new System.Drawing.Size(40, 44);
			this.btnBoletoEmail.TabIndex = 3;
			this.btnBoletoEmail.TabStop = false;
			this.btnBoletoEmail.UseVisualStyleBackColor = true;
			this.btnBoletoEmail.Click += new System.EventHandler(this.btnBoletoEmail_Click);
			// 
			// prnDocConsulta
			// 
			this.prnDocConsulta.DocumentName = "Consulta de Boletos";
			this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
			this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
			this.prnDocConsulta.QueryPageSettings += new System.Drawing.Printing.QueryPageSettingsEventHandler(this.prnDocConsulta_QueryPageSettings);
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
			// FBoletoConsulta
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoConsulta";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoConsulta_FormClosing);
			this.Load += new System.EventHandler(this.FBoletoConsulta_Load);
			this.Shown += new System.EventHandler(this.FBoletoConsulta_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FBoletoConsulta_KeyDown);
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
		private System.Windows.Forms.TextBox txtDataVenctoInicial;
		private System.Windows.Forms.Label lblDataVenctoAte;
		private System.Windows.Forms.Label lblTitPeriodoVencto;
		private System.Windows.Forms.TextBox txtDataVenctoFinal;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.TextBox txtNomeCliente;
		private System.Windows.Forms.Label lblTitNomeCliente;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.Panel pnResultado;
		private System.Windows.Forms.DataGridView gridDados;
		private System.Windows.Forms.Panel pnTotalizacao;
		private System.Windows.Forms.Label lblTotalizacaoValor;
		private System.Windows.Forms.Label lblTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTitTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTitTotalizacaoValor;
		private System.Windows.Forms.Button btnDetalhe;
		private System.Windows.Forms.Button btnPesquisar;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.TextBox txtNumNF;
		private System.Windows.Forms.Label lblTitNumNF;
		private System.Windows.Forms.TextBox txtNumPedido;
		private System.Windows.Forms.Label lblTitNumPedido;
		private System.Windows.Forms.ComboBox cbOcorrencia;
		private System.Windows.Forms.Label lblTitCtrlPagtoStatus;
		private System.Windows.Forms.TextBox txtDataCargaRetornoFinal;
		private System.Windows.Forms.Label lblTitPeriodoCargaRetorno;
		private System.Windows.Forms.TextBox txtDataCargaRetornoInicial;
		private System.Windows.Forms.Label lblDataCargaRetornoAte;
		private System.Windows.Forms.DataGridViewCheckBoxColumn colCheck;
		private System.Windows.Forms.DataGridViewTextBoxColumn cliente;
		private System.Windows.Forms.DataGridViewTextBoxColumn cnpj_cpf_formatado;
		private System.Windows.Forms.DataGridViewTextBoxColumn pedido;
		private System.Windows.Forms.DataGridViewTextBoxColumn num_documento;
		private System.Windows.Forms.DataGridViewTextBoxColumn num_parcela;
		private System.Windows.Forms.DataGridViewTextBoxColumn situacao;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_vencto_formatada;
		private System.Windows.Forms.DataGridViewTextBoxColumn valor_formatado;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto_item;
		private System.Windows.Forms.Button btnBoletoEmail;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnPrintPreview;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label lblTitCedente;
		private System.Windows.Forms.ComboBox cbBoletoCedente;
		private System.Windows.Forms.Button btnDesmarcarTodos;
		private System.Windows.Forms.Button btnMarcarTodos;
	}
}
