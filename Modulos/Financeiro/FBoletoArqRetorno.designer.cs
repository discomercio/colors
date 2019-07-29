namespace Financeiro
{
	partial class FBoletoArqRetorno
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoArqRetorno));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.btnCarregaArqRetorno = new System.Windows.Forms.Button();
			this.lblTitArqRetorno = new System.Windows.Forms.Label();
			this.txtArqRetorno = new System.Windows.Forms.TextBox();
			this.btnSelecionaArqRetorno = new System.Windows.Forms.Button();
			this.gboxBoletos = new System.Windows.Forms.GroupBox();
			this.lblTotalRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalRegistros = new System.Windows.Forms.Label();
			this.grdBoletos = new System.Windows.Forms.DataGridView();
			this.numeroDocumento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dataVenctoTitulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valorTitulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.identificacaoOcorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.motivosRejeicoes = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto_item = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.gboxCedente = new System.Windows.Forms.GroupBox();
			this.lblCedenteDataCredito = new System.Windows.Forms.Label();
			this.lblTitCedenteDataCredito = new System.Windows.Forms.Label();
			this.lblCedenteDataBanco = new System.Windows.Forms.Label();
			this.lblTitCedenteDataBanco = new System.Windows.Forms.Label();
			this.lblCedenteNumAvisoBancario = new System.Windows.Forms.Label();
			this.lblTitCedenteNumAvisoBancario = new System.Windows.Forms.Label();
			this.lblCedenteCodigoEmpresa = new System.Windows.Forms.Label();
			this.lblTitCedenteCodigoEmpresa = new System.Windows.Forms.Label();
			this.lblCedenteConta = new System.Windows.Forms.Label();
			this.lblTitCedenteConta = new System.Windows.Forms.Label();
			this.lblCedenteAgencia = new System.Windows.Forms.Label();
			this.lblTitCedenteAgencia = new System.Windows.Forms.Label();
			this.lblCedenteCarteira = new System.Windows.Forms.Label();
			this.lblTitCedenteCarteira = new System.Windows.Forms.Label();
			this.lblCedenteNome = new System.Windows.Forms.Label();
			this.lblTitCedenteNome = new System.Windows.Forms.Label();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxBoletos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).BeginInit();
			this.gboxMensagensInformativas.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.gboxCedente.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnCarregaArqRetorno);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCarregaArqRetorno, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxBoletos);
			this.pnCampos.Controls.Add(this.gboxCedente);
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.btnSelecionaArqRetorno);
			this.pnCampos.Controls.Add(this.txtArqRetorno);
			this.pnCampos.Controls.Add(this.lblTitArqRetorno);
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
			this.lblTitulo.Text = "Boleto: Carga do Arquivo de Retorno";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnCarregaArqRetorno
			// 
			this.btnCarregaArqRetorno.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCarregaArqRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnCarregaArqRetorno.Image")));
			this.btnCarregaArqRetorno.Location = new System.Drawing.Point(879, 4);
			this.btnCarregaArqRetorno.Name = "btnCarregaArqRetorno";
			this.btnCarregaArqRetorno.Size = new System.Drawing.Size(40, 44);
			this.btnCarregaArqRetorno.TabIndex = 0;
			this.btnCarregaArqRetorno.TabStop = false;
			this.btnCarregaArqRetorno.UseVisualStyleBackColor = true;
			this.btnCarregaArqRetorno.Click += new System.EventHandler(this.btnCarregaArqRetorno_Click);
			// 
			// lblTitArqRetorno
			// 
			this.lblTitArqRetorno.AutoSize = true;
			this.lblTitArqRetorno.Location = new System.Drawing.Point(9, 51);
			this.lblTitArqRetorno.Name = "lblTitArqRetorno";
			this.lblTitArqRetorno.Size = new System.Drawing.Size(99, 13);
			this.lblTitArqRetorno.TabIndex = 1;
			this.lblTitArqRetorno.Text = "Arquivo de Retorno";
			// 
			// txtArqRetorno
			// 
			this.txtArqRetorno.BackColor = System.Drawing.SystemColors.Window;
			this.txtArqRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtArqRetorno.Location = new System.Drawing.Point(109, 48);
			this.txtArqRetorno.Name = "txtArqRetorno";
			this.txtArqRetorno.ReadOnly = true;
			this.txtArqRetorno.Size = new System.Drawing.Size(695, 20);
			this.txtArqRetorno.TabIndex = 2;
			this.txtArqRetorno.DoubleClick += new System.EventHandler(this.txtArqRetorno_DoubleClick);
			this.txtArqRetorno.Enter += new System.EventHandler(this.txtArqRetorno_Enter);
			// 
			// btnSelecionaArqRetorno
			// 
			this.btnSelecionaArqRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSelecionaArqRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArqRetorno.Image")));
			this.btnSelecionaArqRetorno.Location = new System.Drawing.Point(811, 45);
			this.btnSelecionaArqRetorno.Name = "btnSelecionaArqRetorno";
			this.btnSelecionaArqRetorno.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaArqRetorno.TabIndex = 3;
			this.btnSelecionaArqRetorno.UseVisualStyleBackColor = true;
			this.btnSelecionaArqRetorno.Click += new System.EventHandler(this.btnSelecionaArqRetorno_Click);
			// 
			// gboxBoletos
			// 
			this.gboxBoletos.Controls.Add(this.lblTotalRegistros);
			this.gboxBoletos.Controls.Add(this.lblTitTotalRegistros);
			this.gboxBoletos.Controls.Add(this.grdBoletos);
			this.gboxBoletos.Location = new System.Drawing.Point(10, 136);
			this.gboxBoletos.Name = "gboxBoletos";
			this.gboxBoletos.Size = new System.Drawing.Size(995, 262);
			this.gboxBoletos.TabIndex = 4;
			this.gboxBoletos.TabStop = false;
			this.gboxBoletos.Text = "Dados do Arquivo de Retorno";
			// 
			// lblTotalRegistros
			// 
			this.lblTotalRegistros.AutoSize = true;
			this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalRegistros.Location = new System.Drawing.Point(100, 244);
			this.lblTotalRegistros.Name = "lblTotalRegistros";
			this.lblTotalRegistros.Size = new System.Drawing.Size(28, 13);
			this.lblTotalRegistros.TabIndex = 6;
			this.lblTotalRegistros.Text = "999";
			// 
			// lblTitTotalRegistros
			// 
			this.lblTitTotalRegistros.AutoSize = true;
			this.lblTitTotalRegistros.Location = new System.Drawing.Point(12, 244);
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
            this.numeroDocumento,
            this.dataVenctoTitulo,
            this.valorTitulo,
            this.identificacaoOcorrencia,
            this.motivosRejeicoes,
            this.id_boleto_item});
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdBoletos.DefaultCellStyle = dataGridViewCellStyle7;
			this.grdBoletos.Location = new System.Drawing.Point(15, 19);
			this.grdBoletos.MultiSelect = false;
			this.grdBoletos.Name = "grdBoletos";
			this.grdBoletos.ReadOnly = true;
			this.grdBoletos.RowHeadersVisible = false;
			this.grdBoletos.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdBoletos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdBoletos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdBoletos.ShowEditingIcon = false;
			this.grdBoletos.Size = new System.Drawing.Size(965, 223);
			this.grdBoletos.StandardTab = true;
			this.grdBoletos.TabIndex = 0;
			// 
			// numeroDocumento
			// 
			this.numeroDocumento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.numeroDocumento.DefaultCellStyle = dataGridViewCellStyle2;
			this.numeroDocumento.HeaderText = "Nº Documento";
			this.numeroDocumento.MinimumWidth = 120;
			this.numeroDocumento.Name = "numeroDocumento";
			this.numeroDocumento.ReadOnly = true;
			this.numeroDocumento.Width = 120;
			// 
			// dataVenctoTitulo
			// 
			this.dataVenctoTitulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.dataVenctoTitulo.DefaultCellStyle = dataGridViewCellStyle3;
			this.dataVenctoTitulo.HeaderText = "Vencto";
			this.dataVenctoTitulo.MinimumWidth = 80;
			this.dataVenctoTitulo.Name = "dataVenctoTitulo";
			this.dataVenctoTitulo.ReadOnly = true;
			this.dataVenctoTitulo.Width = 80;
			// 
			// valorTitulo
			// 
			this.valorTitulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			this.valorTitulo.DefaultCellStyle = dataGridViewCellStyle4;
			this.valorTitulo.HeaderText = "Valor";
			this.valorTitulo.MinimumWidth = 140;
			this.valorTitulo.Name = "valorTitulo";
			this.valorTitulo.ReadOnly = true;
			this.valorTitulo.Width = 140;
			// 
			// identificacaoOcorrencia
			// 
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.identificacaoOcorrencia.DefaultCellStyle = dataGridViewCellStyle5;
			this.identificacaoOcorrencia.HeaderText = "Ocorrência";
			this.identificacaoOcorrencia.MinimumWidth = 250;
			this.identificacaoOcorrencia.Name = "identificacaoOcorrencia";
			this.identificacaoOcorrencia.ReadOnly = true;
			this.identificacaoOcorrencia.Width = 250;
			// 
			// motivosRejeicoes
			// 
			this.motivosRejeicoes.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.motivosRejeicoes.DefaultCellStyle = dataGridViewCellStyle6;
			this.motivosRejeicoes.HeaderText = "Motivos";
			this.motivosRejeicoes.MinimumWidth = 120;
			this.motivosRejeicoes.Name = "motivosRejeicoes";
			this.motivosRejeicoes.ReadOnly = true;
			// 
			// id_boleto_item
			// 
			this.id_boleto_item.HeaderText = "id_boleto_item";
			this.id_boleto_item.Name = "id_boleto_item";
			this.id_boleto_item.ReadOnly = true;
			this.id_boleto_item.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.id_boleto_item.Visible = false;
			// 
			// openFileDialog
			// 
			this.openFileDialog.AddExtension = false;
			this.openFileDialog.Filter = "Arquivo de retorno|*.RET|Todos os arquivos|*.*";
			this.openFileDialog.InitialDirectory = "\\";
			this.openFileDialog.Title = "Selecionar arquivo de retorno";
			// 
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(10, 404);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(994, 95);
			this.gboxMensagensInformativas.TabIndex = 5;
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
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(10, 505);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(994, 95);
			this.gboxMsgErro.TabIndex = 6;
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
			// gboxCedente
			// 
			this.gboxCedente.Controls.Add(this.lblCedenteDataCredito);
			this.gboxCedente.Controls.Add(this.lblTitCedenteDataCredito);
			this.gboxCedente.Controls.Add(this.lblCedenteDataBanco);
			this.gboxCedente.Controls.Add(this.lblTitCedenteDataBanco);
			this.gboxCedente.Controls.Add(this.lblCedenteNumAvisoBancario);
			this.gboxCedente.Controls.Add(this.lblTitCedenteNumAvisoBancario);
			this.gboxCedente.Controls.Add(this.lblCedenteCodigoEmpresa);
			this.gboxCedente.Controls.Add(this.lblTitCedenteCodigoEmpresa);
			this.gboxCedente.Controls.Add(this.lblCedenteConta);
			this.gboxCedente.Controls.Add(this.lblTitCedenteConta);
			this.gboxCedente.Controls.Add(this.lblCedenteAgencia);
			this.gboxCedente.Controls.Add(this.lblTitCedenteAgencia);
			this.gboxCedente.Controls.Add(this.lblCedenteCarteira);
			this.gboxCedente.Controls.Add(this.lblTitCedenteCarteira);
			this.gboxCedente.Controls.Add(this.lblCedenteNome);
			this.gboxCedente.Controls.Add(this.lblTitCedenteNome);
			this.gboxCedente.Location = new System.Drawing.Point(10, 78);
			this.gboxCedente.Name = "gboxCedente";
			this.gboxCedente.Size = new System.Drawing.Size(995, 50);
			this.gboxCedente.TabIndex = 8;
			this.gboxCedente.TabStop = false;
			this.gboxCedente.Text = "Informações do Arquivo de Retorno";
			// 
			// lblCedenteDataCredito
			// 
			this.lblCedenteDataCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteDataCredito.Location = new System.Drawing.Point(657, 32);
			this.lblCedenteDataCredito.Name = "lblCedenteDataCredito";
			this.lblCedenteDataCredito.Size = new System.Drawing.Size(85, 13);
			this.lblCedenteDataCredito.TabIndex = 15;
			this.lblCedenteDataCredito.Text = "23/05/2011";
			// 
			// lblTitCedenteDataCredito
			// 
			this.lblTitCedenteDataCredito.AutoSize = true;
			this.lblTitCedenteDataCredito.Location = new System.Drawing.Point(571, 32);
			this.lblTitCedenteDataCredito.Name = "lblTitCedenteDataCredito";
			this.lblTitCedenteDataCredito.Size = new System.Drawing.Size(84, 13);
			this.lblTitCedenteDataCredito.TabIndex = 14;
			this.lblTitCedenteDataCredito.Text = "Data do Crédito:";
			// 
			// lblCedenteDataBanco
			// 
			this.lblCedenteDataBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteDataBanco.Location = new System.Drawing.Point(403, 32);
			this.lblCedenteDataBanco.Name = "lblCedenteDataBanco";
			this.lblCedenteDataBanco.Size = new System.Drawing.Size(85, 13);
			this.lblCedenteDataBanco.TabIndex = 13;
			this.lblCedenteDataBanco.Text = "23/05/2011";
			// 
			// lblTitCedenteDataBanco
			// 
			this.lblTitCedenteDataBanco.AutoSize = true;
			this.lblTitCedenteDataBanco.Location = new System.Drawing.Point(319, 32);
			this.lblTitCedenteDataBanco.Name = "lblTitCedenteDataBanco";
			this.lblTitCedenteDataBanco.Size = new System.Drawing.Size(82, 13);
			this.lblTitCedenteDataBanco.TabIndex = 12;
			this.lblTitCedenteDataBanco.Text = "Data no Banco:";
			// 
			// lblCedenteNumAvisoBancario
			// 
			this.lblCedenteNumAvisoBancario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteNumAvisoBancario.Location = new System.Drawing.Point(104, 32);
			this.lblCedenteNumAvisoBancario.Name = "lblCedenteNumAvisoBancario";
			this.lblCedenteNumAvisoBancario.Size = new System.Drawing.Size(76, 13);
			this.lblCedenteNumAvisoBancario.TabIndex = 11;
			this.lblCedenteNumAvisoBancario.Text = "00000";
			// 
			// lblTitCedenteNumAvisoBancario
			// 
			this.lblTitCedenteNumAvisoBancario.AutoSize = true;
			this.lblTitCedenteNumAvisoBancario.Location = new System.Drawing.Point(6, 32);
			this.lblTitCedenteNumAvisoBancario.Name = "lblTitCedenteNumAvisoBancario";
			this.lblTitCedenteNumAvisoBancario.Size = new System.Drawing.Size(96, 13);
			this.lblTitCedenteNumAvisoBancario.TabIndex = 10;
			this.lblTitCedenteNumAvisoBancario.Text = "Nº Aviso Bancário:";
			// 
			// lblCedenteCodigoEmpresa
			// 
			this.lblCedenteCodigoEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteCodigoEmpresa.Location = new System.Drawing.Point(808, 16);
			this.lblCedenteCodigoEmpresa.Name = "lblCedenteCodigoEmpresa";
			this.lblCedenteCodigoEmpresa.Size = new System.Drawing.Size(84, 13);
			this.lblCedenteCodigoEmpresa.TabIndex = 9;
			this.lblCedenteCodigoEmpresa.Text = "0000000";
			// 
			// lblTitCedenteCodigoEmpresa
			// 
			this.lblTitCedenteCodigoEmpresa.AutoSize = true;
			this.lblTitCedenteCodigoEmpresa.Location = new System.Drawing.Point(704, 16);
			this.lblTitCedenteCodigoEmpresa.Name = "lblTitCedenteCodigoEmpresa";
			this.lblTitCedenteCodigoEmpresa.Size = new System.Drawing.Size(102, 13);
			this.lblTitCedenteCodigoEmpresa.TabIndex = 8;
			this.lblTitCedenteCodigoEmpresa.Text = "Código da Empresa:";
			// 
			// lblCedenteConta
			// 
			this.lblCedenteConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteConta.Location = new System.Drawing.Point(611, 16);
			this.lblCedenteConta.Name = "lblCedenteConta";
			this.lblCedenteConta.Size = new System.Drawing.Size(77, 13);
			this.lblCedenteConta.TabIndex = 7;
			this.lblCedenteConta.Text = "22222-3";
			// 
			// lblTitCedenteConta
			// 
			this.lblTitCedenteConta.AutoSize = true;
			this.lblTitCedenteConta.Location = new System.Drawing.Point(571, 16);
			this.lblTitCedenteConta.Name = "lblTitCedenteConta";
			this.lblTitCedenteConta.Size = new System.Drawing.Size(38, 13);
			this.lblTitCedenteConta.TabIndex = 6;
			this.lblTitCedenteConta.Text = "Conta:";
			// 
			// lblCedenteAgencia
			// 
			this.lblCedenteAgencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteAgencia.Location = new System.Drawing.Point(485, 16);
			this.lblCedenteAgencia.Name = "lblCedenteAgencia";
			this.lblCedenteAgencia.Size = new System.Drawing.Size(70, 13);
			this.lblCedenteAgencia.TabIndex = 5;
			this.lblCedenteAgencia.Text = "1111-2";
			// 
			// lblTitCedenteAgencia
			// 
			this.lblTitCedenteAgencia.AutoSize = true;
			this.lblTitCedenteAgencia.Location = new System.Drawing.Point(434, 16);
			this.lblTitCedenteAgencia.Name = "lblTitCedenteAgencia";
			this.lblTitCedenteAgencia.Size = new System.Drawing.Size(49, 13);
			this.lblTitCedenteAgencia.TabIndex = 4;
			this.lblTitCedenteAgencia.Text = "Agência:";
			// 
			// lblCedenteCarteira
			// 
			this.lblCedenteCarteira.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteCarteira.Location = new System.Drawing.Point(367, 16);
			this.lblCedenteCarteira.Name = "lblCedenteCarteira";
			this.lblCedenteCarteira.Size = new System.Drawing.Size(44, 13);
			this.lblCedenteCarteira.TabIndex = 3;
			this.lblCedenteCarteira.Text = "009";
			// 
			// lblTitCedenteCarteira
			// 
			this.lblTitCedenteCarteira.AutoSize = true;
			this.lblTitCedenteCarteira.Location = new System.Drawing.Point(319, 16);
			this.lblTitCedenteCarteira.Name = "lblTitCedenteCarteira";
			this.lblTitCedenteCarteira.Size = new System.Drawing.Size(46, 13);
			this.lblTitCedenteCarteira.TabIndex = 2;
			this.lblTitCedenteCarteira.Text = "Carteira:";
			// 
			// lblCedenteNome
			// 
			this.lblCedenteNome.BackColor = System.Drawing.SystemColors.Control;
			this.lblCedenteNome.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedenteNome.Location = new System.Drawing.Point(58, 16);
			this.lblCedenteNome.Name = "lblCedenteNome";
			this.lblCedenteNome.Size = new System.Drawing.Size(239, 13);
			this.lblCedenteNome.TabIndex = 1;
			this.lblCedenteNome.Text = "Nome da Empresa Ltda";
			// 
			// lblTitCedenteNome
			// 
			this.lblTitCedenteNome.AutoSize = true;
			this.lblTitCedenteNome.Location = new System.Drawing.Point(6, 16);
			this.lblTitCedenteNome.Name = "lblTitCedenteNome";
			this.lblTitCedenteNome.Size = new System.Drawing.Size(50, 13);
			this.lblTitCedenteNome.TabIndex = 0;
			this.lblTitCedenteNome.Text = "Cedente:";
			// 
			// FBoletoArqRetorno
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoArqRetorno";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoArqRetorno_FormClosing);
			this.Load += new System.EventHandler(this.FBoletoArqRetorno_Load);
			this.Shown += new System.EventHandler(this.FBoletoArqRetorno_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxBoletos.ResumeLayout(false);
			this.gboxBoletos.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).EndInit();
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.gboxMsgErro.ResumeLayout(false);
			this.gboxCedente.ResumeLayout(false);
			this.gboxCedente.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Button btnCarregaArqRetorno;
		private System.Windows.Forms.TextBox txtArqRetorno;
		private System.Windows.Forms.Label lblTitArqRetorno;
		private System.Windows.Forms.Button btnSelecionaArqRetorno;
		private System.Windows.Forms.GroupBox gboxBoletos;
		private System.Windows.Forms.DataGridView grdBoletos;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.DataGridViewTextBoxColumn numeroDocumento;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataVenctoTitulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn valorTitulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn identificacaoOcorrencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn motivosRejeicoes;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto_item;
		private System.Windows.Forms.Label lblTotalRegistros;
		private System.Windows.Forms.Label lblTitTotalRegistros;
		private System.Windows.Forms.GroupBox gboxCedente;
		private System.Windows.Forms.Label lblCedenteDataCredito;
		private System.Windows.Forms.Label lblTitCedenteDataCredito;
		private System.Windows.Forms.Label lblCedenteDataBanco;
		private System.Windows.Forms.Label lblTitCedenteDataBanco;
		private System.Windows.Forms.Label lblCedenteNumAvisoBancario;
		private System.Windows.Forms.Label lblTitCedenteNumAvisoBancario;
		private System.Windows.Forms.Label lblCedenteCodigoEmpresa;
		private System.Windows.Forms.Label lblTitCedenteCodigoEmpresa;
		private System.Windows.Forms.Label lblCedenteConta;
		private System.Windows.Forms.Label lblTitCedenteConta;
		private System.Windows.Forms.Label lblCedenteAgencia;
		private System.Windows.Forms.Label lblTitCedenteAgencia;
		private System.Windows.Forms.Label lblCedenteCarteira;
		private System.Windows.Forms.Label lblTitCedenteCarteira;
		private System.Windows.Forms.Label lblCedenteNome;
		private System.Windows.Forms.Label lblTitCedenteNome;

	}
}
