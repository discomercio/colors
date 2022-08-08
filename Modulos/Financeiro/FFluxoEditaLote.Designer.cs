namespace Financeiro
{
	partial class FFluxoEditaLote
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoEditaLote));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblDataCompetencia = new System.Windows.Forms.Label();
			this.txtDataCompetencia = new System.Windows.Forms.TextBox();
			this.btnAtualizar = new System.Windows.Forms.Button();
			this.cbStSemEfeito = new System.Windows.Forms.ComboBox();
			this.lblTitStSemEfeito = new System.Windows.Forms.Label();
			this.cbCtrlPagtoStatus = new System.Windows.Forms.ComboBox();
			this.lblTitCtrlPagtoStatus = new System.Windows.Forms.Label();
			this.gboxCamposLancamento = new System.Windows.Forms.GroupBox();
			this.cbPlanoContasConta = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasConta = new System.Windows.Forms.Label();
			this.cbPlanoContasEmpresa = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasEmpresa = new System.Windows.Forms.Label();
			this.cbContaCorrente = new System.Windows.Forms.ComboBox();
			this.lblContaCorrente = new System.Windows.Forms.Label();
			this.txtComp2 = new System.Windows.Forms.TextBox();
			this.lblTitComp2 = new System.Windows.Forms.Label();
			this.cbStConfirmacaoPendente = new System.Windows.Forms.ComboBox();
			this.lblTitStConfirmacaoPendente = new System.Windows.Forms.Label();
			this.lblTitValorTotal = new System.Windows.Forms.Label();
			this.lblValorTotal = new System.Windows.Forms.Label();
			this.lblQtdeLancamentos = new System.Windows.Forms.Label();
			this.lblTitQtdeLancamentos = new System.Windows.Forms.Label();
			this.gboxLote = new System.Windows.Forms.GroupBox();
			this.grdLote = new Financeiro.DataGridViewEditavel();
			this.colNatureza = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colStSemEfeito = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colConfirmacaoPendente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colContaCorrente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colPlanoContasConta = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDataCompetencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colComp2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colValorLancto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colCnpjCpf = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDescricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colIdLancto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxCamposLancamento.SuspendLayout();
			this.gboxLote.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdLote)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnAtualizar);
			this.pnBotoes.Controls.SetChildIndex(this.btnAtualizar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.lblTitValorTotal);
			this.pnCampos.Controls.Add(this.lblValorTotal);
			this.pnCampos.Controls.Add(this.lblQtdeLancamentos);
			this.pnCampos.Controls.Add(this.lblTitQtdeLancamentos);
			this.pnCampos.Controls.Add(this.gboxLote);
			this.pnCampos.Controls.Add(this.gboxCamposLancamento);
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
			this.lblTitulo.Text = "Edição de Lançamento em Lote";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblDataCompetencia
			// 
			this.lblDataCompetencia.AutoSize = true;
			this.lblDataCompetencia.Location = new System.Drawing.Point(581, 56);
			this.lblDataCompetencia.Name = "lblDataCompetencia";
			this.lblDataCompetencia.Size = new System.Drawing.Size(95, 13);
			this.lblDataCompetencia.TabIndex = 7;
			this.lblDataCompetencia.Text = "Data Competência";
			// 
			// txtDataCompetencia
			// 
			this.txtDataCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetencia.Location = new System.Drawing.Point(682, 51);
			this.txtDataCompetencia.MaxLength = 10;
			this.txtDataCompetencia.Name = "txtDataCompetencia";
			this.txtDataCompetencia.Size = new System.Drawing.Size(99, 23);
			this.txtDataCompetencia.TabIndex = 3;
			this.txtDataCompetencia.Text = "01/01/2000";
			this.txtDataCompetencia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCompetencia.Enter += new System.EventHandler(this.txtDataCompetencia_Enter);
			this.txtDataCompetencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetencia_KeyDown);
			this.txtDataCompetencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetencia_KeyPress);
			this.txtDataCompetencia.Leave += new System.EventHandler(this.txtDataCompetencia_Leave);
			// 
			// btnAtualizar
			// 
			this.btnAtualizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnAtualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizar.Image")));
			this.btnAtualizar.Location = new System.Drawing.Point(879, 4);
			this.btnAtualizar.Name = "btnAtualizar";
			this.btnAtualizar.Size = new System.Drawing.Size(40, 44);
			this.btnAtualizar.TabIndex = 0;
			this.btnAtualizar.TabStop = false;
			this.btnAtualizar.UseVisualStyleBackColor = true;
			this.btnAtualizar.Click += new System.EventHandler(this.btnAtualizar_Click);
			// 
			// cbStSemEfeito
			// 
			this.cbStSemEfeito.BackColor = System.Drawing.SystemColors.Window;
			this.cbStSemEfeito.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbStSemEfeito.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbStSemEfeito.FormattingEnabled = true;
			this.cbStSemEfeito.Location = new System.Drawing.Point(122, 16);
			this.cbStSemEfeito.Name = "cbStSemEfeito";
			this.cbStSemEfeito.Size = new System.Drawing.Size(155, 24);
			this.cbStSemEfeito.TabIndex = 0;
			this.cbStSemEfeito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbStSemEfeito_KeyDown);
			// 
			// lblTitStSemEfeito
			// 
			this.lblTitStSemEfeito.AutoSize = true;
			this.lblTitStSemEfeito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitStSemEfeito.Location = new System.Drawing.Point(82, 21);
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
			this.cbCtrlPagtoStatus.Location = new System.Drawing.Point(122, 51);
			this.cbCtrlPagtoStatus.Name = "cbCtrlPagtoStatus";
			this.cbCtrlPagtoStatus.Size = new System.Drawing.Size(399, 24);
			this.cbCtrlPagtoStatus.TabIndex = 2;
			this.cbCtrlPagtoStatus.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbCtrlPagtoStatus_KeyDown);
			// 
			// lblTitCtrlPagtoStatus
			// 
			this.lblTitCtrlPagtoStatus.AutoSize = true;
			this.lblTitCtrlPagtoStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCtrlPagtoStatus.Location = new System.Drawing.Point(79, 56);
			this.lblTitCtrlPagtoStatus.Name = "lblTitCtrlPagtoStatus";
			this.lblTitCtrlPagtoStatus.Size = new System.Drawing.Size(37, 13);
			this.lblTitCtrlPagtoStatus.TabIndex = 37;
			this.lblTitCtrlPagtoStatus.Text = "Status";
			// 
			// gboxCamposLancamento
			// 
			this.gboxCamposLancamento.Controls.Add(this.cbPlanoContasConta);
			this.gboxCamposLancamento.Controls.Add(this.lblPlanoContasConta);
			this.gboxCamposLancamento.Controls.Add(this.cbPlanoContasEmpresa);
			this.gboxCamposLancamento.Controls.Add(this.lblPlanoContasEmpresa);
			this.gboxCamposLancamento.Controls.Add(this.cbContaCorrente);
			this.gboxCamposLancamento.Controls.Add(this.lblContaCorrente);
			this.gboxCamposLancamento.Controls.Add(this.txtComp2);
			this.gboxCamposLancamento.Controls.Add(this.lblTitComp2);
			this.gboxCamposLancamento.Controls.Add(this.cbStConfirmacaoPendente);
			this.gboxCamposLancamento.Controls.Add(this.lblTitStConfirmacaoPendente);
			this.gboxCamposLancamento.Controls.Add(this.cbCtrlPagtoStatus);
			this.gboxCamposLancamento.Controls.Add(this.lblTitCtrlPagtoStatus);
			this.gboxCamposLancamento.Controls.Add(this.cbStSemEfeito);
			this.gboxCamposLancamento.Controls.Add(this.lblTitStSemEfeito);
			this.gboxCamposLancamento.Controls.Add(this.txtDataCompetencia);
			this.gboxCamposLancamento.Controls.Add(this.lblDataCompetencia);
			this.gboxCamposLancamento.Location = new System.Drawing.Point(10, 45);
			this.gboxCamposLancamento.Name = "gboxCamposLancamento";
			this.gboxCamposLancamento.Size = new System.Drawing.Size(995, 156);
			this.gboxCamposLancamento.TabIndex = 0;
			this.gboxCamposLancamento.TabStop = false;
			// 
			// cbPlanoContasConta
			// 
			this.cbPlanoContasConta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasConta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasConta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasConta.FormattingEnabled = true;
			this.cbPlanoContasConta.Location = new System.Drawing.Point(122, 121);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(399, 24);
			this.cbPlanoContasConta.TabIndex = 7;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblPlanoContasConta
			// 
			this.lblPlanoContasConta.AutoSize = true;
			this.lblPlanoContasConta.Location = new System.Drawing.Point(36, 126);
			this.lblPlanoContasConta.Name = "lblPlanoContasConta";
			this.lblPlanoContasConta.Size = new System.Drawing.Size(80, 13);
			this.lblPlanoContasConta.TabIndex = 47;
			this.lblPlanoContasConta.Text = "Plano de Conta";
			// 
			// cbPlanoContasEmpresa
			// 
			this.cbPlanoContasEmpresa.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasEmpresa.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasEmpresa.FormattingEnabled = true;
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(635, 90);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(319, 24);
			this.cbPlanoContasEmpresa.TabIndex = 6;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(581, 95);
			this.lblPlanoContasEmpresa.Name = "lblPlanoContasEmpresa";
			this.lblPlanoContasEmpresa.Size = new System.Drawing.Size(48, 13);
			this.lblPlanoContasEmpresa.TabIndex = 45;
			this.lblPlanoContasEmpresa.Text = "Empresa";
			// 
			// cbContaCorrente
			// 
			this.cbContaCorrente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbContaCorrente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbContaCorrente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbContaCorrente.FormattingEnabled = true;
			this.cbContaCorrente.Location = new System.Drawing.Point(122, 86);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(399, 24);
			this.cbContaCorrente.TabIndex = 5;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblContaCorrente
			// 
			this.lblContaCorrente.AutoSize = true;
			this.lblContaCorrente.Location = new System.Drawing.Point(38, 91);
			this.lblContaCorrente.Name = "lblContaCorrente";
			this.lblContaCorrente.Size = new System.Drawing.Size(78, 13);
			this.lblContaCorrente.TabIndex = 43;
			this.lblContaCorrente.Text = "Conta Corrente";
			// 
			// txtComp2
			// 
			this.txtComp2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtComp2.Location = new System.Drawing.Point(863, 51);
			this.txtComp2.MaxLength = 7;
			this.txtComp2.Name = "txtComp2";
			this.txtComp2.Size = new System.Drawing.Size(91, 23);
			this.txtComp2.TabIndex = 4;
			this.txtComp2.Text = "01/2000";
			this.txtComp2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtComp2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtComp2_KeyDown);
			this.txtComp2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComp2_KeyPress);
			this.txtComp2.Leave += new System.EventHandler(this.txtComp2_Leave);
			// 
			// lblTitComp2
			// 
			this.lblTitComp2.AutoSize = true;
			this.lblTitComp2.Location = new System.Drawing.Point(817, 56);
			this.lblTitComp2.Name = "lblTitComp2";
			this.lblTitComp2.Size = new System.Drawing.Size(40, 13);
			this.lblTitComp2.TabIndex = 41;
			this.lblTitComp2.Text = "Comp2";
			// 
			// cbStConfirmacaoPendente
			// 
			this.cbStConfirmacaoPendente.BackColor = System.Drawing.SystemColors.Window;
			this.cbStConfirmacaoPendente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbStConfirmacaoPendente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbStConfirmacaoPendente.FormattingEnabled = true;
			this.cbStConfirmacaoPendente.Location = new System.Drawing.Point(787, 16);
			this.cbStConfirmacaoPendente.Name = "cbStConfirmacaoPendente";
			this.cbStConfirmacaoPendente.Size = new System.Drawing.Size(167, 24);
			this.cbStConfirmacaoPendente.TabIndex = 1;
			this.cbStConfirmacaoPendente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbStConfirmacaoPendente_KeyDown);
			// 
			// lblTitStConfirmacaoPendente
			// 
			this.lblTitStConfirmacaoPendente.AutoSize = true;
			this.lblTitStConfirmacaoPendente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitStConfirmacaoPendente.Location = new System.Drawing.Point(666, 21);
			this.lblTitStConfirmacaoPendente.Name = "lblTitStConfirmacaoPendente";
			this.lblTitStConfirmacaoPendente.Size = new System.Drawing.Size(115, 13);
			this.lblTitStConfirmacaoPendente.TabIndex = 39;
			this.lblTitStConfirmacaoPendente.Text = "Confirmação Pendente";
			// 
			// lblTitValorTotal
			// 
			this.lblTitValorTotal.AutoSize = true;
			this.lblTitValorTotal.Location = new System.Drawing.Point(339, 580);
			this.lblTitValorTotal.Name = "lblTitValorTotal";
			this.lblTitValorTotal.Size = new System.Drawing.Size(58, 13);
			this.lblTitValorTotal.TabIndex = 43;
			this.lblTitValorTotal.Text = "Valor Total";
			// 
			// lblValorTotal
			// 
			this.lblValorTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValorTotal.Location = new System.Drawing.Point(386, 580);
			this.lblValorTotal.Name = "lblValorTotal";
			this.lblValorTotal.Size = new System.Drawing.Size(112, 13);
			this.lblValorTotal.TabIndex = 44;
			this.lblValorTotal.Text = "999.999.999,00";
			this.lblValorTotal.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblQtdeLancamentos
			// 
			this.lblQtdeLancamentos.AutoSize = true;
			this.lblQtdeLancamentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeLancamentos.Location = new System.Drawing.Point(147, 580);
			this.lblQtdeLancamentos.Name = "lblQtdeLancamentos";
			this.lblQtdeLancamentos.Size = new System.Drawing.Size(21, 13);
			this.lblQtdeLancamentos.TabIndex = 42;
			this.lblQtdeLancamentos.Text = "00";
			// 
			// lblTitQtdeLancamentos
			// 
			this.lblTitQtdeLancamentos.AutoSize = true;
			this.lblTitQtdeLancamentos.Location = new System.Drawing.Point(47, 580);
			this.lblTitQtdeLancamentos.Name = "lblTitQtdeLancamentos";
			this.lblTitQtdeLancamentos.Size = new System.Drawing.Size(97, 13);
			this.lblTitQtdeLancamentos.TabIndex = 41;
			this.lblTitQtdeLancamentos.Text = "Qtde Lançamentos";
			// 
			// gboxLote
			// 
			this.gboxLote.Controls.Add(this.grdLote);
			this.gboxLote.Location = new System.Drawing.Point(10, 219);
			this.gboxLote.Name = "gboxLote";
			this.gboxLote.Size = new System.Drawing.Size(995, 355);
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
            this.colNatureza,
            this.colStSemEfeito,
            this.colConfirmacaoPendente,
            this.colContaCorrente,
            this.colPlanoContasConta,
            this.colDataCompetencia,
            this.colComp2,
            this.colValorLancto,
            this.colCnpjCpf,
            this.colDescricao,
            this.colIdLancto});
			this.grdLote.Dock = System.Windows.Forms.DockStyle.Fill;
			this.grdLote.Location = new System.Drawing.Point(3, 16);
			this.grdLote.MultiSelect = false;
			this.grdLote.Name = "grdLote";
			this.grdLote.RowHeadersWidth = 35;
			this.grdLote.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdLote.Size = new System.Drawing.Size(989, 336);
			this.grdLote.TabIndex = 0;
			this.grdLote.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdLote_CellValueChanged);
			// 
			// colNatureza
			// 
			this.colNatureza.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colNatureza.DefaultCellStyle = dataGridViewCellStyle2;
			this.colNatureza.Frozen = true;
			this.colNatureza.HeaderText = "C/D";
			this.colNatureza.MinimumWidth = 40;
			this.colNatureza.Name = "colNatureza";
			this.colNatureza.ReadOnly = true;
			this.colNatureza.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colNatureza.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colNatureza.Width = 40;
			// 
			// colStSemEfeito
			// 
			this.colStSemEfeito.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colStSemEfeito.DefaultCellStyle = dataGridViewCellStyle3;
			this.colStSemEfeito.Frozen = true;
			this.colStSemEfeito.HeaderText = "Efeito";
			this.colStSemEfeito.MinimumWidth = 65;
			this.colStSemEfeito.Name = "colStSemEfeito";
			this.colStSemEfeito.ReadOnly = true;
			this.colStSemEfeito.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colStSemEfeito.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colStSemEfeito.Width = 65;
			// 
			// colConfirmacaoPendente
			// 
			this.colConfirmacaoPendente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.colConfirmacaoPendente.DefaultCellStyle = dataGridViewCellStyle4;
			this.colConfirmacaoPendente.Frozen = true;
			this.colConfirmacaoPendente.HeaderText = "Confirmação";
			this.colConfirmacaoPendente.MinimumWidth = 85;
			this.colConfirmacaoPendente.Name = "colConfirmacaoPendente";
			this.colConfirmacaoPendente.ReadOnly = true;
			this.colConfirmacaoPendente.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colConfirmacaoPendente.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colConfirmacaoPendente.Width = 85;
			// 
			// colContaCorrente
			// 
			this.colContaCorrente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colContaCorrente.DefaultCellStyle = dataGridViewCellStyle5;
			this.colContaCorrente.Frozen = true;
			this.colContaCorrente.HeaderText = "C/C";
			this.colContaCorrente.MinimumWidth = 70;
			this.colContaCorrente.Name = "colContaCorrente";
			this.colContaCorrente.ReadOnly = true;
			this.colContaCorrente.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colContaCorrente.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colContaCorrente.Width = 70;
			// 
			// colPlanoContasConta
			// 
			this.colPlanoContasConta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colPlanoContasConta.DefaultCellStyle = dataGridViewCellStyle6;
			this.colPlanoContasConta.Frozen = true;
			this.colPlanoContasConta.HeaderText = "Plano de Conta";
			this.colPlanoContasConta.MinimumWidth = 160;
			this.colPlanoContasConta.Name = "colPlanoContasConta";
			this.colPlanoContasConta.ReadOnly = true;
			this.colPlanoContasConta.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colPlanoContasConta.Width = 160;
			// 
			// colDataCompetencia
			// 
			this.colDataCompetencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colDataCompetencia.DefaultCellStyle = dataGridViewCellStyle7;
			this.colDataCompetencia.Frozen = true;
			this.colDataCompetencia.HeaderText = "Competência";
			this.colDataCompetencia.MaxInputLength = 10;
			this.colDataCompetencia.MinimumWidth = 100;
			this.colDataCompetencia.Name = "colDataCompetencia";
			this.colDataCompetencia.ReadOnly = true;
			this.colDataCompetencia.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colDataCompetencia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// colComp2
			// 
			this.colComp2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colComp2.DefaultCellStyle = dataGridViewCellStyle8;
			this.colComp2.Frozen = true;
			this.colComp2.HeaderText = "Comp2";
			this.colComp2.MaxInputLength = 7;
			this.colComp2.MinimumWidth = 60;
			this.colComp2.Name = "colComp2";
			this.colComp2.ReadOnly = true;
			this.colComp2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colComp2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colComp2.Width = 60;
			// 
			// colValorLancto
			// 
			this.colValorLancto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colValorLancto.DefaultCellStyle = dataGridViewCellStyle9;
			this.colValorLancto.Frozen = true;
			this.colValorLancto.HeaderText = "Valor";
			this.colValorLancto.MaxInputLength = 20;
			this.colValorLancto.MinimumWidth = 100;
			this.colValorLancto.Name = "colValorLancto";
			this.colValorLancto.ReadOnly = true;
			this.colValorLancto.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colValorLancto.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// colCnpjCpf
			// 
			this.colCnpjCpf.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colCnpjCpf.DefaultCellStyle = dataGridViewCellStyle10;
			this.colCnpjCpf.Frozen = true;
			this.colCnpjCpf.HeaderText = "CNPJ/CPF";
			this.colCnpjCpf.MaxInputLength = 18;
			this.colCnpjCpf.MinimumWidth = 115;
			this.colCnpjCpf.Name = "colCnpjCpf";
			this.colCnpjCpf.ReadOnly = true;
			this.colCnpjCpf.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.colCnpjCpf.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colCnpjCpf.Width = 115;
			// 
			// colDescricao
			// 
			this.colDescricao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colDescricao.DefaultCellStyle = dataGridViewCellStyle11;
			this.colDescricao.Frozen = true;
			this.colDescricao.HeaderText = "Descrição";
			this.colDescricao.MaxInputLength = 80;
			this.colDescricao.MinimumWidth = 160;
			this.colDescricao.Name = "colDescricao";
			this.colDescricao.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.colDescricao.Width = 160;
			// 
			// colIdLancto
			// 
			this.colIdLancto.Frozen = true;
			this.colIdLancto.HeaderText = "colIdLancto";
			this.colIdLancto.Name = "colIdLancto";
			this.colIdLancto.Visible = false;
			// 
			// FFluxoEditaLote
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FFluxoEditaLote";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FFluxoEditaLote_FormClosing);
			this.Load += new System.EventHandler(this.FFluxoEditaLote_Load);
			this.Shown += new System.EventHandler(this.FFluxoEditaLote_Shown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FFluxoEditaLote_KeyPress);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxCamposLancamento.ResumeLayout(false);
			this.gboxCamposLancamento.PerformLayout();
			this.gboxLote.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdLote)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Label lblDataCompetencia;
		private System.Windows.Forms.TextBox txtDataCompetencia;
		private System.Windows.Forms.Button btnAtualizar;
		private System.Windows.Forms.ComboBox cbStSemEfeito;
		private System.Windows.Forms.Label lblTitStSemEfeito;
		private System.Windows.Forms.ComboBox cbCtrlPagtoStatus;
		private System.Windows.Forms.Label lblTitCtrlPagtoStatus;
		private System.Windows.Forms.GroupBox gboxCamposLancamento;
		private System.Windows.Forms.ComboBox cbStConfirmacaoPendente;
		private System.Windows.Forms.Label lblTitStConfirmacaoPendente;
		private System.Windows.Forms.Label lblTitValorTotal;
		private System.Windows.Forms.Label lblValorTotal;
		private System.Windows.Forms.Label lblQtdeLancamentos;
		private System.Windows.Forms.Label lblTitQtdeLancamentos;
		private System.Windows.Forms.GroupBox gboxLote;
		private DataGridViewEditavel grdLote;
        private System.Windows.Forms.TextBox txtComp2;
        private System.Windows.Forms.Label lblTitComp2;
		private System.Windows.Forms.ComboBox cbContaCorrente;
		private System.Windows.Forms.Label lblContaCorrente;
		private System.Windows.Forms.ComboBox cbPlanoContasEmpresa;
		private System.Windows.Forms.Label lblPlanoContasEmpresa;
		private System.Windows.Forms.ComboBox cbPlanoContasConta;
		private System.Windows.Forms.Label lblPlanoContasConta;
		private System.Windows.Forms.DataGridViewTextBoxColumn colNatureza;
		private System.Windows.Forms.DataGridViewTextBoxColumn colStSemEfeito;
		private System.Windows.Forms.DataGridViewTextBoxColumn colConfirmacaoPendente;
		private System.Windows.Forms.DataGridViewTextBoxColumn colContaCorrente;
		private System.Windows.Forms.DataGridViewTextBoxColumn colPlanoContasConta;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDataCompetencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn colComp2;
		private System.Windows.Forms.DataGridViewTextBoxColumn colValorLancto;
		private System.Windows.Forms.DataGridViewTextBoxColumn colCnpjCpf;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDescricao;
		private System.Windows.Forms.DataGridViewTextBoxColumn colIdLancto;
	}
}
