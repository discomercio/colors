namespace Financeiro
{
	partial class FBoletoCadastra
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
			this.components = new System.ComponentModel.Container();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoCadastra));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.pnParametros = new System.Windows.Forms.Panel();
			this.txtNumLoja = new System.Windows.Forms.TextBox();
			this.lblTitNumLoja = new System.Windows.Forms.Label();
			this.lblTitCedente = new System.Windows.Forms.Label();
			this.cbBoletoCedente = new System.Windows.Forms.ComboBox();
			this.txtNumPedido = new System.Windows.Forms.TextBox();
			this.lblTitNumPedido = new System.Windows.Forms.Label();
			this.txtNumNF = new System.Windows.Forms.TextBox();
			this.lblTitNumNF = new System.Windows.Forms.Label();
			this.txtNumParcelas = new System.Windows.Forms.TextBox();
			this.lblTitNumParcelas = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.txtNomeCliente = new System.Windows.Forms.TextBox();
			this.lblTitNomeCliente = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.txtDataEmissaoFinal = new System.Windows.Forms.TextBox();
			this.lblTitPeriodoEmissaoNF = new System.Windows.Forms.Label();
			this.txtDataEmissaoInicial = new System.Windows.Forms.TextBox();
			this.lblDataCompetenciaAte = new System.Windows.Forms.Label();
			this.pnResultado = new System.Windows.Forms.Panel();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.dt_cadastro_formatada = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.numero_NF = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.qtde_parcelas_boleto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valor_formatado = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cnpj_cpf_formatado = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.nome = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cnpj_cpf = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_cadastro = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.status = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.vl_total = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dtbNfParcelaPagtoGridBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.dsDataSource = new Financeiro.DsDataSource();
			this.pnTotalizacao = new System.Windows.Forms.Panel();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoValor = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.btnDetalhe = new System.Windows.Forms.Button();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.pnParametros.SuspendLayout();
			this.pnResultado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtbNfParcelaPagtoGridBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dsDataSource)).BeginInit();
			this.pnTotalizacao.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnLimpar);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.Add(this.btnDetalhe);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDetalhe, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnLimpar, 0);
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
			this.btnFechar.TabIndex = 4;
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 3;
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
			this.lblTitulo.Text = "Cadastramento de Boletos";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pnParametros
			// 
			this.pnParametros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnParametros.Controls.Add(this.txtNumLoja);
			this.pnParametros.Controls.Add(this.lblTitNumLoja);
			this.pnParametros.Controls.Add(this.lblTitCedente);
			this.pnParametros.Controls.Add(this.cbBoletoCedente);
			this.pnParametros.Controls.Add(this.txtNumPedido);
			this.pnParametros.Controls.Add(this.lblTitNumPedido);
			this.pnParametros.Controls.Add(this.txtNumNF);
			this.pnParametros.Controls.Add(this.lblTitNumNF);
			this.pnParametros.Controls.Add(this.txtNumParcelas);
			this.pnParametros.Controls.Add(this.lblTitNumParcelas);
			this.pnParametros.Controls.Add(this.txtCnpjCpf);
			this.pnParametros.Controls.Add(this.lblCnpjCpf);
			this.pnParametros.Controls.Add(this.txtValor);
			this.pnParametros.Controls.Add(this.txtNomeCliente);
			this.pnParametros.Controls.Add(this.lblTitNomeCliente);
			this.pnParametros.Controls.Add(this.lblValor);
			this.pnParametros.Controls.Add(this.txtDataEmissaoFinal);
			this.pnParametros.Controls.Add(this.lblTitPeriodoEmissaoNF);
			this.pnParametros.Controls.Add(this.txtDataEmissaoInicial);
			this.pnParametros.Controls.Add(this.lblDataCompetenciaAte);
			this.pnParametros.Dock = System.Windows.Forms.DockStyle.Top;
			this.pnParametros.Location = new System.Drawing.Point(0, 40);
			this.pnParametros.Name = "pnParametros";
			this.pnParametros.Size = new System.Drawing.Size(1014, 109);
			this.pnParametros.TabIndex = 2;
			// 
			// txtNumLoja
			// 
			this.txtNumLoja.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumLoja.Location = new System.Drawing.Point(686, 39);
			this.txtNumLoja.MaxLength = 3;
			this.txtNumLoja.Name = "txtNumLoja";
			this.txtNumLoja.Size = new System.Drawing.Size(91, 23);
			this.txtNumLoja.TabIndex = 6;
			this.txtNumLoja.Text = "999";
			this.txtNumLoja.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumLoja.Enter += new System.EventHandler(this.txtNumLoja_Enter);
			this.txtNumLoja.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumLoja_KeyDown);
			this.txtNumLoja.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumLoja_KeyPress);
			this.txtNumLoja.Leave += new System.EventHandler(this.txtNumLoja_Leave);
			// 
			// lblTitNumLoja
			// 
			this.lblTitNumLoja.AutoSize = true;
			this.lblTitNumLoja.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNumLoja.Location = new System.Drawing.Point(631, 44);
			this.lblTitNumLoja.Name = "lblTitNumLoja";
			this.lblTitNumLoja.Size = new System.Drawing.Size(49, 13);
			this.lblTitNumLoja.TabIndex = 47;
			this.lblTitNumLoja.Text = "Nº Loja";
			// 
			// lblTitCedente
			// 
			this.lblTitCedente.AutoSize = true;
			this.lblTitCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitCedente.Location = new System.Drawing.Point(44, 44);
			this.lblTitCedente.Name = "lblTitCedente";
			this.lblTitCedente.Size = new System.Drawing.Size(54, 13);
			this.lblTitCedente.TabIndex = 45;
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
			this.cbBoletoCedente.Location = new System.Drawing.Point(104, 41);
			this.cbBoletoCedente.MaxDropDownItems = 12;
			this.cbBoletoCedente.Name = "cbBoletoCedente";
			this.cbBoletoCedente.Size = new System.Drawing.Size(452, 21);
			this.cbBoletoCedente.TabIndex = 5;
			this.cbBoletoCedente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbBoletoCedente_KeyDown);
			// 
			// txtNumPedido
			// 
			this.txtNumPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumPedido.Location = new System.Drawing.Point(436, 6);
			this.txtNumPedido.MaxLength = 9;
			this.txtNumPedido.Name = "txtNumPedido";
			this.txtNumPedido.Size = new System.Drawing.Size(120, 23);
			this.txtNumPedido.TabIndex = 2;
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
			this.lblTitNumPedido.Location = new System.Drawing.Point(366, 11);
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
			// txtNumParcelas
			// 
			this.txtNumParcelas.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumParcelas.Location = new System.Drawing.Point(686, 6);
			this.txtNumParcelas.MaxLength = 2;
			this.txtNumParcelas.Name = "txtNumParcelas";
			this.txtNumParcelas.Size = new System.Drawing.Size(91, 23);
			this.txtNumParcelas.TabIndex = 3;
			this.txtNumParcelas.Text = "99";
			this.txtNumParcelas.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumParcelas.Enter += new System.EventHandler(this.txtNumParcelas_Enter);
			this.txtNumParcelas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumParcelas_KeyDown);
			this.txtNumParcelas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumParcelas_KeyPress);
			this.txtNumParcelas.Leave += new System.EventHandler(this.txtNumParcelas_Leave);
			// 
			// lblTitNumParcelas
			// 
			this.lblTitNumParcelas.AutoSize = true;
			this.lblTitNumParcelas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNumParcelas.Location = new System.Drawing.Point(606, 11);
			this.lblTitNumParcelas.Name = "lblTitNumParcelas";
			this.lblTitNumParcelas.Size = new System.Drawing.Size(74, 13);
			this.lblTitNumParcelas.TabIndex = 33;
			this.lblTitNumParcelas.Text = "Nº Parcelas";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(860, 72);
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
			// lblCnpjCpf
			// 
			this.lblCnpjCpf.AutoSize = true;
			this.lblCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCnpjCpf.Location = new System.Drawing.Point(787, 77);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(67, 13);
			this.lblCnpjCpf.TabIndex = 31;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(883, 39);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(122, 23);
			this.txtValor.TabIndex = 7;
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
			this.txtNomeCliente.Location = new System.Drawing.Point(104, 72);
			this.txtNomeCliente.MaxLength = 40;
			this.txtNomeCliente.Name = "txtNomeCliente";
			this.txtNomeCliente.Size = new System.Drawing.Size(452, 23);
			this.txtNomeCliente.TabIndex = 8;
			this.txtNomeCliente.Enter += new System.EventHandler(this.txtNomeCliente_Enter);
			this.txtNomeCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNomeCliente_KeyDown);
			this.txtNomeCliente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNomeCliente_KeyPress);
			this.txtNomeCliente.Leave += new System.EventHandler(this.txtNomeCliente_Leave);
			// 
			// lblTitNomeCliente
			// 
			this.lblTitNomeCliente.AutoSize = true;
			this.lblTitNomeCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitNomeCliente.Location = new System.Drawing.Point(16, 77);
			this.lblTitNomeCliente.Name = "lblTitNomeCliente";
			this.lblTitNomeCliente.Size = new System.Drawing.Size(82, 13);
			this.lblTitNomeCliente.TabIndex = 29;
			this.lblTitNomeCliente.Text = "Nome Cliente";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblValor.Location = new System.Drawing.Point(813, 44);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(64, 13);
			this.lblValor.TabIndex = 28;
			this.lblValor.Text = "Valor (R$)";
			// 
			// txtDataEmissaoFinal
			// 
			this.txtDataEmissaoFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataEmissaoFinal.Location = new System.Drawing.Point(221, 6);
			this.txtDataEmissaoFinal.MaxLength = 10;
			this.txtDataEmissaoFinal.Name = "txtDataEmissaoFinal";
			this.txtDataEmissaoFinal.Size = new System.Drawing.Size(91, 23);
			this.txtDataEmissaoFinal.TabIndex = 1;
			this.txtDataEmissaoFinal.Text = "01/01/2000";
			this.txtDataEmissaoFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataEmissaoFinal.Enter += new System.EventHandler(this.txtDataEmissaoFinal_Enter);
			this.txtDataEmissaoFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataEmissaoFinal_KeyDown);
			this.txtDataEmissaoFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataEmissaoFinal_KeyPress);
			this.txtDataEmissaoFinal.Leave += new System.EventHandler(this.txtDataEmissaoFinal_Leave);
			// 
			// lblTitPeriodoEmissaoNF
			// 
			this.lblTitPeriodoEmissaoNF.AutoSize = true;
			this.lblTitPeriodoEmissaoNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitPeriodoEmissaoNF.Location = new System.Drawing.Point(25, 11);
			this.lblTitPeriodoEmissaoNF.Name = "lblTitPeriodoEmissaoNF";
			this.lblTitPeriodoEmissaoNF.Size = new System.Drawing.Size(73, 13);
			this.lblTitPeriodoEmissaoNF.TabIndex = 10;
			this.lblTitPeriodoEmissaoNF.Text = "Emissão NF";
			// 
			// txtDataEmissaoInicial
			// 
			this.txtDataEmissaoInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataEmissaoInicial.Location = new System.Drawing.Point(104, 6);
			this.txtDataEmissaoInicial.MaxLength = 10;
			this.txtDataEmissaoInicial.Name = "txtDataEmissaoInicial";
			this.txtDataEmissaoInicial.Size = new System.Drawing.Size(91, 23);
			this.txtDataEmissaoInicial.TabIndex = 0;
			this.txtDataEmissaoInicial.Text = "01/01/2000";
			this.txtDataEmissaoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataEmissaoInicial.Enter += new System.EventHandler(this.txtDataEmissaoInicial_Enter);
			this.txtDataEmissaoInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataEmissaoInicial_KeyDown);
			this.txtDataEmissaoInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataEmissaoInicial_KeyPress);
			this.txtDataEmissaoInicial.Leave += new System.EventHandler(this.txtDataEmissaoInicial_Leave);
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
			this.pnResultado.Location = new System.Drawing.Point(0, 149);
			this.pnResultado.Name = "pnResultado";
			this.pnResultado.Size = new System.Drawing.Size(1014, 435);
			this.pnResultado.TabIndex = 4;
			// 
			// gridDados
			// 
			this.gridDados.AllowUserToAddRows = false;
			this.gridDados.AllowUserToDeleteRows = false;
			this.gridDados.AutoGenerateColumns = false;
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
            this.dt_cadastro_formatada,
            this.numero_NF,
            this.qtde_parcelas_boleto,
            this.valor_formatado,
            this.pedido,
            this.cnpj_cpf_formatado,
            this.nome,
            this.id,
            this.cnpj_cpf,
            this.dt_cadastro,
            this.status,
            this.vl_total,
            this.id_cliente});
			this.gridDados.DataSource = this.dtbNfParcelaPagtoGridBindingSource;
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
			this.gridDados.Size = new System.Drawing.Size(1010, 431);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 0;
			this.gridDados.DoubleClick += new System.EventHandler(this.gridDados_DoubleClick);
			this.gridDados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridDados_KeyDown);
			// 
			// dt_cadastro_formatada
			// 
			this.dt_cadastro_formatada.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.dt_cadastro_formatada.DataPropertyName = "dt_cadastro_formatada";
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.dt_cadastro_formatada.DefaultCellStyle = dataGridViewCellStyle2;
			this.dt_cadastro_formatada.HeaderText = "Data";
			this.dt_cadastro_formatada.MinimumWidth = 80;
			this.dt_cadastro_formatada.Name = "dt_cadastro_formatada";
			this.dt_cadastro_formatada.ReadOnly = true;
			this.dt_cadastro_formatada.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.dt_cadastro_formatada.Width = 80;
			// 
			// numero_NF
			// 
			this.numero_NF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.numero_NF.DataPropertyName = "numero_NF";
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.numero_NF.DefaultCellStyle = dataGridViewCellStyle3;
			this.numero_NF.HeaderText = "NF";
			this.numero_NF.MinimumWidth = 80;
			this.numero_NF.Name = "numero_NF";
			this.numero_NF.ReadOnly = true;
			this.numero_NF.Width = 80;
			// 
			// qtde_parcelas_boleto
			// 
			this.qtde_parcelas_boleto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.qtde_parcelas_boleto.DataPropertyName = "qtde_parcelas_boleto";
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.qtde_parcelas_boleto.DefaultCellStyle = dataGridViewCellStyle4;
			this.qtde_parcelas_boleto.HeaderText = "Parcelas";
			this.qtde_parcelas_boleto.MinimumWidth = 85;
			this.qtde_parcelas_boleto.Name = "qtde_parcelas_boleto";
			this.qtde_parcelas_boleto.ReadOnly = true;
			this.qtde_parcelas_boleto.Width = 85;
			// 
			// valor_formatado
			// 
			this.valor_formatado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.valor_formatado.DataPropertyName = "valor_formatado";
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.valor_formatado.DefaultCellStyle = dataGridViewCellStyle5;
			this.valor_formatado.FillWeight = 331.0345F;
			this.valor_formatado.HeaderText = "Valor Total";
			this.valor_formatado.MinimumWidth = 120;
			this.valor_formatado.Name = "valor_formatado";
			this.valor_formatado.ReadOnly = true;
			this.valor_formatado.Width = 120;
			// 
			// pedido
			// 
			this.pedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.pedido.DataPropertyName = "pedido";
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.pedido.DefaultCellStyle = dataGridViewCellStyle6;
			this.pedido.HeaderText = "Pedido";
			this.pedido.MinimumWidth = 115;
			this.pedido.Name = "pedido";
			this.pedido.ReadOnly = true;
			this.pedido.Width = 115;
			// 
			// cnpj_cpf_formatado
			// 
			this.cnpj_cpf_formatado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.cnpj_cpf_formatado.DataPropertyName = "cnpj_cpf_formatado";
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.cnpj_cpf_formatado.DefaultCellStyle = dataGridViewCellStyle7;
			this.cnpj_cpf_formatado.HeaderText = "CNPJ / CPF";
			this.cnpj_cpf_formatado.MinimumWidth = 130;
			this.cnpj_cpf_formatado.Name = "cnpj_cpf_formatado";
			this.cnpj_cpf_formatado.ReadOnly = true;
			this.cnpj_cpf_formatado.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.cnpj_cpf_formatado.Width = 130;
			// 
			// nome
			// 
			this.nome.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.nome.DataPropertyName = "nome";
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.nome.DefaultCellStyle = dataGridViewCellStyle8;
			this.nome.HeaderText = "Cliente";
			this.nome.MinimumWidth = 180;
			this.nome.Name = "nome";
			this.nome.ReadOnly = true;
			// 
			// id
			// 
			this.id.DataPropertyName = "id";
			this.id.HeaderText = "id";
			this.id.Name = "id";
			this.id.ReadOnly = true;
			this.id.Visible = false;
			// 
			// cnpj_cpf
			// 
			this.cnpj_cpf.DataPropertyName = "cnpj_cpf";
			this.cnpj_cpf.HeaderText = "cnpj_cpf";
			this.cnpj_cpf.Name = "cnpj_cpf";
			this.cnpj_cpf.ReadOnly = true;
			this.cnpj_cpf.Visible = false;
			// 
			// dt_cadastro
			// 
			this.dt_cadastro.DataPropertyName = "dt_cadastro";
			this.dt_cadastro.HeaderText = "dt_cadastro";
			this.dt_cadastro.Name = "dt_cadastro";
			this.dt_cadastro.ReadOnly = true;
			this.dt_cadastro.Visible = false;
			// 
			// status
			// 
			this.status.DataPropertyName = "status";
			this.status.HeaderText = "status";
			this.status.Name = "status";
			this.status.ReadOnly = true;
			this.status.Visible = false;
			// 
			// vl_total
			// 
			this.vl_total.DataPropertyName = "vl_total";
			this.vl_total.HeaderText = "vl_total";
			this.vl_total.Name = "vl_total";
			this.vl_total.ReadOnly = true;
			this.vl_total.Visible = false;
			// 
			// id_cliente
			// 
			this.id_cliente.DataPropertyName = "id_cliente";
			this.id_cliente.HeaderText = "id_cliente";
			this.id_cliente.Name = "id_cliente";
			this.id_cliente.ReadOnly = true;
			this.id_cliente.Visible = false;
			// 
			// dtbNfParcelaPagtoGridBindingSource
			// 
			this.dtbNfParcelaPagtoGridBindingSource.DataMember = "DtbNfParcelaPagtoGrid";
			this.dtbNfParcelaPagtoGridBindingSource.DataSource = this.dsDataSource;
			// 
			// dsDataSource
			// 
			this.dsDataSource.DataSetName = "DsDataSource";
			this.dsDataSource.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// pnTotalizacao
			// 
			this.pnTotalizacao.Controls.Add(this.lblTitTotalizacaoRegistros);
			this.pnTotalizacao.Controls.Add(this.lblTitTotalizacaoValor);
			this.pnTotalizacao.Controls.Add(this.lblTotalizacaoValor);
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
			// lblTitTotalizacaoValor
			// 
			this.lblTitTotalizacaoValor.AutoSize = true;
			this.lblTitTotalizacaoValor.Location = new System.Drawing.Point(822, 4);
			this.lblTitTotalizacaoValor.Name = "lblTitTotalizacaoValor";
			this.lblTitTotalizacaoValor.Size = new System.Drawing.Size(61, 13);
			this.lblTitTotalizacaoValor.TabIndex = 4;
			this.lblTitTotalizacaoValor.Text = "Valor Total:";
			// 
			// lblTotalizacaoValor
			// 
			this.lblTotalizacaoValor.AutoSize = true;
			this.lblTotalizacaoValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoValor.Location = new System.Drawing.Point(889, 4);
			this.lblTotalizacaoValor.Name = "lblTotalizacaoValor";
			this.lblTotalizacaoValor.Size = new System.Drawing.Size(96, 13);
			this.lblTotalizacaoValor.TabIndex = 7;
			this.lblTotalizacaoValor.Text = "999.999.999,99";
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
			// btnDetalhe
			// 
			this.btnDetalhe.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnDetalhe.Image = ((System.Drawing.Image)(resources.GetObject("btnDetalhe.Image")));
			this.btnDetalhe.Location = new System.Drawing.Point(879, 4);
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
			this.btnPesquisar.Location = new System.Drawing.Point(789, 4);
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
			this.btnLimpar.Location = new System.Drawing.Point(834, 4);
			this.btnLimpar.Name = "btnLimpar";
			this.btnLimpar.Size = new System.Drawing.Size(40, 44);
			this.btnLimpar.TabIndex = 1;
			this.btnLimpar.TabStop = false;
			this.btnLimpar.UseVisualStyleBackColor = true;
			this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
			// 
			// FBoletoCadastra
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoCadastra";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoCadastra_FormClosing);
			this.Load += new System.EventHandler(this.FBoletoCadastra_Load);
			this.Shown += new System.EventHandler(this.FBoletoCadastra_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FBoletoCadastra_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnParametros.ResumeLayout(false);
			this.pnParametros.PerformLayout();
			this.pnResultado.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtbNfParcelaPagtoGridBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dsDataSource)).EndInit();
			this.pnTotalizacao.ResumeLayout(false);
			this.pnTotalizacao.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Panel pnParametros;
		private System.Windows.Forms.TextBox txtDataEmissaoInicial;
		private System.Windows.Forms.Label lblDataCompetenciaAte;
		private System.Windows.Forms.Label lblTitPeriodoEmissaoNF;
		private System.Windows.Forms.TextBox txtDataEmissaoFinal;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.TextBox txtNomeCliente;
		private System.Windows.Forms.Label lblTitNomeCliente;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private DsDataSource dsDataSource;
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
		private System.Windows.Forms.Label lblTitNumParcelas;
		private System.Windows.Forms.TextBox txtNumParcelas;
		private System.Windows.Forms.TextBox txtNumNF;
		private System.Windows.Forms.Label lblTitNumNF;
		private System.Windows.Forms.TextBox txtNumPedido;
		private System.Windows.Forms.Label lblTitNumPedido;
		private System.Windows.Forms.BindingSource dtbNfParcelaPagtoGridBindingSource;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_cadastro_formatada;
		private System.Windows.Forms.DataGridViewTextBoxColumn numero_NF;
		private System.Windows.Forms.DataGridViewTextBoxColumn qtde_parcelas_boleto;
		private System.Windows.Forms.DataGridViewTextBoxColumn valor_formatado;
		private System.Windows.Forms.DataGridViewTextBoxColumn pedido;
		private System.Windows.Forms.DataGridViewTextBoxColumn cnpj_cpf_formatado;
		private System.Windows.Forms.DataGridViewTextBoxColumn nome;
		private System.Windows.Forms.DataGridViewTextBoxColumn id;
		private System.Windows.Forms.DataGridViewTextBoxColumn cnpj_cpf;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_cadastro;
		private System.Windows.Forms.DataGridViewTextBoxColumn status;
		private System.Windows.Forms.DataGridViewTextBoxColumn vl_total;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_cliente;
		private System.Windows.Forms.ComboBox cbBoletoCedente;
		private System.Windows.Forms.Label lblTitCedente;
		private System.Windows.Forms.TextBox txtNumLoja;
		private System.Windows.Forms.Label lblTitNumLoja;
	}
}
