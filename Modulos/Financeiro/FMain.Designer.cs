namespace Financeiro
{
	partial class FMain
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FMain));
			this.btnConfig = new System.Windows.Forms.Button();
			this.btnFluxoCaixaCredito = new System.Windows.Forms.Button();
			this.btnFluxoCaixaDebito = new System.Windows.Forms.Button();
			this.btnFluxoCaixaConsulta = new System.Windows.Forms.Button();
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente = new System.Windows.Forms.Button();
			this.btnFluxoCaixaRelatorioMovimentoSintetico = new System.Windows.Forms.Button();
			this.btnBoletoCadastra = new System.Windows.Forms.Button();
			this.btnBoletoCadastraAvulsoComPedido = new System.Windows.Forms.Button();
			this.btnBoletoGeraArquivoRemessa = new System.Windows.Forms.Button();
			this.btnBoletoCarregaArquivoRetorno = new System.Windows.Forms.Button();
			this.btnBoletoCadastraAvulsoSemPedido = new System.Windows.Forms.Button();
			this.gboxFluxoCaixa = new System.Windows.Forms.GroupBox();
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo = new System.Windows.Forms.Button();
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico = new System.Windows.Forms.Button();
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico = new System.Windows.Forms.Button();
			this.btnFluxoCaixaRelatorioMovimentoAnalitico = new System.Windows.Forms.Button();
			this.btnFluxoCaixaEditaLote = new System.Windows.Forms.Button();
			this.btnFluxoCaixaCreditoLote = new System.Windows.Forms.Button();
			this.btnFluxoCaixaDebitoLote = new System.Windows.Forms.Button();
			this.gboxBoleto = new System.Windows.Forms.GroupBox();
			this.btnBoletoRelatorioArquivoRemessa = new System.Windows.Forms.Button();
			this.btnBoletoRelatoriosArquivoRetorno = new System.Windows.Forms.Button();
			this.btnBoletoConsulta = new System.Windows.Forms.Button();
			this.btnBoletoOcorrencias = new System.Windows.Forms.Button();
			this.gboxModuloCobranca = new System.Windows.Forms.GroupBox();
			this.btnModuloCobranca = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnPlanilhasPagtosMarketplace = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxFluxoCaixa.SuspendLayout();
			this.gboxBoleto.SuspendLayout();
			this.gboxModuloCobranca.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnConfig);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnConfig, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.groupBox1);
			this.pnCampos.Controls.Add(this.gboxModuloCobranca);
			this.pnCampos.Controls.Add(this.gboxBoleto);
			this.pnCampos.Controls.Add(this.gboxFluxoCaixa);
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
			// btnConfig
			// 
			this.btnConfig.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnConfig.Image = ((System.Drawing.Image)(resources.GetObject("btnConfig.Image")));
			this.btnConfig.Location = new System.Drawing.Point(879, 4);
			this.btnConfig.Name = "btnConfig";
			this.btnConfig.Size = new System.Drawing.Size(40, 44);
			this.btnConfig.TabIndex = 0;
			this.btnConfig.TabStop = false;
			this.btnConfig.UseVisualStyleBackColor = true;
			this.btnConfig.Click += new System.EventHandler(this.btnConfig_Click);
			// 
			// btnFluxoCaixaCredito
			// 
			this.btnFluxoCaixaCredito.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaCredito.ForeColor = System.Drawing.Color.Green;
			this.btnFluxoCaixaCredito.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaCredito.Image")));
			this.btnFluxoCaixaCredito.Location = new System.Drawing.Point(14, 104);
			this.btnFluxoCaixaCredito.Name = "btnFluxoCaixaCredito";
			this.btnFluxoCaixaCredito.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaCredito.TabIndex = 2;
			this.btnFluxoCaixaCredito.Text = "Fluxo de Caixa: Lançamento de Crédito";
			this.btnFluxoCaixaCredito.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaCredito.Click += new System.EventHandler(this.btnFluxoCaixaCredito_Click);
			// 
			// btnFluxoCaixaDebito
			// 
			this.btnFluxoCaixaDebito.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaDebito.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaDebito.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
			this.btnFluxoCaixaDebito.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaDebito.Image")));
			this.btnFluxoCaixaDebito.Location = new System.Drawing.Point(14, 16);
			this.btnFluxoCaixaDebito.Name = "btnFluxoCaixaDebito";
			this.btnFluxoCaixaDebito.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaDebito.TabIndex = 0;
			this.btnFluxoCaixaDebito.Text = "Fluxo de Caixa: Lançamento de Débito";
			this.btnFluxoCaixaDebito.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaDebito.Click += new System.EventHandler(this.btnFluxoCaixaDebito_Click);
			// 
			// btnFluxoCaixaConsulta
			// 
			this.btnFluxoCaixaConsulta.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaConsulta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaConsulta.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaConsulta.Image")));
			this.btnFluxoCaixaConsulta.Location = new System.Drawing.Point(14, 192);
			this.btnFluxoCaixaConsulta.Name = "btnFluxoCaixaConsulta";
			this.btnFluxoCaixaConsulta.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaConsulta.TabIndex = 4;
			this.btnFluxoCaixaConsulta.Text = "Fluxo de Caixa: Consulta";
			this.btnFluxoCaixaConsulta.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaConsulta.Click += new System.EventHandler(this.btnFluxoCaixaConsulta_Click);
			// 
			// btnFluxoCaixaRelatorioSinteticoCtaCorrente
			// 
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioSinteticoCtaCorrente.Image")));
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Location = new System.Drawing.Point(14, 280);
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Name = "btnFluxoCaixaRelatorioSinteticoCtaCorrente";
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.TabIndex = 6;
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Text = "Fluxo de Caixa: Relatório Sintético de Fluxo";
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioSinteticoCtaCorrente.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioSinteticoCtaCorrente_Click);
			// 
			// btnFluxoCaixaRelatorioMovimentoSintetico
			// 
			this.btnFluxoCaixaRelatorioMovimentoSintetico.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioMovimentoSintetico.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioMovimentoSintetico.Image")));
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Location = new System.Drawing.Point(14, 324);
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Name = "btnFluxoCaixaRelatorioMovimentoSintetico";
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Size = new System.Drawing.Size(400, 38);
			this.btnFluxoCaixaRelatorioMovimentoSintetico.TabIndex = 7;
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Text = "Fluxo de Caixa: Relatório Sintético de Movimentos";
			this.btnFluxoCaixaRelatorioMovimentoSintetico.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioMovimentoSintetico.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioMovimentoSintetico_Click);
			// 
			// btnBoletoCadastra
			// 
			this.btnBoletoCadastra.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoCadastra.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoCadastra.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoCadastra.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoCadastra.Image")));
			this.btnBoletoCadastra.Location = new System.Drawing.Point(14, 16);
			this.btnBoletoCadastra.Name = "btnBoletoCadastra";
			this.btnBoletoCadastra.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoCadastra.TabIndex = 0;
			this.btnBoletoCadastra.Text = "Boleto: Cadastramento";
			this.btnBoletoCadastra.UseVisualStyleBackColor = true;
			this.btnBoletoCadastra.Click += new System.EventHandler(this.btnBoletoCadastra_Click);
			// 
			// btnBoletoCadastraAvulsoComPedido
			// 
			this.btnBoletoCadastraAvulsoComPedido.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoCadastraAvulsoComPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoCadastraAvulsoComPedido.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoCadastraAvulsoComPedido.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoCadastraAvulsoComPedido.Image")));
			this.btnBoletoCadastraAvulsoComPedido.Location = new System.Drawing.Point(14, 60);
			this.btnBoletoCadastraAvulsoComPedido.Name = "btnBoletoCadastraAvulsoComPedido";
			this.btnBoletoCadastraAvulsoComPedido.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoCadastraAvulsoComPedido.TabIndex = 1;
			this.btnBoletoCadastraAvulsoComPedido.Text = "Boleto: Cadastramento Avulso (com pedido)";
			this.btnBoletoCadastraAvulsoComPedido.UseVisualStyleBackColor = true;
			this.btnBoletoCadastraAvulsoComPedido.Click += new System.EventHandler(this.btnBoletoCadastraAvulsoComPedido_Click);
			// 
			// btnBoletoGeraArquivoRemessa
			// 
			this.btnBoletoGeraArquivoRemessa.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoGeraArquivoRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoGeraArquivoRemessa.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoGeraArquivoRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoGeraArquivoRemessa.Image")));
			this.btnBoletoGeraArquivoRemessa.Location = new System.Drawing.Point(14, 148);
			this.btnBoletoGeraArquivoRemessa.Name = "btnBoletoGeraArquivoRemessa";
			this.btnBoletoGeraArquivoRemessa.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoGeraArquivoRemessa.TabIndex = 3;
			this.btnBoletoGeraArquivoRemessa.Text = "Boleto: Gera Arquivo de Remessa";
			this.btnBoletoGeraArquivoRemessa.UseVisualStyleBackColor = true;
			this.btnBoletoGeraArquivoRemessa.Click += new System.EventHandler(this.btnBoletoGeraArquivoRemessa_Click);
			// 
			// btnBoletoCarregaArquivoRetorno
			// 
			this.btnBoletoCarregaArquivoRetorno.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoCarregaArquivoRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoCarregaArquivoRetorno.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoCarregaArquivoRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoCarregaArquivoRetorno.Image")));
			this.btnBoletoCarregaArquivoRetorno.Location = new System.Drawing.Point(14, 236);
			this.btnBoletoCarregaArquivoRetorno.Name = "btnBoletoCarregaArquivoRetorno";
			this.btnBoletoCarregaArquivoRetorno.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoCarregaArquivoRetorno.TabIndex = 5;
			this.btnBoletoCarregaArquivoRetorno.Text = "Boleto: Carrega Arquivo de Retorno";
			this.btnBoletoCarregaArquivoRetorno.UseVisualStyleBackColor = true;
			this.btnBoletoCarregaArquivoRetorno.Click += new System.EventHandler(this.btnBoletoCarregaArquivoRetorno_Click);
			// 
			// btnBoletoCadastraAvulsoSemPedido
			// 
			this.btnBoletoCadastraAvulsoSemPedido.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoCadastraAvulsoSemPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoCadastraAvulsoSemPedido.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoCadastraAvulsoSemPedido.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoCadastraAvulsoSemPedido.Image")));
			this.btnBoletoCadastraAvulsoSemPedido.Location = new System.Drawing.Point(14, 104);
			this.btnBoletoCadastraAvulsoSemPedido.Name = "btnBoletoCadastraAvulsoSemPedido";
			this.btnBoletoCadastraAvulsoSemPedido.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoCadastraAvulsoSemPedido.TabIndex = 2;
			this.btnBoletoCadastraAvulsoSemPedido.Text = "Boleto: Cadastramento Avulso (sem pedido)";
			this.btnBoletoCadastraAvulsoSemPedido.UseVisualStyleBackColor = true;
			this.btnBoletoCadastraAvulsoSemPedido.Click += new System.EventHandler(this.btnBoletoCadastraAvulsoSemPedido_Click);
			// 
			// gboxFluxoCaixa
			// 
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioMovimentoRateioSintetico);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioMovimentoAnalitico);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaEditaLote);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaCreditoLote);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaDebitoLote);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaDebito);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaCredito);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaConsulta);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioSinteticoCtaCorrente);
			this.gboxFluxoCaixa.Controls.Add(this.btnFluxoCaixaRelatorioMovimentoSintetico);
			this.gboxFluxoCaixa.Location = new System.Drawing.Point(14, 79);
			this.gboxFluxoCaixa.Name = "gboxFluxoCaixa";
			this.gboxFluxoCaixa.Size = new System.Drawing.Size(478, 507);
			this.gboxFluxoCaixa.TabIndex = 1;
			this.gboxFluxoCaixa.TabStop = false;
			// 
			// btnFluxoCaixaRelatorioMovimentoSinteticoComparativo
			// 
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.FlatAppearance.BorderColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Image")));
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Location = new System.Drawing.Point(424, 324);
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Name = "btnFluxoCaixaRelatorioMovimentoSinteticoComparativo";
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Size = new System.Drawing.Size(40, 38);
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.TabIndex = 8;
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioMovimentoSinteticoComparativo_Click);
			// 
			// btnFluxoCaixaRelatorioMovimentoRateioAnalitico
			// 
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Image")));
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Location = new System.Drawing.Point(14, 456);
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Name = "btnFluxoCaixaRelatorioMovimentoRateioAnalitico";
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.TabIndex = 11;
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Text = "Fluxo de Caixa: Relatório Analítico de Movimentos (Rateio)";
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioMovimentoRateioAnalitico_Click);
			// 
			// btnFluxoCaixaRelatorioMovimentoRateioSintetico
			// 
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioMovimentoRateioSintetico.Image")));
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Location = new System.Drawing.Point(14, 412);
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Name = "btnFluxoCaixaRelatorioMovimentoRateioSintetico";
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.TabIndex = 10;
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Text = "Fluxo de Caixa: Relatório Sintético de Movimentos (Rateio)";
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioMovimentoRateioSintetico.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioMovimentoRateioSintetico_Click);
			// 
			// btnFluxoCaixaRelatorioMovimentoAnalitico
			// 
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaRelatorioMovimentoAnalitico.Image")));
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Location = new System.Drawing.Point(14, 368);
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Name = "btnFluxoCaixaRelatorioMovimentoAnalitico";
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.TabIndex = 9;
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Text = "Fluxo de Caixa: Relatório Analítico de Movimentos";
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaRelatorioMovimentoAnalitico.Click += new System.EventHandler(this.btnFluxoCaixaRelatorioMovimentoAnalitico_Click);
			// 
			// btnFluxoCaixaEditaLote
			// 
			this.btnFluxoCaixaEditaLote.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaEditaLote.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaEditaLote.ForeColor = System.Drawing.Color.Black;
			this.btnFluxoCaixaEditaLote.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaEditaLote.Image")));
			this.btnFluxoCaixaEditaLote.Location = new System.Drawing.Point(14, 236);
			this.btnFluxoCaixaEditaLote.Name = "btnFluxoCaixaEditaLote";
			this.btnFluxoCaixaEditaLote.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaEditaLote.TabIndex = 5;
			this.btnFluxoCaixaEditaLote.Text = "Fluxo de Caixa: Edição em Lote";
			this.btnFluxoCaixaEditaLote.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaEditaLote.Click += new System.EventHandler(this.btnFluxoCaixaEditaLote_Click);
			// 
			// btnFluxoCaixaCreditoLote
			// 
			this.btnFluxoCaixaCreditoLote.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaCreditoLote.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaCreditoLote.ForeColor = System.Drawing.Color.Green;
			this.btnFluxoCaixaCreditoLote.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaCreditoLote.Image")));
			this.btnFluxoCaixaCreditoLote.Location = new System.Drawing.Point(14, 148);
			this.btnFluxoCaixaCreditoLote.Name = "btnFluxoCaixaCreditoLote";
			this.btnFluxoCaixaCreditoLote.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaCreditoLote.TabIndex = 3;
			this.btnFluxoCaixaCreditoLote.Text = "Fluxo de Caixa: Lançamento de Crédito em Lote";
			this.btnFluxoCaixaCreditoLote.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaCreditoLote.Click += new System.EventHandler(this.btnFluxoCaixaCreditoLote_Click);
			this.btnFluxoCaixaCreditoLote.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btnFluxoCaixaCreditoLote_MouseUp);
			// 
			// btnFluxoCaixaDebitoLote
			// 
			this.btnFluxoCaixaDebitoLote.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnFluxoCaixaDebitoLote.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFluxoCaixaDebitoLote.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
			this.btnFluxoCaixaDebitoLote.Image = ((System.Drawing.Image)(resources.GetObject("btnFluxoCaixaDebitoLote.Image")));
			this.btnFluxoCaixaDebitoLote.Location = new System.Drawing.Point(14, 60);
			this.btnFluxoCaixaDebitoLote.Name = "btnFluxoCaixaDebitoLote";
			this.btnFluxoCaixaDebitoLote.Size = new System.Drawing.Size(450, 38);
			this.btnFluxoCaixaDebitoLote.TabIndex = 1;
			this.btnFluxoCaixaDebitoLote.Text = "Fluxo de Caixa: Lançamento de Débito em Lote";
			this.btnFluxoCaixaDebitoLote.UseVisualStyleBackColor = true;
			this.btnFluxoCaixaDebitoLote.Click += new System.EventHandler(this.btnFluxoCaixaDebitoLote_Click);
			this.btnFluxoCaixaDebitoLote.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btnFluxoCaixaDebitoLote_MouseUp);
			// 
			// gboxBoleto
			// 
			this.gboxBoleto.Controls.Add(this.btnBoletoRelatorioArquivoRemessa);
			this.gboxBoleto.Controls.Add(this.btnBoletoRelatoriosArquivoRetorno);
			this.gboxBoleto.Controls.Add(this.btnBoletoConsulta);
			this.gboxBoleto.Controls.Add(this.btnBoletoOcorrencias);
			this.gboxBoleto.Controls.Add(this.btnBoletoCadastra);
			this.gboxBoleto.Controls.Add(this.btnBoletoCadastraAvulsoComPedido);
			this.gboxBoleto.Controls.Add(this.btnBoletoCadastraAvulsoSemPedido);
			this.gboxBoleto.Controls.Add(this.btnBoletoGeraArquivoRemessa);
			this.gboxBoleto.Controls.Add(this.btnBoletoCarregaArquivoRetorno);
			this.gboxBoleto.Location = new System.Drawing.Point(523, 6);
			this.gboxBoleto.Name = "gboxBoleto";
			this.gboxBoleto.Size = new System.Drawing.Size(478, 419);
			this.gboxBoleto.TabIndex = 2;
			this.gboxBoleto.TabStop = false;
			// 
			// btnBoletoRelatorioArquivoRemessa
			// 
			this.btnBoletoRelatorioArquivoRemessa.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoRelatorioArquivoRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoRelatorioArquivoRemessa.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoRelatorioArquivoRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoRelatorioArquivoRemessa.Image")));
			this.btnBoletoRelatorioArquivoRemessa.Location = new System.Drawing.Point(14, 192);
			this.btnBoletoRelatorioArquivoRemessa.Name = "btnBoletoRelatorioArquivoRemessa";
			this.btnBoletoRelatorioArquivoRemessa.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoRelatorioArquivoRemessa.TabIndex = 4;
			this.btnBoletoRelatorioArquivoRemessa.Text = "Boleto: Relatório do Arquivo de Remessa";
			this.btnBoletoRelatorioArquivoRemessa.UseVisualStyleBackColor = true;
			this.btnBoletoRelatorioArquivoRemessa.Click += new System.EventHandler(this.btnBoletoRelatorioArquivoRemessa_Click);
			// 
			// btnBoletoRelatoriosArquivoRetorno
			// 
			this.btnBoletoRelatoriosArquivoRetorno.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoRelatoriosArquivoRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoRelatoriosArquivoRetorno.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoRelatoriosArquivoRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoRelatoriosArquivoRetorno.Image")));
			this.btnBoletoRelatoriosArquivoRetorno.Location = new System.Drawing.Point(14, 280);
			this.btnBoletoRelatoriosArquivoRetorno.Name = "btnBoletoRelatoriosArquivoRetorno";
			this.btnBoletoRelatoriosArquivoRetorno.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoRelatoriosArquivoRetorno.TabIndex = 6;
			this.btnBoletoRelatoriosArquivoRetorno.Text = "Boleto: Relatórios do Arquivo de Retorno";
			this.btnBoletoRelatoriosArquivoRetorno.UseVisualStyleBackColor = true;
			this.btnBoletoRelatoriosArquivoRetorno.Click += new System.EventHandler(this.btnBoletoRelatoriosArquivoRetorno_Click);
			// 
			// btnBoletoConsulta
			// 
			this.btnBoletoConsulta.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoConsulta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoConsulta.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoConsulta.Image")));
			this.btnBoletoConsulta.Location = new System.Drawing.Point(14, 324);
			this.btnBoletoConsulta.Name = "btnBoletoConsulta";
			this.btnBoletoConsulta.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoConsulta.TabIndex = 7;
			this.btnBoletoConsulta.Text = "Boleto: Consulta";
			this.btnBoletoConsulta.UseVisualStyleBackColor = true;
			this.btnBoletoConsulta.Click += new System.EventHandler(this.btnBoletoConsulta_Click);
			// 
			// btnBoletoOcorrencias
			// 
			this.btnBoletoOcorrencias.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnBoletoOcorrencias.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnBoletoOcorrencias.ForeColor = System.Drawing.Color.Black;
			this.btnBoletoOcorrencias.Image = ((System.Drawing.Image)(resources.GetObject("btnBoletoOcorrencias.Image")));
			this.btnBoletoOcorrencias.Location = new System.Drawing.Point(14, 368);
			this.btnBoletoOcorrencias.Name = "btnBoletoOcorrencias";
			this.btnBoletoOcorrencias.Size = new System.Drawing.Size(450, 38);
			this.btnBoletoOcorrencias.TabIndex = 8;
			this.btnBoletoOcorrencias.Text = "Boleto: Ocorrências";
			this.btnBoletoOcorrencias.UseVisualStyleBackColor = true;
			this.btnBoletoOcorrencias.Click += new System.EventHandler(this.btnBoletoOcorrencias_Click);
			// 
			// gboxModuloCobranca
			// 
			this.gboxModuloCobranca.Controls.Add(this.btnModuloCobranca);
			this.gboxModuloCobranca.Location = new System.Drawing.Point(14, 6);
			this.gboxModuloCobranca.Name = "gboxModuloCobranca";
			this.gboxModuloCobranca.Size = new System.Drawing.Size(478, 67);
			this.gboxModuloCobranca.TabIndex = 0;
			this.gboxModuloCobranca.TabStop = false;
			// 
			// btnModuloCobranca
			// 
			this.btnModuloCobranca.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnModuloCobranca.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnModuloCobranca.ForeColor = System.Drawing.Color.Black;
			this.btnModuloCobranca.Image = ((System.Drawing.Image)(resources.GetObject("btnModuloCobranca.Image")));
			this.btnModuloCobranca.Location = new System.Drawing.Point(14, 16);
			this.btnModuloCobranca.Name = "btnModuloCobranca";
			this.btnModuloCobranca.Size = new System.Drawing.Size(450, 38);
			this.btnModuloCobranca.TabIndex = 0;
			this.btnModuloCobranca.Text = "Módulo de Cobrança";
			this.btnModuloCobranca.UseVisualStyleBackColor = true;
			this.btnModuloCobranca.Click += new System.EventHandler(this.btnModuloCobranca_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnPlanilhasPagtosMarketplace);
			this.groupBox1.Location = new System.Drawing.Point(523, 431);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(478, 67);
			this.groupBox1.TabIndex = 3;
			this.groupBox1.TabStop = false;
			// 
			// btnPlanilhasPagtosMarketplace
			// 
			this.btnPlanilhasPagtosMarketplace.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnPlanilhasPagtosMarketplace.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnPlanilhasPagtosMarketplace.ForeColor = System.Drawing.Color.Black;
			this.btnPlanilhasPagtosMarketplace.Image = ((System.Drawing.Image)(resources.GetObject("btnPlanilhasPagtosMarketplace.Image")));
			this.btnPlanilhasPagtosMarketplace.Location = new System.Drawing.Point(14, 16);
			this.btnPlanilhasPagtosMarketplace.Name = "btnPlanilhasPagtosMarketplace";
			this.btnPlanilhasPagtosMarketplace.Size = new System.Drawing.Size(450, 38);
			this.btnPlanilhasPagtosMarketplace.TabIndex = 0;
			this.btnPlanilhasPagtosMarketplace.Text = "Planilhas de Pagamentos Marketplace";
			this.btnPlanilhasPagtosMarketplace.UseVisualStyleBackColor = true;
			this.btnPlanilhasPagtosMarketplace.Click += new System.EventHandler(this.btnPlanilhaPagtosMarketplace_Click);
			// 
			// FMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.Name = "FMain";
			this.Text = "Financeiro  -  1.00 - xx.JUL.2009";
			this.Activated += new System.EventHandler(this.FMain_Activated);
			this.Deactivate += new System.EventHandler(this.FMain_Deactivate);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.fMain_FormClosing);
			this.Shown += new System.EventHandler(this.fMain_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxFluxoCaixa.ResumeLayout(false);
			this.gboxBoleto.ResumeLayout(false);
			this.gboxModuloCobranca.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnConfig;
		private System.Windows.Forms.Button btnFluxoCaixaCredito;
		private System.Windows.Forms.Button btnFluxoCaixaDebito;
		private System.Windows.Forms.Button btnFluxoCaixaConsulta;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioMovimentoSintetico;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioSinteticoCtaCorrente;
		private System.Windows.Forms.Button btnBoletoCadastra;
		private System.Windows.Forms.Button btnBoletoGeraArquivoRemessa;
		private System.Windows.Forms.Button btnBoletoCadastraAvulsoComPedido;
		private System.Windows.Forms.Button btnBoletoCarregaArquivoRetorno;
		private System.Windows.Forms.Button btnBoletoCadastraAvulsoSemPedido;
		private System.Windows.Forms.GroupBox gboxBoleto;
		private System.Windows.Forms.GroupBox gboxFluxoCaixa;
		private System.Windows.Forms.Button btnBoletoConsulta;
		private System.Windows.Forms.Button btnBoletoOcorrencias;
		private System.Windows.Forms.Button btnBoletoRelatoriosArquivoRetorno;
		private System.Windows.Forms.Button btnBoletoRelatorioArquivoRemessa;
		private System.Windows.Forms.GroupBox gboxModuloCobranca;
		private System.Windows.Forms.Button btnModuloCobranca;
		private System.Windows.Forms.Button btnFluxoCaixaEditaLote;
		private System.Windows.Forms.Button btnFluxoCaixaCreditoLote;
		private System.Windows.Forms.Button btnFluxoCaixaDebitoLote;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioMovimentoAnalitico;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioMovimentoRateioAnalitico;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioMovimentoRateioSintetico;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnPlanilhasPagtosMarketplace;
		private System.Windows.Forms.Button btnFluxoCaixaRelatorioMovimentoSinteticoComparativo;
	}
}
