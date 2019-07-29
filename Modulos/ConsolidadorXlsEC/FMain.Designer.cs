namespace ConsolidadorXlsEC
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
            this.gboxMain = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnConferenciaPreco = new System.Windows.Forms.Button();
            this.btnAtualizarPrecosSistema = new System.Windows.Forms.Button();
            this.btnConsolidarDadosPlanilha = new System.Windows.Forms.Button();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.gboxMain);
            this.pnCampos.Size = new System.Drawing.Size(1018, 285);
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // btnFechar
            // 
            this.btnFechar.TabIndex = 1;
            // 
            // btnSobre
            // 
            this.btnSobre.TabIndex = 0;
            // 
            // gboxMain
            // 
            this.gboxMain.Controls.Add(this.button1);
            this.gboxMain.Controls.Add(this.btnConferenciaPreco);
            this.gboxMain.Controls.Add(this.btnAtualizarPrecosSistema);
            this.gboxMain.Controls.Add(this.btnConsolidarDadosPlanilha);
            this.gboxMain.Location = new System.Drawing.Point(264, 17);
            this.gboxMain.Name = "gboxMain";
            this.gboxMain.Size = new System.Drawing.Size(478, 245);
            this.gboxMain.TabIndex = 9;
            this.gboxMain.TabStop = false;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(14, 195);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(450, 38);
            this.button1.TabIndex = 3;
            this.button1.Text = "Integração Marketplace";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnIntegracaoMarketplace_Click);
            // 
            // btnConferenciaPreco
            // 
            this.btnConferenciaPreco.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnConferenciaPreco.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConferenciaPreco.ForeColor = System.Drawing.Color.Black;
            this.btnConferenciaPreco.Image = ((System.Drawing.Image)(resources.GetObject("btnConferenciaPreco.Image")));
            this.btnConferenciaPreco.Location = new System.Drawing.Point(14, 136);
            this.btnConferenciaPreco.Name = "btnConferenciaPreco";
            this.btnConferenciaPreco.Size = new System.Drawing.Size(450, 38);
            this.btnConferenciaPreco.TabIndex = 2;
            this.btnConferenciaPreco.Text = "Conferência de Preços";
            this.btnConferenciaPreco.UseVisualStyleBackColor = true;
            this.btnConferenciaPreco.Click += new System.EventHandler(this.btnConferenciaPreco_Click);
            // 
            // btnAtualizarPrecosSistema
            // 
            this.btnAtualizarPrecosSistema.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAtualizarPrecosSistema.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAtualizarPrecosSistema.ForeColor = System.Drawing.Color.Black;
            this.btnAtualizarPrecosSistema.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizarPrecosSistema.Image")));
            this.btnAtualizarPrecosSistema.Location = new System.Drawing.Point(14, 77);
            this.btnAtualizarPrecosSistema.Name = "btnAtualizarPrecosSistema";
            this.btnAtualizarPrecosSistema.Size = new System.Drawing.Size(450, 38);
            this.btnAtualizarPrecosSistema.TabIndex = 1;
            this.btnAtualizarPrecosSistema.Text = "Atualizar Preços no Sistema";
            this.btnAtualizarPrecosSistema.UseVisualStyleBackColor = true;
            this.btnAtualizarPrecosSistema.Click += new System.EventHandler(this.btnAtualizarPrecosSistema_Click);
            // 
            // btnConsolidarDadosPlanilha
            // 
            this.btnConsolidarDadosPlanilha.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnConsolidarDadosPlanilha.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConsolidarDadosPlanilha.ForeColor = System.Drawing.Color.Black;
            this.btnConsolidarDadosPlanilha.Image = ((System.Drawing.Image)(resources.GetObject("btnConsolidarDadosPlanilha.Image")));
            this.btnConsolidarDadosPlanilha.Location = new System.Drawing.Point(14, 18);
            this.btnConsolidarDadosPlanilha.Name = "btnConsolidarDadosPlanilha";
            this.btnConsolidarDadosPlanilha.Size = new System.Drawing.Size(450, 38);
            this.btnConsolidarDadosPlanilha.TabIndex = 0;
            this.btnConsolidarDadosPlanilha.Text = "Consolidar Dados da Planilha";
            this.btnConsolidarDadosPlanilha.UseVisualStyleBackColor = true;
            this.btnConsolidarDadosPlanilha.Click += new System.EventHandler(this.btnConsolidarDadosPlanilha_Click);
            // 
            // FMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1018, 382);
            this.Name = "FMain";
            this.Text = " - ConsolidadorXlsEC  -  1.04 - 27.OUT.2016";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
            this.Shown += new System.EventHandler(this.FMain_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.gboxMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxMain;
		private System.Windows.Forms.Button btnConsolidarDadosPlanilha;
		private System.Windows.Forms.Button btnAtualizarPrecosSistema;
		private System.Windows.Forms.Button btnConferenciaPreco;
        private System.Windows.Forms.Button button1;
    }
}

