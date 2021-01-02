namespace ADM2
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
            this.gboxIBPT = new System.Windows.Forms.GroupBox();
            this.btnAtualizarDataImportArquivosXML = new System.Windows.Forms.Button();
            this.btnAnotarPedidosRecebidosCliente = new System.Windows.Forms.Button();
            this.btnAtualizarPlanilhaEstoque = new System.Windows.Forms.Button();
            this.btnIbptCarregaArqCsv = new System.Windows.Forms.Button();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxIBPT.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.gboxIBPT);
            this.pnCampos.Size = new System.Drawing.Size(1008, 343);
            // 
            // gboxIBPT
            // 
            this.gboxIBPT.Controls.Add(this.btnAtualizarDataImportArquivosXML);
            this.gboxIBPT.Controls.Add(this.btnAnotarPedidosRecebidosCliente);
            this.gboxIBPT.Controls.Add(this.btnAtualizarPlanilhaEstoque);
            this.gboxIBPT.Controls.Add(this.btnIbptCarregaArqCsv);
            this.gboxIBPT.Location = new System.Drawing.Point(263, 103);
            this.gboxIBPT.Name = "gboxIBPT";
            this.gboxIBPT.Size = new System.Drawing.Size(478, 242);
            this.gboxIBPT.TabIndex = 8;
            this.gboxIBPT.TabStop = false;
            // 
            // btnAtualizarDataImportArquivosXML
            // 
            this.btnAtualizarDataImportArquivosXML.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAtualizarDataImportArquivosXML.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAtualizarDataImportArquivosXML.ForeColor = System.Drawing.Color.Black;
            this.btnAtualizarDataImportArquivosXML.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizarDataImportArquivosXML.Image")));
            this.btnAtualizarDataImportArquivosXML.Location = new System.Drawing.Point(14, 193);
            this.btnAtualizarDataImportArquivosXML.Name = "btnAtualizarDataImportArquivosXML";
            this.btnAtualizarDataImportArquivosXML.Size = new System.Drawing.Size(450, 38);
            this.btnAtualizarDataImportArquivosXML.TabIndex = 3;
            this.btnAtualizarDataImportArquivosXML.Text = "Atualizar Datas de Importação de Arquivos XML";
            this.btnAtualizarDataImportArquivosXML.UseVisualStyleBackColor = true;
            this.btnAtualizarDataImportArquivosXML.Click += new System.EventHandler(this.BtnAtualizarDataImportArquivosXML_Click);
            // 
            // btnAnotarPedidosRecebidosCliente
            // 
            this.btnAnotarPedidosRecebidosCliente.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAnotarPedidosRecebidosCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnotarPedidosRecebidosCliente.ForeColor = System.Drawing.Color.Black;
            this.btnAnotarPedidosRecebidosCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnAnotarPedidosRecebidosCliente.Image")));
            this.btnAnotarPedidosRecebidosCliente.Location = new System.Drawing.Point(14, 134);
            this.btnAnotarPedidosRecebidosCliente.Name = "btnAnotarPedidosRecebidosCliente";
            this.btnAnotarPedidosRecebidosCliente.Size = new System.Drawing.Size(450, 38);
            this.btnAnotarPedidosRecebidosCliente.TabIndex = 2;
            this.btnAnotarPedidosRecebidosCliente.Text = "Anotar Pedidos Recebidos pelo Cliente";
            this.btnAnotarPedidosRecebidosCliente.UseVisualStyleBackColor = true;
            this.btnAnotarPedidosRecebidosCliente.Click += new System.EventHandler(this.btnAnotarPedidosRecebidosCliente_Click);
            // 
            // btnAtualizarPlanilhaEstoque
            // 
            this.btnAtualizarPlanilhaEstoque.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAtualizarPlanilhaEstoque.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAtualizarPlanilhaEstoque.ForeColor = System.Drawing.Color.Black;
            this.btnAtualizarPlanilhaEstoque.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizarPlanilhaEstoque.Image")));
            this.btnAtualizarPlanilhaEstoque.Location = new System.Drawing.Point(14, 75);
            this.btnAtualizarPlanilhaEstoque.Name = "btnAtualizarPlanilhaEstoque";
            this.btnAtualizarPlanilhaEstoque.Size = new System.Drawing.Size(450, 38);
            this.btnAtualizarPlanilhaEstoque.TabIndex = 1;
            this.btnAtualizarPlanilhaEstoque.Text = "Atualizar Planilha do Estoque";
            this.btnAtualizarPlanilhaEstoque.UseVisualStyleBackColor = true;
            this.btnAtualizarPlanilhaEstoque.Click += new System.EventHandler(this.btnAtualizarPlanilhaEstoque_Click);
            // 
            // btnIbptCarregaArqCsv
            // 
            this.btnIbptCarregaArqCsv.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnIbptCarregaArqCsv.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIbptCarregaArqCsv.ForeColor = System.Drawing.Color.Black;
            this.btnIbptCarregaArqCsv.Image = ((System.Drawing.Image)(resources.GetObject("btnIbptCarregaArqCsv.Image")));
            this.btnIbptCarregaArqCsv.Location = new System.Drawing.Point(14, 16);
            this.btnIbptCarregaArqCsv.Name = "btnIbptCarregaArqCsv";
            this.btnIbptCarregaArqCsv.Size = new System.Drawing.Size(450, 38);
            this.btnIbptCarregaArqCsv.TabIndex = 0;
            this.btnIbptCarregaArqCsv.Text = "Carrega Arquivo IBPT";
            this.btnIbptCarregaArqCsv.UseVisualStyleBackColor = true;
            this.btnIbptCarregaArqCsv.Click += new System.EventHandler(this.btnIbptCarregaArqCsv_Click);
            // 
            // FMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1008, 440);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FMain";
            this.Text = "ADM2  -  1.00 - 01.JUN.2013";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
            this.Shown += new System.EventHandler(this.FMain_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.gboxIBPT.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxIBPT;
		private System.Windows.Forms.Button btnIbptCarregaArqCsv;
		private System.Windows.Forms.Button btnAtualizarPlanilhaEstoque;
		private System.Windows.Forms.Button btnAnotarPedidosRecebidosCliente;
        private System.Windows.Forms.Button btnAtualizarDataImportArquivosXML;
    }
}
