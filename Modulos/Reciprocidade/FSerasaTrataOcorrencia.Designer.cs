namespace Reciprocidade
{
    partial class FSerasaTrataOcorrencia
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FSerasaTrataOcorrencia));
			this.groupBoxTitulo = new System.Windows.Forms.GroupBox();
			this.mskNumTitulo = new System.Windows.Forms.MaskedTextBox();
			this.DtpDataPagamento = new System.Windows.Forms.DateTimePicker();
			this.DtpDataVecimento = new System.Windows.Forms.DateTimePicker();
			this.DtpDataEmissao = new System.Windows.Forms.DateTimePicker();
			this.txtValorPago = new System.Windows.Forms.TextBox();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btnOk = new System.Windows.Forms.Button();
			this.btnCancela = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.LstErros = new System.Windows.Forms.ListBox();
			this.groupBoxTitulo.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBoxTitulo
			// 
			this.groupBoxTitulo.Controls.Add(this.mskNumTitulo);
			this.groupBoxTitulo.Controls.Add(this.DtpDataPagamento);
			this.groupBoxTitulo.Controls.Add(this.DtpDataVecimento);
			this.groupBoxTitulo.Controls.Add(this.DtpDataEmissao);
			this.groupBoxTitulo.Controls.Add(this.txtValorPago);
			this.groupBoxTitulo.Controls.Add(this.txtValor);
			this.groupBoxTitulo.Controls.Add(this.label6);
			this.groupBoxTitulo.Controls.Add(this.label5);
			this.groupBoxTitulo.Controls.Add(this.label4);
			this.groupBoxTitulo.Controls.Add(this.label3);
			this.groupBoxTitulo.Controls.Add(this.label2);
			this.groupBoxTitulo.Controls.Add(this.label1);
			this.groupBoxTitulo.Location = new System.Drawing.Point(12, 182);
			this.groupBoxTitulo.Name = "groupBoxTitulo";
			this.groupBoxTitulo.Size = new System.Drawing.Size(389, 224);
			this.groupBoxTitulo.TabIndex = 0;
			this.groupBoxTitulo.TabStop = false;
			this.groupBoxTitulo.Text = "Dados do Título";
			// 
			// mskNumTitulo
			// 
			this.mskNumTitulo.Location = new System.Drawing.Point(135, 35);
			this.mskNumTitulo.Mask = "00000000000-0";
			this.mskNumTitulo.Name = "mskNumTitulo";
			this.mskNumTitulo.ResetOnSpace = false;
			this.mskNumTitulo.Size = new System.Drawing.Size(189, 20);
			this.mskNumTitulo.TabIndex = 15;
			this.mskNumTitulo.TextMaskFormat = System.Windows.Forms.MaskFormat.ExcludePromptAndLiterals;
			// 
			// DtpDataPagamento
			// 
			this.DtpDataPagamento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.DtpDataPagamento.Location = new System.Drawing.Point(135, 153);
			this.DtpDataPagamento.Name = "DtpDataPagamento";
			this.DtpDataPagamento.ShowCheckBox = true;
			this.DtpDataPagamento.Size = new System.Drawing.Size(189, 20);
			this.DtpDataPagamento.TabIndex = 14;
			// 
			// DtpDataVecimento
			// 
			this.DtpDataVecimento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.DtpDataVecimento.Location = new System.Drawing.Point(135, 120);
			this.DtpDataVecimento.Name = "DtpDataVecimento";
			this.DtpDataVecimento.Size = new System.Drawing.Size(189, 20);
			this.DtpDataVecimento.TabIndex = 13;
			// 
			// DtpDataEmissao
			// 
			this.DtpDataEmissao.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.DtpDataEmissao.Location = new System.Drawing.Point(135, 61);
			this.DtpDataEmissao.Name = "DtpDataEmissao";
			this.DtpDataEmissao.Size = new System.Drawing.Size(189, 20);
			this.DtpDataEmissao.TabIndex = 12;
			// 
			// txtValorPago
			// 
			this.txtValorPago.Location = new System.Drawing.Point(135, 184);
			this.txtValorPago.Name = "txtValorPago";
			this.txtValorPago.Size = new System.Drawing.Size(189, 20);
			this.txtValorPago.TabIndex = 11;
			// 
			// txtValor
			// 
			this.txtValor.Location = new System.Drawing.Point(135, 91);
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(189, 20);
			this.txtValor.TabIndex = 8;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(46, 191);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(59, 13);
			this.label6.TabIndex = 5;
			this.label6.Text = "Valor Pago";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(18, 160);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(87, 13);
			this.label5.TabIndex = 4;
			this.label5.Text = "Data Pagamento";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 127);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(89, 13);
			this.label4.TabIndex = 3;
			this.label4.Text = "Data Vencimento";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(74, 98);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 13);
			this.label3.TabIndex = 2;
			this.label3.Text = "Valor";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(33, 68);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 13);
			this.label2.TabIndex = 1;
			this.label2.Text = "Data Emissão";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(15, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(90, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Número do Título";
			// 
			// btnOk
			// 
			this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(89, 439);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 12;
			this.btnOk.Text = "&Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(221, 439);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 13;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.LstErros);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(389, 145);
			this.groupBox1.TabIndex = 14;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Lista de Erros";
			// 
			// LstErros
			// 
			this.LstErros.FormattingEnabled = true;
			this.LstErros.Location = new System.Drawing.Point(3, 16);
			this.LstErros.Name = "LstErros";
			this.LstErros.Size = new System.Drawing.Size(380, 121);
			this.LstErros.TabIndex = 0;
			// 
			// FSerasaTrataOcorrencia
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(416, 491);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.groupBoxTitulo);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FSerasaTrataOcorrencia";
			this.Text = "Formulário para Tratamento de Ocorrência";
			this.Shown += new System.EventHandler(this.FSerasaTrataOcorrencia_Shown);
			this.groupBoxTitulo.ResumeLayout(false);
			this.groupBoxTitulo.PerformLayout();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxTitulo;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TextBox txtValorPago;
		private System.Windows.Forms.TextBox txtValor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker DtpDataPagamento;
        private System.Windows.Forms.DateTimePicker DtpDataVecimento;
        private System.Windows.Forms.DateTimePicker DtpDataEmissao;
        private System.Windows.Forms.Button btnCancela;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox LstErros;
		private System.Windows.Forms.MaskedTextBox mskNumTitulo;
    }
}