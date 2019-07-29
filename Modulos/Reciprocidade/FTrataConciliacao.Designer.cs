namespace Reciprocidade
{
    partial class FTrataConciliacao
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FTrataConciliacao));
			this.groupBoxTitulo = new System.Windows.Forms.GroupBox();
			this.txtDataEmissao = new System.Windows.Forms.TextBox();
			this.DtpDataPagamento = new System.Windows.Forms.DateTimePicker();
			this.DtpDataVecimento = new System.Windows.Forms.DateTimePicker();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.txtNumTitulo = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btnOk = new System.Windows.Forms.Button();
			this.btnCancela = new System.Windows.Forms.Button();
			this.chkExclusaoTitulo = new System.Windows.Forms.CheckBox();
			this.btnLimpaDtPagto = new System.Windows.Forms.Button();
			this.groupBoxTitulo.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBoxTitulo
			// 
			this.groupBoxTitulo.Controls.Add(this.btnLimpaDtPagto);
			this.groupBoxTitulo.Controls.Add(this.txtDataEmissao);
			this.groupBoxTitulo.Controls.Add(this.DtpDataPagamento);
			this.groupBoxTitulo.Controls.Add(this.DtpDataVecimento);
			this.groupBoxTitulo.Controls.Add(this.txtValor);
			this.groupBoxTitulo.Controls.Add(this.txtNumTitulo);
			this.groupBoxTitulo.Controls.Add(this.label5);
			this.groupBoxTitulo.Controls.Add(this.label4);
			this.groupBoxTitulo.Controls.Add(this.label3);
			this.groupBoxTitulo.Controls.Add(this.label2);
			this.groupBoxTitulo.Controls.Add(this.label1);
			this.groupBoxTitulo.Location = new System.Drawing.Point(15, 17);
			this.groupBoxTitulo.Name = "groupBoxTitulo";
			this.groupBoxTitulo.Size = new System.Drawing.Size(389, 205);
			this.groupBoxTitulo.TabIndex = 1;
			this.groupBoxTitulo.TabStop = false;
			this.groupBoxTitulo.Text = "Dados do Título";
			// 
			// txtDataEmissao
			// 
			this.txtDataEmissao.BackColor = System.Drawing.Color.White;
			this.txtDataEmissao.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataEmissao.Location = new System.Drawing.Point(135, 65);
			this.txtDataEmissao.Name = "txtDataEmissao";
			this.txtDataEmissao.ReadOnly = true;
			this.txtDataEmissao.Size = new System.Drawing.Size(189, 22);
			this.txtDataEmissao.TabIndex = 1;
			// 
			// DtpDataPagamento
			// 
			this.DtpDataPagamento.Checked = false;
			this.DtpDataPagamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.DtpDataPagamento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.DtpDataPagamento.Location = new System.Drawing.Point(135, 167);
			this.DtpDataPagamento.Name = "DtpDataPagamento";
			this.DtpDataPagamento.ShowCheckBox = true;
			this.DtpDataPagamento.Size = new System.Drawing.Size(189, 22);
			this.DtpDataPagamento.TabIndex = 4;
			// 
			// DtpDataVecimento
			// 
			this.DtpDataVecimento.Enabled = false;
			this.DtpDataVecimento.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.DtpDataVecimento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.DtpDataVecimento.Location = new System.Drawing.Point(135, 133);
			this.DtpDataVecimento.Name = "DtpDataVecimento";
			this.DtpDataVecimento.Size = new System.Drawing.Size(189, 22);
			this.DtpDataVecimento.TabIndex = 3;
			// 
			// txtValor
			// 
			this.txtValor.BackColor = System.Drawing.Color.White;
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(135, 99);
			this.txtValor.Name = "txtValor";
			this.txtValor.ReadOnly = true;
			this.txtValor.Size = new System.Drawing.Size(189, 22);
			this.txtValor.TabIndex = 2;
			// 
			// txtNumTitulo
			// 
			this.txtNumTitulo.BackColor = System.Drawing.Color.White;
			this.txtNumTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNumTitulo.Location = new System.Drawing.Point(135, 31);
			this.txtNumTitulo.Name = "txtNumTitulo";
			this.txtNumTitulo.ReadOnly = true;
			this.txtNumTitulo.Size = new System.Drawing.Size(189, 22);
			this.txtNumTitulo.TabIndex = 0;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(30, 172);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(87, 13);
			this.label5.TabIndex = 4;
			this.label5.Text = "Data Pagamento";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(28, 138);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(89, 13);
			this.label4.TabIndex = 3;
			this.label4.Text = "Data Vencimento";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(86, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 13);
			this.label3.TabIndex = 2;
			this.label3.Text = "Valor";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(45, 70);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 13);
			this.label2.TabIndex = 1;
			this.label2.Text = "Data Emissão";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(27, 36);
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
			this.btnOk.Location = new System.Drawing.Point(103, 249);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 0;
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
			this.btnCancela.Location = new System.Drawing.Point(225, 249);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 1;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			// 
			// chkExclusaoTitulo
			// 
			this.chkExclusaoTitulo.AutoSize = true;
			this.chkExclusaoTitulo.Location = new System.Drawing.Point(324, 231);
			this.chkExclusaoTitulo.Name = "chkExclusaoTitulo";
			this.chkExclusaoTitulo.Size = new System.Drawing.Size(88, 17);
			this.chkExclusaoTitulo.TabIndex = 17;
			this.chkExclusaoTitulo.Text = "Excluir Título";
			this.chkExclusaoTitulo.UseVisualStyleBackColor = true;
			this.chkExclusaoTitulo.Visible = false;
			// 
			// btnLimpaDtPagto
			// 
			this.btnLimpaDtPagto.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpaDtPagto.Image")));
			this.btnLimpaDtPagto.Location = new System.Drawing.Point(330, 165);
			this.btnLimpaDtPagto.Name = "btnLimpaDtPagto";
			this.btnLimpaDtPagto.Size = new System.Drawing.Size(30, 26);
			this.btnLimpaDtPagto.TabIndex = 5;
			this.btnLimpaDtPagto.UseVisualStyleBackColor = true;
			this.btnLimpaDtPagto.Click += new System.EventHandler(this.btnLimpaDtPagto_Click);
			// 
			// FTrataConciliacao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(416, 305);
			this.Controls.Add(this.chkExclusaoTitulo);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.groupBoxTitulo);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FTrataConciliacao";
			this.Text = "FTrataConciliacao";
			this.Shown += new System.EventHandler(this.FTrataConciliacao_Shown);
			this.groupBoxTitulo.ResumeLayout(false);
			this.groupBoxTitulo.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxTitulo;
        private System.Windows.Forms.DateTimePicker DtpDataPagamento;
        private System.Windows.Forms.DateTimePicker DtpDataVecimento;
        private System.Windows.Forms.TextBox txtValor;
        private System.Windows.Forms.TextBox txtNumTitulo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancela;
        private System.Windows.Forms.TextBox txtDataEmissao;
		private System.Windows.Forms.CheckBox chkExclusaoTitulo;
		private System.Windows.Forms.Button btnLimpaDtPagto;
    }
}