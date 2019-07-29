namespace Financeiro
{
	partial class FBoletoTrataOcorrenciaValaComum
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoTrataOcorrenciaValaComum));
			this.gboxCliente = new System.Windows.Forms.GroupBox();
			this.txtClienteCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblTitClienteCnpjCpf = new System.Windows.Forms.Label();
			this.txtClienteNome = new System.Windows.Forms.TextBox();
			this.lblTitClienteNome = new System.Windows.Forms.Label();
			this.gboxDadosRegistro = new System.Windows.Forms.GroupBox();
			this.txtDadosRegistro = new System.Windows.Forms.TextBox();
			this.gboxIdentificacaoOcorrencia = new System.Windows.Forms.GroupBox();
			this.txtIdentificacaoOcorrencia = new System.Windows.Forms.TextBox();
			this.gboxComentarioOcorrenciaTratada = new System.Windows.Forms.GroupBox();
			this.txtComentarioOcorrenciaTratada = new System.Windows.Forms.TextBox();
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnOk = new System.Windows.Forms.Button();
			this.gboxCedente = new System.Windows.Forms.GroupBox();
			this.lblCedente = new System.Windows.Forms.Label();
			this.gboxCliente.SuspendLayout();
			this.gboxDadosRegistro.SuspendLayout();
			this.gboxIdentificacaoOcorrencia.SuspendLayout();
			this.gboxComentarioOcorrenciaTratada.SuspendLayout();
			this.gboxCedente.SuspendLayout();
			this.SuspendLayout();
			// 
			// gboxCliente
			// 
			this.gboxCliente.Controls.Add(this.txtClienteCnpjCpf);
			this.gboxCliente.Controls.Add(this.lblTitClienteCnpjCpf);
			this.gboxCliente.Controls.Add(this.txtClienteNome);
			this.gboxCliente.Controls.Add(this.lblTitClienteNome);
			this.gboxCliente.Location = new System.Drawing.Point(12, 58);
			this.gboxCliente.Name = "gboxCliente";
			this.gboxCliente.Size = new System.Drawing.Size(758, 78);
			this.gboxCliente.TabIndex = 0;
			this.gboxCliente.TabStop = false;
			this.gboxCliente.Text = "Cliente";
			// 
			// txtClienteCnpjCpf
			// 
			this.txtClienteCnpjCpf.BackColor = System.Drawing.SystemColors.Window;
			this.txtClienteCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtClienteCnpjCpf.Location = new System.Drawing.Point(83, 48);
			this.txtClienteCnpjCpf.MaxLength = 18;
			this.txtClienteCnpjCpf.Name = "txtClienteCnpjCpf";
			this.txtClienteCnpjCpf.ReadOnly = true;
			this.txtClienteCnpjCpf.Size = new System.Drawing.Size(146, 20);
			this.txtClienteCnpjCpf.TabIndex = 1;
			this.txtClienteCnpjCpf.TabStop = false;
			this.txtClienteCnpjCpf.Text = "00.000.000/0000-00";
			this.txtClienteCnpjCpf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			// 
			// lblTitClienteCnpjCpf
			// 
			this.lblTitClienteCnpjCpf.AutoSize = true;
			this.lblTitClienteCnpjCpf.Location = new System.Drawing.Point(12, 51);
			this.lblTitClienteCnpjCpf.Name = "lblTitClienteCnpjCpf";
			this.lblTitClienteCnpjCpf.Size = new System.Drawing.Size(65, 13);
			this.lblTitClienteCnpjCpf.TabIndex = 24;
			this.lblTitClienteCnpjCpf.Text = "CNPJ / CPF";
			// 
			// txtClienteNome
			// 
			this.txtClienteNome.BackColor = System.Drawing.SystemColors.Window;
			this.txtClienteNome.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtClienteNome.ForeColor = System.Drawing.SystemColors.WindowText;
			this.txtClienteNome.Location = new System.Drawing.Point(83, 21);
			this.txtClienteNome.MaxLength = 0;
			this.txtClienteNome.Name = "txtClienteNome";
			this.txtClienteNome.ReadOnly = true;
			this.txtClienteNome.Size = new System.Drawing.Size(659, 20);
			this.txtClienteNome.TabIndex = 0;
			this.txtClienteNome.TabStop = false;
			this.txtClienteNome.Text = "FULANO BELTRANO CICLANO DA SILVA";
			// 
			// lblTitClienteNome
			// 
			this.lblTitClienteNome.AutoSize = true;
			this.lblTitClienteNome.Location = new System.Drawing.Point(42, 24);
			this.lblTitClienteNome.Name = "lblTitClienteNome";
			this.lblTitClienteNome.Size = new System.Drawing.Size(35, 13);
			this.lblTitClienteNome.TabIndex = 22;
			this.lblTitClienteNome.Text = "Nome";
			// 
			// gboxDadosRegistro
			// 
			this.gboxDadosRegistro.Controls.Add(this.txtDadosRegistro);
			this.gboxDadosRegistro.Location = new System.Drawing.Point(12, 242);
			this.gboxDadosRegistro.Name = "gboxDadosRegistro";
			this.gboxDadosRegistro.Size = new System.Drawing.Size(758, 194);
			this.gboxDadosRegistro.TabIndex = 2;
			this.gboxDadosRegistro.TabStop = false;
			this.gboxDadosRegistro.Text = "Dados do registro";
			// 
			// txtDadosRegistro
			// 
			this.txtDadosRegistro.BackColor = System.Drawing.SystemColors.Window;
			this.txtDadosRegistro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDadosRegistro.Location = new System.Drawing.Point(15, 19);
			this.txtDadosRegistro.Multiline = true;
			this.txtDadosRegistro.Name = "txtDadosRegistro";
			this.txtDadosRegistro.ReadOnly = true;
			this.txtDadosRegistro.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDadosRegistro.Size = new System.Drawing.Size(727, 165);
			this.txtDadosRegistro.TabIndex = 0;
			this.txtDadosRegistro.TabStop = false;
			this.txtDadosRegistro.Text = "LINHA 1\r\nLINHA 2\r\nLINHA 3\r\nLINHA 4\r\nLINHA 5\r\nLINHA 6\r\nLINHA 7\r\nLINHA 8\r\nLINHA 9\r\n" +
    "LINHA 10\r\nLINHA 11\r\nLINHA 12";
			// 
			// gboxIdentificacaoOcorrencia
			// 
			this.gboxIdentificacaoOcorrencia.Controls.Add(this.txtIdentificacaoOcorrencia);
			this.gboxIdentificacaoOcorrencia.Location = new System.Drawing.Point(12, 151);
			this.gboxIdentificacaoOcorrencia.Name = "gboxIdentificacaoOcorrencia";
			this.gboxIdentificacaoOcorrencia.Size = new System.Drawing.Size(758, 76);
			this.gboxIdentificacaoOcorrencia.TabIndex = 1;
			this.gboxIdentificacaoOcorrencia.TabStop = false;
			this.gboxIdentificacaoOcorrencia.Text = "Identificação da ocorrência";
			// 
			// txtIdentificacaoOcorrencia
			// 
			this.txtIdentificacaoOcorrencia.BackColor = System.Drawing.SystemColors.Window;
			this.txtIdentificacaoOcorrencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtIdentificacaoOcorrencia.Location = new System.Drawing.Point(15, 19);
			this.txtIdentificacaoOcorrencia.Multiline = true;
			this.txtIdentificacaoOcorrencia.Name = "txtIdentificacaoOcorrencia";
			this.txtIdentificacaoOcorrencia.ReadOnly = true;
			this.txtIdentificacaoOcorrencia.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtIdentificacaoOcorrencia.Size = new System.Drawing.Size(727, 47);
			this.txtIdentificacaoOcorrencia.TabIndex = 0;
			this.txtIdentificacaoOcorrencia.TabStop = false;
			this.txtIdentificacaoOcorrencia.Text = "LINHA 1\r\nLINHA 2\r\nLINHA 3\r\nLINHA 4\r\nLINHA 5";
			// 
			// gboxComentarioOcorrenciaTratada
			// 
			this.gboxComentarioOcorrenciaTratada.Controls.Add(this.txtComentarioOcorrenciaTratada);
			this.gboxComentarioOcorrenciaTratada.Location = new System.Drawing.Point(12, 451);
			this.gboxComentarioOcorrenciaTratada.Name = "gboxComentarioOcorrenciaTratada";
			this.gboxComentarioOcorrenciaTratada.Size = new System.Drawing.Size(758, 90);
			this.gboxComentarioOcorrenciaTratada.TabIndex = 3;
			this.gboxComentarioOcorrenciaTratada.TabStop = false;
			this.gboxComentarioOcorrenciaTratada.Text = "Comentários e/ou observações para registrar junto com a ocorrência";
			// 
			// txtComentarioOcorrenciaTratada
			// 
			this.txtComentarioOcorrenciaTratada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtComentarioOcorrenciaTratada.Location = new System.Drawing.Point(15, 19);
			this.txtComentarioOcorrenciaTratada.MaxLength = 240;
			this.txtComentarioOcorrenciaTratada.Multiline = true;
			this.txtComentarioOcorrenciaTratada.Name = "txtComentarioOcorrenciaTratada";
			this.txtComentarioOcorrenciaTratada.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtComentarioOcorrenciaTratada.Size = new System.Drawing.Size(727, 61);
			this.txtComentarioOcorrenciaTratada.TabIndex = 0;
			this.txtComentarioOcorrenciaTratada.Text = "LINHA 1\r\nLINHA 2\r\nLINHA 3\r\nLINHA 4\r\nLINHA 5";
			this.txtComentarioOcorrenciaTratada.Enter += new System.EventHandler(this.txtComentarioOcorrenciaTratada_Enter);
			this.txtComentarioOcorrenciaTratada.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComentarioOcorrenciaTratada_KeyPress);
			this.txtComentarioOcorrenciaTratada.Leave += new System.EventHandler(this.txtComentarioOcorrenciaTratada_Leave);
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(450, 554);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 5;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnOk
			// 
			this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(233, 554);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 4;
			this.btnOk.Text = "&Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// gboxCedente
			// 
			this.gboxCedente.Controls.Add(this.lblCedente);
			this.gboxCedente.Location = new System.Drawing.Point(12, 6);
			this.gboxCedente.Name = "gboxCedente";
			this.gboxCedente.Size = new System.Drawing.Size(758, 37);
			this.gboxCedente.TabIndex = 6;
			this.gboxCedente.TabStop = false;
			this.gboxCedente.Text = "Cedente";
			// 
			// lblCedente
			// 
			this.lblCedente.AutoSize = true;
			this.lblCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblCedente.Location = new System.Drawing.Point(12, 16);
			this.lblCedente.Name = "lblCedente";
			this.lblCedente.Size = new System.Drawing.Size(160, 13);
			this.lblCedente.TabIndex = 0;
			this.lblCedente.Text = "Nome da Empresa Cedente";
			// 
			// FBoletoTrataOcorrenciaValaComum
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancela;
			this.ClientSize = new System.Drawing.Size(782, 604);
			this.Controls.Add(this.gboxCedente);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.gboxComentarioOcorrenciaTratada);
			this.Controls.Add(this.gboxIdentificacaoOcorrencia);
			this.Controls.Add(this.gboxDadosRegistro);
			this.Controls.Add(this.gboxCliente);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "FBoletoTrataOcorrenciaValaComum";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Boleto - Tratamento de Ocorrência";
			this.Load += new System.EventHandler(this.FBoletoTrataOcorrenciaValaComum_Load);
			this.Shown += new System.EventHandler(this.FBoletoTrataOcorrenciaValaComum_Shown);
			this.gboxCliente.ResumeLayout(false);
			this.gboxCliente.PerformLayout();
			this.gboxDadosRegistro.ResumeLayout(false);
			this.gboxDadosRegistro.PerformLayout();
			this.gboxIdentificacaoOcorrencia.ResumeLayout(false);
			this.gboxIdentificacaoOcorrencia.PerformLayout();
			this.gboxComentarioOcorrenciaTratada.ResumeLayout(false);
			this.gboxComentarioOcorrenciaTratada.PerformLayout();
			this.gboxCedente.ResumeLayout(false);
			this.gboxCedente.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxCliente;
		private System.Windows.Forms.TextBox txtClienteCnpjCpf;
		private System.Windows.Forms.Label lblTitClienteCnpjCpf;
		private System.Windows.Forms.TextBox txtClienteNome;
		private System.Windows.Forms.Label lblTitClienteNome;
		private System.Windows.Forms.GroupBox gboxDadosRegistro;
		private System.Windows.Forms.GroupBox gboxIdentificacaoOcorrencia;
		private System.Windows.Forms.TextBox txtIdentificacaoOcorrencia;
		private System.Windows.Forms.TextBox txtDadosRegistro;
		private System.Windows.Forms.GroupBox gboxComentarioOcorrenciaTratada;
		private System.Windows.Forms.TextBox txtComentarioOcorrenciaTratada;
		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.GroupBox gboxCedente;
		private System.Windows.Forms.Label lblCedente;

	}
}