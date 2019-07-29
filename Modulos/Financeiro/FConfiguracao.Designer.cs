namespace Financeiro
{
	partial class FConfiguracao
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FConfiguracao));
			this.gboxEmail = new System.Windows.Forms.GroupBox();
			this.txtDisplayNameRemetente = new System.Windows.Forms.TextBox();
			this.lblTitDisplayNameRemetente = new System.Windows.Forms.Label();
			this.txtSenhaSmtp = new System.Windows.Forms.TextBox();
			this.lblTitSenhaSmtp = new System.Windows.Forms.Label();
			this.txtUsuarioSmtp = new System.Windows.Forms.TextBox();
			this.lblTitIdUserSmtp = new System.Windows.Forms.Label();
			this.txtEmailRemetente = new System.Windows.Forms.TextBox();
			this.lblTitEmailRemetente = new System.Windows.Forms.Label();
			this.txtServidorSmtp = new System.Windows.Forms.TextBox();
			this.lblTitServidorSmtp = new System.Windows.Forms.Label();
			this.btnConfirma = new System.Windows.Forms.Button();
			this.lblTitServidorSmtpPorta = new System.Windows.Forms.Label();
			this.txtServidorSmtpPorta = new System.Windows.Forms.TextBox();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxEmail.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnConfirma);
			this.pnBotoes.Size = new System.Drawing.Size(778, 55);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnConfirma, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxEmail);
			this.pnCampos.Size = new System.Drawing.Size(778, 219);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(729, 4);
			this.btnFechar.TabIndex = 2;
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(684, 4);
			this.btnSobre.TabIndex = 1;
			// 
			// gboxEmail
			// 
			this.gboxEmail.Controls.Add(this.txtServidorSmtpPorta);
			this.gboxEmail.Controls.Add(this.lblTitServidorSmtpPorta);
			this.gboxEmail.Controls.Add(this.txtDisplayNameRemetente);
			this.gboxEmail.Controls.Add(this.lblTitDisplayNameRemetente);
			this.gboxEmail.Controls.Add(this.txtSenhaSmtp);
			this.gboxEmail.Controls.Add(this.lblTitSenhaSmtp);
			this.gboxEmail.Controls.Add(this.txtUsuarioSmtp);
			this.gboxEmail.Controls.Add(this.lblTitIdUserSmtp);
			this.gboxEmail.Controls.Add(this.txtEmailRemetente);
			this.gboxEmail.Controls.Add(this.lblTitEmailRemetente);
			this.gboxEmail.Controls.Add(this.txtServidorSmtp);
			this.gboxEmail.Controls.Add(this.lblTitServidorSmtp);
			this.gboxEmail.Location = new System.Drawing.Point(35, 34);
			this.gboxEmail.Name = "gboxEmail";
			this.gboxEmail.Size = new System.Drawing.Size(705, 146);
			this.gboxEmail.TabIndex = 0;
			this.gboxEmail.TabStop = false;
			this.gboxEmail.Text = "E-mail usado para envio dos boletos";
			// 
			// txtDisplayNameRemetente
			// 
			this.txtDisplayNameRemetente.Location = new System.Drawing.Point(137, 81);
			this.txtDisplayNameRemetente.MaxLength = 80;
			this.txtDisplayNameRemetente.Name = "txtDisplayNameRemetente";
			this.txtDisplayNameRemetente.Size = new System.Drawing.Size(543, 20);
			this.txtDisplayNameRemetente.TabIndex = 3;
			this.txtDisplayNameRemetente.WordWrap = false;
			this.txtDisplayNameRemetente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDisplayNameRemetente_KeyDown);
			this.txtDisplayNameRemetente.Leave += new System.EventHandler(this.txtDisplayNameRemetente_Leave);
			this.txtDisplayNameRemetente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDisplayNameRemetente_KeyPress);
			this.txtDisplayNameRemetente.Enter += new System.EventHandler(this.txtDisplayNameRemetente_Enter);
			// 
			// lblTitDisplayNameRemetente
			// 
			this.lblTitDisplayNameRemetente.AutoSize = true;
			this.lblTitDisplayNameRemetente.Location = new System.Drawing.Point(46, 84);
			this.lblTitDisplayNameRemetente.Name = "lblTitDisplayNameRemetente";
			this.lblTitDisplayNameRemetente.Size = new System.Drawing.Size(85, 13);
			this.lblTitDisplayNameRemetente.TabIndex = 16;
			this.lblTitDisplayNameRemetente.Text = "Nome remetente";
			// 
			// txtSenhaSmtp
			// 
			this.txtSenhaSmtp.Location = new System.Drawing.Point(500, 107);
			this.txtSenhaSmtp.MaxLength = 80;
			this.txtSenhaSmtp.Name = "txtSenhaSmtp";
			this.txtSenhaSmtp.Size = new System.Drawing.Size(180, 20);
			this.txtSenhaSmtp.TabIndex = 5;
			this.txtSenhaSmtp.UseSystemPasswordChar = true;
			this.txtSenhaSmtp.WordWrap = false;
			this.txtSenhaSmtp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSenhaSmtp_KeyDown);
			this.txtSenhaSmtp.Leave += new System.EventHandler(this.txtSenhaSmtp_Leave);
			this.txtSenhaSmtp.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSenhaSmtp_KeyPress);
			this.txtSenhaSmtp.Enter += new System.EventHandler(this.txtSenhaSmtp_Enter);
			// 
			// lblTitSenhaSmtp
			// 
			this.lblTitSenhaSmtp.AutoSize = true;
			this.lblTitSenhaSmtp.Location = new System.Drawing.Point(393, 110);
			this.lblTitSenhaSmtp.Name = "lblTitSenhaSmtp";
			this.lblTitSenhaSmtp.Size = new System.Drawing.Size(101, 13);
			this.lblTitSenhaSmtp.TabIndex = 14;
			this.lblTitSenhaSmtp.Text = "Senha conta SMTP";
			// 
			// txtUsuarioSmtp
			// 
			this.txtUsuarioSmtp.Location = new System.Drawing.Point(137, 107);
			this.txtUsuarioSmtp.MaxLength = 80;
			this.txtUsuarioSmtp.Name = "txtUsuarioSmtp";
			this.txtUsuarioSmtp.Size = new System.Drawing.Size(180, 20);
			this.txtUsuarioSmtp.TabIndex = 4;
			this.txtUsuarioSmtp.WordWrap = false;
			this.txtUsuarioSmtp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUsuarioSmtp_KeyDown);
			this.txtUsuarioSmtp.Leave += new System.EventHandler(this.txtUsuarioSmtp_Leave);
			this.txtUsuarioSmtp.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUsuarioSmtp_KeyPress);
			this.txtUsuarioSmtp.Enter += new System.EventHandler(this.txtUsuarioSmtp_Enter);
			// 
			// lblTitIdUserSmtp
			// 
			this.lblTitIdUserSmtp.AutoSize = true;
			this.lblTitIdUserSmtp.Location = new System.Drawing.Point(25, 110);
			this.lblTitIdUserSmtp.Name = "lblTitIdUserSmtp";
			this.lblTitIdUserSmtp.Size = new System.Drawing.Size(106, 13);
			this.lblTitIdUserSmtp.TabIndex = 13;
			this.lblTitIdUserSmtp.Text = "Usuário conta SMTP";
			// 
			// txtEmailRemetente
			// 
			this.txtEmailRemetente.Location = new System.Drawing.Point(137, 55);
			this.txtEmailRemetente.MaxLength = 80;
			this.txtEmailRemetente.Name = "txtEmailRemetente";
			this.txtEmailRemetente.Size = new System.Drawing.Size(543, 20);
			this.txtEmailRemetente.TabIndex = 2;
			this.txtEmailRemetente.WordWrap = false;
			this.txtEmailRemetente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEmailRemetente_KeyDown);
			this.txtEmailRemetente.Leave += new System.EventHandler(this.txtEmailRemetente_Leave);
			this.txtEmailRemetente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEmailRemetente_KeyPress);
			this.txtEmailRemetente.Enter += new System.EventHandler(this.txtEmailRemetente_Enter);
			// 
			// lblTitEmailRemetente
			// 
			this.lblTitEmailRemetente.AutoSize = true;
			this.lblTitEmailRemetente.Location = new System.Drawing.Point(46, 58);
			this.lblTitEmailRemetente.Name = "lblTitEmailRemetente";
			this.lblTitEmailRemetente.Size = new System.Drawing.Size(85, 13);
			this.lblTitEmailRemetente.TabIndex = 11;
			this.lblTitEmailRemetente.Text = "E-mail remetente";
			// 
			// txtServidorSmtp
			// 
			this.txtServidorSmtp.Location = new System.Drawing.Point(137, 29);
			this.txtServidorSmtp.MaxLength = 80;
			this.txtServidorSmtp.Name = "txtServidorSmtp";
			this.txtServidorSmtp.Size = new System.Drawing.Size(357, 20);
			this.txtServidorSmtp.TabIndex = 0;
			this.txtServidorSmtp.WordWrap = false;
			this.txtServidorSmtp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtServidorSmtp_KeyDown);
			this.txtServidorSmtp.Leave += new System.EventHandler(this.txtServidorSmtp_Leave);
			this.txtServidorSmtp.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtServidorSmtp_KeyPress);
			this.txtServidorSmtp.Enter += new System.EventHandler(this.txtServidorSmtp_Enter);
			// 
			// lblTitServidorSmtp
			// 
			this.lblTitServidorSmtp.AutoSize = true;
			this.lblTitServidorSmtp.Location = new System.Drawing.Point(52, 32);
			this.lblTitServidorSmtp.Name = "lblTitServidorSmtp";
			this.lblTitServidorSmtp.Size = new System.Drawing.Size(79, 13);
			this.lblTitServidorSmtp.TabIndex = 8;
			this.lblTitServidorSmtp.Text = "Servidor SMTP";
			// 
			// btnConfirma
			// 
			this.btnConfirma.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnConfirma.Image = ((System.Drawing.Image)(resources.GetObject("btnConfirma.Image")));
			this.btnConfirma.Location = new System.Drawing.Point(639, 4);
			this.btnConfirma.Name = "btnConfirma";
			this.btnConfirma.Size = new System.Drawing.Size(40, 44);
			this.btnConfirma.TabIndex = 0;
			this.btnConfirma.TabStop = false;
			this.btnConfirma.UseVisualStyleBackColor = true;
			this.btnConfirma.Click += new System.EventHandler(this.btnConfirma_Click);
			// 
			// lblTitServidorSmtpPorta
			// 
			this.lblTitServidorSmtpPorta.AutoSize = true;
			this.lblTitServidorSmtpPorta.Location = new System.Drawing.Point(592, 32);
			this.lblTitServidorSmtpPorta.Name = "lblTitServidorSmtpPorta";
			this.lblTitServidorSmtpPorta.Size = new System.Drawing.Size(32, 13);
			this.lblTitServidorSmtpPorta.TabIndex = 17;
			this.lblTitServidorSmtpPorta.Text = "Porta";
			// 
			// txtServidorSmtpPorta
			// 
			this.txtServidorSmtpPorta.Location = new System.Drawing.Point(630, 29);
			this.txtServidorSmtpPorta.MaxLength = 5;
			this.txtServidorSmtpPorta.Name = "txtServidorSmtpPorta";
			this.txtServidorSmtpPorta.Size = new System.Drawing.Size(50, 20);
			this.txtServidorSmtpPorta.TabIndex = 1;
			this.txtServidorSmtpPorta.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtServidorSmtpPorta.WordWrap = false;
			this.txtServidorSmtpPorta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtServidorSmtpPorta_KeyDown);
			this.txtServidorSmtpPorta.Leave += new System.EventHandler(this.txtServidorSmtpPorta_Leave);
			this.txtServidorSmtpPorta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtServidorSmtpPorta_KeyPress);
			this.txtServidorSmtpPorta.Enter += new System.EventHandler(this.txtServidorSmtpPorta_Enter);
			// 
			// FConfiguracao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(778, 316);
			this.Name = "FConfiguracao";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FConfiguracao_Load);
			this.Shown += new System.EventHandler(this.FConfiguracao_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxEmail.ResumeLayout(false);
			this.gboxEmail.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxEmail;
		private System.Windows.Forms.Button btnConfirma;
		internal System.Windows.Forms.TextBox txtSenhaSmtp;
		internal System.Windows.Forms.Label lblTitSenhaSmtp;
		internal System.Windows.Forms.TextBox txtUsuarioSmtp;
		internal System.Windows.Forms.Label lblTitIdUserSmtp;
		internal System.Windows.Forms.TextBox txtEmailRemetente;
		internal System.Windows.Forms.Label lblTitEmailRemetente;
		internal System.Windows.Forms.TextBox txtServidorSmtp;
		internal System.Windows.Forms.Label lblTitServidorSmtp;
		internal System.Windows.Forms.TextBox txtDisplayNameRemetente;
		internal System.Windows.Forms.Label lblTitDisplayNameRemetente;
		internal System.Windows.Forms.TextBox txtServidorSmtpPorta;
		internal System.Windows.Forms.Label lblTitServidorSmtpPorta;
	}
}
