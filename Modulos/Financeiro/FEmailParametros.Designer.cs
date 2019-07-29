namespace Financeiro
{
	partial class FEmailParametros
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FEmailParametros));
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnOk = new System.Windows.Forms.Button();
			this.gboxRemetente = new System.Windows.Forms.GroupBox();
			this.lblEmailRemetente = new System.Windows.Forms.Label();
			this.gboxDestinatario = new System.Windows.Forms.GroupBox();
			this.lblTitDestinatarioPara = new System.Windows.Forms.Label();
			this.txtDestinatarioPara = new System.Windows.Forms.TextBox();
			this.lblTitDestinatarioCopia = new System.Windows.Forms.Label();
			this.txtDestinatarioCopia = new System.Windows.Forms.TextBox();
			this.gboxAssunto = new System.Windows.Forms.GroupBox();
			this.txtAssunto = new System.Windows.Forms.TextBox();
			this.gboxRemetente.SuspendLayout();
			this.gboxDestinatario.SuspendLayout();
			this.gboxAssunto.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(514, 241);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 1;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnOk
			// 
			this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(297, 241);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 0;
			this.btnOk.Text = "&Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// gboxRemetente
			// 
			this.gboxRemetente.Controls.Add(this.lblEmailRemetente);
			this.gboxRemetente.Location = new System.Drawing.Point(32, 9);
			this.gboxRemetente.Name = "gboxRemetente";
			this.gboxRemetente.Size = new System.Drawing.Size(853, 42);
			this.gboxRemetente.TabIndex = 8;
			this.gboxRemetente.TabStop = false;
			this.gboxRemetente.Text = "Remetente";
			// 
			// lblEmailRemetente
			// 
			this.lblEmailRemetente.AutoSize = true;
			this.lblEmailRemetente.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblEmailRemetente.Location = new System.Drawing.Point(9, 17);
			this.lblEmailRemetente.Name = "lblEmailRemetente";
			this.lblEmailRemetente.Size = new System.Drawing.Size(52, 17);
			this.lblEmailRemetente.TabIndex = 0;
			this.lblEmailRemetente.Text = "label1";
			// 
			// gboxDestinatario
			// 
			this.gboxDestinatario.Controls.Add(this.txtDestinatarioCopia);
			this.gboxDestinatario.Controls.Add(this.lblTitDestinatarioCopia);
			this.gboxDestinatario.Controls.Add(this.txtDestinatarioPara);
			this.gboxDestinatario.Controls.Add(this.lblTitDestinatarioPara);
			this.gboxDestinatario.Location = new System.Drawing.Point(32, 121);
			this.gboxDestinatario.Name = "gboxDestinatario";
			this.gboxDestinatario.Size = new System.Drawing.Size(853, 101);
			this.gboxDestinatario.TabIndex = 9;
			this.gboxDestinatario.TabStop = false;
			this.gboxDestinatario.Text = "Destinatário";
			// 
			// lblTitDestinatarioPara
			// 
			this.lblTitDestinatarioPara.AutoSize = true;
			this.lblTitDestinatarioPara.Location = new System.Drawing.Point(14, 19);
			this.lblTitDestinatarioPara.Name = "lblTitDestinatarioPara";
			this.lblTitDestinatarioPara.Size = new System.Drawing.Size(29, 13);
			this.lblTitDestinatarioPara.TabIndex = 0;
			this.lblTitDestinatarioPara.Text = "Para";
			// 
			// txtDestinatarioPara
			// 
			this.txtDestinatarioPara.Location = new System.Drawing.Point(49, 19);
			this.txtDestinatarioPara.Multiline = true;
			this.txtDestinatarioPara.Name = "txtDestinatarioPara";
			this.txtDestinatarioPara.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDestinatarioPara.Size = new System.Drawing.Size(792, 32);
			this.txtDestinatarioPara.TabIndex = 0;
			this.txtDestinatarioPara.Leave += new System.EventHandler(this.txtDestinatarioPara_Leave);
			this.txtDestinatarioPara.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDestinatarioPara_KeyPress);
			this.txtDestinatarioPara.Enter += new System.EventHandler(this.txtDestinatarioPara_Enter);
			// 
			// lblTitDestinatarioCopia
			// 
			this.lblTitDestinatarioCopia.AutoSize = true;
			this.lblTitDestinatarioCopia.Location = new System.Drawing.Point(9, 58);
			this.lblTitDestinatarioCopia.Name = "lblTitDestinatarioCopia";
			this.lblTitDestinatarioCopia.Size = new System.Drawing.Size(34, 13);
			this.lblTitDestinatarioCopia.TabIndex = 2;
			this.lblTitDestinatarioCopia.Text = "Cópia";
			// 
			// txtDestinatarioCopia
			// 
			this.txtDestinatarioCopia.Location = new System.Drawing.Point(49, 58);
			this.txtDestinatarioCopia.Multiline = true;
			this.txtDestinatarioCopia.Name = "txtDestinatarioCopia";
			this.txtDestinatarioCopia.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDestinatarioCopia.Size = new System.Drawing.Size(792, 32);
			this.txtDestinatarioCopia.TabIndex = 1;
			this.txtDestinatarioCopia.Leave += new System.EventHandler(this.txtDestinatarioCopia_Leave);
			this.txtDestinatarioCopia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDestinatarioCopia_KeyPress);
			this.txtDestinatarioCopia.Enter += new System.EventHandler(this.txtDestinatarioCopia_Enter);
			// 
			// gboxAssunto
			// 
			this.gboxAssunto.Controls.Add(this.txtAssunto);
			this.gboxAssunto.Location = new System.Drawing.Point(32, 63);
			this.gboxAssunto.Name = "gboxAssunto";
			this.gboxAssunto.Size = new System.Drawing.Size(853, 46);
			this.gboxAssunto.TabIndex = 10;
			this.gboxAssunto.TabStop = false;
			this.gboxAssunto.Text = "Assunto";
			// 
			// txtAssunto
			// 
			this.txtAssunto.Location = new System.Drawing.Point(12, 17);
			this.txtAssunto.Name = "txtAssunto";
			this.txtAssunto.Size = new System.Drawing.Size(829, 20);
			this.txtAssunto.TabIndex = 0;
			this.txtAssunto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAssunto_KeyDown);
			this.txtAssunto.Leave += new System.EventHandler(this.txtAssunto_Leave);
			this.txtAssunto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAssunto_KeyPress);
			this.txtAssunto.Enter += new System.EventHandler(this.txtAssunto_Enter);
			// 
			// FEmailParametros
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(910, 296);
			this.Controls.Add(this.gboxAssunto);
			this.Controls.Add(this.gboxDestinatario);
			this.Controls.Add(this.gboxRemetente);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "FEmailParametros";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "E-mail";
			this.Load += new System.EventHandler(this.FEmailParametros_Load);
			this.Shown += new System.EventHandler(this.FEmailParametros_Shown);
			this.gboxRemetente.ResumeLayout(false);
			this.gboxRemetente.PerformLayout();
			this.gboxDestinatario.ResumeLayout(false);
			this.gboxDestinatario.PerformLayout();
			this.gboxAssunto.ResumeLayout(false);
			this.gboxAssunto.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.GroupBox gboxRemetente;
		private System.Windows.Forms.Label lblEmailRemetente;
		private System.Windows.Forms.GroupBox gboxDestinatario;
		private System.Windows.Forms.TextBox txtDestinatarioCopia;
		private System.Windows.Forms.Label lblTitDestinatarioCopia;
		private System.Windows.Forms.TextBox txtDestinatarioPara;
		private System.Windows.Forms.Label lblTitDestinatarioPara;
		private System.Windows.Forms.GroupBox gboxAssunto;
		private System.Windows.Forms.TextBox txtAssunto;
	}
}