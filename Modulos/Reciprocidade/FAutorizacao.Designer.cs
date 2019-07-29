namespace Reciprocidade
{
	partial class FAutorizacao
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FAutorizacao));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.pnMensagem = new System.Windows.Forms.Panel();
			this.webBrowserMensagem = new System.Windows.Forms.WebBrowser();
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnOk = new System.Windows.Forms.Button();
			this.lblTitSenha = new System.Windows.Forms.Label();
			this.txtSenha = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.pnMensagem.SuspendLayout();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(115, 259);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(83, 83);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 13;
			this.pictureBox1.TabStop = false;
			// 
			// pnMensagem
			// 
			this.pnMensagem.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pnMensagem.Controls.Add(this.webBrowserMensagem);
			this.pnMensagem.Location = new System.Drawing.Point(14, 14);
			this.pnMensagem.Name = "pnMensagem";
			this.pnMensagem.Size = new System.Drawing.Size(755, 230);
			this.pnMensagem.TabIndex = 12;
			// 
			// webBrowserMensagem
			// 
			this.webBrowserMensagem.AllowNavigation = false;
			this.webBrowserMensagem.AllowWebBrowserDrop = false;
			this.webBrowserMensagem.Dock = System.Windows.Forms.DockStyle.Fill;
			this.webBrowserMensagem.Location = new System.Drawing.Point(0, 0);
			this.webBrowserMensagem.MinimumSize = new System.Drawing.Size(20, 20);
			this.webBrowserMensagem.Name = "webBrowserMensagem";
			this.webBrowserMensagem.ScriptErrorsSuppressed = true;
			this.webBrowserMensagem.Size = new System.Drawing.Size(751, 226);
			this.webBrowserMensagem.TabIndex = 0;
			this.webBrowserMensagem.TabStop = false;
			this.webBrowserMensagem.WebBrowserShortcutsEnabled = false;
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(527, 302);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 3;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnOk
			// 
			this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(310, 302);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 2;
			this.btnOk.Text = "&Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// lblTitSenha
			// 
			this.lblTitSenha.AutoSize = true;
			this.lblTitSenha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitSenha.Location = new System.Drawing.Point(243, 266);
			this.lblTitSenha.Name = "lblTitSenha";
			this.lblTitSenha.Size = new System.Drawing.Size(61, 20);
			this.lblTitSenha.TabIndex = 0;
			this.lblTitSenha.Text = "Senha";
			// 
			// txtSenha
			// 
			this.txtSenha.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtSenha.Location = new System.Drawing.Point(310, 259);
			this.txtSenha.MaxLength = 15;
			this.txtSenha.Name = "txtSenha";
			this.txtSenha.Size = new System.Drawing.Size(317, 32);
			this.txtSenha.TabIndex = 1;
			this.txtSenha.Text = "99999999";
			this.txtSenha.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtSenha.UseSystemPasswordChar = true;
			this.txtSenha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSenha_KeyDown);
			// 
			// FAutorizacao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancela;
			this.ClientSize = new System.Drawing.Size(782, 357);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.pnMensagem);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.lblTitSenha);
			this.Controls.Add(this.txtSenha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.Name = "FAutorizacao";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Autorização";
			this.Load += new System.EventHandler(this.FAutorizacao_Load);
			this.Shown += new System.EventHandler(this.FAutorizacao_Shown);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.pnMensagem.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Panel pnMensagem;
		private System.Windows.Forms.WebBrowser webBrowserMensagem;
		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.Label lblTitSenha;
		private System.Windows.Forms.TextBox txtSenha;
	}
}