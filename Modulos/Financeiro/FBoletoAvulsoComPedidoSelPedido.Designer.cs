namespace Financeiro
{
	partial class FBoletoAvulsoComPedidoSelPedido
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoAvulsoComPedidoSelPedido));
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnOk = new System.Windows.Forms.Button();
			this.txtPedido = new System.Windows.Forms.TextBox();
			this.gboxPedido = new System.Windows.Forms.GroupBox();
			this.gboxPedido.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(154, 167);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(100, 40);
			this.btnCancela.TabIndex = 2;
			this.btnCancela.Text = "&Cancela";
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnOk
			// 
			this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(16, 167);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(100, 40);
			this.btnOk.TabIndex = 1;
			this.btnOk.Text = "&Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// txtPedido
			// 
			this.txtPedido.AcceptsReturn = true;
			this.txtPedido.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtPedido.Location = new System.Drawing.Point(41, 21);
			this.txtPedido.Multiline = true;
			this.txtPedido.Name = "txtPedido";
			this.txtPedido.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtPedido.Size = new System.Drawing.Size(157, 115);
			this.txtPedido.TabIndex = 0;
			this.txtPedido.Text = "009611G\r\n009612G-B";
			this.txtPedido.Leave += new System.EventHandler(this.txtPedido_Leave);
			this.txtPedido.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPedido_KeyPress);
			// 
			// gboxPedido
			// 
			this.gboxPedido.Controls.Add(this.txtPedido);
			this.gboxPedido.Location = new System.Drawing.Point(16, 9);
			this.gboxPedido.Name = "gboxPedido";
			this.gboxPedido.Size = new System.Drawing.Size(239, 147);
			this.gboxPedido.TabIndex = 0;
			this.gboxPedido.TabStop = false;
			this.gboxPedido.Text = "Pedido(s)";
			// 
			// FBoletoAvulsoComPedidoSelPedido
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancela;
			this.ClientSize = new System.Drawing.Size(270, 219);
			this.Controls.Add(this.gboxPedido);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "FBoletoAvulsoComPedidoSelPedido";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Cadastramento de Boleto Avulso";
			this.Load += new System.EventHandler(this.FBoletoAvulsoComPedidoSelPedido_Load);
			this.gboxPedido.ResumeLayout(false);
			this.gboxPedido.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.TextBox txtPedido;
		private System.Windows.Forms.GroupBox gboxPedido;
	}
}