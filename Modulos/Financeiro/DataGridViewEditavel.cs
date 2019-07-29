#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
#endregion

namespace Financeiro
{
	class DataGridViewEditavel : DataGridView
	{
		#region [ ProcessDialogKey ]
		[System.Security.Permissions.UIPermission(
			   System.Security.Permissions.SecurityAction.LinkDemand,
			   Window = System.Security.Permissions.UIPermissionWindow.AllWindows)]
		protected override bool ProcessDialogKey(Keys keyData)
		{
			// Extract the key code from the key value. 
			Keys key = (keyData & Keys.KeyCode);

			// Handle the ENTER key as if it were a TAB key. 
			if (key == Keys.Enter)
			{
				return this.ProcessTabKey(keyData);
			}
			else if (key == Keys.Escape)
			{
				this.CancelEdit();
				return this.ProcessEscapeKey(keyData);
			}
			
			return base.ProcessDialogKey(keyData);
		}
		#endregion

		#region [ ProcessDataGridViewKey ]
		[System.Security.Permissions.SecurityPermission(
			System.Security.Permissions.SecurityAction.LinkDemand, Flags =
			System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode)]
		protected override bool ProcessDataGridViewKey(KeyEventArgs e)
		{
			// Handle the ENTER key as if it were a TAB key. 
			if (e.KeyCode == Keys.Enter)
			{
				if (this.ReadOnly)
				{
					return this.ProcessTabKey(e.KeyData);
				}
				else
				{
					focusNextEditableCell();
					e.SuppressKeyPress = true;
					return true;
				}
			}
			
			return base.ProcessDataGridViewKey(e);
		}
		#endregion

		#region [ OnRowPostPaint ]
		protected override void OnRowPostPaint(DataGridViewRowPostPaintEventArgs e)
		{ //this method overrides the DataGridView's RowPostPaint event 
			//in order to automatically draw numbers on the row header cells
			//and to automatically adjust the width of the column containing
			//the row header cells so that it can accommodate the new row
			//numbers,

			//store a string representation of the row number in 'strRowNumber'
			string strRowNumber = (e.RowIndex + 1).ToString();

			//prepend leading zeros to the string if necessary to improve
			//appearance. For example, if there are ten rows in the grid,
			//row seven will be numbered as "07" instead of "7". Similarly, if 
			//there are 100 rows in the grid, row seven will be numbered as "007".
			while (strRowNumber.Length < this.RowCount.ToString().Length) strRowNumber = "0" + strRowNumber;

			//determine the display size of the row number string using
			//the DataGridView's current font.
			SizeF size = e.Graphics.MeasureString(strRowNumber, this.Font);

			//adjust the width of the column that contains the row header cells 
			//if necessary
			if (this.RowHeadersWidth < (int)(size.Width + 20)) this.RowHeadersWidth = (int)(size.Width + 20);

			//this brush will be used to draw the row number string on the
			//row header cell using the system's current ControlText color
			Brush b = SystemBrushes.ControlText;

			//draw the row number string on the current row header cell using
			//the brush defined above and the DataGridView's default font
			e.Graphics.DrawString(strRowNumber, this.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2));

			//call the base object's OnRowPostPaint method
			base.OnRowPostPaint(e);
		} //end OnRowPostPaint method
		#endregion

		#region [ focusNextCell ]
		/// <summary>
		/// Move o foco para a próxima célula do grid, mesmo que seja readonly
		/// </summary>
		/// <returns>
		/// True: transferiu o foco para a próxima célula
		/// False: não há próxima célula para transferir o foco
		/// </returns>
		public bool focusNextCell()
		{
			#region [ Declarações ]
			DataGridViewCell celula = null;
			#endregion

			#region [ Há células? ]
			if (this.Rows.Count == 0) return false;
			if (this.Columns.Count == 0) return false;
			#endregion

			#region [ Obtém a célula selecionada atualmente ]
			celula = this.CurrentCell;
			#endregion

			if (celula == null)
			{
				#region [ Não tem nenhuma célula selecionada, posiciona na primeira ]
				for (int i = 0; i < this.Columns.Count; i++)
				{
					if (this.Columns[i].Visible)
					{
						this.Rows[0].Cells[i].Selected = true;
						return true;
					}
				}
				#endregion
			}
			else
			{
				#region [ Posiciona na próxima célula ]
				for (int i = celula.RowIndex; i < this.Rows.Count; i++)
				{
					for (int j = 0; j < this.Columns.Count; j++)
					{
						if ((i == celula.RowIndex) && (j <= celula.ColumnIndex)) continue;
						if (this.Columns[j].Visible)
						{
							this.Rows[i].Cells[j].Selected = true;
							return true;
						}
					}
				}
				#endregion
			}

			return false;
		}
		#endregion

		#region [ focusNextEditableCell ]
		/// <summary>
		/// Move o foco para a próxima célula do grid que seja editável
		/// </summary>
		/// <returns>
		/// True: transferiu o foco para a próxima célula
		/// False: não há próxima célula para transferir o foco
		/// </returns>
		public bool focusNextEditableCell()
		{
			#region [ Declarações ]
			DataGridViewCell celula = null;
			#endregion

			#region [ Há células? ]
			if (this.Rows.Count == 0) return false;
			if (this.Columns.Count == 0) return false;
			#endregion

			#region [ Obtém a célula selecionada atualmente ]
			celula = this.CurrentCell;
			#endregion

			if (celula == null)
			{
				#region [ Não tem nenhuma célula selecionada, posiciona na primeira ]
				for (int i = 0; i < this.Columns.Count; i++)
				{
					if (this.Columns[i].Visible && (!this.Rows[0].Cells[i].ReadOnly))
					{
						this.Rows[0].Cells[i].Selected = true;
						return true;
					}
				}
				#endregion
			}
			else
			{
				#region [ Posiciona na próxima célula ]
				for (int i = celula.RowIndex; i < this.Rows.Count; i++)
				{
					for (int j = 0; j < this.Columns.Count; j++)
					{
						if ((i == celula.RowIndex) && (j <= celula.ColumnIndex)) continue;
						if (this.Columns[j].Visible && (!this.Rows[i].Cells[j].ReadOnly))
						{
							this.Rows[i].Cells[j].Selected = true;
							return true;
						}
					}
				}
				#endregion
			}

			return false;
		}
		#endregion
	}
}
