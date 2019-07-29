#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Financeiro
{
	public partial class FBoletoTrataOcorrenciaCepInvalido : Form
	{
		#region [ Atributos ]

		#region [ Diversos ]
		private bool _InicializacaoOk;
		private Form _formChamador = null;
		#endregion

		#region [ Dados p/ exibição/edição ]
		private BoletoCedente _boletoCedente;
		private String _nomeCliente;
		private String _cnpjCpf;
		private String _endereco;
		private String _bairro;
		private String _cep;
		private String _cidade;
		private String _uf;

		public String enderecoCorrigido = "";
		public String bairroCorrigido = "";
		public String cepCorrigido = "";
		public String cidadeCorrigido = "";
		public String ufCorrigido = "";
		#endregion

		FCepPesquisa fCepPesquisa;
		#endregion

		#region [ Construtor ]
		public FBoletoTrataOcorrenciaCepInvalido(
							Form formChamador,
							int id_boleto_cedente,
							String nomeCliente,
							String cnpjCpf,
							String endereco,
							String bairro,
							String cep,
							String cidade,
							String uf)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			#region [ Armazena dados ]
			_formChamador = formChamador;
			_boletoCedente = BoletoCedenteDAO.getBoletoCedente(id_boleto_cedente);
			_nomeCliente = nomeCliente;
			_cnpjCpf = cnpjCpf;
			_endereco = endereco;
			_bairro = bairro;
			_cep = cep;
			_cidade = cidade;
			_uf = uf;
			#endregion
		}
		#endregion

		#region [ Métodos Privados ]

		#region[ aviso ]
		public void aviso(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		#endregion

		#region[ avisoErro ]
		public void avisoErro(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			lblCedente.Text = "";
			txtClienteNome.Text = "";
			txtClienteCnpjCpf.Text = "";
			txtEndereco.Text = "";
			txtBairro.Text = "";
			txtCep.Text = "";
			txtCidade.Text = "";
			txtUF.Text = "";
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Endereço ]
			if (txtEndereco.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o endereço do cliente!!");
				txtEndereco.Focus();
				return false;
			}

			if (txtEndereco.Text.Length > Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_ENDERECO)
			{
				avisoErro("É necessário editar o endereço, pois está excedendo o tamanho máximo!!");
				txtEndereco.Focus();
				if (txtEndereco.Text.Length > 0)
				{
					txtEndereco.SelectionStart = txtEndereco.Text.Length;
					txtEndereco.SelectionLength = 0;
				}
				return false;
			}
			#endregion

			#region [ CEP ]
			if (txtCep.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o CEP do cliente!!");
				txtCep.Focus();
				if (txtCep.Text.Length > 0)
				{
					txtCep.SelectionStart = txtCep.Text.Length;
					txtCep.SelectionLength = 0;
				}
				return false;
			}

			if (!Global.isCepOk(txtCep.Text))
			{
				avisoErro("CEP do cliente é inválido!!");
				txtCep.Focus();
				return false;
			}
			#endregion

			#region [ UF ]
			if (txtUF.Text.Length > 0)
			{
				if (!Global.isUfOk(txtUF.Text))
				{
					avisoErro("UF inválida!!");
					txtUF.Focus();
					return false;
				}
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoPesquisaCep ]
		void trataBotaoPesquisaCep()
		{
			#region [ Declarações ]
			String strEspaco = "";
			DialogResult drCep;
			String s1, s2, strEndereco;
			char c1, c2;
			#endregion

			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			#region [ Exibe painel de consulta de CEP ]
			fCepPesquisa = new FCepPesquisa();
			fCepPesquisa.StartPosition = FormStartPosition.Manual;
			fCepPesquisa.Location = _formChamador.Location;
			fCepPesquisa.cepDefault = txtCep.Text;
			drCep = fCepPesquisa.ShowDialog();
			if (drCep != DialogResult.OK) return;
			#endregion

			#region [ Atualiza com o resultado da pesquisa ]
			if (fCepPesquisa.cepSelecionado.Trim().Length > 0) txtCep.Text = fCepPesquisa.cepSelecionado;
			if (fCepPesquisa.bairroSelecionado.Trim().Length > 0) txtBairro.Text = fCepPesquisa.bairroSelecionado;
			if (fCepPesquisa.cidadeSelecionada.Trim().Length > 0) txtCidade.Text = fCepPesquisa.cidadeSelecionada;
			if (fCepPesquisa.ufSelecionado.Trim().Length > 0) txtUF.Text = fCepPesquisa.ufSelecionado;

			s1 = fCepPesquisa.logradouroSelecionado;
			s2 = fCepPesquisa.numeroOuComplementoSelecionado;
			if ((s1.Length > 0) && (s2.Length > 0))
			{
				c1 = s1.Substring(s1.Length - 1, 1)[0];
				c2 = s2.Substring(0, 1)[0];
				if (Global.isAlfaNumerico(c1) && Global.isAlfaNumerico(c2)) strEspaco = " ";
			}
			strEndereco = s1 + strEspaco + s2;
			if (strEndereco.Trim().Length > 0) txtEndereco.Text = strEndereco;
			#endregion
		}
		#endregion

		#region [ trataBotaoCancela ]
		private void trataBotaoCancela()
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#region [ trataBotaoOk ]
		private void trataBotaoOk()
		{
			if (!consisteCampos()) return;

			enderecoCorrigido = txtEndereco.Text.Trim();
			bairroCorrigido = txtBairro.Text.Trim();
			cepCorrigido = Global.digitos(txtCep.Text);
			cidadeCorrigido = txtCidade.Text.Trim();
			ufCorrigido = txtUF.Text.Trim();

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoTrataOcorrenciaCepInvalido ]

		#region [ FBoletoTrataOcorrenciaCepInvalido_Load ]
		private void FBoletoTrataOcorrenciaCepInvalido_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				btnDummy.Top = -200;

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				blnSucesso = true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				if (!blnSucesso) Close();
			}
		}
		#endregion

		#region [ FBoletoTrataOcorrenciaCepInvalido_Shown ]
		private void FBoletoTrataOcorrenciaCepInvalido_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche dados ]
					lblCedente.Text = _boletoCedente.nome_empresa.ToUpper();
					txtClienteNome.Text = _nomeCliente;
					txtClienteCnpjCpf.Text = Global.formataCnpjCpf(_cnpjCpf);
					txtEndereco.Text = _endereco;
					txtBairro.Text = _bairro;
					txtCep.Text = Global.formataCep(_cep);
					txtCidade.Text = _cidade;
					txtUF.Text = _uf;
					#endregion

					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					_InicializacaoOk = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				// Se não inicializou corretamente, assegura-se de que o painel será fechado
				if (!_InicializacaoOk) Close();
			}
		}
		#endregion

		#endregion

		#region [ txtEndereco ]

		#region [ txtEndereco_Enter ]
		private void txtEndereco_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtEndereco_Leave ]
		private void txtEndereco_Leave(object sender, EventArgs e)
		{
			txtEndereco.Text = txtEndereco.Text.Trim();
		}
		#endregion

		#region [ txtEndereco_KeyDown ]
		private void txtEndereco_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtBairro);
		}
		#endregion

		#region [ txtEndereco_KeyPress ]
		private void txtEndereco_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#region [ txtEndereco_TextChanged ]
		private void txtEndereco_TextChanged(object sender, EventArgs e)
		{
			int intTamanho;
			intTamanho = Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_ENDERECO - txtEndereco.Text.Length;
			lblEnderecoTamanhoRestante.Text = "(" + intTamanho.ToString() + ")";
			if (intTamanho > 0)
				lblEnderecoTamanhoRestante.ForeColor = Color.DarkGreen;
			else if (intTamanho < 0)
				lblEnderecoTamanhoRestante.ForeColor = Color.DarkRed;
			else
				lblEnderecoTamanhoRestante.ForeColor = Color.DimGray;
		}
		#endregion

		#endregion

		#region [ txtBairro ]

		#region [ txtBairro_Enter ]
		private void txtBairro_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtBairro_Leave ]
		private void txtBairro_Leave(object sender, EventArgs e)
		{
			txtBairro.Text = txtBairro.Text.Trim();
		}
		#endregion

		#region [ txtBairro_KeyDown ]
		private void txtBairro_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtCep);
		}
		#endregion

		#region [ txtBairro_KeyPress ]
		private void txtBairro_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtCep ]

		#region [ txtCep_Enter ]
		private void txtCep_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtCep_Leave ]
		private void txtCep_Leave(object sender, EventArgs e)
		{
			if (txtCep.Text.Length == 0) return;
			txtCep.Text = Global.formataCep(txtCep.Text);
			if (!Global.isCepOk(txtCep.Text))
			{
				avisoErro("CEP inválido!!");
				txtCep.Focus();
				return;
			}
		}
		#endregion

		#region [ txtCep_KeyDown ]
		private void txtCep_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtCidade);
		}
		#endregion

		#region [ txtCep_KeyPress ]
		private void txtCep_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCep(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtCidade ]

		#region [ txtCidade_Enter ]
		private void txtCidade_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtCidade_Leave ]
		private void txtCidade_Leave(object sender, EventArgs e)
		{
			txtCidade.Text = txtCidade.Text.Trim();
		}
		#endregion

		#region [ txtCidade_KeyDown ]
		private void txtCidade_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtUF);
		}
		#endregion

		#region [ txtCidade_KeyPress ]
		private void txtCidade_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtUF ]

		#region [ txtUF_Enter ]
		private void txtUF_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtUF_Leave ]
		private void txtUF_Leave(object sender, EventArgs e)
		{
			if (txtUF.Text.Length == 0) return;

			txtUF.Text = txtUF.Text.Trim().ToUpper();

			if (!Global.isUfOk(txtUF.Text))
			{
				avisoErro("UF inválida!!");
				txtUF.Focus();
				return;
			}
		}
		#endregion

		#region [ txtUF_KeyDown ]
		private void txtUF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnDummy);
		}
		#endregion

		#region [ txtUF_KeyPress ]
		private void txtUF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoSomenteLetras(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ btnCepPesquisa ]

		#region [ btnCepPesquisa_Click ]
		private void btnCepPesquisa_Click(object sender, EventArgs e)
		{
			trataBotaoPesquisaCep();
		}
		#endregion

		#endregion

		#region [ btnOk ]

		#region [ btnOk_Click ]
		private void btnOk_Click(object sender, EventArgs e)
		{
			trataBotaoOk();
		}
		#endregion

		#endregion

		#region [ btnCancela ]

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			trataBotaoCancela();
		}
		#endregion

		#endregion

		#endregion
	}
}
