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
	public partial class FBoletoTrataOcorrenciaValaComum : Form
	{
		#region [ Atributos ]

		#region [ Diversos ]
		private bool _InicializacaoOk;
		#endregion

		#region [ Dados p/ exibição/edição ]
		private BoletoCedente _boletoCedente;
		private String _nomeCliente;
		private String _cnpjCpf;
		private String _linhaTextoRegistroArquivo;
		private LinhaRegistroTipo1ArquivoRetorno _linhaRegTipo1ArqRetorno;
		public String comentarioOcorrenciaTratada = "";
		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoTrataOcorrenciaValaComum(
							int id_boleto_cedente,
							String nomeCliente,
							String cnpjCpf,
							String linhaTextoRegistroArquivo
							)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			_boletoCedente = BoletoCedenteDAO.getBoletoCedente(id_boleto_cedente);
			_nomeCliente = nomeCliente;
			_cnpjCpf = cnpjCpf;
			_linhaTextoRegistroArquivo = linhaTextoRegistroArquivo;
			_linhaRegTipo1ArqRetorno = new LinhaRegistroTipo1ArquivoRetorno();
			_linhaRegTipo1ArqRetorno.CarregaDados(_linhaTextoRegistroArquivo);
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
			txtIdentificacaoOcorrencia.Text = "";
			txtDadosRegistro.Text = "";
			txtComentarioOcorrenciaTratada.Text = "";
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
			comentarioOcorrenciaTratada = txtComentarioOcorrenciaTratada.Text;
			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoTrataOcorrenciaValaComum ]

		#region [ FBoletoTrataOcorrenciaValaComum_Load ]
		private void FBoletoTrataOcorrenciaValaComum_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
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

		#region [ FBoletoTrataOcorrenciaValaComum_Shown ]
		private void FBoletoTrataOcorrenciaValaComum_Shown(object sender, EventArgs e)
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
					txtIdentificacaoOcorrencia.Text = Global.montaDescricaoOcorrenciaBoleto(_linhaRegTipo1ArqRetorno.identificacaoOcorrencia.valor, _linhaRegTipo1ArqRetorno.motivosRejeicoes.valor, _linhaRegTipo1ArqRetorno.motivoCodigoOcorrencia19.valor).Replace("\n", "\r\n");
					txtDadosRegistro.Text = "Número documento: " + _linhaRegTipo1ArqRetorno.numeroDocumento.valor +
											"\r\n" +
											"Nosso número: " + Global.formataBoletoNossoNumero(_linhaRegTipo1ArqRetorno.nossoNumeroSemDigito.valor, _linhaRegTipo1ArqRetorno.digitoNossoNumero.valor) +
											"\r\n" +
											"Data da ocorrência no banco: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaRegTipo1ArqRetorno.dataOcorrencia.valor)) +
											"\r\n" +
											"Data vencimento: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaRegTipo1ArqRetorno.dataVenctoTitulo.valor)) +
											"\r\n" +
											"Valor do título: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorTitulo.valor)) +
											"\r\n" +
											"Valor pago: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorPago.valor)) +
											"\r\n" +
											"Abatimento concedido: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorAbatimentoConcedido.valor)) +
											"\r\n" +
											"Desconto concedido: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorDescontoConcedido.valor)) +
											"\r\n" +
											"Juros de mora: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorMora.valor)) +
											"\r\n" +
											"Despesas de cobrança: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorDespesasCobranca.valor)) +
											"\r\n" +
											"Outras despesas / custas de protesto: " + Global.formataMoeda(Global.decodificaCampoMonetario(_linhaRegTipo1ArqRetorno.valorOutrasDespesas.valor));
					#endregion

					#region [ Posiciona foco ]
					txtComentarioOcorrenciaTratada.Focus();
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

		#region [ txtComentarioOcorrenciaTratada ]

		#region [ txtComentarioOcorrenciaTratada_Enter ]
		private void txtComentarioOcorrenciaTratada_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtComentarioOcorrenciaTratada_Leave ]
		private void txtComentarioOcorrenciaTratada_Leave(object sender, EventArgs e)
		{
			txtComentarioOcorrenciaTratada.Text = txtComentarioOcorrenciaTratada.Text.Trim();
		}
		#endregion

		#region [ txtComentarioOcorrenciaTratada_KeyPress ]
		private void txtComentarioOcorrenciaTratada_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
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
