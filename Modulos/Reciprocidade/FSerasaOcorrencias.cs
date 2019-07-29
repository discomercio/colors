#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Media;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FSerasaOcorrencias : FModelo
	{
		#region [ Atributos ]
		#region [ Diversos ]
		private bool _atualizacaoAutomaticaPesquisaEmAndamento = false;
		FSerasaTrataOcorrencia _fSerasaTrataOcorrencia;
		Dictionary<String, String> _dictErros;
		#endregion
		#endregion

		#region [ Construtor ]
		public FSerasaOcorrencias()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos ]
		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			int intQtdeRegistros = 0;
			DataTable dtbConsulta = new DataTable();
			DataTable dtbErros = new DataTable();
			DataRow rowConsulta;
			#endregion

			if (_atualizacaoAutomaticaPesquisaEmAndamento) return false;

			_atualizacaoAutomaticaPesquisaEmAndamento = true;
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				dtbConsulta = TituloMovimentoDAO.selecionaBoletosParaTratamento();
				if (dtbConsulta.Rows.Count == 0)
				{
					aviso("Não há ocorrências a serem tratadas!!");
					gridDados.Rows.Clear();
					lblTotalizacaoRegistros.Text = "";
					return false;
				}
				int id_serasa_arq_retorno_normal = BD.readToInt(dtbConsulta.Rows[0]["id_serasa_arq_retorno_normal"]);
				dtbErros = TabErrosDAO.selecionaTabErrosPorArqRetorno(id_serasa_arq_retorno_normal);
				_dictErros = criaDicionarioErros(dtbErros);

				#region [ Exibição dos dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					gridDados.SuspendLayout();

					#region [ Carrega os dados no grid ]
					gridDados.Rows.Clear();
					if (dtbConsulta.Rows.Count > 0) gridDados.Rows.Add(dtbConsulta.Rows.Count);
					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						rowConsulta = dtbConsulta.Rows[i];
						gridDados.Rows[i].Cells["id"].Value = BD.readToString(rowConsulta["id"]);
						gridDados.Rows[i].Cells["nosso_numero"].Value = BD.readToString(rowConsulta["nosso_numero"]) + "-" + BD.readToString(rowConsulta["digito_nosso_numero"]);
						gridDados.Rows[i].Cells["dt_emissao"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_emissao"]));
						gridDados.Rows[i].Cells["dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_vencto"]));
						gridDados.Rows[i].Cells["vl_titulo"].Value = Global.formataMoeda(BD.readToDecimal(rowConsulta["vl_titulo"]));
						gridDados.Rows[i].Cells["dt_pagto"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_pagto"]));
						gridDados.Rows[i].Cells["vl_pago"].Value = Global.formataMoeda(BD.readToDecimal(rowConsulta["vl_pago"]));
						gridDados.Rows[i].Cells["primeira_ocorrencia"].Value = _dictErros[BD.readToString(rowConsulta["retorno_codigos_erro"]).Substring(0, 3)]; //mostra a primeira ocorrencia
						gridDados.Rows[i].Cells["todas_ocorrencias"].Value = BD.readToString(rowConsulta["retorno_codigos_erro"]);

						intQtdeRegistros++;
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < gridDados.Rows.Count; i++)
					{
						if (gridDados.Rows[i].Selected) gridDados.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					gridDados.ResumeLayout();
				}
				#endregion

				#region [ Exibe totalização ]
				lblTotalizacaoRegistros.Text = intQtdeRegistros.ToString();
				#endregion

				gridDados.Focus();

				// Feedback da conclusão da pesquisa
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return false;
			}
			finally
			{
				_atualizacaoAutomaticaPesquisaEmAndamento = false;
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ criaDicionarioErros ]
		private Dictionary<String, String> criaDicionarioErros(DataTable DtbErros)
		{
			Dictionary<String, String> dict = new Dictionary<String, String>();
			for (int i = 0; i < DtbErros.Rows.Count; i++)
			{
				DataRow dr = DtbErros.Rows[i];
				String codErro = BD.readToString(dr["numero_mensagem"]);
				String msgErro = BD.readToString(dr["descricao_msg_erro"]);

				dict.Add(codErro, msgErro);
			}
			return dict;
		}
		#endregion

		#region [ trataOcorrenciaSelecionada ]
		private void trataOcorrenciaSelecionada()
		{
			int id = 0;
			String numBoleto = "";
			DateTime dtEmissao = DateTime.MinValue;
			DateTime dtVencto = DateTime.MinValue;
			Decimal vlTitulo = 0;
			DateTime dtPagto = DateTime.MinValue;
			Decimal vlPago = 0;
			String linhaCodigosErros = "";
			String numBoletoCorrigido;
			DateTime dtEmissaoCorrigido;
			DateTime dtVenctoCorrigido;
			Decimal vlTituloCorrigido;
			DateTime dtPagtoCorrigido;
			Decimal vlPagoCorrigido;
			DialogResult drResultado;
			bool blnSucesso = false;

			#region [ Consistência ]
			if (gridDados.SelectedRows.Count == 0)
			{
				avisoErro("Nenhum registro foi selecionado!!");
				return;
			}

			if (gridDados.SelectedRows.Count > 1)
			{
				avisoErro("Não é permitida a seleção de múltiplos registros!!");
				return;
			}
			#endregion

			try
			{
				#region [ Obtém dados a serem editados do registro selecionado ]
				foreach (DataGridViewRow item in gridDados.SelectedRows)
				{
					id = Convert.ToInt32(item.Cells["id"].Value);
					numBoleto = item.Cells["nosso_numero"].Value.ToString();
					dtEmissao = Global.converteDdMmYyyyParaDateTime(item.Cells["dt_emissao"].Value.ToString());
					dtVencto = Global.converteDdMmYyyyParaDateTime(item.Cells["dt_vencto"].Value.ToString());
					vlTitulo = Global.converteNumeroDecimal(item.Cells["vl_titulo"].Value.ToString());
					dtPagto = Global.converteDdMmYyyyParaDateTime(item.Cells["dt_pagto"].Value.ToString());
					vlPago = Global.converteNumeroDecimal(item.Cells["vl_pago"].Value.ToString());
					linhaCodigosErros = item.Cells["todas_ocorrencias"].Value.ToString();
				}
				#endregion

				#region [ Exibe painel p/ reeditar o endereço ]
				_fSerasaTrataOcorrencia = new FSerasaTrataOcorrencia(id,
																	 numBoleto,
																	 dtEmissao,
																	 dtVencto,
																	 vlTitulo,
																	 dtPagto,
																	 vlPago,
																	 linhaCodigosErros,
																	 _dictErros);

				_fSerasaTrataOcorrencia.StartPosition = FormStartPosition.Manual;
				_fSerasaTrataOcorrencia.Left = this.Left + (this.Width - _fSerasaTrataOcorrencia.Width) / 2;
				_fSerasaTrataOcorrencia.Top = this.Top + (this.Height - _fSerasaTrataOcorrencia.Height) / 2;
				drResultado = _fSerasaTrataOcorrencia.ShowDialog();
				if (drResultado != DialogResult.OK) return;
				#endregion

				#region [ Altera o endereço e reseta o status p/ poder ser enviado novamente no arquivo remessa ]
				numBoletoCorrigido = _fSerasaTrataOcorrencia._numBoletoCorrigido;
				dtEmissaoCorrigido = _fSerasaTrataOcorrencia._dtEmissaoCorrigido;
				dtVenctoCorrigido = _fSerasaTrataOcorrencia._dtVenctoCorrigido;
				vlTituloCorrigido = _fSerasaTrataOcorrencia._vlTituloCorrigido;
				dtPagtoCorrigido = _fSerasaTrataOcorrencia._dtPagtoCorrigido;
				vlPagoCorrigido = _fSerasaTrataOcorrencia._vlPagoCorrigido;

				BD.iniciaTransacao();
				try
				{
					if (!TituloMovimentoDAO.trataOcorrencia(numBoletoCorrigido,
													   dtEmissaoCorrigido,
													   dtVenctoCorrigido,
													   vlTituloCorrigido,
													   dtPagtoCorrigido,
													   vlPagoCorrigido,
													   id))
					{
						throw new Exception("Falha na tentativa de gravação do tratamento da ocorrência!!");
					}

					blnSucesso = true;
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}
				#endregion

				#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
				executaPesquisa();
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
			}
		}
		#endregion
		#endregion

		#region [ Eventos ]
		#region [ Form FSerasaOcorrencias ]
		#region [ FSerasaOcorrencias_FormClosing ]
		private void FSerasaOcorrencias_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain._fMain.Location = this.Location;
			FMain._fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FSerasaOcorrencias_KeyDown ]
		private void FSerasaOcorrencias_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				e.SuppressKeyPress = true;
				executaPesquisa();
				return;
			}
		}
		#endregion

		#region [ FSerasaOcorrencias_Shown ]
		private void FSerasaOcorrencias_Shown(object sender, EventArgs e)
		{
			lblTotalizacaoRegistros.Text = "";
		}
		#endregion
		#endregion

		#region [ gridDados ]
		#region [ gridDados_KeyDown ]
		private void gridDados_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				trataOcorrenciaSelecionada();
				return;
			}
		}
		#endregion

		#region [ gridDados_DoubleClick ]
		private void gridDados_DoubleClick(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion
		#endregion

		#region [ Botões / Menu ]
		#region [ Pesquisar ]
		#region [ btnPesquisar_Click ]
		private void btnPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#region [ menuOcorrenciaPesquisar_Click ]
		private void menuOcorrenciaPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion
		#endregion

		#region [ Tratar Ocorrência ]
		#region [ btnOcorrenciaTratar_Click ]
		private void btnOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion

		#region [ menuOcorrenciaTratar_Click ]
		private void menuOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion
		#endregion
		#endregion
		#endregion
	}
}
