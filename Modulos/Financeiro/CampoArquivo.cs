#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	#region [ Classe CampoArquivo ]
	public class CampoArquivo
	{
		#region [ Enumeradores ]

		#region [ eAlinhamento ]
		public enum eAlinhamento
		{
			DIREITA,
			ESQUERDA
		}
		#endregion

		#region [ ePreenchimento ]
		public enum ePreenchimento
		{
			BRANCO,
			ZERO
		}
		#endregion

		#endregion

		#region [ Atributos ]

		#region [ Getter/Setter: valor ]
		private String _valor = "";
		public String valor
		{
			get { return _valor; }
			set
			{
				if (value.Length > _tamanho)
				{
					throw new FinanceiroException("O valor informado (" + value.Length.ToString() + " posições) excede a capacidade do campo (" + _tamanho.ToString() + " posições)!!\n" + value);
				}
				else
				{
					_valor = value;
					switch (_alinhamento)
					{
						case eAlinhamento.DIREITA:
							_valor = _valor.PadLeft(_tamanho, decodificaCaracterPreenchimento(_preenchimento));
							break;
						case eAlinhamento.ESQUERDA:
							_valor = _valor.PadRight(_tamanho, decodificaCaracterPreenchimento(_preenchimento));
							break;
						default:
							_valor = _valor.PadRight(_tamanho, decodificaCaracterPreenchimento(_preenchimento));
							break;
					}
				}
			}
		}
		#endregion

		#region [ Getter/Setter: tamanho ]
		private int _tamanho;
		public int tamanho
		{
			get { return _tamanho; }
		}
		#endregion

		#region [ Getter/Setter: alinhamento ]
		private eAlinhamento _alinhamento;
		internal eAlinhamento alinhamento
		{
			get { return _alinhamento; }
		}
		#endregion

		#region [ Getter/Setter: preenchimento ]
		private ePreenchimento _preenchimento;
		internal ePreenchimento preenchimento
		{
			get { return _preenchimento; }
		}
		#endregion

		#endregion

		#region [ Construtor ]
		public CampoArquivo(int tamanho, String valorDefault, ePreenchimento preenchimento, eAlinhamento alinhamento)
		{
			if ((preenchimento == ePreenchimento.ZERO) && (alinhamento == eAlinhamento.ESQUERDA))
			{
				throw new Exception("Não é permitido fazer o preenchimento do campo com zeros à direita!!");
			}
			this._tamanho = tamanho;
			this._preenchimento = preenchimento;
			this._alinhamento = alinhamento;
			this.valor = valorDefault;
		}

		public CampoArquivo(int tamanho, ePreenchimento preenchimento, eAlinhamento alinhamento)
		{
			if ((preenchimento == ePreenchimento.ZERO) && (alinhamento == eAlinhamento.ESQUERDA))
			{
				throw new Exception("Não é permitido fazer o preenchimento do campo com zeros à direita!!");
			}
			this._tamanho = tamanho;
			this._preenchimento = preenchimento;
			this._alinhamento = alinhamento;
			this.valor = "";
		}

		public CampoArquivo(int tamanho, String valorDefault, ePreenchimento preenchimento)
		{
			if (preenchimento == ePreenchimento.ZERO)
			{
				throw new Exception("Não é permitido fazer o preenchimento do campo com zeros à direita!!");
			}
			this._tamanho = tamanho;
			this._preenchimento = preenchimento;
			this._alinhamento = eAlinhamento.ESQUERDA;
			this.valor = valorDefault;
		}

		public CampoArquivo(int tamanho, ePreenchimento preenchimento)
		{
			if (preenchimento == ePreenchimento.ZERO)
			{
				throw new Exception("Não é permitido fazer o preenchimento do campo com zeros à direita!!");
			}
			this._tamanho = tamanho;
			this._preenchimento = preenchimento;
			this._alinhamento = eAlinhamento.ESQUERDA;
			this.valor = "";
		}

		public CampoArquivo(int tamanho, String valorDefault)
		{
			this._tamanho = tamanho;
			this._preenchimento = ePreenchimento.BRANCO;
			this._alinhamento = eAlinhamento.ESQUERDA;
			this.valor = valorDefault;
		}

		public CampoArquivo(int tamanho)
		{
			this._tamanho = tamanho;
			this._preenchimento = ePreenchimento.BRANCO;
			this._alinhamento = eAlinhamento.ESQUERDA;
			this.valor = "";
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ decodificaCaracterPreenchimento ]
		private char decodificaCaracterPreenchimento(ePreenchimento opcao)
		{
			switch (opcao)
			{
				case ePreenchimento.BRANCO:
					return '\x20';
				case ePreenchimento.ZERO:
					return '\x30';
				default:
					return '\x20';
			}
		}
		#endregion

		#endregion

		#region [ Métodos Públicos ]

		#region [ ToString ]
		public override string ToString()
		{
			return _valor.ToString();
		}
		#endregion

		#region [ ConsomeCampo ]
		public string ConsomeCampo(string texto)
		{
			if (texto == null)
			{
				valor = "";
				return "";
			}

			if (texto.Length == tamanho)
			{
				valor = texto;
				return "";
			}

			if (texto.Length > tamanho)
			{
				valor = texto.Substring(0, tamanho);
				return texto.Substring(tamanho);
			}
			else
			{
				valor = texto;
				return "";
			}
		}
		#endregion

		#endregion
	}
	#endregion

	#region [ Classe LinhaArquivo ]
	public class LinhaArquivo
	{
		#region [ Atributos ]
		private List<CampoArquivo> listaCampos;
		#endregion

		#region [ Construtor ]
		public LinhaArquivo()
		{
			listaCampos = new List<CampoArquivo>();
		}
		#endregion

		#region [ Métodos Protected ]

		#region [ criaCampo ]
		protected void criaCampo(ref CampoArquivo c, int tamanho)
		{
			c = new CampoArquivo(tamanho);
			listaCampos.Add(c);
		}

		protected void criaCampo(ref CampoArquivo c, int tamanho, String valorDefault)
		{
			c = new CampoArquivo(tamanho, valorDefault);
			listaCampos.Add(c);
		}

		protected void criaCampo(ref CampoArquivo c, int tamanho, CampoArquivo.ePreenchimento preenchimento)
		{
			c = new CampoArquivo(tamanho, preenchimento);
			listaCampos.Add(c);
		}

		protected void criaCampo(ref CampoArquivo c, int tamanho, String valorDefault, CampoArquivo.ePreenchimento preenchimento)
		{
			c = new CampoArquivo(tamanho, valorDefault, preenchimento);
			listaCampos.Add(c);
		}

		protected void criaCampo(ref CampoArquivo c, int tamanho, CampoArquivo.ePreenchimento preenchimento, CampoArquivo.eAlinhamento alinhamento)
		{
			c = new CampoArquivo(tamanho, preenchimento, alinhamento);
			listaCampos.Add(c);
		}

		protected void criaCampo(ref CampoArquivo c, int tamanho, String valorDefault, CampoArquivo.ePreenchimento preenchimento, CampoArquivo.eAlinhamento alinhamento)
		{
			c = new CampoArquivo(tamanho, valorDefault, preenchimento, alinhamento);
			listaCampos.Add(c);
		}
		#endregion

		#endregion

		#region [ Métodos Públicos ]

		#region [ ToString ]
		public override string ToString()
		{
			StringBuilder sbResposta = new StringBuilder("");
			for (int i = 0; i < listaCampos.Count; i++)
			{
				sbResposta.Append(listaCampos[i].valor);
			}
			return sbResposta.ToString();
		}
		#endregion

		#region [ CarregaDados ]
		public void CarregaDados(string linhaDados)
		{
			string strMsg;
			string linhaDadosAux = linhaDados.ToString();

			if (linhaDadosAux.Length != CalculaTamanhoTotal())
			{
				strMsg = "A linha de dados não é compatível com os campos processados!!\nTamanho da linha de dados: " + linhaDadosAux.Length.ToString() + "\nTamanho total dos campos processados: " + CalculaTamanhoTotal().ToString();
				throw new Exception(strMsg);
			}

			for (int i = 0; i < listaCampos.Count; i++)
			{
				linhaDadosAux = listaCampos[i].ConsomeCampo(linhaDadosAux);
			}
		}
		#endregion

		#region [ CalculaTamanhoTotal ]
		public int CalculaTamanhoTotal()
		{
			int intTamanhoResposta = 0;
			for (int i = 0; i < listaCampos.Count; i++)
			{
				intTamanhoResposta += listaCampos[i].tamanho;
			}
			return intTamanhoResposta;
		}
		#endregion

		#endregion
	}
	#endregion

	#region [ Classe LinhaHeaderArquivoRemessa ]
	public class LinhaHeaderArquivoRemessa : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo identificacaoArquivoRemessa;
		public CampoArquivo literalRemessa;
		public CampoArquivo codigoServico;
		public CampoArquivo literalServico;
		public CampoArquivo codigoEmpresa;
		public CampoArquivo nomeEmpresa;
		public CampoArquivo numeroBanco;
		public CampoArquivo nomeBanco;
		public CampoArquivo dataGravacaoArquivo;
		public CampoArquivo filler_1;
		public CampoArquivo identificacaoSistema;
		public CampoArquivo numSequencialRemessa;
		public CampoArquivo filler_2;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaHeaderArquivoRemessa()
			: base()
		{
			criaCampo(ref identificacaoRegistro, 1, "0");
			criaCampo(ref identificacaoArquivoRemessa, 1, "1");
			criaCampo(ref literalRemessa, 7, "REMESSA");
			criaCampo(ref codigoServico, 2, "01");
			criaCampo(ref literalServico, 15, "COBRANCA");
			criaCampo(ref codigoEmpresa, 20, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref nomeEmpresa, 30);
			criaCampo(ref numeroBanco, 3, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref nomeBanco, 15);
			criaCampo(ref dataGravacaoArquivo, 6);
			criaCampo(ref filler_1, 8);
			criaCampo(ref identificacaoSistema, 2, "MX");
			criaCampo(ref numSequencialRemessa, 7, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref filler_2, 277);
			criaCampo(ref numSequencialRegistro, 6, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaRegistroTipo1ArquivoRemessa ]
	public class LinhaRegistroTipo1ArquivoRemessa : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo agenciaDebito;
		public CampoArquivo digitoAgenciaDebito;
		public CampoArquivo razaoContaCorrente;
		public CampoArquivo contaCorrente;
		public CampoArquivo digitoContaCorrente;
		public CampoArquivo identifCedenteFiller;
		public CampoArquivo identifCedenteCarteira;
		public CampoArquivo identifCedenteAgencia;
		public CampoArquivo identifCedenteCtaCorrente;
		public CampoArquivo identifCedenteDigitoCtaCorrente;
		public CampoArquivo numControleParticipante;
		public CampoArquivo codigoBancoDebitado;
		public CampoArquivo campoMulta;
		public CampoArquivo percentualMulta;
		public CampoArquivo nossoNumeroSemDigito;
		public CampoArquivo digitoNossoNumero;
		public CampoArquivo descontoBonificacaoPorDia;
		public CampoArquivo condicaoEmissaoPapeletaCobranca;
		public CampoArquivo identSeEmitePapeletaParaDebAutomatico;
		public CampoArquivo identificacaoOperacaoBanco;
		public CampoArquivo indicadorRateioCredito;
		public CampoArquivo enderecamentoParaAvisoDebAutomaticoEmCtaCorrente;
		public CampoArquivo filler_1;
		public CampoArquivo identificacaoOcorrencia;
		public CampoArquivo numDocumento;
		public CampoArquivo dataVenctoTitulo;
		public CampoArquivo valorTitulo;
		public CampoArquivo bancoEncarregadoCobranca;
		public CampoArquivo agenciaDepositaria;
		public CampoArquivo especieTitulo;
		public CampoArquivo identificacaoAceitoNaoAceito;
		public CampoArquivo dataEmissaoTitulo;
		public CampoArquivo primeiraInstrucao;
		public CampoArquivo segundaInstrucao;
		public CampoArquivo valorPorDiaAtraso;
		public CampoArquivo dataLimiteConcessaoDesconto;
		public CampoArquivo valorDesconto;
		public CampoArquivo valorIOF;
		public CampoArquivo valorAbatimento;
		public CampoArquivo identificacaoTipoInscricaoSacado;
		public CampoArquivo numInscricaoSacado;
		public CampoArquivo nomeSacado;
		public CampoArquivo enderecoCompleto;
		public CampoArquivo primeiraMensagem;
		public CampoArquivo cep;
		public CampoArquivo sufixoCep;
		public CampoArquivo sacadorAvalistaOuSegundaMensagem;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaRegistroTipo1ArquivoRemessa()
		{
			criaCampo(ref identificacaoRegistro, 1, "1");
			criaCampo(ref agenciaDebito, 5, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref digitoAgenciaDebito, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref razaoContaCorrente, 5, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref contaCorrente, 7, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref digitoContaCorrente, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identifCedenteFiller, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identifCedenteCarteira, 3, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identifCedenteAgencia, 5, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identifCedenteCtaCorrente, 7, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identifCedenteDigitoCtaCorrente, 1, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref numControleParticipante, 25);
			criaCampo(ref codigoBancoDebitado, 3, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref campoMulta, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref percentualMulta, 4, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref nossoNumeroSemDigito, 11, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref digitoNossoNumero, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref descontoBonificacaoPorDia, 10, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref condicaoEmissaoPapeletaCobranca, 1, "1");
			criaCampo(ref identSeEmitePapeletaParaDebAutomatico, 1, " ");
			criaCampo(ref identificacaoOperacaoBanco, 10);
			criaCampo(ref indicadorRateioCredito, 1, " ");
			criaCampo(ref enderecamentoParaAvisoDebAutomaticoEmCtaCorrente, 1, "0");
			criaCampo(ref filler_1, 2);
			criaCampo(ref identificacaoOcorrencia, 2, "01");
			criaCampo(ref numDocumento, 10);
			criaCampo(ref dataVenctoTitulo, 6);
			criaCampo(ref valorTitulo, 13, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref bancoEncarregadoCobranca, 3, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref agenciaDepositaria, 5, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref especieTitulo, 2, "01", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identificacaoAceitoNaoAceito, 1, "N");
			criaCampo(ref dataEmissaoTitulo, 6);
			criaCampo(ref primeiraInstrucao, 2);
			criaCampo(ref segundaInstrucao, 2);
			criaCampo(ref valorPorDiaAtraso, 13, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref dataLimiteConcessaoDesconto, 6, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref valorDesconto, 13, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref valorIOF, 13, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref valorAbatimento, 13, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref identificacaoTipoInscricaoSacado, 2);
			criaCampo(ref numInscricaoSacado, 14, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref nomeSacado, 40);
			criaCampo(ref enderecoCompleto, 40);
			criaCampo(ref primeiraMensagem, 12);
			criaCampo(ref cep, 5);
			criaCampo(ref sufixoCep, 3);
			criaCampo(ref sacadorAvalistaOuSegundaMensagem, 60);
			criaCampo(ref numSequencialRegistro, 6, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaRegistroTipo2ArquivoRemessa ]
	public class LinhaRegistroTipo2ArquivoRemessa : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo tipoRegistro;
		public CampoArquivo mensagem_1;
		public CampoArquivo mensagem_2;
		public CampoArquivo mensagem_3;
		public CampoArquivo mensagem_4;
		public CampoArquivo filler_1;
		public CampoArquivo carteira;
		public CampoArquivo agencia;
		public CampoArquivo contaCorrente;
		public CampoArquivo digitoContaCorrente;
		public CampoArquivo nossoNumero;
		public CampoArquivo digitoNossoNumero;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaRegistroTipo2ArquivoRemessa()
		{
			criaCampo(ref tipoRegistro, 1, "2");
			criaCampo(ref mensagem_1, 80);
			criaCampo(ref mensagem_2, 80);
			criaCampo(ref mensagem_3, 80);
			criaCampo(ref mensagem_4, 80);
			criaCampo(ref filler_1, 45);
			criaCampo(ref carteira, 3, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref agencia, 5, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref contaCorrente, 7, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref digitoContaCorrente, 1, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref nossoNumero, 11, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref digitoNossoNumero, 1, "0", CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
			criaCampo(ref numSequencialRegistro, 6, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaTraillerArquivoRemessa ]
	public class LinhaTraillerArquivoRemessa : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo filler_1;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaTraillerArquivoRemessa()
		{
			criaCampo(ref identificacaoRegistro, 1, "9");
			criaCampo(ref filler_1, 393);
			criaCampo(ref numSequencialRegistro, 6, CampoArquivo.ePreenchimento.ZERO, CampoArquivo.eAlinhamento.DIREITA);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaHeaderArquivoRetorno ]
	public class LinhaHeaderArquivoRetorno : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo identificacaoArquivoRetorno;
		public CampoArquivo literalRetorno;
		public CampoArquivo codigoServico;
		public CampoArquivo literalServico;
		public CampoArquivo codigoEmpresa;
		public CampoArquivo nomeEmpresa;
		public CampoArquivo numBanco;
		public CampoArquivo nomeBanco;
		public CampoArquivo dataGravacaoArquivo;
		public CampoArquivo filler_1;
		public CampoArquivo numAvisoBancario;
		public CampoArquivo filler_2;
		public CampoArquivo dataCredito;
		public CampoArquivo filler_3;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaHeaderArquivoRetorno()
		{
			criaCampo(ref identificacaoRegistro, 1);
			criaCampo(ref identificacaoArquivoRetorno, 1);
			criaCampo(ref literalRetorno, 7);
			criaCampo(ref codigoServico, 2);
			criaCampo(ref literalServico, 15);
			criaCampo(ref codigoEmpresa, 20);
			criaCampo(ref nomeEmpresa, 30);
			criaCampo(ref numBanco, 3);
			criaCampo(ref nomeBanco, 15);
			criaCampo(ref dataGravacaoArquivo, 6);
			criaCampo(ref filler_1, 8);
			criaCampo(ref numAvisoBancario, 5);
			criaCampo(ref filler_2, 266);
			criaCampo(ref dataCredito, 6);
			criaCampo(ref filler_3, 9);
			criaCampo(ref numSequencialRegistro, 6);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaRegistroTipo1ArquivoRetorno ]
	public class LinhaRegistroTipo1ArquivoRetorno : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo tipoInscricaoEmpresa;
		public CampoArquivo numInscricaoEmpresa;
		public CampoArquivo filler_1;
		public CampoArquivo identifCedenteFiller;
		public CampoArquivo identifCedenteCarteira;
		public CampoArquivo identifCedenteAgencia;
		public CampoArquivo identifCedenteCtaCorrente;
		public CampoArquivo identifCedenteDigitoCtaCorrente;
		public CampoArquivo numControleParticipante;
		public CampoArquivo filler_2;
		public CampoArquivo nossoNumeroSemDigito;
		public CampoArquivo digitoNossoNumero;
		public CampoArquivo usoDoBanco_1;
		public CampoArquivo usoDoBanco_2;
		public CampoArquivo indicadorRateioCredito;
		public CampoArquivo filler_3;
		public CampoArquivo carteira;
		public CampoArquivo identificacaoOcorrencia;
		public CampoArquivo dataOcorrencia;
		public CampoArquivo numeroDocumento;
		public CampoArquivo identificacaoTitulo;
		public CampoArquivo dataVenctoTitulo;
		public CampoArquivo valorTitulo;
		public CampoArquivo bancoCobrador;
		public CampoArquivo agenciaCobradora;
		public CampoArquivo especieTitulo;
		public CampoArquivo valorDespesasCobranca;
		public CampoArquivo valorOutrasDespesas;
		public CampoArquivo valorJurosOperacaoEmAtraso;
		public CampoArquivo valorIofDevido;
		public CampoArquivo valorAbatimentoConcedido;
		public CampoArquivo valorDescontoConcedido;
		public CampoArquivo valorPago;
		public CampoArquivo valorMora;
		public CampoArquivo valorOutrosCreditos;
		public CampoArquivo filler_4;
		public CampoArquivo motivoCodigoOcorrencia19;
		public CampoArquivo dataCredito;
		public CampoArquivo origemPagamento;
		public CampoArquivo filler_5;
		public CampoArquivo quandoChequeBradesco;
		public CampoArquivo motivosRejeicoes;
		public CampoArquivo filler_6;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaRegistroTipo1ArquivoRetorno()
		{
			criaCampo(ref identificacaoRegistro, 1);
			criaCampo(ref tipoInscricaoEmpresa, 2);
			criaCampo(ref numInscricaoEmpresa, 14);
			criaCampo(ref filler_1, 3);
			criaCampo(ref identifCedenteFiller, 1);
			criaCampo(ref identifCedenteCarteira, 3);
			criaCampo(ref identifCedenteAgencia, 5);
			criaCampo(ref identifCedenteCtaCorrente, 7);
			criaCampo(ref identifCedenteDigitoCtaCorrente, 1);
			criaCampo(ref numControleParticipante, 25);
			criaCampo(ref filler_2, 8);
			criaCampo(ref nossoNumeroSemDigito, 11);
			criaCampo(ref digitoNossoNumero, 1);
			criaCampo(ref usoDoBanco_1, 10);
			criaCampo(ref usoDoBanco_2, 12);
			criaCampo(ref indicadorRateioCredito, 1);
			criaCampo(ref filler_3, 2);
			criaCampo(ref carteira, 1);
			criaCampo(ref identificacaoOcorrencia, 2);
			criaCampo(ref dataOcorrencia, 6);
			criaCampo(ref numeroDocumento, 10);
			criaCampo(ref identificacaoTitulo, 20);
			criaCampo(ref dataVenctoTitulo, 6);
			criaCampo(ref valorTitulo, 13);
			criaCampo(ref bancoCobrador, 3);
			criaCampo(ref agenciaCobradora, 5);
			criaCampo(ref especieTitulo, 2);
			criaCampo(ref valorDespesasCobranca, 13);
			criaCampo(ref valorOutrasDespesas, 13);
			criaCampo(ref valorJurosOperacaoEmAtraso, 13);
			criaCampo(ref valorIofDevido, 13);
			criaCampo(ref valorAbatimentoConcedido, 13);
			criaCampo(ref valorDescontoConcedido, 13);
			criaCampo(ref valorPago, 13);
			criaCampo(ref valorMora, 13);
			criaCampo(ref valorOutrosCreditos, 13);
			criaCampo(ref filler_4, 2);
			criaCampo(ref motivoCodigoOcorrencia19, 1);
			criaCampo(ref dataCredito, 6);
			criaCampo(ref origemPagamento, 3);
			criaCampo(ref filler_5, 10);
			criaCampo(ref quandoChequeBradesco, 4);
			criaCampo(ref motivosRejeicoes, 10);
			criaCampo(ref filler_6, 66);
			criaCampo(ref numSequencialRegistro, 6);
		}
		#endregion
	}
	#endregion

	#region [ Classe LinhaTraillerArquivoRetorno ]
	public class LinhaTraillerArquivoRetorno : LinhaArquivo
	{
		#region [ Atributos ]
		public CampoArquivo identificacaoRegistro;
		public CampoArquivo identificacaoRetorno;
		public CampoArquivo identificacaoTipoRegistro;
		public CampoArquivo codigoBanco;
		public CampoArquivo filler_1;
		public CampoArquivo qtdeTitulosEmCobranca;
		public CampoArquivo valorTotalEmCobranca;
		public CampoArquivo numAvisoBancario;
		public CampoArquivo filler_2;
		public CampoArquivo qtdeRegsOcorrencia02ConfirmacaoEntradas;
		public CampoArquivo valorRegsOcorrencia02ConfirmacaoEntradas;
		public CampoArquivo valorRegsOcorrencia06Liquidacao;
		public CampoArquivo qtdeRegsOcorrencia06Liquidacao;
		public CampoArquivo valorRegsOcorrencia06;
		public CampoArquivo qtdeRegsOcorrencia09e10TitulosBaixados;
		public CampoArquivo valorRegsOcorrencia09e10TitulosBaixados;
		public CampoArquivo qtdeRegsOcorrencia13AbatimentoCancelado;
		public CampoArquivo valorRegsOcorrencia13AbatimentoCancelado;
		public CampoArquivo qtdeRegsOcorrencia14VenctoAlterado;
		public CampoArquivo valorRegsOcorrencia14VenctoAlterado;
		public CampoArquivo qtdeRegsOcorrencia12AbatimentoConcedido;
		public CampoArquivo valorRegsOcorrencia12AbatimentoConcedido;
		public CampoArquivo qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto;
		public CampoArquivo valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto;
		public CampoArquivo filler_3;
		public CampoArquivo valorTotalRateiosEfetuados;
		public CampoArquivo qtdeTotalRateiosEfetuados;
		public CampoArquivo filler_4;
		public CampoArquivo numSequencialRegistro;
		#endregion

		#region [ Construtor ]
		public LinhaTraillerArquivoRetorno()
		{
			criaCampo(ref identificacaoRegistro, 1);
			criaCampo(ref identificacaoRetorno, 1);
			criaCampo(ref identificacaoTipoRegistro, 2);
			criaCampo(ref codigoBanco, 3);
			criaCampo(ref filler_1, 10);
			criaCampo(ref qtdeTitulosEmCobranca, 8);
			criaCampo(ref valorTotalEmCobranca, 14);
			criaCampo(ref numAvisoBancario, 8);
			criaCampo(ref filler_2, 10);
			criaCampo(ref qtdeRegsOcorrencia02ConfirmacaoEntradas, 5);
			criaCampo(ref valorRegsOcorrencia02ConfirmacaoEntradas, 12);
			criaCampo(ref valorRegsOcorrencia06Liquidacao, 12);
			criaCampo(ref qtdeRegsOcorrencia06Liquidacao, 5);
			criaCampo(ref valorRegsOcorrencia06, 12);
			criaCampo(ref qtdeRegsOcorrencia09e10TitulosBaixados, 5);
			criaCampo(ref valorRegsOcorrencia09e10TitulosBaixados, 12);
			criaCampo(ref qtdeRegsOcorrencia13AbatimentoCancelado, 5);
			criaCampo(ref valorRegsOcorrencia13AbatimentoCancelado, 12);
			criaCampo(ref qtdeRegsOcorrencia14VenctoAlterado, 5);
			criaCampo(ref valorRegsOcorrencia14VenctoAlterado, 12);
			criaCampo(ref qtdeRegsOcorrencia12AbatimentoConcedido, 5);
			criaCampo(ref valorRegsOcorrencia12AbatimentoConcedido, 12);
			criaCampo(ref qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto, 5);
			criaCampo(ref valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto, 12);
			criaCampo(ref filler_3, 174);
			criaCampo(ref valorTotalRateiosEfetuados, 15);
			criaCampo(ref qtdeTotalRateiosEfetuados, 8);
			criaCampo(ref filler_4, 9);
			criaCampo(ref numSequencialRegistro, 6);
		}
		#endregion
	}
	#endregion
}
