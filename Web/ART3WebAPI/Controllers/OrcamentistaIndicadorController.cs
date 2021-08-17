using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Domains;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using System.Text;
using System.Web.Script.Serialization;

namespace ART3WebAPI.Controllers
{
	public class OrcamentistaIndicadorController : ApiController
	{
		#region [ Teste ]
		[HttpGet]
		public HttpResponseMessage Teste()
		{
			const string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.Teste()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida";
			Global.gravaLogAtividade(httpRequestId, msg);

			HttpResponseMessage result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ PesquisaPorCnpjCpf ]
		/// <summary>
		/// Pesquisa orçamentista/indicador por CNPJ/CPF
		/// </summary>
		/// <param name="cnpjCpf">CNPJ/CPF a ser pesquisado (parâmetro obrigatório)</param>
		/// <param name="loja">Nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="vendedor">Vendedor a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="status">Status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'dinâmico' opcional)</param>
		/// <param name="parametrosEstaticos">
		/// Parâmetro obrigatório contendo os parâmetros estáticos criptografados para evitar consultas que burlem as restrições impostas pelo perfil de acesso.
		/// Após descriptografar este campo, os seguintes parâmetros devem tratados:
		///		loja = nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		vendedor = vendedor a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		status = status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'estático' opcional)
		///		usuario = usuário que está fazendo a requisição e que será autenticado antes de processar a solicitação (parâmetro 'estático' obrigatório)
		///	Os parâmetros estáticos possuem prioridade sobre os parâmetros 'dinâmicos', pois eles foram criados para garantir as restrições de acesso aos dados.
		///	Os parâmetros estáticos devem ser definidos e criptografados no processamento feito no lado do servidor (server side) p/ que não haja possibilidade do usuário fazer alterações.
		///	Já os parâmetros dinâmicos são os parâmetros que podem ser escolhidos pelo usuário, por exemplo, em relatórios onde se seleciona primeiro um vendedor para o sistema exibir a
		///	lista de todos os indicadores vinculados a esse vendedor.
		/// </param>
		/// <param name="sessionToken">SessionToken da sessão do usuário em andamento no sistema</param>
		/// <returns>Retorna uma lista de orçamentistas/indicadores que possuem o CNPJ/CPF especificado e que atendem também aos demais critérios especificados</returns>
		[HttpGet]
		public HttpResponseMessage PesquisaPorCnpjCpf(string cnpjCpf, string loja, string vendedor, string status, string parametrosEstaticos, string sessionToken)
		{
			#region [ Declarações ]
			Guid opGuid = Guid.NewGuid();
			string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.PesquisaPorCnpjCpf() [" + opGuid.ToString() + "]";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			string lojaParamEstatico;
			string vendedorParamEstatico;
			string statusParamEstatico;
			string usuario;
			string sParamEstaticoDecripto;
			string[] vParamEstatico;
			string sParam;
			string[] vParam;
			string sKey;
			string sValue;
			KeyValuePair<string, string> kvpParamEstatico;
			Usuario usuarioBD;
			List<KeyValuePair<string, string>> listaKvpParamEstatico = new List<KeyValuePair<string, string>>();
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador;
			HttpResponseMessage result;
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: cnpjCpf=" + (cnpjCpf ?? "") + ", loja=" + (loja ?? "") + ", vendedor=" + (vendedor ?? "") + ", status=" + (status ?? "") + ", parametrosEstaticos=" + (parametrosEstaticos ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			#region [ CNPJ/CPF válido? ]
			if ((cnpjCpf ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o CNPJ/CPF para pesquisar o indicador!");
			}

			if (!Global.isCnpjCpfOk(cnpjCpf))
			{
				throw new Exception("CNPJ/CPF informado para pesquisar o indicador é inválido!");
			}
			#endregion

			#region [ Parâmetros estáticos ]
			if ((parametrosEstaticos ?? "").Trim().Length == 0)
			{
				throw new Exception("Os parâmetros estáticos não foram informados!");
			}

			if (!GeralDAO.decriptografaTexto(parametrosEstaticos, out sParamEstaticoDecripto, out msg_erro))
			{
				throw new Exception("Erro no conteúdo dos parâmetros estáticos!");
			}

			// Parâmetros estáticos no formato: loja=999|vendedor=AAAA|status=A|usuario=BBBB
			// A presença do nome dos parâmetros estáticos é obrigatória, ou seja, caso algum deles não seja utilizado a declaração deve ser, por ex: loja=999|vendedor=|status=A|usuario=BBBB
			if (!sParamEstaticoDecripto.Contains('|'))
			{
				throw new Exception("Parâmetros estáticos em formato inválido!");
			}

			vParamEstatico = sParamEstaticoDecripto.Split('|');
			for (int i = 0; i < vParamEstatico.Length; i++)
			{
				sParam = vParamEstatico[i];
				if (!sParam.Contains('='))
				{
					throw new Exception("Parâmetro estático " + (i + 1).ToString() + " está em formato inválido!");
				}

				sKey = "";
				sValue = "";
				vParam = sParam.Split('=');
				if (vParam.Length >= 1) sKey = vParam[0];
				if (vParam.Length >= 2) sValue = vParam[1];
				kvpParamEstatico = new KeyValuePair<string, string>(sKey, sValue);
				listaKvpParamEstatico.Add(kvpParamEstatico);
			}

			// Loja
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("loja"));
				lojaParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'loja' não informado!");
			}

			// Vendedor
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("vendedor"));
				vendedorParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'vendedor' não informado!");
			}

			// Status
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("status"));
				statusParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'status' não informado!");
			}

			// Usuário
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("usuario"));
				usuario = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'usuario' não informado!");
			}
			#endregion

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informada a identificação do usuário!");
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o token da sessão do usuário!");
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				throw new Exception("Falha ao tentar validar usuário!");
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				throw new Exception("Token de sessão inválido!");
			}
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: cnpjCpf=" + (cnpjCpf ?? "") + ", parâmetros estáticos decodificados: loja=" + (lojaParamEstatico ?? "") + ", vendedor=" + (vendedorParamEstatico ?? "") + ", status=" + statusParamEstatico + ", usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			listaOrcamentistaIndicador = OrcamentistaIndicadorDAO.getOrcamentistaIndicadorResumoPesquisaByCnpjCpf(cnpjCpf, loja, lojaParamEstatico, vendedor, vendedorParamEstatico, status, statusParamEstatico, out msg_erro);

			#region [ Há resultado? ]
			if (listaOrcamentistaIndicador != null)
			{
				var serializedResult = JsonConvert.SerializeObject(listaOrcamentistaIndicador);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			else
			{
				var serializedResult = JsonConvert.SerializeObject(new List<OrcamentistaIndicadorResumoPesquisa>());
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			#endregion

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ PesquisaPorCnpjCpfParcial ]
		/// <summary>
		/// Pesquisa orçamentista/indicador pelo início do CNPJ/CPF
		/// </summary>
		/// <param name="cnpjCpfParcial">Parte inicial do CNPJ/CPF a ser pesquisado (parâmetro obrigatório)</param>
		/// <param name="loja">Nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="vendedor">Vendedor a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="status">Status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'dinâmico' opcional)</param>
		/// <param name="parametrosEstaticos">
		/// Parâmetro obrigatório contendo os parâmetros estáticos criptografados para evitar consultas que burlem as restrições impostas pelo perfil de acesso.
		/// Após descriptografar este campo, os seguintes parâmetros devem tratados:
		///		loja = nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		vendedor = vendedor a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		status = status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'estático' opcional)
		///		usuario = usuário que está fazendo a requisição e que será autenticado antes de processar a solicitação (parâmetro 'estático' obrigatório)
		///	Os parâmetros estáticos possuem prioridade sobre os parâmetros 'dinâmicos', pois eles foram criados para garantir as restrições de acesso aos dados.
		///	Os parâmetros estáticos devem ser definidos e criptografados no processamento feito no lado do servidor (server side) p/ que não haja possibilidade do usuário fazer alterações.
		///	Já os parâmetros dinâmicos são os parâmetros que podem ser escolhidos pelo usuário, por exemplo, em relatórios onde se seleciona primeiro um vendedor para o sistema exibir a
		///	lista de todos os indicadores vinculados a esse vendedor.
		/// </param>
		/// <param name="sessionToken">SessionToken da sessão do usuário em andamento no sistema</param>
		/// <returns>Retorna uma lista de orçamentistas/indicadores que possuem o CNPJ/CPF especificado e que atendem também aos demais critérios especificados</returns>
		[HttpGet]
		public HttpResponseMessage PesquisaPorCnpjCpfParcial(string cnpjCpfParcial, string loja, string vendedor, string status, string parametrosEstaticos, string sessionToken)
		{
			#region [ Declarações ]
			Guid opGuid = Guid.NewGuid();
			string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.PesquisaPorCnpjCpfParcial() [" + opGuid.ToString() + "]";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			string lojaParamEstatico;
			string vendedorParamEstatico;
			string statusParamEstatico;
			string usuario;
			string sParamEstaticoDecripto;
			string[] vParamEstatico;
			string sParam;
			string[] vParam;
			string sKey;
			string sValue;
			KeyValuePair<string, string> kvpParamEstatico;
			Usuario usuarioBD;
			List<KeyValuePair<string, string>> listaKvpParamEstatico = new List<KeyValuePair<string, string>>();
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador;
			HttpResponseMessage result;
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: cnpjCpfParcial=" + (cnpjCpfParcial ?? "") + ", loja=" + (loja ?? "") + ", vendedor=" + (vendedor ?? "") + ", status=" + (status ?? "") + ", parametrosEstaticos=" + (parametrosEstaticos ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			#region [ CNPJ/CPF válido? ]
			if ((cnpjCpfParcial ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o CNPJ/CPF parcial para pesquisar o indicador!");
			}

			if (Global.digitos((cnpjCpfParcial ?? "")).Length < 3)
			{
				throw new Exception("O CNPJ/CPF parcial a ser pesquisado deve possuir 3 ou mais dígitos da parte inicial do número!");
			}
			#endregion

			#region [ Parâmetros estáticos ]
			if ((parametrosEstaticos ?? "").Trim().Length == 0)
			{
				throw new Exception("Os parâmetros estáticos não foram informados!");
			}

			if (!GeralDAO.decriptografaTexto(parametrosEstaticos, out sParamEstaticoDecripto, out msg_erro))
			{
				throw new Exception("Erro no conteúdo dos parâmetros estáticos!");
			}

			// Parâmetros estáticos no formato: loja=999|vendedor=AAAA|status=A|usuario=BBBB
			// A presença do nome dos parâmetros estáticos é obrigatória, ou seja, caso algum deles não seja utilizado a declaração deve ser, por ex: loja=999|vendedor=|status=A|usuario=BBBB
			if (!sParamEstaticoDecripto.Contains('|'))
			{
				throw new Exception("Parâmetros estáticos em formato inválido!");
			}

			vParamEstatico = sParamEstaticoDecripto.Split('|');
			for (int i = 0; i < vParamEstatico.Length; i++)
			{
				sParam = vParamEstatico[i];
				if (!sParam.Contains('='))
				{
					throw new Exception("Parâmetro estático " + (i + 1).ToString() + " está em formato inválido!");
				}

				sKey = "";
				sValue = "";
				vParam = sParam.Split('=');
				if (vParam.Length >= 1) sKey = vParam[0];
				if (vParam.Length >= 2) sValue = vParam[1];
				kvpParamEstatico = new KeyValuePair<string, string>(sKey, sValue);
				listaKvpParamEstatico.Add(kvpParamEstatico);
			}

			// Loja
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("loja"));
				lojaParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'loja' não informado!");
			}

			// Vendedor
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("vendedor"));
				vendedorParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'vendedor' não informado!");
			}

			// Status
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("status"));
				statusParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'status' não informado!");
			}

			// Usuário
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("usuario"));
				usuario = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'usuario' não informado!");
			}
			#endregion

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informada a identificação do usuário!");
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o token da sessão do usuário!");
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				throw new Exception("Falha ao tentar validar usuário!");
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				throw new Exception("Token de sessão inválido!");
			}
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: cnpjCpfParcial=" + (cnpjCpfParcial ?? "") + ", parâmetros estáticos decodificados: loja=" + (lojaParamEstatico ?? "") + ", vendedor=" + (vendedorParamEstatico ?? "") + ", status=" + statusParamEstatico + ", usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			listaOrcamentistaIndicador = OrcamentistaIndicadorDAO.getOrcamentistaIndicadorResumoPesquisaByCnpjCpfParcial(cnpjCpfParcial, loja, lojaParamEstatico, vendedor, vendedorParamEstatico, status, statusParamEstatico, out msg_erro);

			#region [ Há resultado? ]
			if (listaOrcamentistaIndicador != null)
			{
				var serializedResult = JsonConvert.SerializeObject(listaOrcamentistaIndicador);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			else
			{
				var serializedResult = JsonConvert.SerializeObject(new List<OrcamentistaIndicadorResumoPesquisa>());
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			#endregion

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ PesquisaPorApelido ]
		/// <summary>
		/// Pesquisa orçamentista/indicador pelo apelido (o apelido deve coincidir exatamente)
		/// </summary>
		/// <param name="apelido">Apelido a ser pesquisado (parâmetro obrigatório)</param>
		/// <param name="loja">Nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="vendedor">Vendedor a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="status">Status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'dinâmico' opcional)</param>
		/// <param name="parametrosEstaticos">
		/// Parâmetro obrigatório contendo os parâmetros estáticos criptografados para evitar consultas que burlem as restrições impostas pelo perfil de acesso.
		/// Após descriptografar este campo, os seguintes parâmetros devem tratados:
		///		loja = nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		vendedor = vendedor a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		status = status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'estático' opcional)
		///		usuario = usuário que está fazendo a requisição e que será autenticado antes de processar a solicitação (parâmetro 'estático' obrigatório)
		///	Os parâmetros estáticos possuem prioridade sobre os parâmetros 'dinâmicos', pois eles foram criados para garantir as restrições de acesso aos dados.
		///	Os parâmetros estáticos devem ser definidos e criptografados no processamento feito no lado do servidor (server side) p/ que não haja possibilidade do usuário fazer alterações.
		///	Já os parâmetros dinâmicos são os parâmetros que podem ser escolhidos pelo usuário, por exemplo, em relatórios onde se seleciona primeiro um vendedor para o sistema exibir a
		///	lista de todos os indicadores vinculados a esse vendedor.
		/// </param>
		/// <param name="sessionToken">SessionToken da sessão do usuário em andamento no sistema</param>
		/// <returns>Retorna os dados do orçamentista/indicador, caso encontrado, que possui o apelido especificado e que atenda também aos demais critérios especificados</returns>
		[HttpGet]
		public HttpResponseMessage PesquisaPorApelido(string apelido, string loja, string vendedor, string status, string parametrosEstaticos, string sessionToken)
		{
			#region [ Declarações ]
			Guid opGuid = Guid.NewGuid();
			string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.PesquisaPorApelido() [" + opGuid.ToString() + "]";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			string lojaParamEstatico;
			string vendedorParamEstatico;
			string statusParamEstatico;
			string usuario;
			string sParamEstaticoDecripto;
			string[] vParamEstatico;
			string sParam;
			string[] vParam;
			string sKey;
			string sValue;
			KeyValuePair<string, string> kvpParamEstatico;
			Usuario usuarioBD;
			List<KeyValuePair<string, string>> listaKvpParamEstatico = new List<KeyValuePair<string, string>>();
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			HttpResponseMessage result;
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: apelido=" + (apelido ?? "") + ", loja=" + (loja ?? "") + ", vendedor=" + (vendedor ?? "") + ", status=" + (status ?? "") + ", parametrosEstaticos=" + (parametrosEstaticos ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			#region [ Apelido válido? ]
			if ((apelido ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o identificador para pesquisar o indicador!");
			}
			#endregion

			#region [ Parâmetros estáticos ]
			if ((parametrosEstaticos ?? "").Trim().Length == 0)
			{
				throw new Exception("Os parâmetros estáticos não foram informados!");
			}

			if (!GeralDAO.decriptografaTexto(parametrosEstaticos, out sParamEstaticoDecripto, out msg_erro))
			{
				throw new Exception("Erro no conteúdo dos parâmetros estáticos!");
			}

			// Parâmetros estáticos no formato: loja=999|vendedor=AAAA|status=A|usuario=BBBB
			// A presença do nome dos parâmetros estáticos é obrigatória, ou seja, caso algum deles não seja utilizado a declaração deve ser, por ex: loja=999|vendedor=|status=A|usuario=BBBB
			if (!sParamEstaticoDecripto.Contains('|'))
			{
				throw new Exception("Parâmetros estáticos em formato inválido!");
			}

			vParamEstatico = sParamEstaticoDecripto.Split('|');
			for (int i = 0; i < vParamEstatico.Length; i++)
			{
				sParam = vParamEstatico[i];
				if (!sParam.Contains('='))
				{
					throw new Exception("Parâmetro estático " + (i + 1).ToString() + " está em formato inválido!");
				}

				sKey = "";
				sValue = "";
				vParam = sParam.Split('=');
				if (vParam.Length >= 1) sKey = vParam[0];
				if (vParam.Length >= 2) sValue = vParam[1];
				kvpParamEstatico = new KeyValuePair<string, string>(sKey, sValue);
				listaKvpParamEstatico.Add(kvpParamEstatico);
			}

			// Loja
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("loja"));
				lojaParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'loja' não informado!");
			}

			// Vendedor
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("vendedor"));
				vendedorParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'vendedor' não informado!");
			}

			// Status
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("status"));
				statusParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'status' não informado!");
			}

			// Usuário
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("usuario"));
				usuario = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'usuario' não informado!");
			}
			#endregion

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informada a identificação do usuário!");
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o token da sessão do usuário!");
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				throw new Exception("Falha ao tentar validar usuário!");
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				throw new Exception("Token de sessão inválido!");
			}
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: apelido=" + (apelido ?? "") + ", parâmetros estáticos decodificados: loja=" + (lojaParamEstatico ?? "") + ", vendedor=" + (vendedorParamEstatico ?? "") + ", status=" + statusParamEstatico + ", usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			orcamentistaIndicador = OrcamentistaIndicadorDAO.getOrcamentistaIndicadorResumoPesquisaByApelido(apelido, loja, lojaParamEstatico, vendedor, vendedorParamEstatico, status, statusParamEstatico, out msg_erro);

			#region [ Há resultado? ]
			if (orcamentistaIndicador != null)
			{
				var serializedResult = JsonConvert.SerializeObject(orcamentistaIndicador);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			else
			{
				var serializedResult = JsonConvert.SerializeObject(new OrcamentistaIndicadorResumoPesquisa());
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			#endregion

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ PesquisaPorApelidoParcial ]
		/// <summary>
		/// Pesquisa orçamentista/indicador pelo apelido que se inicie com o texto especificado
		/// </summary>
		/// <param name="apelidoParcial">Apelido parcial a ser pesquisado (parâmetro obrigatório)</param>
		/// <param name="loja">Nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="vendedor">Vendedor a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="status">Status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'dinâmico' opcional)</param>
		/// <param name="parametrosEstaticos">
		/// Parâmetro obrigatório contendo os parâmetros estáticos criptografados para evitar consultas que burlem as restrições impostas pelo perfil de acesso.
		/// Após descriptografar este campo, os seguintes parâmetros devem tratados:
		///		loja = nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		vendedor = vendedor a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		status = status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'estático' opcional)
		///		usuario = usuário que está fazendo a requisição e que será autenticado antes de processar a solicitação (parâmetro 'estático' obrigatório)
		///	Os parâmetros estáticos possuem prioridade sobre os parâmetros 'dinâmicos', pois eles foram criados para garantir as restrições de acesso aos dados.
		///	Os parâmetros estáticos devem ser definidos e criptografados no processamento feito no lado do servidor (server side) p/ que não haja possibilidade do usuário fazer alterações.
		///	Já os parâmetros dinâmicos são os parâmetros que podem ser escolhidos pelo usuário, por exemplo, em relatórios onde se seleciona primeiro um vendedor para o sistema exibir a
		///	lista de todos os indicadores vinculados a esse vendedor.
		/// </param>
		/// <param name="sessionToken">SessionToken da sessão do usuário em andamento no sistema</param>
		/// <returns>Retorna uma lista de orçamentistas/indicadores que possuem o apelido que se inicie com o texto especificado e que atenda também aos demais critérios especificados</returns>
		[HttpGet]
		public HttpResponseMessage PesquisaPorApelidoParcial(string apelidoParcial, string loja, string vendedor, string status, string parametrosEstaticos, string sessionToken)
		{
			#region [ Declarações ]
			Guid opGuid = Guid.NewGuid();
			string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.PesquisaPorApelidoParcial() [" + opGuid.ToString() + "]";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			string lojaParamEstatico;
			string vendedorParamEstatico;
			string statusParamEstatico;
			string usuario;
			string sParamEstaticoDecripto;
			string[] vParamEstatico;
			string sParam;
			string[] vParam;
			string sKey;
			string sValue;
			KeyValuePair<string, string> kvpParamEstatico;
			Usuario usuarioBD;
			List<KeyValuePair<string, string>> listaKvpParamEstatico = new List<KeyValuePair<string, string>>();
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador;
			HttpResponseMessage result;
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: apelidoParcial=" + (apelidoParcial ?? "") + ", loja=" + (loja ?? "") + ", vendedor=" + (vendedor ?? "") + ", status=" + (status ?? "") + ", parametrosEstaticos=" + (parametrosEstaticos ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			#region [ Apelido válido? ]
			if ((apelidoParcial ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o identificador parcial para pesquisar o indicador!");
			}
			#endregion

			#region [ Parâmetros estáticos ]
			if ((parametrosEstaticos ?? "").Trim().Length == 0)
			{
				throw new Exception("Os parâmetros estáticos não foram informados!");
			}

			if (!GeralDAO.decriptografaTexto(parametrosEstaticos, out sParamEstaticoDecripto, out msg_erro))
			{
				throw new Exception("Erro no conteúdo dos parâmetros estáticos!");
			}

			// Parâmetros estáticos no formato: loja=999|vendedor=AAAA|status=A|usuario=BBBB
			// A presença do nome dos parâmetros estáticos é obrigatória, ou seja, caso algum deles não seja utilizado a declaração deve ser, por ex: loja=999|vendedor=|status=A|usuario=BBBB
			if (!sParamEstaticoDecripto.Contains('|'))
			{
				throw new Exception("Parâmetros estáticos em formato inválido!");
			}

			vParamEstatico = sParamEstaticoDecripto.Split('|');
			for (int i = 0; i < vParamEstatico.Length; i++)
			{
				sParam = vParamEstatico[i];
				if (!sParam.Contains('='))
				{
					throw new Exception("Parâmetro estático " + (i + 1).ToString() + " está em formato inválido!");
				}

				sKey = "";
				sValue = "";
				vParam = sParam.Split('=');
				if (vParam.Length >= 1) sKey = vParam[0];
				if (vParam.Length >= 2) sValue = vParam[1];
				kvpParamEstatico = new KeyValuePair<string, string>(sKey, sValue);
				listaKvpParamEstatico.Add(kvpParamEstatico);
			}

			// Loja
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("loja"));
				lojaParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'loja' não informado!");
			}

			// Vendedor
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("vendedor"));
				vendedorParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'vendedor' não informado!");
			}

			// Status
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("status"));
				statusParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'status' não informado!");
			}

			// Usuário
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("usuario"));
				usuario = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'usuario' não informado!");
			}
			#endregion

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informada a identificação do usuário!");
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o token da sessão do usuário!");
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				throw new Exception("Falha ao tentar validar usuário!");
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				throw new Exception("Token de sessão inválido!");
			}
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: apelidoParcial=" + (apelidoParcial ?? "") + ", parâmetros estáticos decodificados: loja=" + (lojaParamEstatico ?? "") + ", vendedor=" + (vendedorParamEstatico ?? "") + ", status=" + statusParamEstatico + ", usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			listaOrcamentistaIndicador = OrcamentistaIndicadorDAO.getOrcamentistaIndicadorResumoPesquisaByApelidoParcial(apelidoParcial, loja, lojaParamEstatico, vendedor, vendedorParamEstatico, status, statusParamEstatico, out msg_erro);

			#region [ Há resultado? ]
			if (listaOrcamentistaIndicador != null)
			{
				var serializedResult = JsonConvert.SerializeObject(listaOrcamentistaIndicador);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			else
			{
				var serializedResult = JsonConvert.SerializeObject(new List<OrcamentistaIndicadorResumoPesquisa>());
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			#endregion

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ PesquisaPorNomeParcial ]
		/// <summary>
		/// Pesquisa por orçamentista/indicador que tenha o texto especificado no parâmetro em qualquer parte do nome
		/// </summary>
		/// <param name="nomeParcial">Nome parcial a ser pesquisado (parâmetro obrigatório)</param>
		/// <param name="loja">Nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="vendedor">Vendedor a que o orçamentista/indicador está vinculado (parâmetro 'dinâmico' opcional)</param>
		/// <param name="status">Status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'dinâmico' opcional)</param>
		/// <param name="parametrosEstaticos">
		/// Parâmetro obrigatório contendo os parâmetros estáticos criptografados para evitar consultas que burlem as restrições impostas pelo perfil de acesso.
		/// Após descriptografar este campo, os seguintes parâmetros devem tratados:
		///		loja = nº da loja a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		vendedor = vendedor a que o orçamentista/indicador está vinculado (parâmetro 'estático' opcional)
		///		status = status do orçamentista/indicador: A=Ativo, I=Inativo (parâmetro 'estático' opcional)
		///		usuario = usuário que está fazendo a requisição e que será autenticado antes de processar a solicitação (parâmetro 'estático' obrigatório)
		///	Os parâmetros estáticos possuem prioridade sobre os parâmetros 'dinâmicos', pois eles foram criados para garantir as restrições de acesso aos dados.
		///	Os parâmetros estáticos devem ser definidos e criptografados no processamento feito no lado do servidor (server side) p/ que não haja possibilidade do usuário fazer alterações.
		///	Já os parâmetros dinâmicos são os parâmetros que podem ser escolhidos pelo usuário, por exemplo, em relatórios onde se seleciona primeiro um vendedor para o sistema exibir a
		///	lista de todos os indicadores vinculados a esse vendedor.
		/// </param>
		/// <param name="sessionToken">SessionToken da sessão do usuário em andamento no sistema</param>
		/// <returns>Retorna uma lista de orçamentistas/indicadores que possuem o apelido que se inicie com o texto especificado e que atenda também aos demais critérios especificados</returns>
		[HttpGet]
		public HttpResponseMessage PesquisaPorNomeParcial(string nomeParcial, string loja, string vendedor, string status, string parametrosEstaticos, string sessionToken)
		{
			#region [ Declarações ]
			Guid opGuid = Guid.NewGuid();
			string NOME_DESTA_ROTINA = "OrcamentistaIndicadorController.PesquisaPorNomeParcial() [" + opGuid.ToString() + "]";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			string lojaParamEstatico;
			string vendedorParamEstatico;
			string statusParamEstatico;
			string usuario;
			string sParamEstaticoDecripto;
			string[] vParamEstatico;
			string sParam;
			string[] vParam;
			string sKey;
			string sValue;
			KeyValuePair<string, string> kvpParamEstatico;
			Usuario usuarioBD;
			List<KeyValuePair<string, string>> listaKvpParamEstatico = new List<KeyValuePair<string, string>>();
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador;
			HttpResponseMessage result;
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: nomeParcial=" + (nomeParcial ?? "") + ", loja=" + (loja ?? "") + ", vendedor=" + (vendedor ?? "") + ", status=" + (status ?? "") + ", parametrosEstaticos=" + (parametrosEstaticos ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			#region [ Nome parcial válido? ]
			if ((nomeParcial ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o nome parcial para pesquisar o indicador!");
			}

			if ((nomeParcial ?? "").Trim().Length < 4)
			{
				throw new Exception("O nome parcial a ser pesquisado deve possuir 4 ou mais caracteres!");
			}
			#endregion

			#region [ Parâmetros estáticos ]
			if ((parametrosEstaticos ?? "").Trim().Length == 0)
			{
				throw new Exception("Os parâmetros estáticos não foram informados!");
			}

			if (!GeralDAO.decriptografaTexto(parametrosEstaticos, out sParamEstaticoDecripto, out msg_erro))
			{
				throw new Exception("Erro no conteúdo dos parâmetros estáticos!");
			}

			// Parâmetros estáticos no formato: loja=999|vendedor=AAAA|status=A|usuario=BBBB
			// A presença do nome dos parâmetros estáticos é obrigatória, ou seja, caso algum deles não seja utilizado a declaração deve ser, por ex: loja=999|vendedor=|status=A|usuario=BBBB
			if (!sParamEstaticoDecripto.Contains('|'))
			{
				throw new Exception("Parâmetros estáticos em formato inválido!");
			}

			vParamEstatico = sParamEstaticoDecripto.Split('|');
			for (int i = 0; i < vParamEstatico.Length; i++)
			{
				sParam = vParamEstatico[i];
				if (!sParam.Contains('='))
				{
					throw new Exception("Parâmetro estático " + (i + 1).ToString() + " está em formato inválido!");
				}

				sKey = "";
				sValue = "";
				vParam = sParam.Split('=');
				if (vParam.Length >= 1) sKey = vParam[0];
				if (vParam.Length >= 2) sValue = vParam[1];
				kvpParamEstatico = new KeyValuePair<string, string>(sKey, sValue);
				listaKvpParamEstatico.Add(kvpParamEstatico);
			}

			// Loja
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("loja"));
				lojaParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'loja' não informado!");
			}

			// Vendedor
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("vendedor"));
				vendedorParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'vendedor' não informado!");
			}

			// Status
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("status"));
				statusParamEstatico = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'status' não informado!");
			}

			// Usuário
			try
			{
				kvpParamEstatico = listaKvpParamEstatico.Single(x => x.Key.Trim().ToLower().Equals("usuario"));
				usuario = (kvpParamEstatico.Value ?? "").Trim();
			}
			catch (Exception)
			{
				throw new Exception("Parâmetro estático 'usuario' não informado!");
			}
			#endregion

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informada a identificação do usuário!");
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				throw new Exception("Não foi informado o token da sessão do usuário!");
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				throw new Exception("Falha ao tentar validar usuário!");
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				throw new Exception("Token de sessão inválido!");
			}
			#endregion

			#region [ Log atividade ]
			msg = NOME_DESTA_ROTINA + " - Parâmetros: nomeParcial=" + (nomeParcial ?? "") + ", parâmetros estáticos decodificados: loja=" + (lojaParamEstatico ?? "") + ", vendedor=" + (vendedorParamEstatico ?? "") + ", status=" + statusParamEstatico + ", usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);
			#endregion

			listaOrcamentistaIndicador = OrcamentistaIndicadorDAO.getOrcamentistaIndicadorResumoPesquisaByNomeParcial(nomeParcial, loja, lojaParamEstatico, vendedor, vendedorParamEstatico, status, statusParamEstatico, out msg_erro);

			#region [ Há resultado? ]
			if (listaOrcamentistaIndicador != null)
			{
				var serializedResult = JsonConvert.SerializeObject(listaOrcamentistaIndicador);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			else
			{
				var serializedResult = JsonConvert.SerializeObject(new List<OrcamentistaIndicadorResumoPesquisa>());
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			}
			#endregion

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion
	}
}
