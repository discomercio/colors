using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
    public class NFeImagem
    {
        public int id { get; set; }
        public int id_nfe_emitente { get; set; }
        public int NFe_serie_NF { get; set; }
        public int NFe_numero_NF { get; set; }
        public DateTime data { get; set; }
        public DateTime data_hora { get; set; }
        public string usuario { get; set; }
        public string pedido { get; set; }
        public string operacional__email { get; set; }
        public string ide__natOp { get; set; }
        public string ide__indPag { get; set; }
        public string ide__serie { get; set; }
        public string ide__nNF { get; set; }
        public string ide__dEmi { get; set; }
        public string ide__dSaiEnt { get; set; }
        public string ide__tpNF { get; set; }
        public string ide__cMunFG { get; set; }
        public string ide__tpAmb { get; set; }
        public string ide__finNFe { get; set; }
        public string ide__IEST { get; set; }
        public string dest__CNPJ { get; set; }
        public string dest__CPF { get; set; }
        public string dest__xNome { get; set; }
        public string dest__xLgr { get; set; }
        public string dest__nro { get; set; }
        public string dest__xCpl { get; set; }
        public string dest__xBairro { get; set; }
        public string dest__cMun { get; set; }
        public string dest__xMun { get; set; }
        public string dest__UF { get; set; }
        public string dest__CEP { get; set; }
        public string dest__cPais { get; set; }
        public string dest__xPais { get; set; }
        public string dest__fone { get; set; }
        public string dest__IE { get; set; }
        public string dest__ISUF { get; set; }
        public string entrega__CNPJ { get; set; }
        public string entrega__xLgr { get; set; }
        public string entrega__nro { get; set; }
        public string entrega__xCpl { get; set; }
        public string entrega__xBairro { get; set; }
        public string entrega__cMun { get; set; }
        public string entrega__xMun { get; set; }
        public string entrega__UF { get; set; }
        public string total__vBC { get; set; }
        public string total__vICMS { get; set; }
        public string total__vBCST { get; set; }
        public string total__vST { get; set; }
        public string total__vProd { get; set; }
        public string total__vFrete { get; set; }
        public string total__vSeg { get; set; }
        public string total__vDesc { get; set; }
        public string total__vII { get; set; }
        public string total__vIPI { get; set; }
        public string total__vPIS { get; set; }
        public string total__vCOFINS { get; set; }
        public string total__vOutro { get; set; }
        public string total__vNF { get; set; }
        public string transp__modFrete { get; set; }
        public string transporta__CNPJ { get; set; }
        public string transporta__CPF { get; set; }
        public string transporta__xNome { get; set; }
        public string transporta__IE { get; set; }
        public string transporta__xEnder { get; set; }
        public string transporta__xMun { get; set; }
        public string transporta__UF { get; set; }
        public string vol__qVol { get; set; }
        public string vol__esp { get; set; }
        public string vol__marca { get; set; }
        public string vol__nVol { get; set; }
        public string vol__pesoL { get; set; }
        public string vol__pesoB { get; set; }
        public string vol_nLacre { get; set; }
        public string infAdic__infAdFisco { get; set; }
        public string infAdic__infCpl { get; set; }
        public string codigo_retorno_NFe_T1 { get; set; }
        public string msg_retorno_NFe_T1 { get; set; }
        public int st_anulado { get; set; }
        public DateTime dt_anulado { get; set; }
        public DateTime dt_hr_anulado { get; set; }
        public string usuario_anulado { get; set; }
        public string versao_layout_NFe { get; set; }
        public string entrega__CPF { get; set; }
        public string total__vTotTrib { get; set; }
        public string ide__dEmiUTC { get; set; }
        public string ide__idDest { get; set; }
        public string ide__indFinal { get; set; }
        public string ide__indPres { get; set; }
        public string dest__idEstrangeiro { get; set; }
        public string dest__indIEDest { get; set; }
        public string total__vICMSDeson { get; set; }
        public string dest__email { get; set; }
        public string total__vFCPUFDest { get; set; }
        public string total__vICMSUFDest { get; set; }
        public string total__vICMSUFRemet { get; set; }
        public string total__vFCP { get; set; }
        public string total__vFCPST { get; set; }
        public string total__vFCPSTRet { get; set; }
        public string total__vIPIDevol { get; set; }
    }
}
