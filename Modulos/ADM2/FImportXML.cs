using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ADM2
{
    public partial class FImportXML : ADM2.FModelo
    {

        #region [ Atributos ]
        private bool _emProcessamento = false;
        private bool _InicializacaoOk;
        public bool inicializacaoOk
        {
            get { return _InicializacaoOk; }
        }

        private bool _OcorreuExceptionNaInicializacao;
        public bool ocorreuExceptionNaInicializacao
        {
            get { return _OcorreuExceptionNaInicializacao; }
        }

        #endregion

        public FImportXML()
        {
            InitializeComponent();
        }
    }
}
