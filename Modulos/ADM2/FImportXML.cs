using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Media;
using System.Text;
using System.Windows.Forms;

namespace ADM2
{
    public partial class FImportXML : ADM2.FModelo
    {

        #region [ Constantes ]
        const String GRID_COL_ID_ESTOQUE = "colIdEstoque";
        const String GRID_COL_DATA_ENTRADA = "colDataEntrada";
        const String GRID_COL_CD = "colCD";
        const String GRID_COL_DOCUMENTO = "colDocumento";
        const String GRID_COL_FABRICANTE = "colFabricante";
        const String GRID_COL_DESCRICAO = "colDescricao";
        #endregion


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

        #region [ Eventos ]

        private void FImportXML_Load(object sender, EventArgs e)
        {
            if (true) { };
        }

        private void FImportXML_Shown(object sender, EventArgs e)
        {
            if (true) { };
        }

        private void FImportXML_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_emProcessamento)
            {
                SystemSounds.Exclamation.Play();
                e.Cancel = true;
                return;
            }

            FMain.fMain.Location = this.Location;
            FMain.fMain.Visible = true;
            this.Visible = false;

        }

        #endregion
    }
}
