using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.SqlServer.Server;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;
using SLT.ImportacaoNF.Conexao;
using System.Collections;

namespace SLT.ImportacaoNF
{
    [FormAttribute("SLT.ImportacaoNF.frmImportacao", "Forms/frmImportacao.b1f")]
    class frmImportacao : UserFormBase
    {
        #region Attributes

        private SAPbouiCOM.DataTable dtDados;
        private SAPbouiCOM.StaticText lblProcesso;
        private SAPbouiCOM.EditText txtProcesso;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText txtCodigoPedido;

        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText txtTaxaDI;
        private SAPbouiCOM.Button btnSalvar;
        private SAPbouiCOM.Button btnGerarNF;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText txtFreteInternacional;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText txtContainer;
        private SAPbouiCOM.Button btnPesquisar;

        private StaticText StaticText7;
        private StaticText StaticText8;
        private StaticText StaticText9;
        private EditText txtPeso;
        private EditText txtTotalFOB;
        private EditText txtTotalRS;

        private StaticText StaticText4;
        private EditText txtOutraDespesa;
        private EditText txtCodePN;
        private Button btnCalcular;
        private Button btnCancelar;

        private StaticText StaticText10;
        private ComboBox ComboBox0;
        private StaticText StaticText11;
        private EditText txtDataDocumento;

        //private EditText txtCodImp;
        private StaticText StaticText12;
        private StaticText StaticText14;
        private EditText txtCodigoImportacao;
        private Button btnCarregar;

        //private SAPbouiCOM.Grid gridDados;

        private StaticText StaticText13;
        private EditText txtTotal_II;



        private Matrix matrixData;

        #endregion attributes

        public frmImportacao()
        {
            CreateMatrix();
        }

        public void CreateMatrix()
        {
            var dataTableId = "dtDados";
            this.dtDados = this.UIAPIRawForm.DataSources.DataTables.Item(dataTableId);

            string pn = txtCodePN.Value;
            string pedido = txtCodigoPedido.Value;
            string processo = txtProcesso.Value;
            string importacao = txtCodigoImportacao.Value;
            string queryPesquisar = QueryPesquisar(pn, pedido, processo);
            this.dtDados.ExecuteQuery(queryPesquisar);

            //matrixData.Item.Width = 500;
            //matrixData.Item.Height = 300;
            var columns = matrixData.Columns;

            // setup columns
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_CHECK_BOX, "#", "#", "Selected");
            AddMatrixLinkedButtonColumn(columns, BoLinkedObject.lf_PurchaseOrder, dataTableId, "DocEntry", "Pedido de Compra", "DocEntry", false);
            AddMatrixLinkedButtonColumn(columns, BoLinkedObject.lf_Items, dataTableId, "ItemCode", "Cód. Item", "ItemCode", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "Dscription", "Descrição do Item", "Dscription", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "LineNum", "Nº Linha do Pedido", "LineNum", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "Price", "Preço Unitário", "Price", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "Quantity", "Qtd.", "Quantity", true);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "LineTotal", "Total da Linha", "LineTotal", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "OpenQty", "Qtd. Aberto", "OpenQty", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "UnitMsr", "UM", "UnitMsr", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "Peso", "Peso", "Peso", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "NumPerMsr", "NumPerMsr", "NumPerMsr", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "NcmCode", "Cód. NCM", "NcmCode", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "VisOrder", "VisOrder", "VisOrder", false);
            AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_EDIT, "TaxRate", "TaxRate", "TaxRate", false);

            this.matrixData.SelectionMode = BoMatrixSelect.ms_Auto;
            matrixData.ClickAfter += matrixData_ClickAfter;


            // load the data into the rows
            matrixData.LoadFromDataSource();
            matrixData.AutoResizeColumns();
        }

        void matrixData_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID.Equals("#"))
            {
                this.dtDados.SetValue("Selected", pVal.Row, pVal.ActionSuccess ? "Y" : "N");
                Calcular();
            }
        }

        void Calcular()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Atualizando os valores, aguarde alguns instantes!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

            var total = 0;
            var columns = this.matrixData.Columns;
            var value = string.Empty;

            bool selected = false;

            for (int i = 0; i < this.dtDados.Rows.Count; i++)
            {
                selected = this.dtDados.GetValue("Selected", i).ToString().Equals("Y");

                if (selected)
                {
                    txtPeso.Value = 1.ToString("N3");
                    txtTotal_II.Value = 1.ToString("N2");
                    txtTotalFOB.Value = 1.ToString("N2");
                    total += (1 * 1) + 1 + 1;
                }
            }
            
            txtTotalRS.Value = total.ToString("N2");
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Valores atualizados!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        public Column AddMatrixColumn(Columns columns, string dataTableId, BoFormItemTypes formItemTypes, string columnName, string columnTitleCaption = null, string dataSourceColumnName = null, bool editable = true)
        {
            Column column = columns.Add(columnName, formItemTypes);

            if (String.IsNullOrWhiteSpace(columnTitleCaption))
                columnName = columnTitleCaption;

            if (String.IsNullOrWhiteSpace(dataSourceColumnName))
                dataSourceColumnName = columnName;

            column.TitleObject.Caption = columnTitleCaption;
            column.DataBind.Bind(dataTableId, dataSourceColumnName);

            if (!editable)
                column.Editable = editable;

            return column;
        }

        public Column AddMatrixLinkedButtonColumn(Columns columns, BoLinkedObject linkedObject, string dataTableId, string columnName, string columnTitleCaption = null, string dataSourceColumnName = null, bool editable = true)
        {
            var column = AddMatrixColumn(columns, dataTableId, BoFormItemTypes.it_LINKED_BUTTON, columnName, columnTitleCaption, dataSourceColumnName, editable);

            SAPbouiCOM.LinkedButton linkedButton = (SAPbouiCOM.LinkedButton)column.ExtendedObject;
            linkedButton.LinkedObject = linkedObject;

            return column;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.lblProcesso = ((SAPbouiCOM.StaticText)(this.GetItem("lblProc").Specific));
            this.txtProcesso = ((SAPbouiCOM.EditText)(this.GetItem("txtProces").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lblPed").Specific));
            this.txtCodigoPedido = ((SAPbouiCOM.EditText)(this.GetItem("txtPedido").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lblCdPN").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.txtTaxaDI = ((SAPbouiCOM.EditText)(this.GetItem("txtTxID").Specific));
            this.btnSalvar = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            this.btnSalvar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnSalvar_ClickBefore);
            this.btnGerarNF = ((SAPbouiCOM.Button)(this.GetItem("Item_13").Specific));
            this.btnGerarNF.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnGerarNF_ClickBefore);
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.txtFreteInternacional = ((SAPbouiCOM.EditText)(this.GetItem("txtFrtInt").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.txtContainer = ((SAPbouiCOM.EditText)(this.GetItem("txtConta").Specific));
            this.btnPesquisar = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.btnPesquisar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnPesquisar_ClickBefore);

            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.txtOutraDespesa = ((SAPbouiCOM.EditText)(this.GetItem("txtODesp").Specific));
            this.txtCodePN = ((SAPbouiCOM.EditText)(this.GetItem("txtCodPN").Specific));
            this.btnCalcular = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.btnCalcular.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCalcular_ClickBefore);
            this.btnCancelar = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.btnCancelar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCancelar_ClickBefore);
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.txtPeso = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.txtTotalFOB = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.txtTotalRS = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.txtDataDocumento = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            // this.txtCodImp = ((SAPbouiCOM.EditText)(this.GetItem("txtCodImp").Specific));
            this.txtDataDocumento.Value = DateTime.Now.ToShortDateString();
            //     Numero da Importação
            // this.txtCodImp.Value = this.RetornaNrImportacao();
            // this.ComboBox0.Select("Aberto", typeof(SAPbouiCOM.BoSearchKey).psk_ByDescription);
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.txtCodigoImportacao = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            this.btnCarregar = ((SAPbouiCOM.Button)(this.GetItem("btnCarreg").Specific));
            this.btnCarregar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCarregar_ClickBefore);
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.txtTotal_II = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.matrixData = ((SAPbouiCOM.Matrix)(this.GetItem("Item_22").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.ResizeAfter += frmImportacao_ResizeAfter;
        }

        private void OnCustomInitialize()
        {
            //this.dtDados = this.UIAPIRawForm.DataSources.DataTables.Item("dtDados");
            //this.UIAPIRawForm.EnableMenu("1281", false);
            //this.UIAPIRawForm.EnableMenu("1282", false);
            //this.UIAPIRawForm.EnableMenu("1288", false);
            //this.UIAPIRawForm.EnableMenu("1289", false);
            //this.UIAPIRawForm.EnableMenu("1290", false);
            //this.UIAPIRawForm.EnableMenu("1291", false);
            //this.UIAPIRawForm.EnableMenu("1304", false);
        }


        #region Events

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //matrixData.Clear();

            //SAPbouiCOM.DBDataSource oDBDataSource = frmImportacao.DataSources.DBDataSources.Item(1);
            //SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //oDBDataSource.Clear();

            //for (int row = 0; row < oRSet.RecordCount; row++)
            //{
            //    int offset = oDBDataSource.Size;
            //    //add values here 
            //    oDBDataSource.SetValue("U_metodika", offset, oRSet.Fields.Item(1).Value.ToString());
            //    oDBDataSource.InsertRecord(row);
            //    oRSet.MoveNext();
            //}

            //matrixData.LoadFromDataSource();
        }


        void frmImportacao_ResizeAfter(SBOItemEventArg pVal)
        {
            matrixData.LoadFromDataSource();
            matrixData.AutoResizeColumns();
        }

        private void btnPesquisar_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisa em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                string vCodPn = txtCodePN.Value;
                string vCodPedido = txtCodigoPedido.Value;
                string vCodProcesso = txtProcesso.Value;
                string queryPesquisar = QueryPesquisar(vCodPn, vCodPedido, vCodProcesso);

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Executando a consulta...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                this.dtDados.Rows.Clear();

                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRSet.DoQuery(queryPesquisar);

                if (oRSet.RecordCount == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Nenhum registro encontrado!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Consulta concluida, processando os dados para a exibição...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);


                //this.gridDados.DataTable = this.dtDados;
                //this.gridDados.Columns.Item("ColCheck").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                //this.gridDados.Columns.Item("DocEntry").Editable = false;
                //this.gridDados.Columns.Item("DocNum").Editable = false;
                //this.gridDados.Columns.Item("VisOrder").Editable = false;
                //this.gridDados.Columns.Item("NcmCode").Editable = false;
                //this.gridDados.Columns.Item("ItemCode").Editable = false;
                //this.gridDados.Columns.Item("Dscription").Editable = false;
                //this.gridDados.Columns.Item("Price").Editable = false;
                //this.gridDados.Columns.Item("LineTotal").Editable = false;
                //this.gridDados.Columns.Item("OpenQty").Editable = false;
                //this.gridDados.Columns.Item("unitMsr").Editable = false;
                //this.gridDados.Columns.Item("Peso").Editable = false;
                //this.gridDados.Columns.Item("Frete").Editable = false;
                //this.gridDados.Columns.Item("OtrDesp").Editable = false;
                //this.gridDados.Columns.Item("NumPerMsr").Editable = false;
                //this.gridDados.Columns.Item("LineNum").Visible = false;
                //this.gridDados.Columns.Item("TaxRate").Editable = false;

                //SAPbouiCOM.EditTextColumn colCodInterno = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("DocEntry");
                //colCodInterno.LinkedObjectType = "22";

                //SAPbouiCOM.EditTextColumn colCodItem = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("ItemCode");
                //colCodItem.LinkedObjectType = "4";

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }

        }

        private void btnCarregar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisa de importacao em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            try
            {
                //SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                String vCodPN = txtCodePN.Value;
                String vCodPedido = txtCodigoPedido.Value;
                String vCodProcesso = txtProcesso.Value;

                if (!String.IsNullOrWhiteSpace(vCodPN))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Código do PN através do botão pesquisa");
                    return;
                }

                if (!String.IsNullOrWhiteSpace(vCodPedido))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Pedido através do botão pesquisa");
                    return;
                }

                if (!String.IsNullOrWhiteSpace(vCodProcesso))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Nº do proceso através do botão pesquisa");
                    return;
                }

                String vCodImp = txtCodigoImportacao.Value;

                if (vCodImp == "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Informar número do documento");
                    return;
                }


                string Query = QueryCarregar(vCodImp);

                oRSet.DoQuery(Query);

                this.dtDados.Rows.Clear();

                if (oRSet.RecordCount == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Nenhum registro encontrado!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                while (!oRSet.EoF)
                {
                    this.dtDados.Rows.Add();
                    int i = this.dtDados.Rows.Count - 1;

                    this.dtDados.SetValue("DocEntry", i, oRSet.Fields.Item("DocEntry").Value.ToString());
                    this.dtDados.SetValue("DocNum", i, oRSet.Fields.Item("DocNum").Value.ToString());
                    this.dtDados.SetValue("VisOrder", i, Int32.Parse(oRSet.Fields.Item("VisOrder").Value.ToString()) + 1);
                    this.dtDados.SetValue("NcmCode", i, oRSet.Fields.Item("NcmCode").Value.ToString());
                    this.dtDados.SetValue("ItemCode", i, oRSet.Fields.Item("ItemCode").Value.ToString());
                    this.dtDados.SetValue("Dscription", i, RetornaDescItem(oRSet.Fields.Item("ItemCode").Value.ToString()));
                    this.dtDados.SetValue("Price", i, Convert.ToDouble(oRSet.Fields.Item("Price").Value).ToString("N4"));
                    this.dtDados.SetValue("LineTotal", i, Convert.ToDouble(oRSet.Fields.Item("LineTotal").Value).ToString("N2"));
                    this.dtDados.SetValue("Quantity", i, oRSet.Fields.Item("Quantity").Value.ToString());
                    this.dtDados.SetValue("OpenQty", i, oRSet.Fields.Item("OpenQty").Value.ToString());
                    this.dtDados.SetValue("unitMsr", i, oRSet.Fields.Item("unitMsr").Value.ToString());
                    this.dtDados.SetValue("Peso", i, Convert.ToDouble(oRSet.Fields.Item("Peso").Value).ToString("N3"));
                    this.dtDados.SetValue("NumPerMsr", i, oRSet.Fields.Item("NumPerMsr").Value.ToString());
                    this.dtDados.SetValue("LineNum", i, oRSet.Fields.Item("LineNum").Value.ToString());
                    this.dtDados.SetValue("TaxRate", i, oRSet.Fields.Item("TaxRate").Value.ToString());

                    //this.gridDados.DataTable.SetValue(0, i, "Y");
                    //this.gridDados.DataTable.SetValue(13, i, Convert.ToDouble(oRSet.Fields.Item("FreteLinha").Value).ToString("N4"));
                    //this.gridDados.DataTable.SetValue(14, i, Convert.ToDouble(oRSet.Fields.Item("OtrDespLinha").Value).ToString("N4"));

                    txtTaxaDI.Value = ParseGlobalization(oRSet.Fields.Item("TXDI").Value);
                    txtFreteInternacional.Value = ParseGlobalization(oRSet.Fields.Item("FRTINT").Value);
                    txtOutraDespesa.Value = ParseGlobalization(oRSet.Fields.Item("OTRDESP").Value);
                    txtPeso.Value = Convert.ToDouble(oRSet.Fields.Item("TTPESO").Value).ToString("N3");
                    txtTotalFOB.Value = Convert.ToDouble(oRSet.Fields.Item("TTME").Value).ToString("N2");
                    txtTotal_II.Value = Convert.ToDouble(oRSet.Fields.Item("TaxValLine").Value).ToString("N2");
                    txtTotalRS.Value = Convert.ToDouble(oRSet.Fields.Item("TOTAL").Value).ToString("N2");

                    oRSet.MoveNext();
                }

                //this.gridDados.Columns.Item("ColCheck").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                //this.gridDados.Columns.Item("DocEntry").Editable = false;
                //this.gridDados.Columns.Item("DocNum").Editable = false;
                //this.gridDados.Columns.Item("LineNum").Editable = false;
                //this.gridDados.Columns.Item("NcmCode").Editable = false;
                //this.gridDados.Columns.Item("ItemCode").Editable = false;
                //this.gridDados.Columns.Item("Dscription").Editable = false;
                //this.gridDados.Columns.Item("Price").Editable = false;
                //this.gridDados.Columns.Item("LineTotal").Editable = false;
                //this.gridDados.Columns.Item("OpenQty").Editable = false;
                //this.gridDados.Columns.Item("unitMsr").Editable = false;
                //this.gridDados.Columns.Item("Peso").Editable = false;
                //this.gridDados.Columns.Item("Frete").Editable = false;
                //this.gridDados.Columns.Item("OtrDesp").Editable = false;
                //this.gridDados.Columns.Item("NumPerMsr").Editable = false;
                //this.gridDados.Columns.Item("LineNum").Visible = false;
                //this.gridDados.Columns.Item("VisOrder").Editable = false;
                //this.gridDados.Columns.Item("TaxRate").Editable = false;

                //SAPbouiCOM.EditTextColumn colCodInterno = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("DocEntry");
                //colCodInterno.LinkedObjectType = "22";

                //SAPbouiCOM.EditTextColumn colCodItem = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("ItemCode");
                //colCodItem.LinkedObjectType = "4";

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }
        }

        private void btnSalvar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Salvar em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            double vTxtTxID;
            if (!String.IsNullOrWhiteSpace(txtTaxaDI.Value))
            {
                vTxtTxID = Convert.ToDouble(txtTaxaDI.Value.Replace(".", ","));

            }
            else
            {
                vTxtTxID = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                return;
            }

            double vFrtInt;
            if (!String.IsNullOrWhiteSpace(txtFreteInternacional.Value))
            {
                vFrtInt = Convert.ToDouble(txtFreteInternacional.Value.Replace(".", ","));

            }
            else
            {
                vFrtInt = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional.");
                return;
            }

            double vOtrDesp;
            if (!String.IsNullOrWhiteSpace(txtOutraDespesa.Value))
            {
                vOtrDesp = Convert.ToDouble(txtOutraDespesa.Value.ToString().Replace(".", ","));
            }
            else
            {
                vOtrDesp = 0;
            }

            double vContainer;
            if (!String.IsNullOrWhiteSpace(txtContainer.Value))
            {
                vContainer = Convert.ToDouble(txtContainer.Value.ToString().Replace(".", ","));
            }
            else
            {
                vContainer = 0;
            }

            Double vTotalPeso = Convert.ToDouble(txtPeso.Value.PadLeft(3, '0'));
            Double vTotalME = Convert.ToDouble(txtTotalFOB.Value.PadLeft(2, '0'));
            Double vTotal = Convert.ToDouble(txtTotalRS.Value.PadLeft(2, '0'));
            Double vTotalII = Convert.ToDouble(txtTotal_II.Value.PadLeft(2, '0'));

            //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("2");
            //string Hoje = "2021/03/24";
            string Hoje = DateTime.Now.ToShortDateString();

            string vCodImp = txtCodigoImportacao.Value;
            string vRetornoNrImp;
            if (!String.IsNullOrWhiteSpace(vCodImp))
            {
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }
            else
            {
                vCodImp = txtCodigoImportacao.Value.ToString();
                //int CodExiste = Int32.Parse(txtCodImp.ToString()) - 1;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }

            if (!String.IsNullOrWhiteSpace(vRetornoNrImp))
            {

                //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Esse documento já foi salvo.");
                //return;

                // Inserir Cabeçalho de Importacao
                DeleteCabecalhoImportacao(Int32.Parse(vRetornoNrImp));
                DeleteLinhaImportacao(Int32.Parse(vRetornoNrImp));
                InserirCabecalhoImportacao(Int32.Parse(vRetornoNrImp), vTxtTxID, vFrtInt, vOtrDesp, vContainer, vTotalPeso, vTotalME, vTotal, "ABERTO", Hoje, vTotalII);
            }
            else
            {
                // Inserir Cabeçalho de Importacao
                InserirCabecalhoImportacao(0, vTxtTxID, vFrtInt, vOtrDesp, vContainer, vTotalPeso, vTotalME, vTotal, "ABERTO", Hoje, vTotalII);
            }

            for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
            {
                //if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                //{
                //    try
                //    {
                //        int vPedido = Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString());
                //        int vPedNumDoc = Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", y).ToString());
                //        int vVisOrder = Int32.Parse(this.gridDados.DataTable.GetValue("VisOrder", y).ToString());
                //        string vProduto = this.gridDados.DataTable.GetValue("ItemCode", y).ToString();
                //        //string vDescricao = this.gridDados.DataTable.GetValue(6, y).ToString().Replace("'", " ").Replace("&", " ");
                //        string vDescricao = "";
                //        double vPrecoUnit = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", y).ToString());
                //        double vPrecoTotal = Convert.ToDouble(this.gridDados.DataTable.GetValue("LineTotal", y).ToString());
                //        double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", y).ToString());
                //        double vQuantidadeAberta = Convert.ToDouble(this.gridDados.DataTable.GetValue("OpenQty", y).ToString());
                //        string vUm = this.gridDados.DataTable.GetValue("unitMsr", y).ToString();
                //        double vPeso = Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", y).ToString());
                //        double vFrete = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", y).ToString());
                //        double vOutrasDespesas = Convert.ToDouble(this.gridDados.DataTable.GetValue("OtrDesp", y).ToString());
                //        int vItens = Int32.Parse(this.gridDados.DataTable.GetValue("NumPerMsr", y).ToString());
                //        int vLinPed = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", y).ToString());
                //        double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", y).ToString());

                //        string vDeposito = "01";

                //        if (!String.IsNullOrWhiteSpace(vRetornoNrImp))
                //        {
                //            // Inserir Linha de Importacao
                //            InserirLinhaImportacao(Int32.Parse(vRetornoNrImp), vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vVisOrder, vTaxRate);
                //        }
                //        else
                //        {
                //            // Inserir Linha de Importacao
                //            InserirLinhaImportacao(0, vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vVisOrder, vTaxRate);
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        throw;
                //    }

                //}
            }

            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Importação foi salva com sucesso.");

        }

        private void btnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            int intRetorno;
            intRetorno = SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Informações não gravadas serão perdidas. Continuar?", 1, "Sim", "Não", "");

            if (intRetorno == 1)
            {
                BubbleEvent = true;
                SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                oForm.Close();
            }
            else
            {
                BubbleEvent = false;
            }
        }

        private void btnCalcular_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                double vTxtTxID;
                if (!String.IsNullOrWhiteSpace(txtTaxaDI.Value))
                {
                    vTxtTxID = Convert.ToDouble(txtTaxaDI.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vTxtTxID = 0.00;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                    return;
                }

                double vFrtInt;
                if (!String.IsNullOrWhiteSpace(txtFreteInternacional.Value))
                {
                    vFrtInt = Convert.ToDouble(txtFreteInternacional.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vFrtInt = 0;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional.");
                    return;
                }

                double vOtrDesp;
                if (!String.IsNullOrWhiteSpace(txtOutraDespesa.Value))
                {
                    vOtrDesp = Convert.ToDouble(txtOutraDespesa.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vOtrDesp = 0;
                }

                double TotalPeso = 0;
                double TotalME = 0;
                double TotalII = 0;
                double Total = 0;
                int contador = 0;

                /*
                for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
                {

                    if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                    {
                        Double vTotalFrete = ((vFrtInt / TotalPeso) * Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", y).ToString()));
                        this.gridDados.DataTable.SetValue("Frete", y, vTotalFrete.ToString("N4"));
                        Double vCalOtrDesp = (vOtrDesp / contador);
                        this.gridDados.DataTable.SetValue("OtrDesp", y, (vCalOtrDesp.ToString("N4")));

                        txtPeso.Value = TotalPeso.ToString("N3");
                        txtTotalFOB.Value = TotalME.ToString("N2");

                    }
                }
                */

                //for (int z = 0; z <= this.dtDados.Rows.Count - 1; z++)
                //{

                //    if (this.gridDados.DataTable.GetValue(0, z).ToString() == "Y")
                //    {
                //        TotalPeso = TotalPeso + Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", z).ToString());
                //        if (TotalPeso == 0)
                //        {
                //            TotalPeso = 1;
                //        }
                //    }
                //}

                //for (int x = 0; x <= this.dtDados.Rows.Count - 1; x++)
                //{

                //    if (this.gridDados.DataTable.GetValue(0, x).ToString() == "Y")
                //    {
                //        Double vPesoLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", x).ToString());
                //        Double vTotalFrete = ((vFrtInt / TotalPeso) * vPesoLinha);
                //        this.gridDados.DataTable.SetValue("Frete", x, vTotalFrete.ToString("N4"));

                //        TotalME = TotalME + Convert.ToDouble(this.gridDados.DataTable.GetValue("LineTotal", x).ToString());
                //        contador = contador + 1;
                //        Double vCalOtrDesp = (vOtrDesp / contador);
                //        this.gridDados.DataTable.SetValue("OtrDesp", x, (vCalOtrDesp.ToString("N4")));

                //        Double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", x).ToString());
                //        Double vPreco = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", x).ToString());
                //        Double vFreteLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", x).ToString());
                //        Double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", x).ToString());

                //        TotalII = TotalII + (((vTxtTxID * (vPreco * vQuantidade)) + (vFreteLinha)) * (vTaxRate / 100));

                //    }

                //}

                txtPeso.Value = TotalPeso.ToString("N3");
                txtTotal_II.Value = TotalII.ToString("N2");
                txtTotalFOB.Value = TotalME.ToString("N2");

                Total = (((vTxtTxID * TotalME) + vFrtInt) + TotalII);
                txtTotalRS.Value = Total.ToString("N2");

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }

        }

        private void btnGerarNF_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Geração em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            //Processo para Salvar antes de inserir o documento
            string vRetornoNrImp;
            string vCodImp = txtCodigoImportacao.Value;
            //if (string.IsNullOrWhiteSpace(vCodImp) || vCodImp !String.IsNullOrWhiteSpace())
            if (!String.IsNullOrWhiteSpace(vCodImp))
            {
                // vCodImp = txtCodImp.Value;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
                if (!String.IsNullOrWhiteSpace(vRetornoNrImp) && !String.IsNullOrWhiteSpace(vCodImp))
                {
                    vRetornoNrImp = "";
                }
                else
                {
                    vRetornoNrImp = "0";
                }
                //vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }
            else
            {
                vCodImp = txtCodigoImportacao.Value;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }


            double vTxtTxID;
            if (!String.IsNullOrWhiteSpace(txtTaxaDI.Value))
            {
                vTxtTxID = Convert.ToDouble(txtTaxaDI.Value.ToString().Replace(".", ","));
            }
            else
            {
                vTxtTxID = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                return;
            }

            if ((!String.IsNullOrWhiteSpace(vRetornoNrImp)) || (vRetornoNrImp == "0"))
            //if (vRetornoNrImp !String.IsNullOrWhiteSpace())
            {

                double vFrtInt;
                if (!String.IsNullOrWhiteSpace(txtFreteInternacional.Value))
                {
                    vFrtInt = Convert.ToDouble(txtFreteInternacional.Value.ToString().Replace(".", ","));

                }
                else
                {
                    vFrtInt = 0;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional");
                    return;
                }

                double vOtrDesp;
                if (!String.IsNullOrWhiteSpace(txtOutraDespesa.Value))
                {
                    vOtrDesp = Convert.ToDouble(txtOutraDespesa.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vOtrDesp = 0;
                }

                double vContainer;
                if (!String.IsNullOrWhiteSpace(txtContainer.Value))
                {
                    vContainer = Convert.ToDouble(txtContainer.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vContainer = 0;
                }

                Double vTotalPeso = Convert.ToDouble(txtPeso.Value.PadLeft(3, '0'));
                Double vTotalME = Convert.ToDouble(txtTotalFOB.Value.PadLeft(2, '0'));
                Double vTotalII = Convert.ToDouble(txtTotal_II.Value.PadLeft(2, '0'));
                Double vTotal = Convert.ToDouble(txtTotalRS.Value.PadLeft(2, '0'));

                if ((!String.IsNullOrWhiteSpace(vRetornoNrImp)) || (vRetornoNrImp == "0"))
                {
                    // Inserir Cabeçalho de Importacao
                    InserirCabecalhoImportacao(0, vTxtTxID, vFrtInt, vOtrDesp, vContainer, vTotalPeso, vTotalME, vTotal, "ABERTO", DateTime.Now.ToShortDateString(), vTotalII);

                    //for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
                    //{
                    //    if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                    //    {

                    //        int vPedido = Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString());
                    //        int vPedNumDoc = Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", y).ToString());
                    //        int vLinPed = Int32.Parse(this.gridDados.DataTable.GetValue("VisOrder", y).ToString()) - 1;
                    //        string vProduto = this.gridDados.DataTable.GetValue("ItemCode", y).ToString();
                    //        //string vDescricao = this.gridDados.DataTable.GetValue(6, y).ToString();
                    //        string vDescricao = "";
                    //        double vPrecoUnit = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", y).ToString());
                    //        double vPrecoTotal = Convert.ToDouble(this.gridDados.DataTable.GetValue("LineTotal", y).ToString());
                    //        double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", y).ToString());
                    //        double vQuantidadeAberta = Convert.ToDouble(this.gridDados.DataTable.GetValue("OpenQty", y).ToString());
                    //        string vUm = this.gridDados.DataTable.GetValue("unitMsr", y).ToString();
                    //        double vPeso = Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", y).ToString());
                    //        double vFrete = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", y).ToString());
                    //        double vOutrasDespesas = Convert.ToDouble(this.gridDados.DataTable.GetValue("OtrDesp", y).ToString());
                    //        int vItens = Int32.Parse(this.gridDados.DataTable.GetValue("NumPerMsr", y).ToString());
                    //        int vLineNum = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", y).ToString());
                    //        double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", y).ToString());
                    //        string vDeposito = "01";

                    //        if ((!String.IsNullOrWhiteSpace(vRetornoNrImp)))
                    //        {
                    //            // Inserir Linha de Importacao
                    //            InserirLinhaImportacao(Int32.Parse(vRetornoNrImp), vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vLineNum, vTaxRate);
                    //        }
                    //        else
                    //        {
                    //            // Inserir Linha de Importacao
                    //            InserirLinhaImportacao(0, vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vLineNum, vTaxRate);
                    //        }

                    //    }
                    //}
                }

                // Fim Processo de Salvar Documento
            }

            // Geração da Nota Fiscal
            var vEsbocoNFRecebimento = (SAPbobsCOM.Documents)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.oDrafts);
            vEsbocoNFRecebimento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;

            //for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
            //{
            //    if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
            //    {
            //        vEsbocoNFRecebimento.CardCode = RetornaFornecedor(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString()));
            //        vEsbocoNFRecebimento.DocDate = DateTime.Today;
            //        vEsbocoNFRecebimento.DocDueDate = DateTime.Today;
            //        vEsbocoNFRecebimento.TaxDate = DateTime.Today;
            //        vEsbocoNFRecebimento.DocCurrency = "R$";

            //        vEsbocoNFRecebimento.Comments = "Recebimento gerado por integração no dia: " + DateTime.Now.ToString();
            //        vEsbocoNFRecebimento.BPL_IDAssignedToInvoice = 1;
            //        //vEsbocoNFRecebimento.UserFields.Fields.Item("U_MW_PROJETO").Value = 99001000;

            //        // Condição de PAgamento no SAP
            //        vEsbocoNFRecebimento.GroupNumber = 100;

            //        string utilizacao = RetornaUtizacao(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString()));

            //        for (Int32 i = 0; i <= this.dtDados.Rows.Count - 1; i++)
            //        {
            //            if (this.gridDados.DataTable.GetValue(0, i).ToString() == "Y")
            //            {
            //                vEsbocoNFRecebimento.Lines.ItemCode = this.gridDados.DataTable.GetValue("ItemCode", i).ToString();
            //                vEsbocoNFRecebimento.Lines.Quantity = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", i).ToString());

            //                // Converter de Dolar para Real usando TX ID - Multiplicando. Somar o peso pelas linhas selecionada e depois dividir frete pelo total e multiplicar pelo peso da linha   
            //                Double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", i).ToString());
            //                Double vPreco = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", i).ToString());
            //                Double vFreteLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", i).ToString());
            //                Double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", i).ToString());

            //                vEsbocoNFRecebimento.Lines.UnitPrice = (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) + (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) * (vTaxRate / 100)));

            //                vEsbocoNFRecebimento.Lines.ShipDate = DateTime.Today;
            //                vEsbocoNFRecebimento.Lines.TaxCode = RetornaCodImposto(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", i).ToString()), Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", i).ToString()));

            //                //Utilização
            //                vEsbocoNFRecebimento.Lines.Usage = utilizacao;

            //                string vDocEntryPedido = RetornaPedDocEntry(Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", i).ToString()));
            //                //string vDocEntryPedido = RetornaPedDocEntry(Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", i).ToString()));

            //                //Amarração com Pedido de Compra
            //                vEsbocoNFRecebimento.Lines.BaseType = 22;
            //                vEsbocoNFRecebimento.Lines.BaseEntry = Int32.Parse(vDocEntryPedido);
            //                vEsbocoNFRecebimento.Lines.BaseLine = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", i).ToString());

            //                vEsbocoNFRecebimento.Lines.Add();
            //            }
            //        }

            //        int vRetorno = vEsbocoNFRecebimento.Add();

            //        if (vRetorno != 0)
            //        {
            //            string MessagemErro = ConexaoSAP.Company.GetLastErrorDescription();
            //            throw new Exception(MessagemErro);
            //        }
            //        else
            //        {
            //            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Esboço de Recebimento criado com sucesso.");
            //            string draftDocEntry = RetornaCodDraft();
            //            SAPbouiCOM.Framework.Application.SBO_Application.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
            //            return;
            //        }
            //    }
            //}
        }

        #endregion Events


        #region Methods

        #region Repository

        public string GetRecordsetAsString(string query)
        {
            Recordset recordset = null;

            try
            {
                recordset = (Recordset)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);

                if (!recordset.EoF)
                    return recordset.Fields.Item(0).Value.ToString().Trim();

                return string.Empty;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (recordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                recordset = null;
            }
        }

        public void ExecuteQuery(string query)
        {
            Recordset recordset = null;

            try
            {
                recordset = (Recordset)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (recordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                recordset = null;
            }
        }

        private string QueryPesquisar(string pnCode, string pedidoCodigo, string processoCodigo)
        {
            string query = "SELECT " +
                           "'N' AS Selected, " +
                           "OPOR.DocEntry, " +
                           "OPOR.DocNum, " +
                           "POR1.LineNum, " +
                           "POR1.ItemCode, " +
                           "POR1.Dscription, " +
                           "POR1.Price, " +
                           "POR1.TotalFrgn as LineTotal, " +
                           "POR1.OpenQty, " +
                           "POR1.Quantity, " +
                           "POR1.unitMsr, " +
                           "POR1.Weight1 as Peso, " +
                           "POR1.NumPerMsr, " +
                           "ONCM.NcmCode, " +
                           "Por1.VisOrder, " +
                           "ISNULL(POR4.TaxRate,1) AS TaxRate" +
                       " FROM OPOR " +
                           " INNER JOIN POR1 ON POR1.DocEntry = OPOR.DocEntry " +
                           " INNER JOIN OITM ON OITM.ItemCode = POR1.ItemCode " +
                           " LEFT JOIN POR4 ON POR1.DocEntry = POR4.DocEntry and POR1.LineNum = POR4.LineNum and POR4.staType = 23 " +
                           " LEFT JOIN ONCM ON ONCM.AbsEntry = OITM.NcmCode " +
                       " WHERE " +
                           " OPOR.DocStatus <> 'C'" +
                           " AND POR1.LineStatus <> 'C'" +
                           " AND POR1.Currency <> 'R$' " +
                            (String.IsNullOrWhiteSpace(pnCode) ? string.Empty : " AND (OPOR.CardCode = '" + pnCode + "' )") +
                            (String.IsNullOrWhiteSpace(pedidoCodigo) ? string.Empty : " AND (OPOR.DocNum = " + pedidoCodigo + ") ") +
                            (String.IsNullOrWhiteSpace(processoCodigo) ? string.Empty : " AND (OPOR.U_ALFT_NumPrc = " + processoCodigo + ") ") +
                       " ORDER BY " +
                           " OPOR.DocNum desc, " +
                           " Por1.VisOrder asc ";

            return query;
        }

        private string QueryCarregar(String codigoImportacao)
        {
            string query = "SELECT " +
                                " T1.U_PEDLINITEM as LineNum, " +
                                " T1.U_Pedido as DocEntry, " +
                                " T1.U_PEDDOCNUM as DocNum, " +
                                " T1.U_CODPRODUTO as ItemCode, " +
                                " T1.U_DSCPRODUTO as Dscription, " +
                                " T1.U_PRECO as Price, " +
                                " T1.U_PRECOTOTAL as LineTotal, " +
                                " T1.U_CODPRODUTO as ItemCode, " +
                                " T1.U_DSCPRODUTO as Dscription, " +
                                " T1.U_PRECO as Price, " +
                                " T1.U_QTDPED as Quantity, " +
                                " T1.U_QTDABERTA as OpenQty, " +
                                " T1.U_UM AS unitMsr, " +
                                " T1.U_PESO as Peso, " +
                                " T1.U_ITMNUM AS NumPerMsr, " +
                                " T0.U_TXID AS TXDI, " +
                                " T0.U_FRTINT AS FRTINT, " +
                                " T0.U_OTRDESP AS OTRDESP, " +
                                " T0.U_TTPESO AS TTPESO, " +
                                " T0.U_TTME AS TTME, " +
                                " T0.U_TOTAL AS TOTAL, " +
                                " T1.U_FRETE as FreteLinha, " +
                                " T1.U_OTRDESP as OtrDespLinha, " +
                                " T3.NcmCode as NcmCode, " +
                                " T1.U_PEDLINORDER as VisOrder, " +
                                " T1.U_TAXORDER as TaxRate, " +
                                " T0.U_TAXVALLINE as TaxValLine " +
                            " FROM " +
                                " [@ALFT_IMPORT] as T0 " +
                                " INNER JOIN [@ALFT_IMPORT1] AS T1 ON T1.U_DocEntry = T0.U_DocEntry " +
                                " INNER JOIN OITM AS T2 ON T2.ItemCode = T1.U_CODPRODUTO " +
                                " LEFT JOIN ONCM AS T3 ON T3.AbsEntry = T2.NcmCode " +
                            " WHERE " +
                                " T0.U_DocEntry = " + codigoImportacao;
            return query;
        }


        #endregion Repository

        #region SELECT

        public string RetornaFornecedor(int docEntry)
        {
            return GetRecordsetAsString(String.Format("SELECT CardCode As CODFORNECEDOR FROM OPOR WHERE DocEntry = {0}", docEntry));
        }

        public string RetornaUtizacao(int docEntry)
        {
            return GetRecordsetAsString(String.Format("SELECT MainUsage as UTILIZACAO FROM POR12 WHERE DocEntry = {0}", docEntry));
        }

        public string RetornaPedDocEntry(int docNum)
        {
            return GetRecordsetAsString(String.Format("SELECT DocEntry as CODIGO FROM OPOR WHERE DocNum = {0}", docNum));
        }

        public string RetornaDescItem(string itemCode)
        {
            return GetRecordsetAsString(String.Format("SELECT ItemName as DESCRICAO FROM OITM WHERE ItemCode = '{0}'", itemCode));
        }

        public string RetornaCodImposto(int docEntry, int lineNum)
        {
            return GetRecordsetAsString(String.Format("SELECT POR1.TaxCode as TAXCODE FROM OPOR INNER JOIN POR1 ON POR1.DocEntry = OPOR.DocEntry WHERE POR1.DocEntry = {0} AND POR1.LineNum = {1}", docEntry, lineNum));
        }

        public string RetornaCodDraft()
        {
            return GetRecordsetAsString("SELECT TOP 1 DocEntry as CODIGO FROM ODRF ORDER BY DocEntry DESC");
        }

        public string RetornaCodeLog()
        {
            return GetRecordsetAsString("SELECT Top 1 Convert(int, Code) + 1 AS NEWCODE FROM [@ALFT_IMPORT] order by Convert(int,code) desc");
        }

        public string RetornaCodeLogLinha()
        {
            return GetRecordsetAsString("SELECT TOP 1 CONVERT(int, Code) as Codigo FROM [@ALFT_IMPORT1] order by Codigo desc ");
        }

        public string RetornaNrImportacao()
        {
            return GetRecordsetAsString("SELECT TOP 1 (U_DocEntry)+1 AS NRIMPORT FROM [@ALFT_IMPORT] order by Convert(int,code) desc");
        }

        public string RetornaExisteImportacao(int docEntry)
        {
            return GetRecordsetAsString(String.Format("SELECT ISNULL(U_DocEntry,0) AS NRIMPORT FROM [@ALFT_IMPORT] WHERE U_DocEntry = {0}", docEntry));
        }

        #endregion Consultas

        #region Insert

        public void InserirCabecalhoImportacao(int pDocEntry, double pTxId, double pFrtInt, double pOtrDesp, double pContainer, double pTotalPeso, double pTotalMe, double pTotal, string pStatus, string pData, double pTotalII)
        {
            string code;
            string vDocEntry;
            if (pDocEntry == 0)
            {
                code = RetornaCodeLog();
                vDocEntry = code;
            }
            else
            {
                int Retorno = Int32.Parse(RetornaCodeLog()) + 1;
                code = Retorno.ToString();
                vDocEntry = pDocEntry.ToString();
            }

            var insertQuery = String.Format("INSERT INTO [@ALFT_IMPORT] " +
                              "VALUES('{0}', '{0}', '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11})",
                                       code, vDocEntry, ParseGlobalization(pTxId), ParseGlobalization(pFrtInt), ParseGlobalization(pOtrDesp), ParseGlobalization(pContainer), ParseGlobalization(pTotalPeso), ParseGlobalization(pTotalMe),
                                       ParseGlobalization(pTotal), pStatus, pData, ParseGlobalization(pTotalII));

            ExecuteQuery(insertQuery);
        }

        public void InserirLinhaImportacao(int pDocEntry, int pPedido, string pProduto, string pDescricao, double pPreco, double pPrecoTotal, double pQtdPedida, double pQtdAberta, double pQtdFat, string pUm, double pPeso, double pFrete, double pOutraDesp, int pItmUm, string pDeposito, int pPedNumDoc, int pLinPed, int pVisOrder, double pTaxOrder)
        {
            int code;
            int vDocEntry;
            int codelinha;
            if (pDocEntry == 0)
            {
                string Retorno = RetornaCodeLogLinha();
                if (!String.IsNullOrWhiteSpace(Retorno))
                {
                    codelinha = Int32.Parse(RetornaCodeLogLinha()) + 1;
                }
                else
                {
                    codelinha = 1;
                }

                code = Int32.Parse(RetornaCodeLog());
                vDocEntry = code - 1;

            }
            else
            {
                codelinha = Int32.Parse(RetornaCodeLogLinha()) + 1;
                //code = Int32.Parse(RetornaCodeLog());
                vDocEntry = pDocEntry;
            }

            var insertQuery = "insert into [@ALFT_IMPORT1] values(\'" + codelinha + "\',\'" + codelinha + "\',\'" + vDocEntry + "\',\'" + pPedido + "\',\'" + pProduto + "\',\'" + pDescricao + "\', " + pPreco.ToString().Replace(",", ".") + ", " + pPrecoTotal.ToString().Replace(",", ".") + ", " + pQtdPedida.ToString().Replace(",", ".") + ", " + pQtdAberta.ToString().Replace(",", ".") + ", " + pQtdFat.ToString().Replace(",", ".") + ", \'" + pUm + "\', " + pPeso.ToString().Replace(",", ".") + ", " + pFrete.ToString().Replace(",", ".") + " , " + pOutraDesp.ToString().Replace(",", ".") + ",\'" + pItmUm + "\',\'" + pDeposito + "\',\'" + pPedNumDoc + "\',\'" + pLinPed + "\',\'" + pVisOrder + "\', " + pTaxOrder.ToString().Replace(",", ".") + ")";
            ExecuteQuery(insertQuery);

        }

        #endregion

        #region Delete

        public void DeleteCabecalhoImportacao(int vDocEntry)
        {
            var ret = 0;
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                string code = RetornaCodeLog();

                var sql = "DELETE FROM [@ALFT_IMPORT] WHERE U_DocEntry = " + vDocEntry;
                //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("3 " + sql);
                oRs.DoQuery(sql);

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (oRs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }

        public void DeleteLinhaImportacao(int vDocEntry)
        {
            string code = RetornaCodeLog();

            var query = "DELETE FROM [@ALFT_IMPORT1] WHERE U_DocEntry = " + vDocEntry;
            //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("3 " + sql);
            ExecuteQuery(query);
        }

        #endregion

        #region Utils

        public string ParseGlobalization(double value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        public string ParseGlobalization(object value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        #endregion



        #endregion Methods
    }
}