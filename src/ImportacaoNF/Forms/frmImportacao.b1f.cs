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
using SLT.ImportacaoNF.UserDefinedObjects.Models;

namespace SLT.ImportacaoNF
{
    [FormAttribute("OSLT_IMPORT", "Forms/frmImportacao.b1f")]
    class frmImportacao : UserFormBase
    {
        #region Attributes

        /*
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
        //private SAPbouiCOM.Button btnPesquisar;

        private StaticText StaticText7;
        private StaticText StaticText8;
        private StaticText StaticText9;
        private EditText txtPeso;
        private EditText txtTotalFOB;
        private EditText txtTotalRS;

        private StaticText StaticText4;
        private EditText txtOutraDespesa;
        private EditText txtCodePN;
        //private Button btnCalcular;
        private Button btnCancelar;

        private StaticText StaticText10;
        private ComboBox cbStatus;
        private StaticText StaticText11;
        private EditText txtDataDocumento;

        private EditText txtDocEntry;
        private StaticText StaticText12;
        //private StaticText StaticText14;
        //private EditText txtCodigoImportacao;
        //private Button btnCarregar;

        //private SAPbouiCOM.Grid gridDados;

        private StaticText StaticText13;
        private EditText txtTotal_II;

        private Matrix matrixData;
         */

        private SAPbouiCOM.Application SAPApp = null;
        private SAPbouiCOM.EditText oEditText = null;
        private SAPbouiCOM.ComboBox oComboBox = null;
        private SAPbouiCOM.Item oItem = null;
        private SAPbouiCOM.Form oForm = null;
        private SAPbouiCOM.Matrix oMatrix = null;

        private string filtroCodigoPN = string.Empty;
        private string filtroPedido = string.Empty;
        private string filtroProcesso = string.Empty;


        #endregion attributes

        public frmImportacao()
        {
            SAPApp = SAPbouiCOM.Framework.Application.SBO_Application;
            oForm = GetForm();
            return;

            //this.UIAPIRawForm.DataSources.DBDataSources.Item("@SLTIMPORT");
            //this.UIAPIRawForm.Mode = BoFormMode.fm_FIND_MODE;
            //CreateEmptyMatrix();
        }



        private SAPbouiCOM.Form GetForm()
        {
            try
            {
                SAPApp.Forms.Item("SLT_Importacao").Close();

                return SAPApp.Forms.Item("SLT_Importacao");
            }
            catch (Exception ex)
            {
                return CriarFormulario();
            }
        }

        public SAPbouiCOM.Form CriarFormulario()
        {
            SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)SAPApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "SLT_Importacao";
            creationPackage.FormType = "SLT_Importacao";
            creationPackage.ObjectType = "OSLT_IMPORT";
            creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

            oForm = SAPApp.Forms.AddEx(creationPackage);
            oForm.Freeze(true);
            oForm.Title = "Importação";
            oForm.Width = 900;
            oForm.Height = 700;
            SAPApp.ItemEvent += SAPApp_ItemEvent;

            CriarFormularioDefault();
            CriarFormularioFiltros();
            CriarFormularioMatrix();
            CriarFormularioFooter();

            oForm.Visible = true;
            oForm.Freeze(false);

            oForm.Mode = BoFormMode.fm_FIND_MODE;

            return oForm;
        }

        public void SAPApp_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            if (pVal.BeforeAction == false && FormUID == "SLT_Importacao")
            {
                HandleEvent(pVal.EventType, pVal);
            }

            BubbleEvent = true;
        }

        public void HandleEvent(BoEventTypes eventType, ItemEvent pVal)
        {
            switch (eventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;
                case BoEventTypes.et_B1I_SERVICE_COMPLETE:
                    break;
                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    OnChooseFromList(pVal);
                    break;
                case BoEventTypes.et_CLICK:
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_DATASOURCE_LOAD:
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    break;
                case BoEventTypes.et_Drag:
                    break;
                case BoEventTypes.et_EDIT_REPORT:
                    break;
                case BoEventTypes.et_FORMAT_SEARCH_COMPLETED:
                    break;
                case BoEventTypes.et_FORM_ACTIVATE:
                    break;
                case BoEventTypes.et_FORM_CLOSE:
                    break;
                case BoEventTypes.et_FORM_DATA_ADD:
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    break;
                case BoEventTypes.et_FORM_DEACTIVATE:
                    break;
                case BoEventTypes.et_FORM_DRAW:
                    break;
                case BoEventTypes.et_FORM_KEY_DOWN:
                    break;
                case BoEventTypes.et_FORM_LOAD:
                    break;
                case BoEventTypes.et_FORM_MENU_HILIGHT:
                    break;
                case BoEventTypes.et_FORM_RESIZE:
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                    break;
                case BoEventTypes.et_FORM_UNLOAD:
                    break;
                case BoEventTypes.et_FORM_VISIBLE:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_GRID_SORT:
                    break;
                case BoEventTypes.et_ITEM_PRESSED:
                    break;
                case BoEventTypes.et_ITEM_WEBMESSAGE:
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    break;
                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;
                case BoEventTypes.et_MATRIX_LOAD:
                    break;
                case BoEventTypes.et_MENU_CLICK:
                    break;
                case BoEventTypes.et_PICKER_CLICKED:
                    break;
                case BoEventTypes.et_PRINT:
                    break;
                case BoEventTypes.et_PRINT_DATA:
                    break;
                case BoEventTypes.et_PRINT_LAYOUT_KEY:
                    break;
                case BoEventTypes.et_RIGHT_CLICK:
                    break;
                case BoEventTypes.et_UDO_FORM_BUILD:
                    break;
                case BoEventTypes.et_UDO_FORM_OPEN:
                    break;
                case BoEventTypes.et_VALIDATE:
                    break;

                default:
                    break;
            }
        }

        public SAPbouiCOM.Matrix CriarFormularioMatrix()
        {
            oForm.DataSources.DataTables.Add("oMatrixDT");

            oItem = oForm.Items.Add("oMtrx1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);

            oItem = oForm.Items.Item("oMtrx1");
            oItem.Top = 70;
            oItem.Left = 15;
            oItem.Width = oForm.Width - 30;
            oItem.Height = 350;

            oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

            oForm.DataSources.DataTables.Item("oMatrixDT").Clear();

            string sSQL = " SELECT              " +
                          "         DocEntry    " +
                          "    ,'Y' AS [Selected] " +
                          "    ,[LineId]        " +
                          "    ,[U_PedidoId] " +
                          "    ,[U_ItemNum] " +
                          "    ,[U_ItemCode] " +
                          "    ,[U_QtdPed] " +
                          "    ,[U_QtdDisp] " +
                          "    ,[U_QtdFat] " +
                          "    ,[U_Peso] " +
                          "    ,[U_Frete] " +
                          "    ,[U_OutroDes] " +
                          "    ,[U_Deposito] " +
                          "    ,[U_PedTax] " +
                          " FROM [dbo].[@SLTIMPRT1] " +
                          " WHERE DocEntry = -1 ";// + oEditText.Value;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(sSQL);

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("oMtrx1").Specific;
            SAPbouiCOM.Column oColumn = null;

            oColumn = oMatrix.Columns.Add("oClmn0", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            //oColumn.DataBind.SetBound(true, string.Empty, "Selected");
            oColumn.TitleObject.Caption = "#";


            oColumn = oMatrix.Columns.Add("oClmn5", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PedidoId");
            oColumn.TitleObject.Caption = "Nº do Pedido";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_PurchaseOrder;


            oColumn = oMatrix.Columns.Add("oClmn1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "LineId");
            oColumn.TitleObject.Caption = "Nº Lin. Ped.";

            oColumn = oMatrix.Columns.Add("oClmn6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_ItemNum");
            oColumn.TitleObject.Caption = "Nº do item";

            oColumn = oMatrix.Columns.Add("oClmn7", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_ItemCode");
            oColumn.TitleObject.Caption = "Descrição do Item";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_Items;

            oColumn = oMatrix.Columns.Add("oClmn8", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdPed");
            oColumn.TitleObject.Caption = "Qtd. Pedido";

            oColumn = oMatrix.Columns.Add("oClmn9", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdDisp");
            oColumn.TitleObject.Caption = "Qtd Disponível";

            oColumn = oMatrix.Columns.Add("oClmn10", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdFat");
            oColumn.TitleObject.Caption = "Qtd. Faturada";

            oColumn = oMatrix.Columns.Add("oClmn11", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Peso");
            oColumn.TitleObject.Caption = "Peso";

            oColumn = oMatrix.Columns.Add("oClmn12", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Frete");
            oColumn.TitleObject.Caption = "Frete";

            oColumn = oMatrix.Columns.Add("oClmn13", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_OutroDes");
            oColumn.TitleObject.Caption = "Outras Desp.";

            oColumn = oMatrix.Columns.Add("oClmn14", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Deposito");
            oColumn.TitleObject.Caption = "Depósito";

            oColumn = oMatrix.Columns.Add("oClmn15", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PedTax");
            oColumn.TitleObject.Caption = "Taxa";

            //oMatrix.Columns.Item("oClmn0").DataBind.UnBind();
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1); // SLTIMPRT1;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("oMatrixDT");
            oDBDataSource.Clear();


            for (int row = 0; row < oDataTable.Rows.Count; row++)
            {
                int offset = oDBDataSource.Size;
                oDBDataSource.InsertRecord(row);

                //oDBDataSource.SetValue("Selected", offset, oDataTable.GetValue("Selected", row).ToString());
                oDBDataSource.SetValue("LineId", offset, oDataTable.GetValue("LineId", row).ToString());
                //oDBDataSource.SetValue("U_PedidoId", offset, oDataTable.GetValue("U_PedidoId", row).ToString());
                oDBDataSource.SetValue("U_ItemNum", offset, oDataTable.GetValue("U_ItemNum", row).ToString());
                oDBDataSource.SetValue("U_ItemCode", offset, oDataTable.GetValue("U_ItemCode", row).ToString());
                oDBDataSource.SetValue("U_QtdPed", offset, oDataTable.GetValue("U_QtdPed", row).ToString());
                oDBDataSource.SetValue("U_QtdDisp", offset, oDataTable.GetValue("U_QtdDisp", row).ToString());
                oDBDataSource.SetValue("U_QtdFat", offset, oDataTable.GetValue("U_QtdFat", row).ToString());
                oDBDataSource.SetValue("U_Peso", offset, oDataTable.GetValue("U_Peso", row).ToString());
                oDBDataSource.SetValue("U_Frete", offset, oDataTable.GetValue("U_Frete", row).ToString());
                oDBDataSource.SetValue("U_OutroDes", offset, oDataTable.GetValue("U_OutroDes", row).ToString());
                oDBDataSource.SetValue("U_Deposito", offset, oDataTable.GetValue("U_Deposito", row).ToString());
                oDBDataSource.SetValue("U_PedTax", offset, oDataTable.GetValue("U_PedTax", row).ToString());
            }

            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
            oMatrix.ClickAfter += matrixData_ClickAfter;

            return oMatrix;
        }

        void oColumn_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPApp.StatusBar.SetText("Adahuda uahd uadhu", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
        }

        public void OnChooseFromList(ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL;
            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
            
            if (oCFLEvento.BeforeAction == false)
            {
                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
            }
        }

        public void CriarFormularioFiltros()
        {
            oItem = oForm.Items.Add("lblCodPN", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Código do Fornecedor";
            oItem.Top = 5;
            oItem.Left = 15;
            oItem.Width = 110;

            oItem = oForm.Items.Add("txtCodPN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = 5;
            oItem.Left = 135;
            oItem.Width = 150;

            oItem = oForm.Items.Add("lkbCodPN", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItem.LinkTo = "txtCodPN";
            oItem.Top = 5;
            oItem.Left = 115;
            oItem.Width = 20;
            ((SAPbouiCOM.LinkedButton)oItem.Specific).LinkedObject = BoLinkedObject.lf_BusinessPartner;



            oItem = oForm.Items.Add("lblPedido", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Código do Pedido";
            oItem.Top = 25;
            oItem.Left = 15;
            oItem.Width = 100;

            oItem = oForm.Items.Add("cflPedido", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oItem = oForm.Items.Add("txtPedido", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = 25;
            oItem.Left = 135;
            oItem.Width = 150;
            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SAPApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = true;
            oCFLCreationParams.ObjectType = "22";
            oCFLCreationParams.UniqueID = "CFL_PO";
            SAPbouiCOM.ChooseFromList oCFL = oCFLs.Add(oCFLCreationParams);

            oForm.DataSources.DBDataSources.Add("@SLTIMPRT1");

            oEditText = ((SAPbouiCOM.EditText)oItem.Specific);
            oEditText.DataBind.SetBound(true, "@SLTIMPRT1", "Object");
            oEditText.ChooseFromListUID = "CFL_PO";
            oEditText.ChooseFromListAlias = "DocNum";
            oEditText.ChooseFromListAfter += oColumn_ChooseFromListAfter;


            oItem = oForm.Items.Add("lkbPedido", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItem.LinkTo = "cflPedido";
            oItem.Top = 25;
            oItem.Left = 115;
            oItem.Width = 20;
            ((SAPbouiCOM.LinkedButton)oItem.Specific).LinkedObject = BoLinkedObject.lf_PurchaseOrder;




            //oItem = oForm.Items.Add("cflPedido", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oItem.Top = 25;
            //oItem.Left = 290;
            //oItem.Width = 50;


            //SAPbouiCOM.ChooseFromListCreationParams CFL_PO = new ChooseFromListCreationParams();
            //CFL_PO.ObjectType = "";

            //SAPbouiCOM.Conditions oCons = null;
            //SAPbouiCOM.Condition oCon = null;
            
        }

        public void CriarFormularioDefault()
        {
            oItem = oForm.Items.Add("lblEntry", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Nº";
            oItem.Top = 5;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = 5;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            // Now bind Columns to UDO Objects in Add Mode
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "DocEntry");
            //oMatrix.Columns.Item("oClmn0").DataBind.SetBound(true, "@SLTIMPRT1", "DocEntry");
            oForm.DataBrowser.BrowseBy = "3";



            oItem = oForm.Items.Add("lblStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Status";
            oItem.Top = 25;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("ddlStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Top = 25;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            oComboBox.ValidValues.Add(string.Empty, " - ");
            oComboBox.ValidValues.Add("O", "Aberto");
            oComboBox.ValidValues.Add("C", "Fechado");
            oComboBox.DataBind.SetBound(true, "@SLTIMPORT", "Status");



            oItem = oForm.Items.Add("lblData", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Data";
            oItem.Top = 45;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtData", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = 45;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "CreateDate");



            oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oForm.Height - 70;
            oItem.Left = 10;

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oForm.Height - 70;
            oItem.Left = 120;

            oItem = oForm.Items.Add("btnNF", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oForm.Height - 70;
            oItem.Left = oForm.Width - 85;
            var btn = (SAPbouiCOM.Button)oItem.Specific;
            btn.Caption = "Nota Fiscal";
            btn.ClickAfter += btn_ClickAfter;
        }

        public void CriarFormularioFooter()
        {
            var top_position_base = oMatrix.Item.Top + oMatrix.Item.Height + 10;
            var left_position_base = 15;

            oItem = oForm.Items.Add("lblTaxaDI", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Taxa DI";
            oItem.Top = top_position_base;
            oItem.Left = left_position_base;
            oItem.Width = 100;

            oItem = oForm.Items.Add("txtTaxaDI", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = left_position_base + 115;
            oItem.Width = 150;

            oItem = oForm.Items.Add("lblPesoTot", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Peso Total";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtPesoTot", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;




            top_position_base += 25;

            oItem = oForm.Items.Add("lblFrete", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Frete Internacional R$";
            oItem.Top = top_position_base;
            oItem.Left = left_position_base;
            oItem.Width = 105;

            oItem = oForm.Items.Add("txtFrete", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = left_position_base + 115;
            oItem.Width = 150;


            oItem = oForm.Items.Add("lblFOB", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "FOB Total";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtFOB", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;



            top_position_base += 25;
            oItem = oForm.Items.Add("lblOutDesp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Outras Despesas";
            oItem.Top = top_position_base;
            oItem.Left = left_position_base;
            oItem.Width = 100;

            oItem = oForm.Items.Add("txtOutDesp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = left_position_base + 115;
            oItem.Width = 150;



            oItem = oForm.Items.Add("lblTotal2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Total II";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtTotal2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            top_position_base += 25;
            oItem = oForm.Items.Add("lblCntner", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Container";
            oItem.Top = top_position_base;
            oItem.Left = left_position_base;
            oItem.Width = 100;

            oItem = oForm.Items.Add("txtCntner", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = left_position_base + 115;
            oItem.Width = 150;


            oItem = oForm.Items.Add("lblTotal", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Total";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtTotal", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
        }

        void btn_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPApp.StatusBar.SetText("Clicou no NF", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
        }

        #region Events

        void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
        }

        /*
        void matrixData_ClickAfterOld(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID.Equals("oClmn0"))
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);
                oDBDataSource.SetValue("Selected", pVal.Row, pVal.ActionSuccess ? "Y" : "N");
                Calcular();
            }
        }
         */

        void matrixData_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID.Equals("oClmn0"))
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);
                oDBDataSource.SetValue("Selected", pVal.Row, pVal.ActionSuccess ? "Y" : "N");
                Calcular();
            }
        }

        private void Calcular()
        {
            SAPApp.StatusBar.SetText("Atualizando os valores, aguarde alguns instantes!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            SAPApp.Forms.ActiveForm.Freeze(true);

            var columns = oMatrix.Columns;
            double total = 0;
            double price = 0;
            double qtd = 0;
            double peso = 0;

            bool selected = false;
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);

            for (int i = 0; i < oDBDataSource.Size; i++)
            {
                selected = oDBDataSource.GetValue("Selected", i).ToString().Equals("Y");

                if (selected)
                {
                    //price += Convert.ToDouble(this.dtDados.GetValue("U_PedTax", i));
                    //qtd += Convert.ToDouble(this.dtDados.GetValue("U_QtdDisp", i));
                    //peso += Convert.ToDouble(this.dtDados.GetValue("U_Peso", i));
                    total += price * qtd;
                }
            }

            //txtPeso.Value = peso.ToString("N3");
            //txtTotal_II.Value = price.ToString("N2");
            //txtTotalFOB.Value = price.ToString("N2");
            //txtTotalRS.Value = total.ToString("N2");

            SAPApp.StatusBar.SetText("Valores atualizados!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            SAPApp.Forms.ActiveForm.Freeze(false);
        }





























        /*
        private void txtCodePN_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (filtroCodigoPN != txtCodePN.Value)
            {
                if (!ChangeFilter())
                    txtCodePN.Value = string.Empty;
                else
                    filtroCodigoPN = txtCodePN.Value;
            }
        }

        private void txtCodigoPedido_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (filtroPedido != txtCodigoPedido.Value)
            {
                if (!ChangeFilter())
                    txtCodigoPedido.Value = string.Empty;
                else
                    filtroPedido = txtCodigoPedido.Value;
            }
        }

        private void txtProcesso_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (filtroProcesso != txtProcesso.Value)
            {
                if (!ChangeFilter())
                    txtProcesso.Value = string.Empty;
                else
                    filtroProcesso = txtProcesso.Value;
            }
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.lblProcesso = ((SAPbouiCOM.StaticText)(this.GetItem("lblProc").Specific));
            this.txtProcesso = ((SAPbouiCOM.EditText)(this.GetItem("txtProces").Specific));
            this.txtProcesso.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.txtProcesso_LostFocusAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lblPed").Specific));
            this.txtCodigoPedido = ((SAPbouiCOM.EditText)(this.GetItem("txtPedido").Specific));
            this.txtCodigoPedido.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.txtCodigoPedido_LostFocusAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lblCdPN").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.txtTaxaDI = ((SAPbouiCOM.EditText)(this.GetItem("txtTxID").Specific));
            this.btnSalvar = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.btnSalvar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnSalvar_ClickBefore);
            //        this.btnSalvar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnSalvar_ClickBefore);
            this.btnGerarNF = ((SAPbouiCOM.Button)(this.GetItem("Item_13").Specific));
            //        this.btnGerarNF.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnGerarNF_ClickBefore);
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.txtFreteInternacional = ((SAPbouiCOM.EditText)(this.GetItem("txtFrtInt").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.txtContainer = ((SAPbouiCOM.EditText)(this.GetItem("txtConta").Specific));
            //        this.btnPesquisar = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            //        this.btnPesquisar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnPesquisar_ClickBefore);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.txtOutraDespesa = ((SAPbouiCOM.EditText)(this.GetItem("txtODesp").Specific));
            this.txtCodePN = ((SAPbouiCOM.EditText)(this.GetItem("txtCodPN").Specific));
            this.txtCodePN.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.txtCodePN_LostFocusAfter);
            //        this.btnCalcular = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            //        this.btnCalcular.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCalcular_ClickBefore);
            this.btnCancelar = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.btnCancelar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCancelar_ClickBefore);
            //        this.btnCancelar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCancelar_ClickBefore);
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.txtPeso = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.txtTotalFOB = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.txtTotalRS = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.cbStatus = ((SAPbouiCOM.ComboBox)(this.GetItem("cbStatus").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.txtDataDocumento = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.txtDocEntry = ((SAPbouiCOM.EditText)(this.GetItem("txtEntry").Specific));
            //       this.txtDataDocumento.Value = typeof(System.DateTime).Now.ToShortDateString();
            //             Numero da Importação
            //         this.txtCodImp.Value = this.RetornaNrImportacao();
            //         this.ComboBox0.Select("Aberto", typeof(SAPbouiCOM.BoSearchKey).psk_ByDescription);
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            //        this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            //        this.txtCodigoImportacao = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            //        this.btnCarregar = ((SAPbouiCOM.Button)(this.GetItem("btnCarreg").Specific));
            //        this.btnCarregar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCarregar_ClickBefore);
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.txtTotal_II = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.matrixData = ((SAPbouiCOM.Matrix)(this.GetItem("Item_22").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkbPN").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkbPedido").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.ActivateAfter += new ActivateAfterHandler(this.Form_ActivateAfter);

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

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            this.GetItem("lblCdPN").Click();

            if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                EnableField("txtCodPN");
                EnableField("txtPedido");
                EnableField("txtProces");
                EnableField("txtEntry");
                txtDataDocumento.Value = string.Empty;

                cbStatus.Select(1, BoSearchKey.psk_Index);
                EnableField("cbStatus", true);
            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                EnableField("txtCodPN");
                EnableField("txtPedido");
                EnableField("txtProces");
                EnableField("txtEntry", false);
                txtDocEntry.Value = ConexaoSAP.Company.GetNewObjectKey();
                txtDataDocumento.Value = DateTime.Today.ToShortDateString();

                cbStatus.Select(1, BoSearchKey.psk_Index);
                EnableField("cbStatus");

            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE)
            {
                EnableField("txtCodPN", false);
                EnableField("txtPedido", false);
                EnableField("txtProces", false);
                EnableField("txtEntry", false);
                txtDocEntry.Value = string.Empty;
                txtDataDocumento.Value = string.Empty;
                cbStatus.Select(0, BoSearchKey.psk_Index);
                EnableField("cbStatus");
            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                EnableField("txtCodPN", false);
                EnableField("txtPedido", false);
                EnableField("txtProces", false);
                EnableField("txtEntry", false);
                txtDataDocumento.Value = string.Empty;
                EnableField("cbStatus");
            }
        }

        private void EnableField(string fieldname, bool enable = true)
        {
            try
            {
                this.GetItem(fieldname).Enabled = enable;
            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        private void btnSalvar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var formMode = (SAPbouiCOM.BoFormMode)pVal.FormMode;

            if (formMode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || formMode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || formMode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE)
            {
                Salvar();
            }
            else if (formMode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                FindData();
            }
        }

        private void btnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Cancelar();
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

        private string QueryCarregar(int docEntry)
        {
            string query = "SELECT " +
                           " 'Y' AS Selected" +
                           " ,[DocEntry]    " +
                           " ,[U_PedidoId]  AS DocNum " +
                           " ,[U_ItemNum]   AS LineNum " +
                           " ,[U_ItemCode]  AS ItemCode, " +
                           " '' AS Dscription, " +
                           " '' AS Price, " +
                           " '' AS TotalFrgn, " +
                           " '' AS LineTotal, " +
                           " '' AS OpenQty, " +
                           " '' AS Quantity, " +
                           " '' AS unitMsr, " +
                           " '' AS Peso, " +
                           " '' AS NumPerMsr, " +
                           " '' AS NcmCode, " +
                           " '' AS VisOrder, " +
                           " '' AS TaxRate" +
                //" ,[LineId]      " +
                //" ,[VisOrder]    " +
                //" ,[Object]      " +
                //" ,[LogInst]     " +
                //" ,[U_QtdPed]    " +
                //" ,[U_QtdDisp]   " +
                //" ,[U_QtdFat]    " +
                //" ,[U_Peso]     as Peso " +
                //" ,[U_Frete]     " +
                //" ,[U_OutroDes]  " +
                //" ,[U_Deposito]  " +
                //" ,[U_PedTax]    " +
                           " FROM [@SLTIMPRT1] " +
                           " WHERE DocEntry = " + docEntry +
                           " ORDER BY " +
                           " LineId ";

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
         * 
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
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                string code = RetornaCodeLog();

                var sql = "DELETE FROM [@ALFT_IMPORT] WHERE U_DocEntry = " + vDocEntry;
                //SAPAppMessageBox("3 " + sql);
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
            //SAPApp.MessageBox("3 " + sql);
            ExecuteQuery(query);
        }

        #endregion
        
        */

        #region Utils

        public string ParseGlobalization(double value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        private LinkedButton LinkedButton0;
        private LinkedButton LinkedButton1;

        public string ParseGlobalization(object value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        #endregion

        /*

        private void CreateEmptyMatrix()
        {
            var dataTableId = "dtDados";
            this.dtDados = this.UIAPIRawForm.DataSources.DataTables.Item(dataTableId);
            var columns = matrixData.Columns;

            // setup columns
            AddMatrixColumn(columns, BoFormItemTypes.it_CHECK_BOX, "#", "Selected");
            AddMatrixLinkedButtonColumn(columns, BoLinkedObject.lf_PurchaseOrder, "DocEntry", "Pedido de Compra", false);
            AddMatrixLinkedButtonColumn(columns, BoLinkedObject.lf_Items, "ItemCode", "Cód. Item", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "Dscription", "Descrição do Item", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "LineNum", "Nº Linha do Pedido", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "Price", "Preço Unitário", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "Quantity", "Qtd.", true);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "LineTotal", "Total da Linha", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "OpenQty", "Qtd. Aberto", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "UnitMsr", "UM", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "Peso", "Peso", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "NumPerMsr", "NumPerMsr", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "NcmCode", "Cód. NCM", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "VisOrder", "VisOrder", false);
            AddMatrixColumn(columns, BoFormItemTypes.it_EDIT, "TaxRate", "TaxRate", false);

            this.matrixData.SelectionMode = BoMatrixSelect.ms_Auto;
            matrixData.ClickAfter += matrixData_ClickAfterOld;

            matrixData.AutoResizeColumns();
        }

        private void BindData(Columns columns, string dataTableId)
        {
            BindMatrixColumn(columns.Item(0), dataTableId, "Selected");
            BindMatrixColumn(columns.Item(1), dataTableId, "DocEntry");
            BindMatrixColumn(columns.Item(2), dataTableId, "ItemCode");
            BindMatrixColumn(columns.Item(3), dataTableId, "Dscription");
            BindMatrixColumn(columns.Item(4), dataTableId, "LineNum");
            BindMatrixColumn(columns.Item(5), dataTableId, "Price");
            BindMatrixColumn(columns.Item(6), dataTableId, "Quantity");
            BindMatrixColumn(columns.Item(7), dataTableId, "LineTotal");
            BindMatrixColumn(columns.Item(8), dataTableId, "OpenQty");
            BindMatrixColumn(columns.Item(9), dataTableId, "UnitMsr");
            BindMatrixColumn(columns.Item(10), dataTableId, "Peso");
            BindMatrixColumn(columns.Item(11), dataTableId, "NumPerMsr");
            BindMatrixColumn(columns.Item(12), dataTableId, "NcmCode");
            BindMatrixColumn(columns.Item(13), dataTableId, "VisOrder");
            BindMatrixColumn(columns.Item(14), dataTableId, "TaxRate");
        }

        private void FindData()
        {
            SAPApp.Forms.ActiveForm.Freeze(true);

            if (!String.IsNullOrWhiteSpace(txtDocEntry.Value))
            {
                LoadData();
                return;
            }

            // load the data into the rows
            string pn = txtCodePN.Value;
            string pedido = txtCodigoPedido.Value;
            string processo = txtProcesso.Value;

            if (String.IsNullOrWhiteSpace(pn) && String.IsNullOrWhiteSpace(pedido) && String.IsNullOrWhiteSpace(processo))
            {
                SAPApp.StatusBar.SetText("É necessário informar ao menos um dos filtros [Parceiro de Negócio], [Pedido], [Processo]!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }


            var dataTableId = "dtDados";
            var columns = matrixData.Columns;

            string queryPesquisar = QueryPesquisar(pn, pedido, processo);
            this.dtDados.ExecuteQuery(queryPesquisar);

            BindData(columns, dataTableId);
            matrixData.LoadFromDataSource();
            matrixData.AutoResizeColumns();

            SAPApp.Forms.ActiveForm.Freeze(false);
        }

    

        private Column AddMatrixColumn(Columns columns, BoFormItemTypes formItemTypes, string columnName, string columnTitleCaption = null, bool editable = true)
        {
            Column column = columns.Add(columnName, formItemTypes);

            if (String.IsNullOrWhiteSpace(columnTitleCaption))
                columnName = columnTitleCaption;

            column.TitleObject.Caption = columnTitleCaption;

            if (!editable)
                column.Editable = editable;

            return column;
        }

        private Column AddMatrixLinkedButtonColumn(Columns columns, BoLinkedObject linkedObject, string columnName, string columnTitleCaption = null, bool editable = true)
        {
            var column = AddMatrixColumn(columns, BoFormItemTypes.it_LINKED_BUTTON, columnName, columnTitleCaption, editable);

            SAPbouiCOM.LinkedButton linkedButton = (SAPbouiCOM.LinkedButton)column.ExtendedObject;
            linkedButton.LinkedObject = linkedObject;

            return column;
        }

        private void BindMatrixColumn(Column column, string dataTableId, string dataSourceColumnName)
        {
            column.DataBind.Bind(dataTableId, dataSourceColumnName);
        }

        private bool ChangeFilter()
        {
            var columns = this.matrixData.Columns;
            bool hasItemSelected = false;
            bool continuar = true;

            for (int i = 0; i < this.dtDados.Rows.Count; i++)
            {
                hasItemSelected = this.dtDados.GetValue("Selected", i).ToString().Equals("Y");

                if (hasItemSelected)
                    break;
            }

            if (hasItemSelected)
                continuar = SAPApp.MessageBox("Itens já foram selecionados, ao carregar novamente os dados não são salvos serão perdidos. Deseja continuar?", 1, "Continuar", "Cancelar", "") == 1;

            if (!continuar)
                SAPApp.StatusBar.SetText("Atualização de dados cancelada!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            else
                FindData();

            return continuar;
        }

        */

        private void Salvar()
        {
            SAPApp.StatusBar.SetText("Salvando, aguarde alguns instantes...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

            var columns = oMatrix.Columns;
            var docEntry = 0;
            bool selected = false;

            GeneralService oGeneralService = ConexaoSAP.Company.GetCompanyService().GetGeneralService("OSLT_IMPORT");
            GeneralData oGeneralData = (GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

            if (docEntry > 0)
            {
                GeneralDataParams headerParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                headerParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(headerParams);
            }

            GeneralDataCollection generalServiceItem = oGeneralData.Child("SLTIMPRT1");

            oGeneralData.SetProperty("U_TaxId", "0");
            oGeneralData.SetProperty("U_FreteInt", "0");
            oGeneralData.SetProperty("U_OutDesp", "0");
            oGeneralData.SetProperty("U_Container", "0");
            oGeneralData.SetProperty("U_TotalPes", "0");
            oGeneralData.SetProperty("U_TotalFOB", "0");
            oGeneralData.SetProperty("U_Total", "0");
            oGeneralData.SetProperty("U_TaxLine", "0");

            //for (int i = 0; i < this.dtDados.Rows.Count; i++)
            for (int i = 0; i < 0; i++)
            {
                // selected = this.dtDados.GetValue("Selected", i).ToString().Equals("Y");

                if (selected)
                {
                    var item = generalServiceItem.Add();
                    item.SetProperty("U_PedidoId", "0");
                    item.SetProperty("U_ItemNum", "0");
                    item.SetProperty("U_ItemCode", "0");
                    // item.SetProperty("U_Descript", "0"); Dscription
                    //item.SetProperty("U_PrecoUni", "0");
                    //item.SetProperty("U_PrecoTot", "0");
                    item.SetProperty("U_QtdPed", "0");
                    item.SetProperty("U_QtdDisp", "0");
                    item.SetProperty("U_QtdFat", "0");
                    //item.SetProperty("U_Unit", "0");
                    item.SetProperty("U_Peso", "0");
                    item.SetProperty("U_Frete", "0");
                    item.SetProperty("U_OutroDes", "0");
                    item.SetProperty("U_Deposito", "0");
                    item.SetProperty("U_PedTax", "0");
                }
            }


            if (docEntry == 0)
            {
                //oGeneralData.SetProperty("Period", "N");
                //oGeneralData.SetProperty("Instance", "0");
                //oGeneralData.SetProperty("Handwrtten", "N");
                //oGeneralData.SetProperty("Canceled", "N");
                //oGeneralData.SetProperty("Object", "O");
                //oGeneralData.SetProperty("LogInst", "0");
                //oGeneralData.SetProperty("UserSign", "0");
                //oGeneralData.SetProperty("Transfered", "N");
                //oGeneralData.SetProperty("Status", "O");
                //oGeneralData.SetProperty("CreateDate", txtDataDocumento.Value);
                //oGeneralData.SetProperty("CreateTime", DateTime.Now.ToShortTimeString());
                oGeneralService.Add(oGeneralData);
            }
            else
            {
                //oGeneralData.SetProperty("UpdateDate", txtDataDocumento.Value);
                //oGeneralData.SetProperty("UpdateTime", DateTime.Now.ToShortTimeString());
                oGeneralService.Update(oGeneralData);
            }

            SAPApp.StatusBar.SetText("Salvo com sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void LoadData()
        {
            var columns = oMatrix.Columns;
            int docEntry = 0;

            // if (int.TryParse(txtDocEntry.Value, out docEntry))
            if (int.TryParse("", out docEntry))
            {
                SAPApp.StatusBar.SetText("Carregando, aguarde alguns instantes...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                GeneralService oGeneralService = ConexaoSAP.Company.GetCompanyService().GetGeneralService("OSLT_IMPORT");
                GeneralData oGeneralData = (GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                if (docEntry > 0)
                {
                    GeneralDataParams headerParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    headerParams.SetProperty("DocEntry", docEntry);
                    oGeneralData = oGeneralService.GetByParams(headerParams);
                }

                this.UIAPIRawForm.Mode = BoFormMode.fm_OK_MODE;

                GeneralDataCollection generalServiceItem = oGeneralData.Child("SLTIMPRT1");
                oMatrix.Clear();

                //this.dtDados.ExecuteQuery(this.QueryCarregar(docEntry));

                var dataTableId = "dtDados";
                //BindData(columns, dataTableId);
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();







                //oGeneralData.SetProperty("U_TaxId", "0");
                //oGeneralData.SetProperty("U_FreteInt", "0");
                //oGeneralData.SetProperty("U_OutDesp", "0");
                //oGeneralData.SetProperty("U_Container", "0");
                //oGeneralData.SetProperty("U_TotalPes", "0");
                //oGeneralData.SetProperty("U_TotalFOB", "0");
                //oGeneralData.SetProperty("U_Total", "0");
                //oGeneralData.SetProperty("U_TaxLine", "0");

                //for (int i = 0; i < this.dtDados.Rows.Count; i++)
                //{
                //    selected = this.dtDados.GetValue("Selected", i).ToString().Equals("Y");

                //    if (selected)
                //    {
                //        var item = generalServiceItem.Add();
                //        item.SetProperty("U_PedidoId", "0");
                //        item.SetProperty("U_ItemNum", "0");
                //        item.SetProperty("U_ItemCode", "0");
                //        // item.SetProperty("U_Descript", "0"); Dscription
                //        //item.SetProperty("U_PrecoUni", "0");
                //        //item.SetProperty("U_PrecoTot", "0");
                //        item.SetProperty("U_QtdPed", "0");
                //        item.SetProperty("U_QtdDisp", "0");
                //        item.SetProperty("U_QtdFat", "0");
                //        //item.SetProperty("U_Unit", "0");
                //        item.SetProperty("U_Peso", "0");
                //        item.SetProperty("U_Frete", "0");
                //        item.SetProperty("U_OutroDes", "0");
                //        item.SetProperty("U_Deposito", "0");
                //        item.SetProperty("U_PedTax", "0");
                //    }
                //}

                SAPApp.StatusBar.SetText("Carregado com sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
        }

        private void Cancelar()
        {

        }

        #endregion Methods








        #region REMOVER

        /*
         * 
        private void btnPesquisar_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPApp.StatusBar.SetText("Pesquisa em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            try
            {
                SAPApp.Forms.ActiveForm.Freeze(true);

                string vCodPn = txtCodePN.Value;
                string vCodPedido = txtCodigoPedido.Value;
                string vCodProcesso = txtProcesso.Value;
                string queryPesquisar = QueryPesquisar(vCodPn, vCodPedido, vCodProcesso);

                SAPApp.StatusBar.SetText("Executando a consulta...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                this.dtDados.Rows.Clear();

                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRSet.DoQuery(queryPesquisar);

                if (oRSet.RecordCount == 0)
                {
                    SAPApp.StatusBar.SetText("Nenhum registro encontrado!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                SAPApp.StatusBar.SetText("Consulta concluida, processando os dados para a exibição...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);


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
                SAPApp.Forms.ActiveForm.Freeze(false);
            }

        }

        private void btnCarregar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPApp.StatusBar.SetText("Pesquisa de importacao em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            try
            {
                //SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SAPApp.Company.GetDICompany();
                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPApp.Forms.ActiveForm.Freeze(true);

                String vCodPN = txtCodePN.Value;
                String vCodPedido = txtCodigoPedido.Value;
                String vCodProcesso = txtProcesso.Value;

                if (!String.IsNullOrWhiteSpace(vCodPN))
                {
                    SAPApp.MessageBox("Realizar consulta pelo Código do PN através do botão pesquisa");
                    return;
                }

                if (!String.IsNullOrWhiteSpace(vCodPedido))
                {
                    SAPApp.MessageBox("Realizar consulta pelo Pedido através do botão pesquisa");
                    return;
                }

                if (!String.IsNullOrWhiteSpace(vCodProcesso))
                {
                    SAPApp.MessageBox("Realizar consulta pelo Nº do proceso através do botão pesquisa");
                    return;
                }

                String vCodImp = txtCodigoImportacao.Value;

                if (vCodImp == "")
                {
                    SAPApp.MessageBox("Informar número do documento");
                    return;
                }


                string Query = QueryCarregar(vCodImp);

                oRSet.DoQuery(Query);

                this.dtDados.Rows.Clear();

                if (oRSet.RecordCount == 0)
                {
                    SAPApp.StatusBar.SetText("Nenhum registro encontrado!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
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
                SAPApp.Forms.ActiveForm.Freeze(false);
            }
        }

        private void btnSalvar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPApp.StatusBar.SetText("Salvar em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            double vTxtTxID;
            if (!String.IsNullOrWhiteSpace(txtTaxaDI.Value))
            {
                vTxtTxID = Convert.ToDouble(txtTaxaDI.Value.Replace(".", ","));

            }
            else
            {
                vTxtTxID = 0;
                SAPApp.MessageBox("Necessário informar taxa DI.");
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
                SAPApp.MessageBox("Necessário informar frete internacional.");
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

            //SAPApp.MessageBox("2");
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

                //SAPApp.MessageBox("Esse documento já foi salvo.");
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

            SAPApp.MessageBox("Importação foi salva com sucesso.");

        }

        private void btnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            int intRetorno;
            intRetorno = SAPApp.MessageBox("Informações não gravadas serão perdidas. Continuar?", 1, "Sim", "Não", "");

            if (intRetorno == 1)
            {
                BubbleEvent = true;
                SAPbouiCOM.Form oForm = SAPApp.Forms.ActiveForm;
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
                SAPApp.Forms.ActiveForm.Freeze(true);

                double vTxtTxID;
                if (!String.IsNullOrWhiteSpace(txtTaxaDI.Value))
                {
                    vTxtTxID = Convert.ToDouble(txtTaxaDI.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vTxtTxID = 0.00;
                    SAPApp.MessageBox("Necessário informar taxa DI.");
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
                    SAPApp.MessageBox("Necessário informar frete internacional.");
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

                
                //for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
                //{

                //    if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                //    {
                //        Double vTotalFrete = ((vFrtInt / TotalPeso) * Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", y).ToString()));
                //        this.gridDados.DataTable.SetValue("Frete", y, vTotalFrete.ToString("N4"));
                //        Double vCalOtrDesp = (vOtrDesp / contador);
                //        this.gridDados.DataTable.SetValue("OtrDesp", y, (vCalOtrDesp.ToString("N4")));

                //        txtPeso.Value = TotalPeso.ToString("N3");
                //        txtTotalFOB.Value = TotalME.ToString("N2");

                //    }
                //}
                

                for (int z = 0; z <= this.dtDados.Rows.Count - 1; z++)
                {

                    if (this.gridDados.DataTable.GetValue(0, z).ToString() == "Y")
                    {
                        TotalPeso = TotalPeso + Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", z).ToString());
                        if (TotalPeso == 0)
                        {
                            TotalPeso = 1;
                        }
                    }
                }

                for (int x = 0; x <= this.dtDados.Rows.Count - 1; x++)
                {

                    if (this.gridDados.DataTable.GetValue(0, x).ToString() == "Y")
                    {
                        Double vPesoLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", x).ToString());
                        Double vTotalFrete = ((vFrtInt / TotalPeso) * vPesoLinha);
                        this.gridDados.DataTable.SetValue("Frete", x, vTotalFrete.ToString("N4"));

                        TotalME = TotalME + Convert.ToDouble(this.gridDados.DataTable.GetValue("LineTotal", x).ToString());
                        contador = contador + 1;
                        Double vCalOtrDesp = (vOtrDesp / contador);
                        this.gridDados.DataTable.SetValue("OtrDesp", x, (vCalOtrDesp.ToString("N4")));

                        Double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", x).ToString());
                        Double vPreco = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", x).ToString());
                        Double vFreteLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", x).ToString());
                        Double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", x).ToString());

                        TotalII = TotalII + (((vTxtTxID * (vPreco * vQuantidade)) + (vFreteLinha)) * (vTaxRate / 100));

                    }

                }

                txtPeso.Value = TotalPeso.ToString("N3");
                txtTotal_II.Value = TotalII.ToString("N2");
                txtTotalFOB.Value = TotalME.ToString("N2");

                Total = (((vTxtTxID * TotalME) + vFrtInt) + TotalII);
                txtTotalRS.Value = Total.ToString("N2");

            }
            finally
            {
                SAPApp.Forms.ActiveForm.Freeze(false);
            }

        }

        private void btnGerarNF_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPApp.StatusBar.SetText("Geração em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

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
                SAPApp.MessageBox("Necessário informar taxa DI.");
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
                    SAPApp.MessageBox("Necessário informar frete internacional");
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

                    for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
                    {
                        if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                        {

                            int vPedido = Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString());
                            int vPedNumDoc = Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", y).ToString());
                            int vLinPed = Int32.Parse(this.gridDados.DataTable.GetValue("VisOrder", y).ToString()) - 1;
                            string vProduto = this.gridDados.DataTable.GetValue("ItemCode", y).ToString();
                            //string vDescricao = this.gridDados.DataTable.GetValue(6, y).ToString();
                            string vDescricao = "";
                            double vPrecoUnit = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", y).ToString());
                            double vPrecoTotal = Convert.ToDouble(this.gridDados.DataTable.GetValue("LineTotal", y).ToString());
                            double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", y).ToString());
                            double vQuantidadeAberta = Convert.ToDouble(this.gridDados.DataTable.GetValue("OpenQty", y).ToString());
                            string vUm = this.gridDados.DataTable.GetValue("unitMsr", y).ToString();
                            double vPeso = Convert.ToDouble(this.gridDados.DataTable.GetValue("Peso", y).ToString());
                            double vFrete = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", y).ToString());
                            double vOutrasDespesas = Convert.ToDouble(this.gridDados.DataTable.GetValue("OtrDesp", y).ToString());
                            int vItens = Int32.Parse(this.gridDados.DataTable.GetValue("NumPerMsr", y).ToString());
                            int vLineNum = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", y).ToString());
                            double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", y).ToString());
                            string vDeposito = "01";

                            if ((!String.IsNullOrWhiteSpace(vRetornoNrImp)))
                            {
                                // Inserir Linha de Importacao
                                InserirLinhaImportacao(Int32.Parse(vRetornoNrImp), vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vLineNum, vTaxRate);
                            }
                            else
                            {
                                // Inserir Linha de Importacao
                                InserirLinhaImportacao(0, vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vLineNum, vTaxRate);
                            }

                        }
                    }
                }

                // Fim Processo de Salvar Documento
            }

            // Geração da Nota Fiscal
            var vEsbocoNFRecebimento = (SAPbobsCOM.Documents)ConexaoSAP.Company.GetBusinessObject(BoObjectTypes.oDrafts);
            vEsbocoNFRecebimento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;

            for (int y = 0; y <= this.dtDados.Rows.Count - 1; y++)
            {
                if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                {
                    vEsbocoNFRecebimento.CardCode = RetornaFornecedor(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString()));
                    vEsbocoNFRecebimento.DocDate = DateTime.Today;
                    vEsbocoNFRecebimento.DocDueDate = DateTime.Today;
                    vEsbocoNFRecebimento.TaxDate = DateTime.Today;
                    vEsbocoNFRecebimento.DocCurrency = "R$";

                    vEsbocoNFRecebimento.Comments = "Recebimento gerado por integração no dia: " + DateTime.Now.ToString();
                    vEsbocoNFRecebimento.BPL_IDAssignedToInvoice = 1;
                    //vEsbocoNFRecebimento.UserFields.Fields.Item("U_MW_PROJETO").Value = 99001000;

                    // Condição de PAgamento no SAP
                    vEsbocoNFRecebimento.GroupNumber = 100;

                    string utilizacao = RetornaUtizacao(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString()));

                    for (Int32 i = 0; i <= this.dtDados.Rows.Count - 1; i++)
                    {
                        if (this.gridDados.DataTable.GetValue(0, i).ToString() == "Y")
                        {
                            vEsbocoNFRecebimento.Lines.ItemCode = this.gridDados.DataTable.GetValue("ItemCode", i).ToString();
                            vEsbocoNFRecebimento.Lines.Quantity = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", i).ToString());

                            // Converter de Dolar para Real usando TX ID - Multiplicando. Somar o peso pelas linhas selecionada e depois dividir frete pelo total e multiplicar pelo peso da linha   
                            Double vQuantidade = Convert.ToDouble(this.gridDados.DataTable.GetValue("Quantity", i).ToString());
                            Double vPreco = Convert.ToDouble(this.gridDados.DataTable.GetValue("Price", i).ToString());
                            Double vFreteLinha = Convert.ToDouble(this.gridDados.DataTable.GetValue("Frete", i).ToString());
                            Double vTaxRate = Convert.ToDouble(this.gridDados.DataTable.GetValue("TaxRate", i).ToString());

                            vEsbocoNFRecebimento.Lines.UnitPrice = (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) + (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) * (vTaxRate / 100)));

                            vEsbocoNFRecebimento.Lines.ShipDate = DateTime.Today;
                            vEsbocoNFRecebimento.Lines.TaxCode = RetornaCodImposto(int.Parse(this.gridDados.DataTable.GetValue("DocEntry", i).ToString()), Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", i).ToString()));

                            //Utilização
                            vEsbocoNFRecebimento.Lines.Usage = utilizacao;

                            string vDocEntryPedido = RetornaPedDocEntry(Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", i).ToString()));
                            //string vDocEntryPedido = RetornaPedDocEntry(Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", i).ToString()));

                            //Amarração com Pedido de Compra
                            vEsbocoNFRecebimento.Lines.BaseType = 22;
                            vEsbocoNFRecebimento.Lines.BaseEntry = Int32.Parse(vDocEntryPedido);
                            vEsbocoNFRecebimento.Lines.BaseLine = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", i).ToString());

                            vEsbocoNFRecebimento.Lines.Add();
                        }
                    }

                    int vRetorno = vEsbocoNFRecebimento.Add();

                    if (vRetorno != 0)
                    {
                        string MessagemErro = ConexaoSAP.Company.GetLastErrorDescription();
                        throw new Exception(MessagemErro);
                    }
                    else
                    {
                        SAPApp.MessageBox("Esboço de Recebimento criado com sucesso.");
                        string draftDocEntry = RetornaCodDraft();
                        SAPApp.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
                        return;
                    }
                }
            }
        }

        */

        #endregion REMOVER

    }
}