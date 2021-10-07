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
using System.Text;
using SLT.ImportacaoNF.Utils;

namespace SLT.ImportacaoNF
{
    [FormAttribute("OSLT_IMPORT", "Forms/frmImportacao.b1f")]
    class frmImportacao : UserFormBase
    {
        #region Attributes

        private SAPbouiCOM.Application SAPApp = null;
        private SAPbouiCOM.EditText oEditText = null;
        private SAPbouiCOM.ComboBox oComboBox = null;
        private SAPbouiCOM.Item oItem = null;
        private SAPbouiCOM.Form oForm = null;
        private SAPbouiCOM.Matrix oMatrix = null;
        private SAPbouiCOM.ChooseFromList oCFL = null;
        private const string menu_nf_entrada = "2308";

        #endregion attributes

        public frmImportacao()
        {
            SAPApp = SAPbouiCOM.Framework.Application.SBO_Application;

            oForm = GetForm();
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

            CriarDataSources();
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

        public void CriarDataSources()
        {
            oForm.DataSources.DBDataSources.Add("@SLTIMPRT1");
        }

        public SAPbouiCOM.Matrix CriarFormularioMatrix()
        {
            oForm.DataSources.DataTables.Add("oMatrixDT");
            oForm.DataSources.DataTables.Add("oTempDT");

            oItem = oForm.Items.Add("oMtrxImp", SAPbouiCOM.BoFormItemTypes.it_MATRIX);

            oItem = oForm.Items.Item("oMtrxImp");
            oItem.Top = 70;
            oItem.Left = 15;
            oItem.Width = oForm.Width - 30;
            oItem.Height = 350;

            oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("oMtrxImp").Specific;
            SAPbouiCOM.Column oColumn = null;

            oColumn = oMatrix.Columns.Add("clSelected", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Selected");
            oColumn.TitleObject.Caption = "#";

            oColumn = oMatrix.Columns.Add("clCardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_CardCode");
            oColumn.TitleObject.Caption = "Fornecedor";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_BusinessPartner;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("clPedidoNr", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PedidoNr");
            oColumn.TitleObject.Caption = "Nº do Pedido";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_PurchaseOrder;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("clPedidoId", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PedidoId");
            oColumn.TitleObject.Caption = "Pedido de Compra";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_PurchaseOrder;
            oColumn.Editable = false;
            // oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("clItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_ItemCode");
            oColumn.TitleObject.Caption = "Cód. do item";
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_Items;
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("clDescript", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Descript");
            oColumn.TitleObject.Caption = "Descrição do Item";
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("clItemNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_ItemNum");
            oColumn.TitleObject.Caption = "Nº do item";
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("clPrecoUni", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PrecoUni");
            oColumn.TitleObject.Caption = "Preço Unitário";
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("clPrecoInt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_PrecoInt");
            oColumn.TitleObject.Caption = "Valor Total US$";
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("clQtdPed", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdPed");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Qtd. Pedido";

            oColumn = oMatrix.Columns.Add("clQtdDisp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdDisp");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Qtd Disponível";

            oColumn = oMatrix.Columns.Add("clQtdFat", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_QtdFat");
            oColumn.TitleObject.Caption = "Qtd. Faturada";

            oColumn = oMatrix.Columns.Add("clUnit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Unit");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Un. Méd.";

            oColumn = oMatrix.Columns.Add("clPeso", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Peso");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Peso";

            oColumn = oMatrix.Columns.Add("clNCM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_NcmCode");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "NCM";

            oColumn = oMatrix.Columns.Add("clFrete", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Frete");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Frete";

            oColumn = oMatrix.Columns.Add("clOutroDes", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_OutroDes");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Outras Desp.";

            oColumn = oMatrix.Columns.Add("clDeposito", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Deposito");
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.lf_Warehouses;
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Depósito";

            oColumn = oMatrix.Columns.Add("clPedTaxN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Imposto");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Cód. Imposto";
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("clPedtax", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Aliquota");
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Alíquota";            

            oColumn = oMatrix.Columns.Add("clUtiliza", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.DataBind.SetBound(true, "@SLTIMPRT1", "U_Utiliza");
            //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = BoLinkedObject.usage;
            oColumn.Editable = false;
            oColumn.TitleObject.Caption = "Utilização";
            oColumn.Visible = false;

            //oMatrix.Columns.Item("oClmn0").DataBind.UnBind();
            oForm.DataSources.DataTables.Item("oMatrixDT").Clear();

            string sSQL = " SELECT                      " +
                //"     DocNum                  " +
                          "    [U_Selected]            " +
                          "    ,[U_CardCode]            " +
                          "    ,[U_PedidoId]            " +
                          "    ,[U_PedidoNr]            " +
                          "    ,[U_ItemNum]             " +
                          "    ,[U_ItemCode]            " +
                          "    ,[U_Descript]            " +
                          "    ,[U_PrecoUni]            " +
                          "    ,[U_PrecoInt]            " +
                          "    ,[U_QtdPed]              " +
                          "    ,[U_QtdDisp]             " +
                          "    ,[U_QtdFat]              " +
                          "    ,[U_Unit]                " +
                          "    ,[U_Peso]                " +
                          "    ,[U_NcmCode]             " +
                          "    ,[U_Frete]               " +
                          "    ,[U_OutroDes]            " +
                          "    ,[U_Deposito]            " +
                          "    ,[U_Imposto]              " +
                          "    ,[U_Aliquota]              " +
                          "    ,[U_Utiliza]              " +                          
                          " FROM [dbo].[@SLTIMPRT1] " +
                          " WHERE DocEntry = -1 ";// + oEditText.Value;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(sSQL);

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("oMatrixDT");
            BindMatrixData(oDataTable);

            oMatrix.ChooseFromListBefore += oMatrix_ChooseFromListBefore;
            oMatrix.ClickAfter += matrixData_ClickAfter;

            return oMatrix;
        }

        void BindMatrixData(SAPbouiCOM.DataTable oDataTable)
        {
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1); // SLTIMPRT1;
            oDBDataSource.Clear();

            for (int row = 0; row < oDataTable.Rows.Count; row++)
            {
                int offset = oDBDataSource.Size;
                oDBDataSource.InsertRecord(row);

                oDBDataSource.SetValue("U_Selected", offset, oDataTable.GetValue("U_Selected", row).ToString());
                oDBDataSource.SetValue("U_CardCode", offset, oDataTable.GetValue("U_CardCode", row).ToString());
                oDBDataSource.SetValue("U_PedidoId", offset, oDataTable.GetValue("U_PedidoId", row).ToString());
                oDBDataSource.SetValue("U_PedidoNr", offset, oDataTable.GetValue("U_PedidoNr", row).ToString());
                oDBDataSource.SetValue("U_ItemCode", offset, oDataTable.GetValue("U_ItemCode", row).ToString());
                oDBDataSource.SetValue("U_Descript", offset, oDataTable.GetValue("U_Descript", row).ToString());
                oDBDataSource.SetValue("U_ItemNum", offset, oDataTable.GetValue("U_ItemNum", row).ToString());
                oDBDataSource.SetValue("U_PrecoUni", offset, oDataTable.GetValue("U_PrecoUni", row).ToString());
                oDBDataSource.SetValue("U_PrecoInt", offset, oDataTable.GetValue("U_PrecoInt", row).ToString());
                oDBDataSource.SetValue("U_QtdPed", offset, oDataTable.GetValue("U_QtdPed", row).ToString());
                oDBDataSource.SetValue("U_QtdDisp", offset, oDataTable.GetValue("U_QtdDisp", row).ToString());
                oDBDataSource.SetValue("U_QtdFat", offset, oDataTable.GetValue("U_QtdFat", row).ToString());
                oDBDataSource.SetValue("U_Unit", offset, oDataTable.GetValue("U_Unit", row).ToString());
                oDBDataSource.SetValue("U_Peso", offset, oDataTable.GetValue("U_Peso", row).ToString());
                oDBDataSource.SetValue("U_NcmCode", offset, oDataTable.GetValue("U_NcmCode", row).ToString());
                oDBDataSource.SetValue("U_Frete", offset, oDataTable.GetValue("U_Frete", row).ToString());
                oDBDataSource.SetValue("U_OutroDes", offset, oDataTable.GetValue("U_OutroDes", row).ToString());
                oDBDataSource.SetValue("U_Deposito", offset, oDataTable.GetValue("U_Deposito", row).ToString());
                oDBDataSource.SetValue("U_Imposto", offset, oDataTable.GetValue("U_Imposto", row).ToString());
                oDBDataSource.SetValue("U_Aliquota", offset, oDataTable.GetValue("U_Aliquota", row).ToString());
                oDBDataSource.SetValue("U_Utiliza", offset, oDataTable.GetValue("U_Utiliza", row).ToString());                
            }

            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
        }

        public void OnChooseFromList(ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL;
            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

            if (oCFLEvento.BeforeAction)
            {
                Conditions conditions = oCFL.GetConditions();
                Condition condition = null;
                condition = conditions.Add();
                condition.Alias = "CardCode";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "3600410"; //"cflPedido";
                condition.Relationship = BoConditionRelationship.cr_AND;

                //condition = conditions.Add();
                //condition.Alias = "DocStatus";
                //condition.Operation = BoConditionOperation.co_EQUAL;
                //condition.CondVal = "O";
                //condition.Relationship = BoConditionRelationship.cr_AND;

                oCFL.SetConditions(conditions);
            }


            SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

            if (oDataTable != null && oDataTable.Rows.Count > 0)
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1); // SLTIMPRT1;
                SAPbouiCOM.DataTable oDataTableMatrix = oForm.DataSources.DataTables.Item("oMatrixDT");
                oDBDataSource.Clear();

                SAPbouiCOM.DataTable dtPedido = oForm.DataSources.DataTables.Item("oTempDT");
                dtPedido.ExecuteQuery(this.GetPurchaseOrderItems(oDataTable));

                BindMatrixData(dtPedido);
                Calcular();
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
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SAPApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = true;
            oCFLCreationParams.ObjectType = "22";
            oCFLCreationParams.UniqueID = "CFL_PO";

            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
            oCFL = oCFLs.Add(oCFLCreationParams);



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
        }

        public void CriarFormularioDefault()
        {
            oItem = oForm.Items.Add("lblEntry", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Nº";
            oItem.Top = 5;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;
            oItem.Click();

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
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_TaxaDI");
            //oEditText.Value = "0,00";

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
            oItem.Enabled = false;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_TotalPes");
            //oEditText.Value = "0,000";



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
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_FreteInt");
            //oEditText.Value = "0,00";

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
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_TotalFOB");
            //oEditText.Value = "0,00";


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
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_OutDesp");



            oItem = oForm.Items.Add("lblTotal2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Total II";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtTotal2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_Total2");

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
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_Container");


            oItem = oForm.Items.Add("lblTotal", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            ((StaticText)oItem.Specific).Caption = "Total R$";
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 245;
            oItem.Width = 50;

            oItem = oForm.Items.Add("txtTotal", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = top_position_base;
            oItem.Left = oForm.Width - 180;
            oItem.Width = 150;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@SLTIMPORT", "U_Total");
            //oEditText.Value = "0,00";
        }



        #region Events

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
                    //oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                    break;
                case BoEventTypes.et_FORM_UNLOAD:
                    break;
                case BoEventTypes.et_FORM_VISIBLE:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    OnGotFocus(pVal);
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
                    OnLostFocus(pVal);
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

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            oMatrix.AutoResizeColumns();
        }

        private void matrixData_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID.Equals("clSelected"))
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);
                CheckBox chkSelected = (CheckBox)oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row);
                oDBDataSource.SetValue("U_Selected", pVal.Row - 1, chkSelected.Checked ? "Y" : "N");
                Calcular();
            }
        }

        private void Calcular()
        {
            SAPApp.StatusBar.SetText("Atualizando os valores, aguarde alguns instantes!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            SAPApp.Forms.ActiveForm.Freeze(true);

            var columns = oMatrix.Columns;
            double valorTotal = 0;
            double pesoTotal = 0;
            double qtd = 0;

            double valorUnitario = 0;
            double pesoUnitario = 0;

            bool selected = false;
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);

            for (int i = 0; i < oDBDataSource.Size; i++)
            {
                selected = oDBDataSource.GetValue("U_Selected", i).ToString().Trim().Equals("Y");

                if (selected)
                {
                    qtd = ParseGlobalization(oDBDataSource.GetValue("U_QtdFat", i));
                    valorUnitario = ParseGlobalization(oDBDataSource.GetValue("U_PrecoUni", i));
                    pesoUnitario = ParseGlobalization(oDBDataSource.GetValue("U_Peso", i));

                    valorTotal += valorUnitario * qtd;
                    pesoTotal += pesoUnitario * qtd;
                }
            }

            double taxaDI = GetEditTextValue("txtTaxaDI");
            double freteInternacional = GetEditTextValue("txtFrete");
            double fob = GetEditTextValue("txtFOB");
            double outraDesp = GetEditTextValue("txtOutDesp");
            double totalOutros = taxaDI + freteInternacional + fob + outraDesp;

            GetEditText("txtPesoTot").Value = pesoTotal.ToString("N3");
            GetEditText("txtTotal2").Value = valorTotal.ToString("N2");
            GetEditText("txtTotal").Value = (valorTotal + totalOutros).ToString("N2");


            SAPApp.StatusBar.SetText("Valores atualizados!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            SAPApp.Forms.ActiveForm.Freeze(false);
        }

        private EditText GetEditText(string controlName)
        {
            return oForm.Items.Item(controlName).Specific as EditText;
        }

        private double GetEditTextValue(string controlName)
        {
            return ParseGlobalization(GetEditText(controlName).Value);
        }

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            //this.GetItem("lblCdPN").Click();

            if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                //EnableField("txtCodPN");
                //EnableField("txtPedido");
                //EnableField("txtProces");
                //EnableField("txtEntry");
                //txtDataDocumento.Value = string.Empty;

                //cbStatus.Select(1, BoSearchKey.psk_Index);
                //EnableField("cbStatus", true);
            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                //EnableField("txtCodPN");
                //EnableField("txtPedido");
                //EnableField("txtProces");
                //EnableField("txtEntry", false);
                //txtDocEntry.Value = ConexaoSAP.Company.GetNewObjectKey();
                //txtDataDocumento.Value = DateTime.Today.ToShortDateString();

                //cbStatus.Select(1, BoSearchKey.psk_Index);
                //EnableField("cbStatus");

            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE)
            {
                //EnableField("txtCodPN", false);
                //EnableField("txtPedido", false);
                //EnableField("txtProces", false);
                //EnableField("txtEntry", false);
                //txtDocEntry.Value = string.Empty;
                //txtDataDocumento.Value = string.Empty;
                //cbStatus.Select(0, BoSearchKey.psk_Index);
                //EnableField("cbStatus");
            }
            else if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                //EnableField("txtCodPN", false);
                //EnableField("txtPedido", false);
                //EnableField("txtProces", false);
                //EnableField("txtEntry", false);
                //txtDataDocumento.Value = string.Empty;
                //EnableField("cbStatus");
            }
        }

        private void btnSalvar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var formMode = (SAPbouiCOM.BoFormMode)pVal.FormMode;

            if (formMode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || formMode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || formMode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE)
            {
                // Salvar();
            }
            else if (formMode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                // FindData();
            }
        }

        private void btnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        void btn_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPApp.StatusBar.SetText("Clicou no NF", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            // SAPApp.Menus.Item(menu_nf_entrada).Activate();

            // Form formNF = SAPApp.Forms.ActiveForm;

            CriarEsbocoNF();
        }

        private void CriarEsbocoNF()
        {
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);
            Dictionary<string, List<int>> dicFornecedoresItems = ObteDadosParaEsboco(oDBDataSource);
            
            if (dicFornecedoresItems.Count == 0)
            {
                SAPApp.StatusBar.SetText("Nenhum item selecionado para o Esboço da Nota Fiscal.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                return;
            }

            GerarEscoboDocument(oDBDataSource, dicFornecedoresItems);
        }

        private void GerarEscoboDocument(SAPbouiCOM.DBDataSource oDBDataSource, Dictionary<string, List<int>> dicFornecedoresItems)
        {
            try
            {
                string cardCode = string.Empty;
                int sucesso = 0;

                foreach (KeyValuePair<string, List<int>> fornecedor in dicFornecedoresItems)
                {
                    SAPbobsCOM.Documents draft = (SAPbobsCOM.Documents)ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    draft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                    //draft.DocType = BoDocumentTypes.dDocument_Items;
                    draft.HandWritten = BoYesNoEnum.tNO;
                    draft.CardCode = fornecedor.Key;
                    draft.DocDate = DateTime.Today;
                    draft.DocDueDate = DateTime.Today;
                    draft.TaxDate = DateTime.Today;
                    draft.DocCurrency = "R$";

                    draft.Comments = "Esboço de NF de Entrada criado através do Addon de Importação em " + DateTime.Now.ToString();
                    draft.BPL_IDAssignedToInvoice = 1;
                    draft.GroupNumber = 100; // condição de pagamento no SAP

                    foreach (var i in fornecedor.Value)
                    {
                        // Converter de Dolar para Real usando TX ID - Multiplicando. Somar o peso pelas linhas selecionada e depois dividir frete pelo total e multiplicar pelo peso da linha   
                        var freteLinha = ParseGlobalization(oDBDataSource.GetValue("U_Frete", i));
                        var imposto = oDBDataSource.GetValue("U_Imposto", i).Trim();
                        var aliquota = ParseGlobalization(oDBDataSource.GetValue("U_Aliquota", i));
                        var precoUnitario = ParseGlobalization(oDBDataSource.GetValue("U_PrecoUni", i));
                        var qtd = ParseGlobalization(oDBDataSource.GetValue("U_QtdFat", i));
                        var utilizacao = oDBDataSource.GetValue("U_Utiliza", i).Trim();
                        var unidadeMedida = Convert.ToInt32(oDBDataSource.GetValue("U_Unit", i).Trim());
                        var ncmCode = Convert.ToInt32(oDBDataSource.GetValue("U_NcmCode", i).Trim());

                        draft.Lines.ItemCode = oDBDataSource.GetValue("U_ItemCode", i).Trim();
                        draft.Lines.Quantity = qtd;
                        draft.Lines.UnitPrice = (((aliquota * precoUnitario) + (freteLinha / qtd)) + (((aliquota * precoUnitario) + (freteLinha / qtd)) * (aliquota / 100)));
                        draft.Lines.ShipDate = DateTime.Today;
                        draft.Lines.TaxCode = imposto;
                        draft.Lines.Usage = utilizacao;
                        //draft.Lines.ItemDescription = oDBDataSource.GetValue("U_Descript", i).Trim();
                        //draft.UnitPrice = (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) + (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) * (vTaxRate / 100)));

                        draft.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                        draft.Lines.BaseEntry = Convert.ToInt32(oDBDataSource.GetValue("U_PedidoId", i).Trim());
                        draft.Lines.BaseLine = Convert.ToInt32(oDBDataSource.GetValue("U_ItemNum", i).Trim());
                        draft.Lines.Add();
                    }

                    sucesso = draft.Add();
                    if (sucesso != 0)
                    {
                        Mensagem.ExibirErro("Falha ao gerar o esboço da Nota Fiscal");
                        return;
                    }
                    
                    var docNum = ConexaoSAP.Company.GetNewObjectKey();
                    SAPApp.MessageBox("O(s) esboço(s) foram criado(s) com sucesso!");

                    var draftForm = (BoFormObjectEnum)112;
                    SAPApp.OpenForm(draftForm, "", docNum);
                }                
            }
            catch (Exception ex)
            {
                Mensagem.ExibirErro(ex.Message);
                throw ex;
            }
        }

        private Dictionary<string, List<int>> ObteDadosParaEsboco(SAPbouiCOM.DBDataSource oDBDataSource)
        {
            string cardCode = string.Empty;
            Dictionary<string, List<int>> dicFornecedoresItems = new Dictionary<string, List<int>>();
            bool selecionado = false;

            for (int i = 0; i < oDBDataSource.Size; i++)
            {
                cardCode = oDBDataSource.GetValue("U_CardCode", i).Trim();
                selecionado = oDBDataSource.GetValue("U_Selected", i).Trim().Equals("Y");

                if (!selecionado)
                    continue;

                if (dicFornecedoresItems.ContainsKey(cardCode))
                    dicFornecedoresItems[cardCode].Add(i);
                else
                    dicFornecedoresItems.Add(cardCode, new List<int>() { i });
            }

            return dicFornecedoresItems;
        }

        void oMatrix_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            SAPApp.StatusBar.SetText("Não é permitido selecionar pedidos de fornecedores diferentes!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);

            BubbleEvent = false;
        }

        void oColumn_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPApp.StatusBar.SetText("Pedidos selecionados...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
        }

        private void OnLostFocus(ItemEvent pVal)
        {
            if (pVal.ItemUID == "oMtrxImp")
            {
                if (pVal.ColUID == "clQtdFat")
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(1);
                    EditText txtQtdFaturamento = (EditText)oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row);
                    double qtdFaturamento = ParseGlobalization(txtQtdFaturamento.Value);
                    double qtdDisponivel = ParseGlobalization(((EditText)oMatrix.GetCellSpecific("clQtdDisp", pVal.Row)).Value);

                    if (qtdFaturamento <= qtdDisponivel)
                        oDBDataSource.SetValue("U_QtdFat", pVal.Row - 1, qtdFaturamento.ToString());
                    else
                    {
                        SAPApp.StatusBar.SetText("A quantidade disponível é menor que a faturar. Valor corrigido.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                        oDBDataSource.SetValue("U_QtdFat", pVal.Row - 1, qtdDisponivel.ToString());
                        txtQtdFaturamento.Value = qtdDisponivel.ToString();
                        return;
                    }

                    Calcular();
                }
            }
            else
            {
                if (oForm.Mode == BoFormMode.fm_ADD_MODE ||
                    oForm.Mode == BoFormMode.fm_UPDATE_MODE ||
                    oForm.Mode == BoFormMode.fm_EDIT_MODE)
                {
                    if (pVal.ItemUID == "txtTaxaDI" || pVal.ItemUID == "txtFrete" ||
                         pVal.ItemUID == "txtOutDesp" || pVal.ItemUID == "txtFOB")
                    {
                        EditText txtValor = oForm.Items.Item(pVal.ItemUID).Specific as EditText;
                        double value = ParseGlobalization(txtValor.Value);

                        if (value >= 0)
                            txtValor.Value = value.ToString("N2");
                        else
                            txtValor.Value = "0,00";

                        Calcular();
                    }
                }
            }
        }

        private void OnGotFocus(ItemEvent pVal)
        {

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

        public string GetPurchaseOrderItems(SAPbouiCOM.DataTable pedidos)
        {
            bool ultimo = false;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT                       ");
            sb.Append("'N'                                     AS U_Selected,     ");
            sb.Append("[OPOR].CardCode                         AS U_CardCode,     ");
            sb.Append("[OPOR].DocEntry                         AS U_PedidoId,     ");
            sb.Append("[OPOR].DocNum                           AS U_PedidoNr,    ");
            sb.Append("[POR1].ItemCode                         AS U_ItemCode,     ");
            sb.Append("[POR1].Dscription                       AS U_Descript,     ");
            sb.Append("[POR1].LineNum                          AS U_ItemNum,      ");
            sb.Append("[POR1].Price                            AS U_PrecoUni,     ");
            sb.Append("[POR1].TotalFrgn                        AS U_PrecoInt,     ");
            sb.Append("[POR1].Quantity                         AS U_QtdPed,       ");
            sb.Append("[POR1].OpenQty                          AS U_QtdDisp,      ");
            sb.Append("[POR1].Quantity                         AS U_QtdFat,       ");
            sb.Append("[POR1].UnitMsr                          AS U_Unit,         ");
            sb.Append("[POR1].Weight1                          AS U_Peso,         ");
            sb.Append("[OITM].NcmCode                          AS U_NcmCode,      ");
            sb.Append("0                                       AS U_Frete,        ");
            sb.Append("0                                       AS U_OutroDes,     ");
            sb.Append("WhsCode                                 AS U_Deposito,     ");
            sb.Append("[POR1].TaxCode                          AS U_Imposto,      ");
            sb.Append("ISNULL(POR4.TaxRate,1)                  AS U_Aliquota,     ");
            sb.Append("[POR12].MainUsage                       AS U_Utiliza       ");
            sb.Append("FROM [POR1]                  ");
            sb.Append("JOIN [OPOR] AS OPOR            ");
            sb.Append("  ON POR1.DocEntry = OPOR.DocEntry ");
            sb.Append("JOIN [OITM] AS OITM            ");
            sb.Append("  ON POR1.ItemCode = OITM.ItemCode ");
            sb.Append("LEFT JOIN [POR4] AS POR4            ");
            sb.Append("  ON POR1.DocEntry = POR4.DocEntry  ");
            sb.Append(" AND POR1.LineNum = POR4.LineNum   ");
            sb.Append(" AND POR4.staType = 23  ");
            sb.Append("LEFT JOIN [POR12] AS POR12            ");
            sb.Append("  ON OPOR.DocEntry = POR12.DocEntry  ");
            sb.Append("WHERE LineStatus = 'O'        ");
            sb.Append(" AND DocStatus = 'O'        ");
            sb.Append(" AND OPOR.DocEntry IN (          ");

            for (int i = 0; i < pedidos.Rows.Count; i++)
            {
                ultimo = i == pedidos.Rows.Count - 1;

                sb.Append(pedidos.GetValue("DocEntry", i).ToString());

                if (ultimo)
                    sb.Append(")");
                else
                    sb.Append(",");
            }

            return sb.ToString();
        }

        #endregion Repository

        #region Utils

        public string ParseGlobalization(double value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        public string ParseGlobalization(object value)
        {
            return value.ToString().Trim().Replace(",", ".");
        }

        public double ParseGlobalization(string value)
        {
            value = value.ToString().Trim().Replace(".", ",");

            if (!String.IsNullOrWhiteSpace(value))
            {
                double _out = 0;

                if (double.TryParse(value, out _out))
                    return Convert.ToDouble(value, System.Globalization.CultureInfo.CurrentCulture);
                else
                    return _out;
            }

            return 0;
        }

        #endregion


        #endregion Methods

    }
}