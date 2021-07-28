using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.SqlServer.Server;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;

namespace ImportacaoNF
{

    #region Identificação do Arquivo
    /********************************************************/
    /* Desenvolvimento Addon de Importação                  */
    /* Empresa: Quintale Serviços de Informática LTDA       */
    /* Data: 01/03/2021                                     */
    /* Responsável: Daniel Quintale de Jesus                */
    /********************************************************/
    #endregion

    [FormAttribute("ImportacaoNF.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.DataTable dtDados;

        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lblProc").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txtProc").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lblPed").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txtPedido").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lblCdPN").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txtTxID").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_13").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("txtFrtInt").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("txtConta").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.gridDados = ((SAPbouiCOM.Grid)(this.GetItem("gridDados").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txtODesp").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("txtCodPN").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button3.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button3_ClickBefore);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button4.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button4_ClickBefore);
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("txtCodImp").Specific));
            string Hoje = DateTime.Now.ToShortDateString();
            this.EditText10.Value = Hoje;
            //   Numero da Importação
            this.EditText13.Value = RetornaNrImportacao();
            this.ComboBox0.Select("Aberto", SAPbouiCOM.BoSearchKey.psk_ByDescription);
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("Item_22").Specific));
            this.Button6.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button6_ClickBefore);
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            this.dtDados = this.UIAPIRawForm.DataSources.DataTables.Item("DT_0");

            this.UIAPIRawForm.EnableMenu("1281", false);
            this.UIAPIRawForm.EnableMenu("1282", false);
            this.UIAPIRawForm.EnableMenu("1288", false);
            this.UIAPIRawForm.EnableMenu("1289", false);
            this.UIAPIRawForm.EnableMenu("1290", false);
            this.UIAPIRawForm.EnableMenu("1291", false);
            this.UIAPIRawForm.EnableMenu("1304", false);

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText2;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.Button Button2;

        private StaticText StaticText7;
        private StaticText StaticText8;
        private StaticText StaticText9;
        private EditText EditText1;
        private EditText EditText5;
        private EditText EditText8;

        private StaticText StaticText4;
        private EditText EditText3;
        private EditText EditText9;
        private Button Button3;
        private Button Button4;

        private StaticText StaticText10;
        private ComboBox ComboBox0;
        private StaticText StaticText11;
        private EditText EditText10;

        private EditText EditText13;
        private StaticText StaticText12;
        private StaticText StaticText14;
        private EditText EditText12;
        private Button Button6;

        private SAPbouiCOM.Grid gridDados;

        #region EVENTO CLICK BOTÃO PESQUISAR
        // EVENTO CLICK BOTÃO PESQUISAR
        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisa em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            
            try
            {
                //SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)Program._Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                string Query = "";
                string vCodPn = EditText9.Value;
                string vCodPedido = EditText2.Value;
                string vCodProcesso = EditText0.Value;
                string vCodImp = EditText12.Value;


                if ((vCodPn != "") && (vCodPedido != "") && (vCodProcesso != ""))
                {
                    Query = "SELECT " +
                                "POR1.LineNum, " +
                                "OPOR.DocEntry, " +
                                "OPOR.DocNum, " +
                                "POR1.ItemCode, " +
                                "POR1.Dscription, " +
                                "POR1.Price, " +
                                "POR1.TotalFrgn as LineTotal, " +
                                "POR1.Quantity, " +
                                "POR1.OpenQty, " +
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
                                " AND OPOR.CardCode = '" + vCodPn + "'" +  
                                " AND OPOR.DocNum = " + vCodPedido + 
                                " AND OPOR.U_ALFT_NumPrc = " + vCodProcesso +
                            " ORDER BY " +
                                " OPOR.DocNum desc, " +
                                " Por1.VisOrder asc ";
                }
                else
                {
                    if ((vCodPn != "") && (vCodPedido != ""))
                    {
                    Query = "SELECT " +
                                "POR1.LineNum, " +
                                "OPOR.DocEntry, " +
                                "OPOR.DocNum, " +
                                "POR1.ItemCode, " +
                                "POR1.Dscription, " +
                                "POR1.Price, " +
                                "POR1.TotalFrgn as LineTotal, " +
                                "POR1.Quantity, " +
                                "POR1.OpenQty, " +
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
                                " AND OPOR.CardCode = '" + vCodPn + "'" +  
                                " AND OPOR.DocNum = " + vCodPedido + 
                            " ORDER BY " +
                                " OPOR.DocNum desc, " +
                                " Por1.VisOrder asc ";
                    } 
                    else
                    {
                        if ((vCodPn != ""))
                        {
                            Query = "SELECT " +
                                        "POR1.LineNum, " +
                                        "OPOR.DocEntry, " +
                                        "OPOR.DocNum, " +
                                        "POR1.ItemCode, " +
                                        "POR1.Dscription, " +
                                        "POR1.Price, " +
                                        "POR1.TotalFrgn as LineTotal, " +
                                        "POR1.Quantity, " +
                                        "POR1.OpenQty, " +
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
                                        " AND OPOR.CardCode = '" + vCodPn + "'" +
                                    " ORDER BY " +
                                        " OPOR.DocNum desc, " +
                                        " Por1.VisOrder asc ";
                        }
                        else
                        {
                            if (vCodPedido != "")
                            {
                                Query = "SELECT " +
                                            "POR1.LineNum, " +
                                            "OPOR.DocEntry, " +
                                            "OPOR.DocNum, " +
                                            "POR1.ItemCode, " +
                                            "POR1.Dscription, " +
                                            "POR1.Price, " +
                                            "POR1.TotalFrgn as LineTotal, " +
                                            "POR1.Quantity, " +
                                            "POR1.OpenQty, " +
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
                                            " AND OPOR.DocNum= "  + vCodPedido +
                                        " ORDER BY " +
                                            " OPOR.DocNum desc, " +
                                            " Por1.VisOrder asc ";
                            }
                            else
                            {
                                if (vCodProcesso != "")
                                {
                                    Query = "SELECT " +
                                                "POR1.LineNum, " +
                                                "OPOR.DocEntry, " +
                                                "OPOR.DocNum, " +
                                                "POR1.ItemCode, " +
                                                "POR1.Dscription, " +
                                                "POR1.Price, " +
                                                "POR1.TotalFrgn as LineTotal, " +
                                                "POR1.Quantity, " +
                                                "POR1.OpenQty, " +
                                                "POR1.unitMsr, " +
                                                "POR1.Weight1 as Peso, " +
                                                "POR1.NumPerMsr, " +
                                                "ONCM.NcmCode, " +
                                                "Por1.VisOrder, " +
                                                "ISNULL(POR4.TaxRate,1)" +
                                            " FROM OPOR " +
                                                " INNER JOIN POR1 ON POR1.DocEntry = OPOR.DocEntry " +
                                                " INNER JOIN OITM ON OITM.ItemCode = POR1.ItemCode " +
                                                " LEFT JOIN POR4 ON POR1.DocEntry = POR4.DocEntry and POR1.LineNum = POR4.LineNum and POR4.staType = 23 " +
                                                " LEFT JOIN ONCM ON ONCM.AbsEntry = OITM.NcmCode " +
                                            " WHERE " +
                                                " OPOR.DocStatus <> 'C'" +
                                                " AND POR1.LineStatus <> 'C'" +
                                                " AND POR1.Currency <> 'R$' " +
                                                " AND OPOR.U_ALFT_NumPrc = " + vCodProcesso +
                                            " ORDER BY " +
                                                " OPOR.DocNum desc, " +
                                                " Por1.VisOrder asc ";
                                }
                                else
                                {
                                    Query = "";
                                }
                            }
                            
                        }
                    }
                }

                oRSet.DoQuery(Query);
                if (oRSet.RecordCount > 0)
                {
                    this.dtDados.Rows.Clear();

/*                    while (!oRSet.EoF)
                    {
                        this.dtDados.Rows.Add();
                        int i = this.dtDados.Rows.Count - 1;

                        this.dtDados.SetValue("DocEntry", i, oRSet.Fields.Item("DocEntry").Value);
                        this.dtDados.SetValue("DocNum", i, oRSet.Fields.Item("DocNum").Value);
                        this.dtDados.SetValue("VisOrder", i, oRSet.Fields.Item("VisOrder").Value);
                        this.dtDados.SetValue("NcmCode", i, oRSet.Fields.Item("NcmCode").Value);
                        this.dtDados.SetValue("ItemCode", i, oRSet.Fields.Item("ItemCode").Value);
                        this.dtDados.SetValue("Dscription", i, "");//RetornaDescItem(oRSet.Fields.Item("ItemCode").Value));
                        //Double vPrice = (oRSet.Fields.Item("Price").Value);
                        this.dtDados.SetValue("Price", i, oRSet.Fields.Item("Price").Value);// vPrice.ToString("N4"));
                        //Double vLineTotal = (oRSet.Fields.Item("LineTotal").Value);
                        this.dtDados.SetValue("LineTotal", i,oRSet.Fields.Item("LineTotal").Value);// vLineTotal.ToString("N2"));
                        this.dtDados.SetValue("Quantity", i, oRSet.Fields.Item("Quantity").Value);
                        this.dtDados.SetValue("OpenQty", i, oRSet.Fields.Item("OpenQty").Value);
                        this.dtDados.SetValue("unitMsr", i, oRSet.Fields.Item("unitMsr").Value);
                        //Double vPeso = (oRSet.Fields.Item("Peso").Value);
                        this.dtDados.SetValue("Peso", i,oRSet.Fields.Item("Peso").Value);// vPeso.ToString("N3"));
                        this.dtDados.SetValue("NumPerMsr", i, oRSet.Fields.Item("NumPerMsr").Value);
                        this.dtDados.SetValue("LineNum", i, oRSet.Fields.Item("LineNum").Value);
                        //Double vTaxRate = (oRSet.Fields.Item("TaxRate").Value);
                        this.dtDados.SetValue("TaxRate", i, oRSet.Fields.Item("TaxRate").Value);// vTaxRate.ToString("N2"));
                       

                        oRSet.MoveNext();
                    }
*/
                    this.gridDados.DataTable = this.dtDados;
                    this.gridDados.Columns.Item("ColCheck").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    this.gridDados.Columns.Item("DocEntry").Editable = false;
                    this.gridDados.Columns.Item("DocNum").Editable = false;
                    this.gridDados.Columns.Item("VisOrder").Editable = false;
                    this.gridDados.Columns.Item("NcmCode").Editable = false;
                    this.gridDados.Columns.Item("ItemCode").Editable = false;
                    this.gridDados.Columns.Item("Dscription").Editable = false;
                    this.gridDados.Columns.Item("Price").Editable = false;
                    this.gridDados.Columns.Item("LineTotal").Editable = false;
                    this.gridDados.Columns.Item("OpenQty").Editable = false;
                    this.gridDados.Columns.Item("unitMsr").Editable = false;
                    this.gridDados.Columns.Item("Peso").Editable = false;
                    this.gridDados.Columns.Item("Frete").Editable = false;
                    this.gridDados.Columns.Item("OtrDesp").Editable = false;
                    this.gridDados.Columns.Item("NumPerMsr").Editable = false;
                    this.gridDados.Columns.Item("LineNum").Visible = false;
                    this.gridDados.Columns.Item("TaxRate").Editable = false;

                    SAPbouiCOM.EditTextColumn colCodInterno = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("DocEntry");
                    colCodInterno.LinkedObjectType = "22";

                    SAPbouiCOM.EditTextColumn colCodItem = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("ItemCode");
                    colCodItem.LinkedObjectType = "4";

                }

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }

        }
        #endregion 

        #region Evento Click - Botão Pesquisar Importação
        // Evento Click - Botão Pesquisar Importação
        private void Button6_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Pesquisa de importacao em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            try
            {
                //SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Recordset oRSet = (SAPbobsCOM.Recordset)Program._Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                String vCodPN = EditText9.Value;
                String vCodPedido = EditText2.Value;
                String vCodProcesso = EditText0.Value;

                if (vCodPN != "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Código do PN através do botão pesquisa");
                    return;
                }

                if (vCodPedido != "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Pedido através do botão pesquisa");
                    return;
                }

                if (vCodProcesso != "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Realizar consulta pelo Nº do proceso através do botão pesquisa");
                    return;
                }

                String vCodImp = EditText12.Value;

                if (vCodImp == "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Informar número do documento");
                    return;
                }


                string Query = "SELECT " +
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
	                                " \"@ALFT_IMPORT\" as T0 " +
	                                " INNER JOIN \"@ALFT_IMPORT1\" AS T1 ON T1.U_DocEntry = T0.U_DocEntry " + 
	                                " INNER JOIN OITM AS T2 ON T2.ItemCode = T1.U_CODPRODUTO " + 
	                                " LEFT JOIN ONCM AS T3 ON T3.AbsEntry = T2.NcmCode " + 
                                " WHERE " +
	                                " T0.U_DocEntry = " + vCodImp;

                oRSet.DoQuery(Query);

                if (oRSet.RecordCount > 0)
                {
                    this.dtDados.Rows.Clear();

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
                        Double vPrice = Double.Parse(oRSet.Fields.Item("Price").Value.ToString());
                        this.dtDados.SetValue("Price", i, vPrice.ToString("N4"));
                        Double vLineTotal = Double.Parse(oRSet.Fields.Item("LineTotal").Value.ToString());
                        this.dtDados.SetValue("LineTotal", i, vLineTotal.ToString("N2"));
                        this.dtDados.SetValue("Quantity", i, oRSet.Fields.Item("Quantity").Value.ToString());
                        this.dtDados.SetValue("OpenQty", i, oRSet.Fields.Item("OpenQty").Value.ToString());
                        this.dtDados.SetValue("unitMsr", i, oRSet.Fields.Item("unitMsr").Value.ToString());
                        Double vPeso = Double.Parse(oRSet.Fields.Item("Peso").Value.ToString());
                        this.dtDados.SetValue("Peso", i, vPeso.ToString("N3"));
                        this.dtDados.SetValue("NumPerMsr", i, oRSet.Fields.Item("NumPerMsr").Value.ToString());
                        this.dtDados.SetValue("LineNum", i, oRSet.Fields.Item("LineNum").Value.ToString());
                        this.dtDados.SetValue("TaxRate", i, oRSet.Fields.Item("TaxRate").Value.ToString());

                        this.gridDados.DataTable.SetValue(0, i, "Y");
                        Double vFreteLinha = Double.Parse(oRSet.Fields.Item("FreteLinha").Value.ToString());
                        this.gridDados.DataTable.SetValue(13, i, vFreteLinha.ToString("N4"));

                        Double vOtrDespLinha = Double.Parse(oRSet.Fields.Item("OtrDespLinha").Value.ToString());
                        this.gridDados.DataTable.SetValue(14, i, vOtrDespLinha.ToString("N4"));

                        String TXDI = oRSet.Fields.Item("TXDI").Value.ToString().Replace(",", ".");
                        EditText4.Value = TXDI;

                        String FRTINT = oRSet.Fields.Item("FRTINT").Value.ToString().Replace(",", ".");
                        EditText6.Value = FRTINT;

                        String OTRDESP = oRSet.Fields.Item("OTRDESP").Value.ToString().Replace(",", ".");
                        EditText3.Value = OTRDESP;

                        Double TTPESO = Double.Parse(oRSet.Fields.Item("TTPESO").Value.ToString());
                        EditText1.Value = TTPESO.ToString("N3");

                        Double TTME = Double.Parse(oRSet.Fields.Item("TTME").Value.ToString());
                        EditText5.Value = TTME.ToString("N2");

                        Double TOTALII = Double.Parse(oRSet.Fields.Item("TaxValLine").Value.ToString());
                        EditText11.Value = TOTALII.ToString("N2");

                        Double TOTAL = Double.Parse(oRSet.Fields.Item("TOTAL").Value.ToString());
                        EditText8.Value = TOTAL.ToString("N2");

                        oRSet.MoveNext();
                    }

                    this.gridDados.Columns.Item("ColCheck").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    this.gridDados.Columns.Item("DocEntry").Editable = false;
                    this.gridDados.Columns.Item("DocNum").Editable = false;
                    this.gridDados.Columns.Item("LineNum").Editable = false;
                    this.gridDados.Columns.Item("NcmCode").Editable = false;
                    this.gridDados.Columns.Item("ItemCode").Editable = false;
                    this.gridDados.Columns.Item("Dscription").Editable = false;
                    this.gridDados.Columns.Item("Price").Editable = false;
                    this.gridDados.Columns.Item("LineTotal").Editable = false;
                    this.gridDados.Columns.Item("OpenQty").Editable = false;
                    this.gridDados.Columns.Item("unitMsr").Editable = false;
                    this.gridDados.Columns.Item("Peso").Editable = false;
                    this.gridDados.Columns.Item("Frete").Editable = false;
                    this.gridDados.Columns.Item("OtrDesp").Editable = false;
                    this.gridDados.Columns.Item("NumPerMsr").Editable = false;
                    this.gridDados.Columns.Item("LineNum").Visible = false;
                    this.gridDados.Columns.Item("VisOrder").Editable = false;
                    this.gridDados.Columns.Item("TaxRate").Editable = false;

                    SAPbouiCOM.EditTextColumn colCodInterno = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("DocEntry");
                    colCodInterno.LinkedObjectType = "22";

                    SAPbouiCOM.EditTextColumn colCodItem = (SAPbouiCOM.EditTextColumn)this.gridDados.Columns.Item("ItemCode");
                    colCodItem.LinkedObjectType = "4";

                }

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }
        }
        #endregion

        #region Evento Click - Botão Salvar
        // Evento Click - Botão Salvar 
        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Salvar em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            double vTxtTxID;
            if (EditText4.Value != "")
            {
                vTxtTxID = Double.Parse(EditText4.Value.Replace(".", ","));

            }
            else
            {
                vTxtTxID = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                return;
            }

            double vFrtInt;
            if (EditText6.Value != "")
            {
                vFrtInt = Double.Parse(EditText6.Value.Replace(".", ","));

            }
            else
            {
                vFrtInt = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional.");
                return;
            }

            double vOtrDesp;
            if (EditText3.Value != "")
            {
                vOtrDesp = Double.Parse(EditText3.Value.ToString().Replace(".", ","));
            }
            else
            {
                vOtrDesp = 0;
            }

            double vContainer;
            if (EditText7.Value != "")
            {
                vContainer = Double.Parse(EditText7.Value.ToString().Replace(".", ","));
            }
            else
            {
                vContainer = 0;
            }

            Double vTotalPeso = Double.Parse(EditText1.Value.PadLeft(3, '0'));
            Double vTotalME = Double.Parse(EditText5.Value.PadLeft(2, '0'));
            Double vTotal = Double.Parse(EditText8.Value.PadLeft(2, '0'));
            Double vTotalII = Double.Parse(EditText11.Value.PadLeft(2, '0'));

            //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("2");
            //string Hoje = "2021/03/24";
            string Hoje = DateTime.Now.ToShortDateString();

            string vCodImp = EditText12.Value;
            string vRetornoNrImp;
            if (vCodImp != "") 
            { 
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }
            else
            {
                vCodImp = EditText13.Value.ToString();
                //int CodExiste = Int32.Parse(EditText13.ToString()) - 1;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }
                
            if ((vRetornoNrImp != ""))
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
                if (this.gridDados.DataTable.GetValue(0, y).ToString() == "Y")
                {
                    try
                    {
                        int vPedido = Int32.Parse(this.gridDados.DataTable.GetValue("DocEntry", y).ToString());
                        int vPedNumDoc = Int32.Parse(this.gridDados.DataTable.GetValue("DocNum", y).ToString());
                        int vVisOrder = Int32.Parse(this.gridDados.DataTable.GetValue("VisOrder", y).ToString());
                        string vProduto = this.gridDados.DataTable.GetValue("ItemCode", y).ToString();
                        //string vDescricao = this.gridDados.DataTable.GetValue(6, y).ToString().Replace("'", " ").Replace("&", " ");
                        string vDescricao = "";
                        double vPrecoUnit = double.Parse(this.gridDados.DataTable.GetValue("Price", y).ToString());
                        double vPrecoTotal = double.Parse(this.gridDados.DataTable.GetValue("LineTotal", y).ToString());
                        double vQuantidade = double.Parse(this.gridDados.DataTable.GetValue("Quantity", y).ToString());
                        double vQuantidadeAberta = double.Parse(this.gridDados.DataTable.GetValue("OpenQty", y).ToString());
                        string vUm = this.gridDados.DataTable.GetValue("unitMsr", y).ToString();
                        double vPeso = double.Parse(this.gridDados.DataTable.GetValue("Peso", y).ToString());
                        double vFrete = double.Parse(this.gridDados.DataTable.GetValue("Frete", y).ToString());
                        double vOutrasDespesas = double.Parse(this.gridDados.DataTable.GetValue("OtrDesp", y).ToString());
                        int vItens = Int32.Parse(this.gridDados.DataTable.GetValue("NumPerMsr", y).ToString());
                        int vLinPed = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", y).ToString());
                        double vTaxRate = double.Parse(this.gridDados.DataTable.GetValue("TaxRate", y).ToString());

                        string vDeposito = "01";

                        if ((vRetornoNrImp != ""))
                        {
                            // Inserir Linha de Importacao
                            InserirLinhaImportacao(Int32.Parse(vRetornoNrImp), vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vVisOrder, vTaxRate);
                        }
                        else
                        {
                            // Inserir Linha de Importacao
                            InserirLinhaImportacao(0, vPedido, vProduto, vDescricao, vPrecoUnit, vPrecoTotal, vQuantidade, vQuantidadeAberta, vQuantidade, vUm, vPeso, vFrete, vOutrasDespesas, vItens, vDeposito, vPedNumDoc, vLinPed, vVisOrder, vTaxRate);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw;
                    }

                }
            }

            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Importação foi salva com sucesso.");

        }
        #endregion

        #region Evento Click Botão Cancelar
        // Evento Click Botão Cancelar
        private void Button4_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
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
        #endregion

        #region EVENTO CLICK BOTÃO CALCULAR
        // EVENTO CLICK BOTÃO CALCULAR
        private void Button3_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                double vTxtTxID;
                if (EditText4.Value != "")
                {
                    vTxtTxID = Double.Parse(EditText4.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vTxtTxID = 0.00;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                    return;
                }

                double vFrtInt;
                if (EditText6.Value != "")
                {
                    vFrtInt = Double.Parse(EditText6.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vFrtInt = 0;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional.");
                    return;
                }

                double vOtrDesp;
                if (EditText3.Value != "")
                {
                    vOtrDesp = Double.Parse(EditText3.Value.ToString().Replace(".", ","));
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
                        Double vTotalFrete = ((vFrtInt / TotalPeso) * Double.Parse(this.gridDados.DataTable.GetValue("Peso", y).ToString()));
                        this.gridDados.DataTable.SetValue("Frete", y, vTotalFrete.ToString("N4"));
                        Double vCalOtrDesp = (vOtrDesp / contador);
                        this.gridDados.DataTable.SetValue("OtrDesp", y, (vCalOtrDesp.ToString("N4")));

                        EditText1.Value = TotalPeso.ToString("N3");
                        EditText5.Value = TotalME.ToString("N2");

                    }
                }
                */

                for (int z = 0; z <= this.dtDados.Rows.Count - 1; z++)
                {

                    if (this.gridDados.DataTable.GetValue(0, z).ToString() == "Y")
                    {
                        TotalPeso = TotalPeso + Double.Parse(this.gridDados.DataTable.GetValue("Peso", z).ToString());
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
                        Double vPesoLinha = Double.Parse(this.gridDados.DataTable.GetValue("Peso", x).ToString());
                        Double vTotalFrete = ((vFrtInt / TotalPeso) * vPesoLinha);
                        this.gridDados.DataTable.SetValue("Frete", x, vTotalFrete.ToString("N4"));

                        TotalME = TotalME + Double.Parse(this.gridDados.DataTable.GetValue("LineTotal", x).ToString());
                        contador = contador + 1;
                        Double vCalOtrDesp = (vOtrDesp / contador);
                        this.gridDados.DataTable.SetValue("OtrDesp", x, (vCalOtrDesp.ToString("N4")));

                        Double vQuantidade = Double.Parse(this.gridDados.DataTable.GetValue("Quantity", x).ToString());
                        Double vPreco = Double.Parse(this.gridDados.DataTable.GetValue("Price", x).ToString());
                        Double vFreteLinha = Double.Parse(this.gridDados.DataTable.GetValue("Frete", x).ToString());
                        Double vTaxRate = Double.Parse(this.gridDados.DataTable.GetValue("TaxRate", x).ToString());

                        TotalII = TotalII + (((vTxtTxID * (vPreco*vQuantidade)) + (vFreteLinha)) * (vTaxRate / 100));

                    }

                }

                EditText1.Value = TotalPeso.ToString("N3");
                EditText11.Value = TotalII.ToString("N2");
                EditText5.Value = TotalME.ToString("N2");

                Total = (((vTxtTxID * TotalME) + vFrtInt) + TotalII);
                EditText8.Value = Total.ToString("N2");

            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }

        }
        #endregion

        #region Geração de Nota Fiscal - Recebimento de Mercadoria Esboço
        // Geração de Nota Fiscal - Recebimento de Mercadoria Esboço
        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Geração em andamento.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

            //Processo para Salvar antes de inserir o documento
            string vRetornoNrImp;
            string vCodImp = EditText12.Value;
            //if (string.IsNullOrEmpty(vCodImp) || vCodImp != "")
            if (vCodImp != "")
            {
               // vCodImp = EditText13.Value;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
                if ((vRetornoNrImp != "") && (vCodImp != ""))
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
                vCodImp = EditText13.Value;
                vRetornoNrImp = RetornaExisteImportacao(Int32.Parse(vCodImp));
            }


            double vTxtTxID;
            if (EditText4.Value != "")
            {
                vTxtTxID = Double.Parse(EditText4.Value.ToString().Replace(".", ","));
            }
            else
            {
                vTxtTxID = 0;
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar taxa DI.");
                return;
            }

            if ((vRetornoNrImp != "") || (vRetornoNrImp == "0"))
            //if (vRetornoNrImp != "")
            {

                double vFrtInt;
                if (EditText6.Value != "")
                {
                    vFrtInt = Double.Parse(EditText6.Value.ToString().Replace(".", ","));

                }
                else
                {
                    vFrtInt = 0;
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Necessário informar frete internacional");
                    return;
                }

                double vOtrDesp;
                if (EditText3.Value != "")
                {
                    vOtrDesp = Double.Parse(EditText3.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vOtrDesp = 0;
                }

                double vContainer;
                if (EditText7.Value != "")
                {
                    vContainer = Double.Parse(EditText7.Value.ToString().Replace(".", ","));
                }
                else
                {
                    vContainer = 0;
                }

                Double vTotalPeso = Double.Parse(EditText1.Value.PadLeft(3, '0'));
                Double vTotalME = Double.Parse(EditText5.Value.PadLeft(2, '0'));
                Double vTotalII = Double.Parse(EditText11.Value.PadLeft(2, '0'));
                Double vTotal = Double.Parse(EditText8.Value.PadLeft(2, '0'));

                string Hoje = "2021/03/25";
                //string Hoje = DateTime.Now.ToShortDateString();

                if((vRetornoNrImp != "") || (vRetornoNrImp == "0")) 
                {
                    // Inserir Cabeçalho de Importacao
                    InserirCabecalhoImportacao(0, vTxtTxID, vFrtInt, vOtrDesp, vContainer, vTotalPeso, vTotalME, vTotal, "ABERTO", Hoje, vTotalII);

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
                            double vPrecoUnit = double.Parse(this.gridDados.DataTable.GetValue("Price", y).ToString());
                            double vPrecoTotal = double.Parse(this.gridDados.DataTable.GetValue("LineTotal", y).ToString());
                            double vQuantidade = double.Parse(this.gridDados.DataTable.GetValue("Quantity", y).ToString());
                            double vQuantidadeAberta = double.Parse(this.gridDados.DataTable.GetValue("OpenQty", y).ToString());
                            string vUm = this.gridDados.DataTable.GetValue("unitMsr", y).ToString();
                            double vPeso = double.Parse(this.gridDados.DataTable.GetValue("Peso", y).ToString());
                            double vFrete = double.Parse(this.gridDados.DataTable.GetValue("Frete", y).ToString());
                            double vOutrasDespesas = double.Parse(this.gridDados.DataTable.GetValue("OtrDesp", y).ToString());
                            int vItens = Int32.Parse(this.gridDados.DataTable.GetValue("NumPerMsr", y).ToString());
                            int vLineNum = Int32.Parse(this.gridDados.DataTable.GetValue("LineNum", y).ToString());
                            double vTaxRate = double.Parse(this.gridDados.DataTable.GetValue("TaxRate", y).ToString());
                            string vDeposito = "01";

                            if ((vRetornoNrImp != ""))
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
            var vEsbocoNFRecebimento = (SAPbobsCOM.Documents)Program._Company.GetBusinessObject(BoObjectTypes.oDrafts);
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
                            vEsbocoNFRecebimento.Lines.Quantity = Double.Parse(this.gridDados.DataTable.GetValue("Quantity", i).ToString());

                            // Converter de Dolar para Real usando TX ID - Multiplicando. Somar o peso pelas linhas selecionada e depois dividir frete pelo total e multiplicar pelo peso da linha   
                            Double vQuantidade = Double.Parse(this.gridDados.DataTable.GetValue("Quantity", i).ToString());
                            Double vPreco = Double.Parse(this.gridDados.DataTable.GetValue("Price", i).ToString());
                            Double vFreteLinha = Double.Parse(this.gridDados.DataTable.GetValue("Frete", i).ToString());
                            Double vTaxRate = Double.Parse(this.gridDados.DataTable.GetValue("TaxRate", i).ToString());

                            vEsbocoNFRecebimento.Lines.UnitPrice = (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) + (((vTxtTxID * vPreco) + (vFreteLinha / vQuantidade)) * (vTaxRate/100)));

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
                        string MessagemErro = Program._Company.GetLastErrorDescription();
                        throw new Exception(MessagemErro);
                        return;
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Esboço de Recebimento criado com sucesso.");
                        string draftDocEntry = RetornaCodDraft();
                        SAPbouiCOM.Framework.Application.SBO_Application.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
                        return;
                    }
                }
            }
        }
        #endregion

        #region Metodo Para Retornar Fornecedor através do Código de Pedido
        public static String RetornaFornecedor(int pDocEntry)
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT \"CardCode\" As CODFORNECEDOR FROM OPOR WHERE \"DocEntry\" = " + pDocEntry;

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("CODFORNECEDOR").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Para Retornar o Código de Utilização
        public static String RetornaUtizacao(int pDocEntry)
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT \"MainUsage\" as UTILIZACAO FROM \"POR12\" WHERE \"DocEntry\" =" + pDocEntry;

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("UTILIZACAO").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Pra retornar o docentry
        public static String RetornaPedDocEntry(int pDocNum)
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT \"DocEntry\" as CODIGO FROM \"OPOR\" WHERE \"DocNum\" =" + pDocNum;

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("CODIGO").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Pra retornar a descrição do item por itemcoode
        public static String RetornaDescItem(string pItemCode)
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT \"ItemName\" as DESCRICAO FROM \"OITM\" WHERE \"ItemCode\" = '" + pItemCode + "'";

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("DESCRICAO").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Para Retornar o Código do Imposto através do Código do Pedido
        public static String RetornaCodImposto(int pDocEntry, int pLineNum)
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT \"POR1\".\"TaxCode\" as TAXCODE FROM \"OPOR\" INNER JOIN \"POR1\" ON \"POR1\".\"DocEntry\" = \"OPOR\".\"DocEntry\" WHERE \"POR1\".\"DocEntry\" =" + pDocEntry + " AND \"POR1\".\"LineNum\" =" + pLineNum ;

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("TAXCODE").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Para Retornar Código do DRAFS
        public static String RetornaCodDraft()
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT TOP 1 \"DocEntry\" as CODIGO FROM \"ODRF\" ORDER BY \"DocEntry\" DESC";

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("CODIGO").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Para Retornar Code da Tabela Log
        public static String RetornaCodeLog()
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT Top 1 Convert(int, Code) + 1 AS NEWCODE FROM \"@ALFT_IMPORT\" order by Convert(int,code) desc";

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("NEWCODE").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo Para Retornar Code da Tabela Log
        public static String RetornaCodeLogLinha()
        {
            Recordset recordset = null;

            try
            {

                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT TOP 1 CONVERT(int, Code) as Codigo FROM \"@ALFT_IMPORT1\" order by Codigo desc ";

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("Codigo").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo para Inserir Registros na Tabela
        public static void InserirCabecalhoImportacao(int pDocEntry, double pTxId, double pFrtInt, double pOtrDesp, double pContainer, double pTotalPeso, double pTotalMe, double pTotal, string pStatus, string pData, double pTotalII)
        {
            var ret = 0;
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                string code;
                string vDocEntry;
                if (pDocEntry == 0) { 
                    code = RetornaCodeLog();
                    vDocEntry = code;
                }
                else
                {
                    int Retorno = Int32.Parse(RetornaCodeLog()) + 1;
                    code = Retorno.ToString();
                    vDocEntry = pDocEntry.ToString();
                }
                var sql = "insert into \"@ALFT_IMPORT\" values(\'" + code + "\',\'" + code + "\',\'" + vDocEntry + "\', " + pTxId.ToString().Replace(",", ".") + ", " + pFrtInt.ToString().Replace(",", ".") + ", " + pOtrDesp.ToString().Replace(",", ".") + ", " + pContainer.ToString().Replace(",", ".") + ", " + pTotalPeso.ToString().Replace(",", ".") + ", " + pTotalMe.ToString().Replace(",", ".") + ", " + pTotal.ToString().Replace(",", ".") + ", \'" + pStatus + "\',\'" + pData + "\', " + pTotalII.ToString().Replace(",", ".") + ")";

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
        #endregion

        #region Metodo para Inserir Registros na Tabela
        public static void InserirLinhaImportacao(int pDocEntry, int pPedido, string pProduto, string pDescricao, double pPreco, double pPrecoTotal, double pQtdPedida, double pQtdAberta, double pQtdFat, string pUm, double pPeso, double pFrete, double pOutraDesp, int pItmUm, string pDeposito, int pPedNumDoc, int pLinPed, int pVisOrder, double pTaxOrder)
        {
            var ret = 0;
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                int code;
                int vDocEntry;
                int codelinha;
                if (pDocEntry == 0)
                {
                    string Retorno = RetornaCodeLogLinha();
                    if (Retorno != "")
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

                var sql = "insert into \"@ALFT_IMPORT1\" values(\'" + codelinha + "\',\'" + codelinha + "\',\'" + vDocEntry + "\',\'" + pPedido + "\',\'" + pProduto + "\',\'" + pDescricao + "\', " + pPreco.ToString().Replace(",", ".") + ", " + pPrecoTotal.ToString().Replace(",", ".") + ", " + pQtdPedida.ToString().Replace(",", ".") + ", " + pQtdAberta.ToString().Replace(",", ".") + ", " + pQtdFat.ToString().Replace(",", ".") + ", \'" + pUm + "\', " + pPeso.ToString().Replace(",", ".") + ", " + pFrete.ToString().Replace(",", ".") + " , " + pOutraDesp.ToString().Replace(",", ".") + ",\'" + pItmUm + "\',\'" + pDeposito + "\',\'" + pPedNumDoc + "\',\'" + pLinPed + "\',\'" + pVisOrder + "\', " + pTaxOrder.ToString().Replace(",", ".") + ")";
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
        #endregion

        #region Metodo Para Retornar o Numero da Ultima Importação
        public static String RetornaNrImportacao()
        {
            Recordset recordset = null;

            try
            {
                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);


                string query = "SELECT TOP 1 (\"U_DocEntry\")+1 AS NRIMPORT FROM \"@ALFT_IMPORT\" order by Convert(int,code) desc";
                //string query = "SELECT MAX(\"U_DocEntry\")+1 AS NRIMPORT FROM \"@ALFT_IMPORT\"";

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("NRIMPORT").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Retorna Se Existe a Importação
        public static String RetornaExisteImportacao(int pDocEntry)
        {
            Recordset recordset = null;

            try
            {
                recordset = (Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT ISNULL(\"U_DocEntry\",0) AS NRIMPORT FROM \"@ALFT_IMPORT\" WHERE \"U_DocEntry\" = " + pDocEntry;

                recordset.DoQuery(query);

                if (!recordset.EoF)
                {
                    return recordset.Fields.Item("NRIMPORT").Value.ToString().Trim();
                }
            }
            catch { }
            finally
            {
                recordset = null;
            }

            return String.Empty;
        }
        #endregion

        #region Metodo para Delete Registros na Tabela
        public static void DeleteCabecalhoImportacao(int vDocEntry)
        {
            var ret = 0;
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                string code = RetornaCodeLog();

                var sql = "DELETE FROM \"@ALFT_IMPORT\" WHERE \"U_DocEntry\" = " + vDocEntry;
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
        #endregion

        #region Metodo para Delete Registros na Tabela
        public static void DeleteLinhaImportacao(int vDocEntry)
        {
            var ret = 0;
            Recordset oRs = null;

            try
            {
                oRs = ((Recordset)Program._Company.GetBusinessObject(BoObjectTypes.BoRecordset));
                string code = RetornaCodeLog();

                var sql = "DELETE FROM \"@ALFT_IMPORT1\" WHERE \"U_DocEntry\" = " + vDocEntry;
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
        #endregion

        private StaticText StaticText13;
        private EditText EditText11;
        
    }
}