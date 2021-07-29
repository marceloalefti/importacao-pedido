using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportacaoNF.Conexao
{
    static class ConexaoSAP
    {
        public static SAPbouiCOM.Application oApplication;

        public static SAPbobsCOM.Company oCompany;

        public static SAPbouiCOM.SboGuiApi SboGuiApi = null;

        private static bool oCompanyConnected = false;

        public static void SetApplication()
        {
            string sConnectionString = null;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            try
            {
                SboGuiApi.Connect(sConnectionString);

            }
            catch (Exception e)
            {
                ConexaoSAP.oApplication.MessageBox("Erro na Conexão: " + e.Message, 1, "Ok", "", "");
                throw e;
            }

            oApplication = SboGuiApi.GetApplication(-1);
        }

        public static void InitializeCompany()
        {
            int lRetCode;

            try
            {

                if (oCompanyConnected == true) return;
                int setConnectionContextReturn = 0;
                string sCookie = null;
                string sConnectionContext = null;

                oCompany = new SAPbobsCOM.Company();

                //oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Portuguese_Br;
                sCookie = oCompany.GetContextCookie();
                sConnectionContext = oApplication.Company.GetConnectionContext(sCookie);

                if (oCompany.Connected == true)
                {
                    oCompany.Disconnect();
                }

                //oApplication.StatusBar.SetText("Inicializando Companhia", SAPbouiCOM.BoMessageTime.bmt_Short, (SAPbouiCOM.BoStatusBarMessageType)SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                setConnectionContextReturn = oCompany.SetSboLoginContext(sConnectionContext);
                oCompanyConnected = true;
                lRetCode = oCompany.Connect();

                return;

            }
            catch (Exception e)
            {
                oApplication.MessageBox(e.Message + e.StackTrace);
            }
        }
    }
}
