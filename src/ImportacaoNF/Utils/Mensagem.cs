using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.Utils
{
    public static class Mensagem
    {
        public static void ExibirMensagem(string msg, SAPbouiCOM.BoMessageTime tempo, SAPbouiCOM.BoStatusBarMessageType tipo)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(msg, tempo, tipo);
        }

        public static void ExibirErro()
        {
            ExibirErro(string.Empty);
        }

        public static void ExibirErro(string prefixo)
        {
            ExibirErro(prefixo, string.Empty);
        }

        public static void ExibirErro(string prefixo, string sufixo)
        {
            string msg = String.Concat(prefixo, Conexao.ConexaoSAP.Company.GetLastErrorCode(), ": ", Conexao.ConexaoSAP.Company.GetLastErrorDescription(), sufixo);
            ExibirMensagem(msg, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
        }
    }
}
