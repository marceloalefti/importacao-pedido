using SAPbouiCOM;
using SLT.ImportacaoNF.UserDefinedObjects.Models;
using SLT.ImportacaoNF.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.UserDefinedObjects
{
    public class UserTableRepository
    {
        public void Salvar(UserTable Tabela)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = (SAPbobsCOM.UserTablesMD)Conexao.ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            string msg = string.Empty;
            int lRetCode = 0;


            bool existe = oUserTablesMD.GetByKey(Tabela.TableName);

            if (!existe)
            {
                oUserTablesMD.TableName = Tabela.TableName;
                oUserTablesMD.TableDescription = Tabela.TableDescription;
                oUserTablesMD.TableType = Tabela.TableType;

                lRetCode = oUserTablesMD.Add();
            }
            else
            {
                oUserTablesMD.TableDescription = Tabela.TableDescription;
                lRetCode = oUserTablesMD.Update();
            }

            //Verifica erros na criação da tabela
            if (lRetCode != 0)
            {
                msg = "Falha ao salvar a tabela: " + Tabela.TableName;
                Mensagem.ExibirErro(msg);
                throw new Exception(msg);
            }
            else
            {
                Mensagem.ExibirMensagem("Tabela: '" + Tabela.TableDescription + "' criada com sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            oUserTablesMD = null;
            GC.Collect();


            SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Conexao.ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            UserFieldRepository userFieldRepository = new UserFieldRepository();

            foreach (UserField field in Tabela.Campos)
            {
                userFieldRepository.Salvar(oUserFieldsMD, field);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
        }

        public bool TabelaExiste(string tabela)
        {
            try
            {
                SAPbobsCOM.UserTablesMD oUserTablesMD = (SAPbobsCOM.UserTablesMD)Conexao.ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                var exists = oUserTablesMD.GetByKey(tabela) == true;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.Collect();

                return exists;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
