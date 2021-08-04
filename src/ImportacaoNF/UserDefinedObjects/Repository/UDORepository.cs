using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SLT.ImportacaoNF.UserDefinedObjects.Models;
using SLT.ImportacaoNF.Utils;
using SAPbouiCOM;

namespace SLT.ImportacaoNF.UserDefinedObjects.Repository
{
    public class UDORepository
    {
        public UDO Obter(string udo_name, int DocEntry)
        {
            GeneralData gd = Conexao.ConexaoSAP.ObterRegistroPorID(udo_name, DocEntry);

            UDO udo = null;

            if (gd != null)
            {
                udo = new UDO();
                gd = null;
            }


            return udo;
        }

        public void Salvar(UDO udo)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectsMD;
            int lRetCode = 0;

            try
            {
                oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Conexao.ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                oUserObjectsMD.CanFind = udo.CanFind;
                oUserObjectsMD.CanDelete = udo.CanDelete;
                oUserObjectsMD.CanCancel = udo.CanCancel;
                oUserObjectsMD.CanYearTransfer = udo.CanYearTransfer;
                oUserObjectsMD.CanClose = udo.CanClose;
                oUserObjectsMD.CanCreateDefaultForm = udo.CanCreateDefaultForm;
                oUserObjectsMD.CanLog = udo.CanLog;
                oUserObjectsMD.ManageSeries = udo.ManageSeries;

                foreach (var coluna in udo.CamposBusca)
                {
                    oUserObjectsMD.FindColumns.ColumnAlias = "U_" + coluna.Name;
                    oUserObjectsMD.FindColumns.Add();
                }

                foreach (var tabela in udo.Filhos)
                {
                    oUserObjectsMD.ChildTables.TableName = tabela.TableName;
                    oUserObjectsMD.ChildTables.Add();
                }

                oUserObjectsMD.Code = udo.Code;
                oUserObjectsMD.Name = udo.Name;
                oUserObjectsMD.ObjectType = udo.ObjectType;
                oUserObjectsMD.TableName = udo.TableName;

                if (!oUserObjectsMD.GetByKey(oUserObjectsMD.Code))
                    lRetCode = oUserObjectsMD.Add();
                else
                    lRetCode = oUserObjectsMD.Update();

                if (lRetCode != 0)
                {
                    Mensagem.ExibirErro("Falha na criação de tabelas: ");
                }
                else
                {
                    Mensagem.ExibirMensagem("UDO: '" + oUserObjectsMD.Name + "' criado com sucesso!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                }


                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);
                oUserObjectsMD = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                Mensagem.ExibirErro(ex.Message);
            }
        }

    }
}
