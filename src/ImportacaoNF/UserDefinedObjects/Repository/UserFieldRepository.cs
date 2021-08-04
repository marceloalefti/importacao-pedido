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
    public class UserFieldRepository
    {
        public void Salvar(SAPbobsCOM.UserFieldsMD oUserFieldsMD, UserField field)
        {
            string msg = string.Empty;
            int lRetCode = 0;

            oUserFieldsMD.TableName = field.Tabela.TableName;
            oUserFieldsMD.Name = field.Name;
            oUserFieldsMD.Description = field.Descricao;
            oUserFieldsMD.Type = field.Type;
            oUserFieldsMD.SubType = field.SubType;
            oUserFieldsMD.EditSize = field.Size;
            oUserFieldsMD.Mandatory = field.Mandatory;
            oUserFieldsMD.DefaultValue = field.DefaultValue;

            int i = 0;
            int zNumDel = oUserFieldsMD.ValidValues.Count;
            for (int t = 1; t <= zNumDel; t++)
            {
                oUserFieldsMD.ValidValues.Delete();
            }

            if (field.ValoresList != null)
            {
                foreach (UserFieldValue p1 in field.ValoresList)
                {
                    i++;
                    if (i >= 2) { oUserFieldsMD.ValidValues.Add(); }
                    oUserFieldsMD.ValidValues.Value = p1.VlCode;
                    oUserFieldsMD.ValidValues.Description = p1.VlName;
                }
            }

            int FieldID = CodigoCampo(field);
            bool existe = oUserFieldsMD.GetByKey(field.Tabela.TableName, FieldID);

            if (!existe)
                lRetCode = oUserFieldsMD.Add();
            else
                lRetCode = oUserFieldsMD.Update();


            if (lRetCode != 0)
            {
                msg = "Falha na criação do campo  [" + field.Name + "/" + field.Descricao + "] da tabela '" + field.Tabela.TableDescription + "'.";
                Mensagem.ExibirErro(msg);
                throw new Exception(msg);
            }
            else
            {
                Mensagem.ExibirMensagem("O campo " + field.Descricao + " da tabela " + field.Tabela.TableDescription + " foi criado com Sucesso!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }            
        }

        public static Int32 CodigoCampo(UserField c)
        {
            return Convert.ToInt32(ProcurarCampo("FieldID", "CUFD", "TAbleId = '@" + c.Tabela + "' and AliasId = '" + c.Name + "'"));
        }

        public static string ProcurarCampo(string zField, string zTabela, string zWhere)
        {
            string Criterio;
            SAPbobsCOM.Recordset Rs = (SAPbobsCOM.Recordset)Conexao.ConexaoSAP.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            Criterio = "SELECT " + zField + " FROM " + zTabela + " WHERE " + zWhere;
            
            Rs.DoQuery(Criterio);
            if (!Rs.EoF)
            {
                Rs.MoveFirst();
                Criterio = Rs.Fields.Item(0).Value.ToString();
            }
            else
            {
                Criterio = "";
            }
            if (Criterio == "") Criterio = "0";

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Rs);
            Rs = null;
            GC.Collect();

            return Criterio;
        }

    }
}
