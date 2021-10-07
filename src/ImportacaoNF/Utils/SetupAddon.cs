using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SLT.ImportacaoNF.UserDefinedObjects;
using SLT.ImportacaoNF.UserDefinedObjects.Models;
using SLT.ImportacaoNF.Conexao;
using SLT.ImportacaoNF.UserDefinedObjects.Repository;

namespace SLT.ImportacaoNF.Utils
{
    public class SetupAddon
    {
        public void InstalarEAtualizar()
        {
            try
            {
                Conexao.ConexaoSAP.Company.StartTransaction();

                if (!new UserTableRepository().TabelaExiste("SLTSETUP"))
                    Instalacao();

                if (SetupAddon.GetVersion("1.0.0.0"))
                    v1();

                Conexao.ConexaoSAP.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                Conexao.ConexaoSAP.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                Mensagem.ExibirErro(ex.Message);
                throw ex;
            }
        }



        #region Primeira Versão

        private void Instalacao()
        {
            Mensagem.ExibirMensagem("Aguarde enquanto são criadas as estruturas necessárias para o funcionamento do add-on.",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            TabelasConfiguracao();

            Mensagem.ExibirMensagem("Instalação do Add-on realizada com sucesso!",
                        SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        private void TabelasConfiguracao()
        {
            UserTable tabela = new UserTable();
            tabela.TableName = "SLTSETUP";
            tabela.TableDescription = "SLT: Setup Addon";
            tabela.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject;

            UserField c = new UserField();
            c.DefaultValue = "1.0.0.0";
            c.Descricao = "Campo com a versão atual";
            c.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            c.Name = "Versao";
            c.Size = 10;
            c.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            c.Tabela = tabela;
            c.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;

            tabela.Campos.Add(c);


            new UserTableRepository().Salvar(tabela);

            SetupAddon.SetVersion(c.DefaultValue);
        }

        #endregion Primeira Versão



        #region Segunda Versão

        private void v1()
        {
            Mensagem.ExibirMensagem("Aguarde enquanto as configurações de Importação são realizadas.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            ImportacaoUDO();
            SetupAddon.SetVersion("2.0.0.1");

            Mensagem.ExibirMensagem("As configurações para as Importações foram realizada com sucesso!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        private void ImportacaoUDO()
        {
            // UDO
            UDO udo = new UDO();
            udo.Code = "OSLT_IMPORT";
            udo.Name = "Importação de Pedido";
            udo.TableName = "SLTIMPORT";
            udo.CamposBusca = new List<UserField>();
            udo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;

            ImportacaoTable(udo);
            UserTable table = ImportacaoTableItem(udo);

            udo.Filhos.Add(table);
            new UDORepository().Salvar(udo);
        }

        private void ImportacaoTable(UDO udo)
        {
            // TABELA
            UserTable table = new UserTable();
            table.TableName = "SLTIMPORT";
            table.TableDescription = "SLT: Importação";
            table.TableType = SAPbobsCOM.BoUTBTableType.bott_Document;

            // CAMPOS
            UserField field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Taxa DI";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "TaxaDI";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field);

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Frete Internacional";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "FreteInt";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field); 

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Outras despesas";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "OutDesp";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field); 

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Container";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO;
            field.Name = "Container";
            field.Size = 16;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field); 

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Peso Total";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "TotalPes";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
            table.Campos.Add(field); 

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Total FOB";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "TotalFOB";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field); 

            field = null;
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Total II";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Total2";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field); 

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Total";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Total";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field); 

            new UserTableRepository().Salvar(table);
        }

        private UserTable ImportacaoTableItem(UDO udo)
        {
            // TABELA
            UserTable table = new UserTable();
            table.TableName = "SLTIMPRT1";
            table.TableDescription = "SLT: Item da Importação";
            table.TableType = SAPbobsCOM.BoUTBTableType.bott_DocumentLines;

            // CAMPOS
            UserField field = null;

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Item Selecionado";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Selected";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            field.ValoresList = new List<UserFieldValue>();
            field.ValoresList.Add(new UserFieldValue("Y", "Yes"));
            field.ValoresList.Add(new UserFieldValue("N", "No"));
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Código do Fornecedor";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "CardCode";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Código do Pedido";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "PedidoId";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Número do Pedido";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "PedidoNr";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field); 

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Número do Item no Pedido";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "ItemNum";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Item Code";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "ItemCode";
            field.Size = 100;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Item Description";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Descript";
            field.Size = 200;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Preço Unitário";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "PrecoUni";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field);
            
            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Preço Internacional";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "PrecoInt";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Quantidade do Pedido";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "QtdPed";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Quantidade Disponivel";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "QtdDisp";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Quantidade Faturada";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "QtdFat";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Unidade de Medida";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Unit";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Peso";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Peso";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Código NCM";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "NcmCode";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Frete";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Frete";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Outras Despesas";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "OutroDes";
            field.Size = 19;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Depósito";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Deposito";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Imposto do Pedido";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Imposto";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Alíquota do Imposto";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Aliquota";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Float;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            table.Campos.Add(field);

            field = new UserField();
            field.Tabela = table;
            field.Descricao = "Utilização do Item";
            field.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
            field.Name = "Utiliza";
            field.Size = 8;
            field.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            field.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
            table.Campos.Add(field);           

            new UserTableRepository().Salvar(table);
            return table;
        }


        #endregion Segunda Versão


        public static bool GetVersion(string Version)
        {
            if (Conexao.ConexaoSAP.ExecuteScalar("U_Versao", "[@SLTSETUP]", "Name = '" + System.Windows.Forms.Application.ProductName + "' and U_Versao = '" + Version + "'") != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static void SetVersion(string Version)
        {
            SAPbobsCOM.UserTable oUserTable;
            oUserTable = Conexao.ConexaoSAP.Company.UserTables.Item("SLTSETUP");
            try
            {
                if (!oUserTable.GetByKey("001"))
                {
                    oUserTable.Code = "001";
                    oUserTable.Name = System.Windows.Forms.Application.ProductName;
                    oUserTable.UserFields.Fields.Item("U_Versao").Value = Version;
                    oUserTable.Add();
                }
                else
                {
                    oUserTable.UserFields.Fields.Item("U_Versao").Value = Version;
                    oUserTable.Update();
                }
            }
            catch (Exception ex)
            {
                Mensagem.ExibirMensagem(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
        }
    }
}
