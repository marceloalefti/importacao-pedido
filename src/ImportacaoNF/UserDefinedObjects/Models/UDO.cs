using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.UserDefinedObjects.Models
{
    public class UDO
    {
        public UDO()
        {
            CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
            CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
            CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
            CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
            CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
            CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
            ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;

            Filhos = new List<UserTable>();
        }

        public SAPbobsCOM.BoYesNoEnum CanFind { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanDelete { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCancel { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanYearTransfer { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanClose { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanLog { get; set; }
        public SAPbobsCOM.BoYesNoEnum ManageSeries { get; set; }

        public string Code { get; set; }
        public string Name { get; set; }
        public SAPbobsCOM.BoUDOObjType ObjectType { get; set; }
        public string TableName { get; set; }

        public List<UserTable> Filhos { get; set; }
        public List<UserField> CamposBusca { get; set; }

    }
}
