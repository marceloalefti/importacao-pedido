using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.UserDefinedObjects.Models
{
    public class UserField
    {
        public UserTable Tabela { get; set; }
        public string Name { get; set; }
        public string Descricao { get; set; }
        public SAPbobsCOM.BoFieldTypes Type { get; set; }
        public SAPbobsCOM.BoFldSubTypes SubType { get; set; }
        public int Size { get; set; }
        public SAPbobsCOM.BoYesNoEnum Mandatory { get; set; }
        public string DefaultValue { get; set; }
        public List<UserFieldValue> ValoresList { get; set; }
    }
}
