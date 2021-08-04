using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.UserDefinedObjects.Models
{
    public class UserTable
    {
        public UserTable()
        {
            Campos = new List<UserField>();
        }

        public string TableName { get; set; }
        public string TableDescription { get; set; }
        public SAPbobsCOM.BoUTBTableType TableType { get; set; }
        public List<UserField> Campos { get; set; }
    }
}
