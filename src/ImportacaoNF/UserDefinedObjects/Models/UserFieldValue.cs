using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLT.ImportacaoNF.UserDefinedObjects.Models
{
    public class UserFieldValue
    {
        public UserFieldValue(string vlCode, string vlName)
        {
            this.VlCode = vlCode;
            this.VlName = vlName;
        }

        public string VlCode { get; set; }
        public string VlName { get; set; }
    }
}
