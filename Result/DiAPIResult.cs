using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xeneff.SAPB1.DiAPI.Common;

namespace Xeneff.SAPB1.DiAPI.Result
{
   public class DiAPIResult
    {
        public int Code { get; set; } = Constants.DefaultDiApiResult;
        public string Message { get; set; } = "Operation failed!";
    }
}
