using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xeneff.SAPB1.DiAPI.DTO;
using Xeneff.SAPB1.DiAPI.Operations;
using Xeneff.SAPB1.DiAPI.Result;

namespace Xeneff.SAPB1.DiAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            DiAPIOperations operation = new DiAPIOperations();
            DiAPIResult addItem = operation.AddItem();
            List<UserDto> users = operation.GetUsers();
            DiAPIResult addInvoice = operation.AddInvoice();
            DiAPIResult addSalesOrder = operation.AddSalesOrder();
            DiAPIResult addJournalEntry = operation.AddJournalEntry();
        }
    }
}
