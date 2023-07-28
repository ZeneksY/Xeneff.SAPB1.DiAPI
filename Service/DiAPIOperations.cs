using SAPbobsCOM;
using System;
using System.Collections.Generic;
using Xeneff.SAPB1.DiAPI.Common;
using Xeneff.SAPB1.DiAPI.DataAccess;
using Xeneff.SAPB1.DiAPI.DTO;
using Xeneff.SAPB1.DiAPI.Result;

namespace Xeneff.SAPB1.DiAPI.Operations
{
    public class DiAPIOperations : DiAPIContext
    {
        public DiAPIResult AddItem()
        {
            DiAPIResult result = null;
            try
            {
                if (Connect() == Constants.DiApiSuccess)
                {
                    Company oCompany = GetCompany();
                    Items oItems = oCompany.GetBusinessObject(BoObjectTypes.oItems);

                    oItems.ItemCode = "ItemCode10";
                    oItems.ItemName = "ItemName10";
                    oItems.ItemsGroupCode = 100;
                    oItems.UoMGroupEntry = 1;
                    oItems.InventoryItem = BoYesNoEnum.tYES;
                    oItems.SalesItem = BoYesNoEnum.tYES;
                    oItems.PurchaseItem = BoYesNoEnum.tYES;
                    oItems.Valid = BoYesNoEnum.tYES;
                    oItems.Add();
                    result.Code = oCompany.GetLastErrorCode();
                    result.Message = oCompany.GetLastErrorDescription();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

        public List<UserDto> GetUsers()
        {
            Recordset oRecordSet = null;
            List<UserDto> response = new List<UserDto>();
            try
            {
                if (Connect() == Constants.DiApiSuccess)
                {
                    Company oCompany = GetCompany();
                    string query = string.Format(@"SELECT empId,firstName,lastName,middleName FROM OHEM");
                    oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery(query);
                    while (!oRecordSet.EoF)
                    {
                        UserDto currentItem = new UserDto();
                        currentItem.EmpId = oRecordSet.Fields.Item(0).Value;
                        currentItem.FirstName = oRecordSet.Fields.Item(1).Value;
                        currentItem.MiddleName = oRecordSet.Fields.Item(2).Value;
                        currentItem.LastName = oRecordSet.Fields.Item(3).Value;
                        response.Add(currentItem);
                        oRecordSet.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseSapBobsObject(oRecordSet);
            }
            return response;
        }
        public DiAPIResult AddInvoice()
        {
            DiAPIResult result = new DiAPIResult();
            try
            {
                if (Connect() == Constants.DiApiSuccess)
                {
                    Company oCompany = GetCompany();
                    Documents oInvoice = ((Documents)(oCompany.GetBusinessObject(BoObjectTypes.oInvoices)));
                    Documents oSalesOrder = ((Documents)(oCompany.GetBusinessObject(BoObjectTypes.oOrders)));
                    if (oSalesOrder.GetByKey(2470)) //DocEntry of existing sales order
                    {
                        oInvoice.Series = 4;
                        oInvoice.CardCode = oSalesOrder.CardCode;
                        oInvoice.CardName = oSalesOrder.CardName;
                        oInvoice.HandWritten = BoYesNoEnum.tNO;
                        oInvoice.PaymentGroupCode = oSalesOrder.PaymentGroupCode;
                        oInvoice.DocDate = DateTime.Now;
                        oInvoice.Address = oSalesOrder.Address;
                        oInvoice.Address2 = oSalesOrder.Address2;
                        oInvoice.Comments = oSalesOrder.Comments;

                        oSalesOrder.Lines.SetCurrentLine(0);
                        oInvoice.Lines.ItemCode = oSalesOrder.Lines.ItemCode;
                        oInvoice.Lines.ItemDescription = oSalesOrder.Lines.ItemDescription;
                        oInvoice.Lines.PriceAfterVAT = oSalesOrder.Lines.PriceAfterVAT;
                        oInvoice.Lines.Price = oSalesOrder.Lines.Price;
                        oInvoice.Lines.TaxTotal = oSalesOrder.Lines.TaxTotal;
                        oInvoice.Lines.Quantity = oSalesOrder.Lines.Quantity;
                        oInvoice.Lines.DiscountPercent = oSalesOrder.Lines.DiscountPercent;
                        oInvoice.Lines.BaseLine = 0;
                        oInvoice.Lines.BaseType = oSalesOrder.Lines.BaseType;
                        oInvoice.Lines.WarehouseCode = oSalesOrder.Lines.WarehouseCode;

                        oInvoice.Lines.BatchNumbers.BatchNumber = "BatchNumber1";
                        oInvoice.Lines.BatchNumbers.Quantity = 1;
                        oInvoice.Lines.BatchNumbers.ItemCode = oSalesOrder.Lines.ItemCode;

                        oInvoice.Lines.BatchNumbers.Add();
                        oInvoice.Lines.Add();
                        oInvoice.Add();
                        result.Code = oCompany.GetLastErrorCode();
                        result.Message = oCompany.GetLastErrorDescription();

                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }
        public DiAPIResult AddSalesOrder()
        {
            DiAPIResult result = new DiAPIResult();
            try
            {
                if (Connect() == Constants.DiApiSuccess)
                {
                    Company oCompany = GetCompany();
                    Documents oOrder = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oOrders);
                    oOrder.CardCode = "CardCode1";
                    oOrder.CardName = "CardName1";
                    oOrder.DocDate = DateTime.Now;
                    oOrder.DocDueDate = DateTime.Now.AddDays(3);
                    oOrder.Comments = "Comment1";
                    oOrder.Address2 = "Address2";
                    oOrder.ShipToCode = "Invoice Address";

                    oOrder.Lines.ItemCode = "ItemCode1";
                    oOrder.Lines.ItemDescription = "ItemName1";
                    oOrder.Lines.WarehouseCode = "WareHouse1";
                    oOrder.Lines.Quantity = 1;
                    oOrder.Lines.ShipDate = DateTime.Now.AddDays(5);
                    oOrder.Lines.Price = 10;
                    oOrder.Lines.Currency = "TRY";
                    oOrder.Lines.Rate = 1;
                    oOrder.Lines.DiscountPercent = 0;
                    oOrder.Lines.Add();
                    oOrder.Add();
                    result.Code = oCompany.GetLastErrorCode();
                    result.Message = oCompany.GetLastErrorDescription();

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }
        public DiAPIResult AddJournalEntry()
        {
            DiAPIResult result = new DiAPIResult();
            try
            {
                if (Connect() == Constants.DiApiSuccess)
                {
                    Company oCompany = GetCompany();

                    JournalEntries journalEntry = (JournalEntries)oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

                    journalEntry.ReferenceDate = DateTime.Now;
                    journalEntry.DueDate = DateTime.Now;
                    journalEntry.TaxDate = DateTime.Now;
                    journalEntry.AutoVAT = BoYesNoEnum.tNO;

                    journalEntry.Lines.AccountCode = "101010100";
                    journalEntry.Lines.ContraAccount = "252000000";
                    journalEntry.Lines.Credit = 0;
                    journalEntry.Lines.Debit = 150;
                    journalEntry.Lines.ShortName = "101010100";
                    journalEntry.Lines.Add();
                    journalEntry.Add();


                    result.Code = oCompany.GetLastErrorCode();
                    result.Message = oCompany.GetLastErrorDescription();


                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }
    }
}
