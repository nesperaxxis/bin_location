using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Reflection;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Data.SqlClient;
using Application = SAPbouiCOM.Framework.Application;
using SAP_Helper;
using BIN_Location_Auto_Select_Addon.Forms;

namespace BINLocationAutoSelect.Helpers
{
    class GlobalVar
    {
        #region Variables
        public static SAP_Helper.SAP _oSAP { get; set; }
        public static SAPbobsCOM.Company iCompany { get; set; }
        public static SAPbobsCOM.Recordset oRS { get; set; }
        public static SAPbobsCOM.Recordset oRSExec { get; set; }
        public static SAPbobsCOM.Recordset oRSQuery { get; set; }
        public static SAPbobsCOM.Recordset oRSHdr { get; set; }
        public static SAPbobsCOM.Recordset oRSDtl { get; set; }
        public static SAPbobsCOM.JournalEntries oJounalEntry { get; set; }
        public static SAPbobsCOM.Payments oPayments { get; set; }
        public static vbHelper_Library.Windows vbWindowsHelper { get; set; }
        public static string oQuery { get; set; }

        public static SAPbobsCOM.Company myCompany { get; set; }
        public static SAPbobsCOM.Recordset myRS { get; set; }
        public static SAPbouiCOM.DBDataSource myDBds { get; set; }

        public static string oSERVERNAME { get; set; }
        public static string oSERVERDB { get; set; }
        public static string oSERVERUN { get; set; }
        public static string oSERVERPW { get; set; }
        public static string oSERVERTYPE { get; set; }
        public static string oB1UN { get; set; }
        public static string oB1PW { get; set; }

        public static Boolean oSuperUser { get; set; }

        static string strQuery { get; set; }
        static string lastMessage { get; set; }
        public static Boolean _Contains { get; set; }
        public static Boolean HasError = false;
        public static Boolean GlobalHasError = false;
        public static string DocMsgLine { get; set; }
        public static string oExcelFile { get; set; }
        public static string oLogPath { get; set; }
        public static string oExcelFileOrigPath { get; set; }
        public static string oSuccessPath { get; set; }
        public static string oFailedPath { get; set; }
        //public static SAPbouiCOM.Form _frmBinReceipt { get; set; }
        public static SAPbouiCOM.Form _frmGrpo { get; set; }
        public static bool _frmAutoPop { get; set; }
        #endregion

        #region "Properties"
        public string LastErrorMessage
        {
            get
            {
                return lastMessage;
            }
        }

        public static object Iif(bool expression, object truePart, object falsePart)
        { return expression ? truePart : falsePart; }

        public static object TrimData(string oValue)
        { return oValue.Replace("'", "''"); }

        public static object TrimStringData(string oValue)
        { return oValue.Replace(" ", ""); }

        public static bool CheckDate(String date)
        {
            try
            {
                DateTime iDateTIme = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static SAPbouiCOM.Item GetItem(string name, string form = "")
        {
            try
            {
                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Form oForm;

                if (form == "")
                {
                    oForm = Application.SBO_Application.Forms.ActiveForm;
                }
                else
                {
                    oForm = Application.SBO_Application.Forms.GetForm(form, 1);
                }

                return oItem = oForm.Items.Item(name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static void releaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion

        #region "Setting Value"
        public static void GetServerCredential()
        {
            try
            {
                oRS = (Recordset)myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oQuery = "SELECT * FROM \"@OCRED\" " + Environment.NewLine +
                "WHERE \"Code\" = '" + myCompany.CompanyDB + "'";
                oRS.DoQuery(Helpers.GlobalVar.oQuery);
                if (oRS.RecordCount > 0)
                {
                    oSERVERDB = myCompany.CompanyDB;
                    oSERVERTYPE = oRS.Fields.Item("U_ServerType").Value.ToString();
                    oSERVERNAME = oRS.Fields.Item("U_ServerName").Value.ToString();
                    oSERVERUN = oRS.Fields.Item("U_ServerUN").Value.ToString();
                    oSERVERPW = oRS.Fields.Item("U_ServerPW").Value.ToString();
                    oB1UN = oRS.Fields.Item("U_B1UN").Value.ToString();
                    oB1PW = oRS.Fields.Item("U_B1PW").Value.ToString();
                }
            }
            catch (Exception)
            { }
        }

        public static string GetPath()
        {
            string oPath = string.Empty;
            Helpers.GlobalVar.oRSQuery = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            Helpers.GlobalVar.oQuery = "SELECT * FROM " + Helpers.GlobalVar.Iif(Helpers.GlobalVar.myCompany.DbServerType == BoDataServerTypes.dst_HANADB, "\"BINLocationAutoSelect\".\"axxis_tb_TextFilePath\"", "\"BINLocationAutoSelect\"..\"axxis_tb_TextFilePath\"");
            Helpers.GlobalVar.oRSQuery.DoQuery(Helpers.GlobalVar.oQuery);

            if (Helpers.GlobalVar.oRSQuery.RecordCount > 0)
            { oPath = Helpers.GlobalVar.oRSQuery.Fields.Item("Path").Value.ToString(); }
            return oPath;
        }

        public static bool CreateFolderAutomatically(string oPath = "")
        {
            try
            {
                string iPath = oPath;
                _Contains = false;

                if (iPath == "")
                {
                    Helpers.GlobalVar.oRSQuery = (Recordset)myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Helpers.GlobalVar.oQuery = "SELECT * FROM " + Helpers.GlobalVar.Iif(Helpers.GlobalVar.myCompany.DbServerType == BoDataServerTypes.dst_HANADB, "\"BINLocationAutoSelect\".\"axxis_tb_TextFilePath\"", "\"BINLocationAutoSelect\"..\"axxis_tb_TextFilePath\"") + "";
                    Helpers.GlobalVar.oRSQuery.DoQuery(Helpers.GlobalVar.oQuery);

                    if (Helpers.GlobalVar.oRSQuery.RecordCount > 0)
                    {
                        iPath = Helpers.GlobalVar.oRSQuery.Fields.Item("Path").Value.ToString();
                    }
                }

                if (iPath[iPath.Length - 1].ToString() == @"\")
                    _Contains = true;

                if (!string.IsNullOrEmpty(iPath))
                {
                    if (Directory.Exists(iPath))
                    {
                        Helpers.GlobalVar.oRSQuery = (Recordset)myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        Helpers.GlobalVar.oQuery = "SELECT * FROM " + Helpers.GlobalVar.Iif(Helpers.GlobalVar.myCompany.DbServerType == BoDataServerTypes.dst_HANADB, "\"BINLocationAutoSelect\".\"axxis_vw_DTWDocuments\"", "\"BINLocationAutoSelect\"..\"axxis_vw_DTWDocuments\" WHERE \"Code\" != ''");
                        Helpers.GlobalVar.oRSQuery.DoQuery(Helpers.GlobalVar.oQuery);

                        while (!Helpers.GlobalVar.oRSQuery.EoF)
                        {
                            string oDocumentType = Helpers.GlobalVar.oRSQuery.Fields.Item("Name").Value.ToString();
                            string completePath = Iif(_Contains == true, iPath, iPath + @"\") + myCompany.CompanyDB + @"\" + oDocumentType;

                            string iSuccessFolder = completePath + @"\" + "Success";
                            if (!Directory.Exists(iSuccessFolder))
                            {
                                Application.SBO_Application.SetStatusBarMessage("Creating folder(Success) for Document:" + oDocumentType + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                Directory.CreateDirectory(iSuccessFolder);
                            }

                            string iFailedFolder = completePath + @"\" + "Failed";
                            if (!Directory.Exists(iFailedFolder))
                            {
                                Application.SBO_Application.SetStatusBarMessage("Creating folder(Failed) for Document:" + oDocumentType + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                Directory.CreateDirectory(iFailedFolder);
                            }

                            string iLogFolder = completePath + @"\" + "Log";
                            if (!Directory.Exists(iLogFolder))
                            {
                                Application.SBO_Application.SetStatusBarMessage("Creating folder(Log) for Document:" + oDocumentType + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                Directory.CreateDirectory(iLogFolder);
                            }
                            Helpers.GlobalVar.oRSQuery.MoveNext();
                        }
                    }
                    return true;
                }
                else
                { return false; }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "OK");
                return false;
            }
        }

        public static bool IsSuperUser()
        {
            try
            {
                Boolean oSuperUser = false;
                oRS = (Recordset)myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oQuery = "SELECT \"SUPERUSER\" FROM \"OUSR\" " + Environment.NewLine +
                "WHERE \"USER_CODE\" = '" + myCompany.UserName + "'";
                oRS.DoQuery(Helpers.GlobalVar.oQuery);
                if (oRS.RecordCount > 0)
                {
                    if (oRS.Fields.Item("SUPERUSER").Value.ToString() == "Y")
                    { oSuperUser = true; }
                }
                return oSuperUser;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "OK");
                return false;
            }
        }

        public static void MoveFilesToFolder(string oOrigPath, string oMovePath, string oLogPath, string oExcelFile)
        {
            try
            {
                // Ensure that the target does not exist.
                if (File.Exists(oMovePath))
                    File.Delete(oMovePath);

                // Move the file.
                File.Move(oOrigPath, oMovePath);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "OK");
            }
        }

        public static void WriteLog(string FilePath, string ExcelFile, string msg)
        {
            try
            {
                string fullPath = FilePath + ExcelFile.Replace(".xlsx", "") + ".txt";
                using (StreamWriter writer = new StreamWriter(fullPath, false))
                {
                    if (msg == "----------------------------------------------------------------")
                    {
                        writer.WriteLine(msg);
                    }
                    else
                    {
                        writer.WriteLine(msg + " || timestamp: " + DateTime.Now.ToString("HH:mm:ss"));
                    }
                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        #endregion

        #region "Helpers"

        public static void GetTransactionInfo(string ObjType = "0", string ExcelFile = "")
        {
            try
            {
                switch (ObjType)
                {
                    case "13":
                        oQuery = "SELECT \"DocEntry\",\"DocNum\" FROM \"OINV\" WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "' AND \"CANCELED\" = 'N'";
                        break;
                    case "18":
                        oQuery = "SELECT \"DocEntry\",\"DocNum\" FROM \"OPCH\" WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "' AND \"CANCELED\" = 'N'";
                        break;
                    case "30":
                        oQuery = "SELECT \"TransId\",\"BaseRef\" FROM \"OJDT\" WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "'";
                        break;
                }
                oRSQuery.DoQuery(oQuery);
                while (!oRSQuery.EoF)
                {
                    switch (ObjType)
                    {
                        case "13":
                            DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Internal No: " + oRSQuery.Fields.Item("DocEntry").Value.ToString() + " and Document No: " + oRSQuery.Fields.Item("DocNum").Value.ToString() + " successfuly created in SAP Business One (A/R Invoice)." + Environment.NewLine;
                            break;
                        case "18":
                            DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Internal No: " + oRSQuery.Fields.Item("DocEntry").Value.ToString() + " and Document No: " + oRSQuery.Fields.Item("DocNum").Value.ToString() + " successfuly created in SAP Business One (A/P Invoice)." + Environment.NewLine;
                            break;
                        case "30":
                            DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Trans No: " + oRSQuery.Fields.Item("TransId").Value.ToString() + " and Origin No: " + oRSQuery.Fields.Item("BaseRef").Value.ToString() + " successfuly created in SAP Business One (Journal Entry)." + Environment.NewLine;
                            break;
                    }
                    oRSQuery.MoveNext();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateExcelFile(string ObjType = "0", string ExcelFile = "")
        {
            try
            {
                switch (ObjType)
                {
                    case "13":
                        oQuery = "SELECT \"DocEntry\",\"DocNum\" FROM \"OINV\" " + Environment.NewLine +
                        "WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "' AND \"CANCELED\" = 'N'";
                        break;
                    case "18":
                        oQuery = "SELECT \"DocEntry\",\"DocNum\" FROM \"OPCH\" " + Environment.NewLine +
                        "WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "' AND \"CANCELED\" = 'N'";
                        break;
                    case "30":
                        oQuery = "SELECT \"TransId\",\"BaseRef\" FROM \"OJDT\" " + Environment.NewLine +
                        "WHERE \"U_ExcelFile\" = '" + TrimData(ExcelFile) + "'";
                        break;
                }
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount > 0)
                {
                    while (!oRSQuery.EoF)
                    {
                        switch (ObjType)
                        {
                            case "13":
                                DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Internal No: " + oRSQuery.Fields.Item("DocEntry").Value.ToString() + " and Document No: " + oRSQuery.Fields.Item("DocNum").Value.ToString() + " already exist in SAP Business One (A/R Invoice)." + Environment.NewLine;
                                break;
                            case "18":
                                DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Internal No: " + oRSQuery.Fields.Item("DocEntry").Value.ToString() + " and Document No: " + oRSQuery.Fields.Item("DocNum").Value.ToString() + " already exist in SAP Business One (A/P Invoice)." + Environment.NewLine;
                                break;
                            case "30":
                                DocMsgLine += "Excel File: " + TrimData(ExcelFile) + " with Trans No: " + oRSQuery.Fields.Item("TransId").Value.ToString() + " and Origin No: " + oRSQuery.Fields.Item("BaseRef").Value.ToString() + " already exist in SAP Business One (Journal Entry)." + Environment.NewLine;
                                break;
                        }
                        oRSQuery.MoveNext();
                    }
                }
                return DocMsgLine;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateSeries(string ObjType = "0", string Series = "")
        {
            try
            {
                string oSeries = string.Empty;
                oQuery = "SELECT \"Series\",\"SeriesName\" FROM \"NNM1\" WHERE \"ObjectCode\" = '" + TrimData(ObjType) + "' AND \"Series\" = '" + Series + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oSeries = "Series: " + Series + " does not exist in SAP Business One (Document Numbering - Setup)."; }
                else
                    oSeries = oRSQuery.Fields.Item("Series").Value.ToString();

                return oSeries;
            }
            catch (Exception ex)
            {
                return "";
                throw ex;
            }
        }

        public static string ValidateBP(string CardCode = "", string CardType = "")
        {
            try
            {
                string iCardType = string.Empty;
                oQuery = "SELECT \"CardCode\",\"CardType\" FROM \"OCRD\" " + Environment.NewLine +
                 "WHERE \"CardCode\" = '" + TrimData(CardCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    if (CardType == "C")
                    { iCardType = "Customer: "; }
                    else if (CardType == "S")
                    { iCardType = "Supplier: "; }

                    return iCardType + CardCode + " does not exist in SAP Business One (Business Partner Master Data - Setup).";
                }
                else
                    return oRSQuery.Fields.Item("CardCode").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateItem(string ItemCode = "")
        {
            try
            {
                oQuery = "SELECT \"ItemCode\" FROM \"OITM\" " + Environment.NewLine +
                "WHERE \"ItemCode\" = '" + TrimData(ItemCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    return "Item Code: " + ItemCode + " does not exist in SAP Business One (Item Master Data).";
                }
                else
                    return oRSQuery.Fields.Item("ItemCode").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateGLAccount(string AcctCode = "")
        {
            try
            {
                oQuery = "SELECT \"AcctCode\" FROM \"OACT\" " + Environment.NewLine +
                "WHERE \"AcctCode\" = '" + TrimData(AcctCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    return "Account Code: " + AcctCode + " does not exist in SAP Business One (Chart of Accounts).";
                }
                else
                    return oRSQuery.Fields.Item("AcctCode").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateTaxGroup(string VatGroup = "")
        {
            try
            {
                oQuery = "SELECT \"Code\" FROM \"OVTG\" " + Environment.NewLine +
                "WHERE \"Code\" = '" + TrimData(VatGroup) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    return "Tax Group: " + VatGroup + " does not exist in SAP Business One (Tax Groups - Setup).";
                }
                else
                    return oRSQuery.Fields.Item("Code").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateProject(string ProjectCode = "")
        {
            try
            {
                oQuery = "SELECT \"PrjCode\" FROM \"OPRJ\" " + Environment.NewLine +
                "WHERE \"PrjCode\" = '" + TrimData(ProjectCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    return "Project Code: " + ProjectCode + " does not exist in SAP Business One (Projects - Setup).";
                }
                else
                    return oRSQuery.Fields.Item("PrjCode").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ValidateWarehouse(string Warehouse = "")
        {
            string oWarehouse = string.Empty;
            try
            {
                oQuery = "SELECT \"WhsCode\" FROM \"OWHS\" " + Environment.NewLine +
                "WHERE \"WhsCode\" = '" + TrimData(Warehouse) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oWarehouse = "Warehouse: " + Warehouse + " does not exist in SAP Business One (Warehouse - Setup)."; }
                else
                { oWarehouse = oRSQuery.Fields.Item("WhsCode").Value.ToString(); }

                return oWarehouse;
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return oWarehouse; }
        }

        public static string ValidateSalePerson(string SlpCode = "")
        {
            string oSlpCode = "0";
            try
            {
                oQuery = "SELECT \"SlpCode\" FROM \"OSLP\" " + Environment.NewLine +
                "WHERE \"SlpCode\" = '" + TrimData(SlpCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oSlpCode = "Sales Person: " + SlpCode + " does not exist in SAP Business One (Sales Employee/Buyers - Setup)."; }
                else
                { oSlpCode = oRSQuery.Fields.Item("SlpCode").Value.ToString(); }

                return oSlpCode;
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return SlpCode.ToString(); }
        }

        public static string ValidateContactPerson(string CardCode = "", string CntctCode = "")
        {
            string oCntctCode = "0";
            try
            {
                oQuery = "SELECT \"CntctCode\",\"CardCode\" FROM \"OCPR\" " + Environment.NewLine +
                "WHERE \"CardCode\" = '" + TrimData(CardCode) + "' AND \"CntctCode\" = '" + TrimData(CntctCode) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oCntctCode = "Contact Person Code: " + CntctCode + " does not exist in SAP Business One (Contact Person)."; }
                else
                { oCntctCode = oRSQuery.Fields.Item("CntctCode").Value.ToString(); }

                return oCntctCode;
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return oCntctCode.ToString(); }
        }

        public static string ValidateCostCenter(string CostCenter = "")
        {
            string oCostCenter = string.Empty;
            try
            {
                oQuery = "SELECT \"PrcCode\" FROM \"OPRC\" " + Environment.NewLine +
                "WHERE \"PrcCode\" = '" + TrimData(CostCenter) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oCostCenter = "Cost Center:" + CostCenter + " does not exist in SAP Business One (Cost Center - Setup)."; }
                else
                { oCostCenter = oRSQuery.Fields.Item("PrcCode").Value.ToString(); }

                return oCostCenter;
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return oCostCenter; }
        }

        public static string ValidateDistributionRules(string OcrCode = "", int DimCode = 0)
        {
            string oOcrCode = string.Empty;
            try
            {
                oQuery = "SELECT \"OcrCode\" FROM \"OOCR\" " + Environment.NewLine +
                "WHERE \"OcrCode\" = '" + TrimData(OcrCode) + "' AND \"DimCode\" = " + DimCode + "";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { oOcrCode = "Distribution Rule:" + OcrCode + " does not exist in SAP Business One (Distribution Rules - Setup)."; }
                else
                { oOcrCode = oRSQuery.Fields.Item("OcrCode").Value.ToString(); }

                return oOcrCode;
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return oOcrCode; }
        }

        public static string ValidateUoM(string UomEntry = "")
        {
            int iUomEntry = 0;
            try
            {
                oQuery = "SELECT \"UomEntry\" FROM \"OUOM\" " + Environment.NewLine +
                "WHERE \"UomEntry\" = '" + TrimData(UomEntry) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                { return "Uom Entry: " + UomEntry + " does not exist in SAP Business One (Unit of Measure Setup)."; }
                else
                { iUomEntry = Convert.ToInt16(oRSQuery.Fields.Item("UomEntry").Value.ToString()); }
                return UomEntry.ToString();
            }
            catch (Exception ex)
            { Console.WriteLine(ex.ToString()); return UomEntry.ToString(); }
        }

        public static string ValidatePayTerms(string GroupNum = "")
        {
            try
            {
                oQuery = "SELECT \"GroupNum\" FROM \"OCTG\" " + Environment.NewLine +
                 "WHERE \"GroupNum\" = '" + TrimData(GroupNum) + "'";
                oRSQuery.DoQuery(oQuery);
                if (oRSQuery.RecordCount == 0)
                {
                    return "GroupNum: " + GroupNum + " does not exist in SAP Business One (Payment Terms).";
                }
                else
                    return oRSQuery.Fields.Item("GroupNum").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string ConvertToSAPDate(string Date)
        {
            string month, day, year;
            year = Date.Substring(0, 4);
            month = Date.Substring(4, 2);
            day = Date.Substring(6, 2);
            return day + "/" + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToInt16(month)) + "/" + year;
        }

        public static string ConvertToValidDate(string val)
        {
            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            int ctr = 0;

            string phrase = val;
            string[] words = phrase.Split('.');

            foreach (var word in words)
            {
                string ival = word.ToString();

                if (ctr == 0)
                { day = ival; }
                else if (ctr == 1)
                { month = ival; }
                else if (ctr == 2)
                { year = ival.Substring(0, 4); }
                ctr++;
            }
            return month + "." + day + "." + year;
        }

        public static string GetValidDate(string val)
        {
            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            int ctr = 0;

            string phrase = val;
            string[] words = phrase.Split('.');

            foreach (var word in words)
            {
                string ival = word.ToString();

                if (ctr == 0)
                { day = ival; }
                else if (ctr == 1)
                { month = ival; }
                else if (ctr == 2)
                { year = ival.Substring(0, 4); }
                ctr++;
            }
            return year + month + day;
        }

        public static string GetSubMenuID(string menuID, string subMenuDesc)
        {
            SAP_Helper.SAP _oSAP = new SAP_Helper.SAP();
            return _oSAP.GetSubMenuID_UserDefinedWindows(menuID, subMenuDesc);
        }

        public static void callBinReceipt(string frmParentType, int frmParentCnt, int mtxRow)
        {
            _frmAutoPop = true;
            _frmGrpo = Application.SBO_Application.Forms.GetForm(frmParentType, frmParentCnt);
            SAPbouiCOM.Matrix mtxContents = ((SAPbouiCOM.Matrix)(_frmGrpo.Items.Item("38").Specific));
            SAPbouiCOM.Column ColBin = (SAPbouiCOM.Column)(mtxContents.Columns.Item("1470002149"));
            ColBin.Cells.Item(mtxRow).Click(SAPbouiCOM.BoCellClickType.ct_Linked,0);            
        }

        public static void PopBinReceipt(string frmParentType, int frmParentCnt, string bincode, string qty)
        {
            //frmBINAllocationReceipt _frmBinReceipt = new frmBINAllocationReceipt();
            //_frmBinReceipt.UIAPIRawForm.VisibleEx = true;
            //SAPbouiCOM.Button btnOk = ((SAPbouiCOM.Button)(_frmBinReceipt.UIAPIRawForm.Items.Item("1470000001").Specific));
            //SAPbouiCOM.Matrix mxBinRcpt = ((SAPbouiCOM.Matrix)(_frmBinReceipt.UIAPIRawForm.Items.Item("1470000019").Specific));
            SAPbouiCOM.Form _frmBinReceipt = Application.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Button btnOk = ((SAPbouiCOM.Button)(_frmBinReceipt.Items.Item("1470000001").Specific));
            SAPbouiCOM.Matrix mxBinRcpt = ((SAPbouiCOM.Matrix)(_frmBinReceipt.Items.Item("1470000019").Specific));
            SAPbouiCOM.Column mxColBinCode = (SAPbouiCOM.Column)(mxBinRcpt.Columns.Item("1470000001")); //BinCode
            SAPbouiCOM.Column mxColBinQty = (SAPbouiCOM.Column)(mxBinRcpt.Columns.Item("1470000003"));  //BinQty
            SAPbouiCOM.EditText txtBinCode = (SAPbouiCOM.EditText)(mxColBinCode.Cells.Item(1).Specific);
            SAPbouiCOM.EditText txtBinQty = (SAPbouiCOM.EditText)(mxColBinQty.Cells.Item(1).Specific);

            txtBinCode.Value = bincode;
            txtBinQty.Value = qty;

            if (btnOk.Caption.ToUpper() == "OK" && _frmAutoPop)
            {
                btnOk.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular); //caption ok only
            }
            else if (btnOk.Caption == "Update" && _frmAutoPop)
            {
                btnOk.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);//caption update
                btnOk.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular); //caption ok
            }

            _frmAutoPop = false;
            //_frmBinReceipt.Update();
            //_frmBinReceipt.SupportedModes = 1;
            //_frmBinReceipt.Close();
        }

        #endregion

        #region "SBO Class"
        public static Boolean ConnectDI(string oServerType, string oServerName, string oServerUN, string oServerPW, string oCompanyDB, string oLicenseServer, string oB1Name, string oB1Password)
        {
            iCompany = new SAPbobsCOM.Company();

            switch (oServerType)
            {
                case "dst_MSSQL2005":
                    iCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                    break;
                case "dst_MSSQL2008":
                    iCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                    break;
                case "dst_MSSQL2012":
                    iCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                    break;
                case "dst_MSSQL2014":
                    iCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                    break;
                case "dst_HANADB":
                    iCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                    break;
            }

            iCompany.Server = oServerName;
            iCompany.CompanyDB = oCompanyDB;
            iCompany.DbUserName = oServerUN;
            iCompany.DbPassword = oServerPW;
            iCompany.UserName = oB1Name;
            iCompany.Password = oB1Password;
            iCompany.UseTrusted = false;

            if (iCompany.Connect() != 0)
            {
                Application.SBO_Application.SetStatusBarMessage(iCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            else
            {
                if (iCompany.Connected == true)
                {
                    iCompany.Disconnect();
                    return true;
                }
            }
            return false;
        }

        public static void CreateJournalEntry(string TextFileName = "", string DocEntry = "", string DocNum = "", string ReceiptNo = "", string DocDate = "")
        {
            try
            {
                oRSDtl = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oQuery = "select \"x\".* from( " + Environment.NewLine +
                "select \"DocEntry\",\"DocNum\",\"DocTotal\", \"U_AccountCode\",\"U_ReceiptNo\", " + Environment.NewLine +
                "(select top 1 \"U_AccountA\" from \"@ARINVGLACCT\") \"AcctCode\", " + Environment.NewLine +
                "\"DocTotal\" \"Debit\", " + Environment.NewLine +
                "0.00 \"Credit\" " + Environment.NewLine +
                "from \"ORCT\" " + Environment.NewLine +
                "where \"Canceled\" = 'N' and convert(nvarchar(max),\"U_TextFileName\") = '" + TextFileName + "' " + Environment.NewLine +
                "and \"DocEntry\" = '" + DocEntry + "' and \"U_ReceiptNo\" = '" + ReceiptNo + "' " + Environment.NewLine +
                "group by \"DocEntry\",\"DocNum\",\"DocTotal\",\"U_AccountCode\",\"U_ReceiptNo\" " + Environment.NewLine +
                "union all " + Environment.NewLine +
                "select \"DocEntry\",\"DocNum\",\"DocTotal\", \"U_AccountCode\",\"U_ReceiptNo\", " + Environment.NewLine +
                "(select top 1 \"U_AccountS\" from \"@ARINVGLACCT\") \"AcctCode\", " + Environment.NewLine +
                "0.00 \"Debit\", " + Environment.NewLine +
                "\"DocTotal\" \"Credit\" " + Environment.NewLine +
                "from \"ORCT\" " + Environment.NewLine +
                "where \"Canceled\" = 'N' and convert(nvarchar(max),\"U_TextFileName\") = '" + TextFileName + "' " + Environment.NewLine +
                "and \"DocEntry\" = '" + DocEntry + "' and \"U_ReceiptNo\" = '" + ReceiptNo + "' " + Environment.NewLine +
                "group by \"DocEntry\",\"DocNum\",\"DocTotal\",\"U_AccountCode\",\"U_ReceiptNo\") \"x\" " + Environment.NewLine +
                "order by \"x\".\"DocEntry\"";
                oRSDtl.DoQuery(oQuery);

                int Ctr = 0;
                for (int i = 0; i < oRSDtl.RecordCount; i++)
                {
                    if (Ctr == 0)
                    {
                        Application.SBO_Application.SetStatusBarMessage("Please wait creating Journal Entry for Incoming Payment:" + DocNum + " with Receipt No.:" + ReceiptNo + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        oJounalEntry = (SAPbobsCOM.JournalEntries)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        oJounalEntry.ReferenceDate = Convert.ToDateTime(DocDate);
                        oJounalEntry.TaxDate = Convert.ToDateTime(DocDate);
                        oJounalEntry.Memo = "Auto created JE for Incoming Payment No.:" + DocNum;
                    }

                    string AcctCode = oRSDtl.Fields.Item("AcctCode").Value.ToString();
                    string iDebit = oRSDtl.Fields.Item("Debit").Value.ToString();
                    string iCredit = oRSDtl.Fields.Item("Credit").Value.ToString();

                    if (iDebit == "")
                        iDebit = "0";

                    if (iCredit == "")
                        iCredit = "0";

                    double Debit = Convert.ToDouble(iDebit);
                    double Credit = Convert.ToDouble(iCredit);

                    if (AcctCode != "")
                        oJounalEntry.Lines.AccountCode = AcctCode;

                    if (Debit != 0)
                        oJounalEntry.Lines.Debit = Debit;

                    if (Credit != 0)
                        oJounalEntry.Lines.Credit = Credit;

                    if (Ctr == 0)
                        oJounalEntry.Lines.Add();

                    Ctr += 1;
                    oRSDtl.MoveNext();
                }

                int lRetCode = oJounalEntry.Add();
                if (lRetCode == 0)
                {
                    string oJENo = myCompany.GetNewObjectKey();
                    oPayments = (Payments)myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                    if (Helpers.GlobalVar.oPayments.GetByKey(Convert.ToInt32(DocEntry)))
                    {
                        oPayments.UserFields.Fields.Item("U_JENo").Value = oJENo;
                        int oRetCode = oPayments.Update();
                        if (oRetCode != 0)
                        {
                            Application.SBO_Application.SetStatusBarMessage("Error updating Incoming Payment No: " + DocNum + " with Receipt No.:" + ReceiptNo + ", " + myCompany.GetLastErrorDescription() + ".", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage("Successfully created Journal Entry No." + oJENo + " for Incoming Payment No: " + DocNum + " with Receipt No.:" + ReceiptNo + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                }
                else
                {
                    Application.SBO_Application.SetStatusBarMessage("Error creating Journal Entry for Incoming Payment No: " + DocNum + " with Receipt No.:" + ReceiptNo + ", " + myCompany.GetLastErrorDescription() + ".", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }
        #endregion
    }

    public static class SBOstrManipulation
    {
        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string BeforeCharacter(this string value, string a)
        {
            int posA = value.IndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }

        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string AfterCharacter(this string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }
    }



}