using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace BIN_Location_Auto_Select_Addon.Forms
{
    [FormAttribute("1470000007", "Forms/frmBINAllocationIssue.b1f")]
    class frmBINAllocationIssue : SystemFormBase
    {

        #region Variables
        public static string Series { get; set; }
     
        #endregion Variables

        #region SAP Objects
        public SAPbouiCOM.EditText edDocNum { get; set; }
        public SAPbouiCOM.EditText edRow { get; set; }
        public SAPbouiCOM.EditText edTotAlloc { get; set; }
        public SAPbouiCOM.ComboBox cbAutoAlloc { get; set; }
        public SAPbouiCOM.Button btTrans { get; set; }
        public SAPbouiCOM.Matrix mxAlloc { get; set; }
        #endregion SAP Objects

        public frmBINAllocationIssue()
        {
            try
            {

            }
            catch (Exception ex)
            { throw ex; }
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            try
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                string DocNum = string.Empty;
                string RowNum = string.Empty;
                string BINLocation = string.Empty;
                string Allocated = string.Empty;

                this.edDocNum = ((SAPbouiCOM.EditText)(this.GetItem("1470000010").Specific));
                this.edRow = ((SAPbouiCOM.EditText)(this.GetItem("1980000024-1980000004").Specific));
                this.edTotAlloc = ((SAPbouiCOM.EditText)(this.GetItem("1470000015").Specific));
                this.cbAutoAlloc = ((SAPbouiCOM.ComboBox)(this.GetItem("1470000022").Specific));
                this.btTrans = ((SAPbouiCOM.Button)(this.GetItem("1470000001").Specific));
                this.mxAlloc = ((SAPbouiCOM.Matrix)(this.GetItem("1470000023").Specific));

                DocNum = this.edDocNum.Value;
                RowNum = this.edRow.Value;
                Series = BINLocationAutoSelect.Program.SelectedSeries;

                if (this.edTotAlloc.Value == "0.0")
                    this.cbAutoAlloc.Select("FIFO", SAPbouiCOM.BoSearchKey.psk_ByValue);

                if (btTrans.Caption == "Update" && this.edTotAlloc.Value != "0.0")
                    btTrans.Item.Click();

                if (mxAlloc.RowCount > 0)
                {
                    for (int i = 1; i <= mxAlloc.RowCount; i++)
                    {
                        try
                        {
                            Allocated = ((SAPbouiCOM.EditText)(mxAlloc.GetCellSpecific("1470000005", i))).Value;

                            if (Allocated != "0.0")
                                BINLocation = ((SAPbouiCOM.EditText)(mxAlloc.GetCellSpecific("1470000001", i))).Value;
                        }
                        catch { }
                    }
                    try
                    {
                        BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "DELETE FROM \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                        "WHERE \"CompanyDB\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "' " + Environment.NewLine +
                        "AND \"UserSign\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "' " + Environment.NewLine +
                        "AND \"Series\" = '" + Series + "' " + Environment.NewLine +
                        "AND \"DocNum\" = '" + DocNum + "' " + Environment.NewLine +
                        "AND \"RowNum\" = '" + RowNum + "' " + Environment.NewLine +
                        "AND \"ObjType\" = '15'";
                            BINLocationAutoSelect.Helpers.GlobalVar.oRSExec.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                            BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "INSERT INTO \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                            "(\"CompanyDB\",\"UserSign\",\"Series\",\"DocNum\",\"BINCode\",\"RowNum\",\"ObjType\") VALUES " + Environment.NewLine +
                            "('" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "', " + Environment.NewLine +
                            "'" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "', " + Environment.NewLine +
                            "'" + Series + "', " + Environment.NewLine +
                            "'" + DocNum + "', " + Environment.NewLine +
                            "'" + BINLocation + "', " + Environment.NewLine +
                            "'" + RowNum + "', " + Environment.NewLine +
                            "'15')";
                        BINLocationAutoSelect.Helpers.GlobalVar.oRSExec.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);
                    }
                    catch
                    { }
                    
                }

                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);

                //if (btTrans.Caption == "OK" && this.edTotAlloc.Value != "0.0")
                //    btTrans.Item.Click();
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                throw ex;
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            //this.UnloadAfter += new SAPbouiCOM.Framework.FormBase.UnloadAfterHandler(this.Form_UnloadAfter);
            //this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            //this.DataAddBefore += new DataAddBeforeHandler(this.Form_DataAddBefore);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_Resize);
        }

        private void Form_Resize(SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private void Form_UnloadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //try
            //{
            //    BINLocation = string.Empty;
            //    mxRow = string.Empty;
            //}
            //catch (Exception ex)
            //{ throw ex; }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //try
            //{
            //    for (int i = 1; i < mxContents.RowCount; i++)
            //    {
            //        mxAxcBin = (SAPbouiCOM.EditText)mxAxcColBin.Cells.Item(i).Specific;
            //        mxAxcBin.Value = "";
            //    }
            //}
            //catch (Exception)
            //{ //throw ex; 
            //}

        }


        private void Form_DataAddBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;

            try
            {
                if ((btTrans.Caption == "N" || btTrans.Caption == "V" || btTrans.Caption == "C") && !Application.SBO_Application.Forms.ActiveForm.Title.Contains("Cancellation"))
                {
                    string DocNum = string.Empty;
                    string RowNum = string.Empty;
                    string BINLocation = string.Empty;
                    string Allocated = string.Empty;
                    //BINLocation = string.Empty;
                    //mxRow = string.Empty;
                    this.edDocNum = ((SAPbouiCOM.EditText)(this.GetItem("1470000010").Specific));
                    this.edRow = ((SAPbouiCOM.EditText)(this.GetItem("1980000024-1980000004").Specific));
                    this.edTotAlloc = ((SAPbouiCOM.EditText)(this.GetItem("1470000015").Specific));
                    this.cbAutoAlloc = ((SAPbouiCOM.ComboBox)(this.GetItem("1470000022").Specific));
                    this.btTrans = ((SAPbouiCOM.Button)(this.GetItem("1470000001").Specific));
                    this.mxAlloc = ((SAPbouiCOM.Matrix)(this.GetItem("1470000023").Specific));

                    DocNum = this.edDocNum.Value;
                    RowNum = this.edRow.Value;


                    //Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                    //for (int i = 1; i < mxContents.RowCount; i++)
                    //{
                    //    Series = this.cbSeries.Selected.Value;
                    //    DocNum = this.edDocNum.Value;

                    //    mxColBin.Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Linked, 0);
                    //    BIN_Location_Auto_Select_Addon.Forms.frmBINAllocationIssue.Series = Series;

                    //    BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "SELECT * FROM \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                    //      "WHERE \"CompanyDB\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "' " + Environment.NewLine +
                    //      "AND \"UserSign\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "' " + Environment.NewLine +
                    //      "AND \"Series\" = '" + Series + "' " + Environment.NewLine +
                    //      "AND \"DocNum\" = '" + DocNum + "' " + Environment.NewLine +
                    //      "AND \"ObjType\" = '15'";
                    //    BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                    //    if (BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.RecordCount > 0)
                    //    {
                    //        while (!BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.EoF)
                    //        {
                    //            mxRow = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("RowNum").Value.ToString();
                    //            BINLocation = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("BINCode").Value.ToString();

                    //            if (!string.IsNullOrEmpty(BINLocation) && !string.IsNullOrEmpty(mxRow))
                    //            {
                    //                ((SAPbouiCOM.EditText)(mxContents.GetCellSpecific("U_AXC_BINLocation", Convert.ToInt16(mxRow)))).Value = BINLocation;
                    //            }

                    //            BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.MoveNext();
                    //        }
                    //    }
                    //}
                    //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                    if (mxAlloc.RowCount > 0)
                    {
                        for (int i = 1; i <= mxAlloc.RowCount; i++)
                        {
                            try
                            {
                                Allocated = ((SAPbouiCOM.EditText)(mxAlloc.GetCellSpecific("1470000005", i))).Value;

                                if (Allocated != "0.0")
                                    BINLocation = ((SAPbouiCOM.EditText)(mxAlloc.GetCellSpecific("1470000001", i))).Value;
                            }
                            catch { }
                        }
                        try
                        {
                            BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "DELETE FROM \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                            "WHERE \"CompanyDB\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "' " + Environment.NewLine +
                            "AND \"UserSign\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "' " + Environment.NewLine +
                            "AND \"Series\" = '" + Series + "' " + Environment.NewLine +
                            "AND \"DocNum\" = '" + DocNum + "' " + Environment.NewLine +
                            "AND \"RowNum\" = '" + RowNum + "' " + Environment.NewLine +
                            "AND \"ObjType\" = '15'";
                            BINLocationAutoSelect.Helpers.GlobalVar.oRSExec.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                            BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "INSERT INTO \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                            "(\"CompanyDB\",\"UserSign\",\"Series\",\"DocNum\",\"BINCode\",\"RowNum\",\"ObjType\") VALUES " + Environment.NewLine +
                            "('" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "', " + Environment.NewLine +
                            "'" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "', " + Environment.NewLine +
                            "'" + Series + "', " + Environment.NewLine +
                            "'" + DocNum + "', " + Environment.NewLine +
                            "'" + BINLocation + "', " + Environment.NewLine +
                            "'" + RowNum + "', " + Environment.NewLine +
                            "'15')";
                            BINLocationAutoSelect.Helpers.GlobalVar.oRSExec.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);
                        }
                        catch
                        { }

                    }
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                throw ex;
            }
        }

    }
}
