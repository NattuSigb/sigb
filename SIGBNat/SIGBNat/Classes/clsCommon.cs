using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using SIGBNat;
//using DepartmentWiseWharehouseManagement;
using General.Extensions;
using System.Globalization;
//using DepartmentWiseWharehouseManagement.Classes;
using static General.Classes.@enum;
using SAPbouiCOM;
namespace General.Classes
{
    class clsCommon : Connection
    {
        #region ConstruBPNamer
        public clsCommon()
        {
        }
        #endregion

        #region Methods

        #region AddMenu

        public void AddMenu(SAPbouiCOM.BoMenuType boMenuType, string FatherID, string UniqueID, string Name, string imageName, int Position)
        {
            try
            {
                SAPbouiCOM.Menus oMenus = null;
                SAPbouiCOM.MenuItem oMenuItem = null;

                oMenus = oApplication.Menus;

                SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams));
                oMenuItem = oApplication.Menus.Item(FatherID);
                oCreationPackage.Type = boMenuType;
                oCreationPackage.UniqueID = UniqueID;
                oCreationPackage.String = Name;
                oCreationPackage.Position = Position;
                oCreationPackage.Enabled = true;
                if (imageName != string.Empty)
                {
                    string imageFilePath = System.Windows.Forms.Application.StartupPath + "\\Images\\" + imageName + ".bmp";
                    oCreationPackage.Image = System.Windows.Forms.Application.StartupPath + "\\Images\\" + imageName + ".bmp";
                }
                oMenus = oMenuItem.SubMenus;
                if (oMenus.Exists(UniqueID) == false)
                    oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public void RemoveMenu(string menuUID)
        {
            try
            {
                SAPbouiCOM.Menus oMenus;
                oMenus = oApplication.Menus;
                if (oMenus.Exists(menuUID))
                {
                    oMenus.RemoveEx(menuUID);
                }
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        #endregion




        #region Database

        public void CreateDataBase()
        {

            SAPMain.logger.DebugFormat("> {0}", nameof(CreateDataBase));

            oApplication.StatusBar.SetText("Please Wait....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            CreateTables();
            CreateFields();
            CreateObjects();

        }


        private void CreateTables()
        {


            // #region DWHM
            #region BPIM
            CreateTable("BPIM", "BusinessPartnerItemMaster", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateTable("BPIM1", "BusinessPartnerItemMasterRows", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            #endregion
            CreateTable("CCBI", "customer header", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateTable("CCBI1", "customer row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            CreateTable("DocTyp", "Document Types", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            //CreateTable("DocTyp", "Document Types", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            /*
                        #region No Object
                        CreateTable("ITMT", "itemmanagement", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                        CreateTable("DEPT", "department", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                        #endregion
            */
        }

        public void CreateFields()
        {
            #region Standard Tables

           //FieldDetails(TableType.StandardTable, "ORDR", "U_Aproval", "AprovalDept", UDFFieldType.YN, false, 30, "N", "", "");

            #endregion

            #region Custom Tables
            FieldDetails(TableType.CustomTable, "BPIM", "BPCode", "BPCode", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM", "BPName", "BPName", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM", "DocNum", "DocNum", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM", "BPT", "BPT", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM", "BPIDocDate", "BPT", UDFFieldType.Date, false, 15, "", "", "");


            FieldDetails(TableType.CustomTable, "CCBI", "code", "BPCode", UDFFieldType.Alpha, false, 40, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "cname", "BPName", UDFFieldType.Alpha, false, 100, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "Type", "DocNum", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "Balance", "BPT", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "typ", "BPT", UDFFieldType.Alpha, false, 35, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "group", "BPT", UDFFieldType.Alpha, false, 45, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "DocNum", "DocNum", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "SDocDate", "SDocDate", UDFFieldType.Date, false, 15, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "DocTy", "DocumentType", UDFFieldType.Alpha, false, 40, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "DocStd", "Document Status", UDFFieldType.Alpha, false, 40, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI", "Rec", "Recount Count", UDFFieldType.Alpha, false, 40, "", "", "");

            FieldDetails(TableType.CustomTable, "BPIM1", "INo", "INo", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM1", "Dscription", "Dscription", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM1", "PPqty", "PPqty", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "BPIM1", "SPqty", "SalesPlanqty", UDFFieldType.Numeric, false, 30, "", "", "");

            FieldDetails(TableType.CustomTable, "DocTyp", "child", "child", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "DocTyp", "ObjT", "ObjT", UDFFieldType.Alpha, false, 30, "", "", "");

            FieldDetails(TableType.CustomTable, "CCBI1", "CIno", "CIno", UDFFieldType.Alpha, false, 70, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "CDesc", "CDesc", UDFFieldType.Alpha, false, 100, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "Cqty", "Cqty", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "Doc", "DocNum", UDFFieldType.Alpha, false, 100, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "Price", "Price", UDFFieldType.Alpha, false, 50, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "ItemType", "ItemType", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "UomCode1", "UomCode", UDFFieldType.Alpha, false, 30, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "whsCode1", "whsCode", UDFFieldType.Alpha, false, 50, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "ta", "TaxCode", UDFFieldType.Alpha, false, 50, "", "", "");
            FieldDetails(TableType.CustomTable, "CCBI1", "DocDate1", "DocDate", UDFFieldType.Alpha, false, 100, "", "", "");
            //FieldDetails(TableType.CustomTable, "CCBI1", "CDesc", "CDesc", UDFFieldType.Alpha, false, 50, "", "", "");
            //FieldDetails(TableType.CustomTable, "CCBI1", "Cqty", "Cqty", UDFFieldType.Alpha, false, 30, "", "", "");
            /*s
                        #region Custom Tables//
                        FieldDetails(TableType.CustomTable, "DWHM", "BPCode", "From", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "BPName", "TO", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "EDocNum", "DocNum", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "EDocDate", "DocDate", UDFFieldType.Date, false, 15, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "BPT", "BPT", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "BPT1", "BPT1", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "BPCode", "FromBin", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM", "BPT", "ToBin", UDFFieldType.Alpha, false, 30, "", "", "");

                        FieldDetails(TableType.CustomTable, "DWHM1", "ino", "itemno", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM1", "desc", "description", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM1", "quantity", "quantity", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM1", "ucode", "ucode", UDFFieldType.Alpha, false, 30, "", "", "");
                        FieldDetails(TableType.CustomTable, "DWHM1", "uname", "uname", UDFFieldType.Alpha, false, 30, "", "", "");//Avlstk
                        FieldDetails(TableType.CustomTable, "DWHM1", "Avlstk", "Avlstk", UDFFieldType.Alpha, false, 30, "", "", "");
                        //#region No Object

                         FieldDetails(TableType.CustomTable, "ITMT", "Document", "Document", UDFFieldType.Alpha, false, 30, "Y", "", "");
                        FieldDetails(TableType.CustomTable, "ITMT", "itemcode", "Itemcode", UDFFieldType.Alpha, false, 30, "Y", "", "");
                        FieldDetails(TableType.CustomTable, "ITMT", "quantity", "Quantity", UDFFieldType.Alpha, false, 10, "", "", "");
                        FieldDetails(TableType.CustomTable, "ITMT", "baseUOM", "BaseUOM",UDFFieldType.Alpha,false,10,"","","");

                        //  #endregion
*/
            #endregion

        }

        private void CreateObjects()
        {
            CreateUserObjectDocuments("BPIM", "BPIM", "BPIM", "BPIM1", "", "", "", false);
            CreateUserObjectDocuments("CCBI", "CCBI", "CCBI", "CCBI1", "", "", "", false);
            CreateUserObjectDocuments("DocTyp", "DocTyp", "DocTyp", "", "", "", "", false);
            /*           CreateUserObjectDocuments("BPIM", "BPIM", "BPIM", "BPIM1", "", "", "", false);
                       CreateUserObjectDocuments("DEPT", "DEPT", "DEPT", "", "", "", "", false);
                       CreateUserObjectDocuments("ITMT", "ITMT", "ITMT", "", "", "", "", false);
            */
        }




        Boolean CreateTable(string tableName, string tableDesc, SAPbobsCOM.BoUTBTableType tableType)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                if (oUserTablesMD == null)
                {
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                }
                if (oUserTablesMD.GetByKey(tableName) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserTablesMD.TableName = tableName;
                oUserTablesMD.TableDescription = tableDesc;
                oUserTablesMD.TableType = tableType;
                int err = oUserTablesMD.Add();
                string errMsg=" ";
                oCompany.GetLastError(out err, out errMsg);
                if (err == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                { 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        public Boolean FieldDetails(TableType tableType, string TableName, string FieldName, string FieldDesc, UDFFieldType FieldType, bool Mandatory, int FieldSize, string DefaultVal, string LinkedTable, string LinkedUDO)
        {
            string ErrMsg;
            int errCode;
            int IRetCode;
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            SAPbobsCOM.Recordset oRecordSet;

            try
            {
                oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlQuery;
                if (tableType == TableType.CustomTable)
                {
                    sqlQuery = string.Format("SELECT T0.\"TableID\",T0.\"FieldID\" FROM CUFD T0 WHERE T0.\"TableID\" = '@{0}' AND T0.\"AliasID\" = '{1}'", TableName, FieldName);
                }
                else
                {
                    sqlQuery = string.Format("SELECT T0.\"TableID\",T0.\"FieldID\" FROM CUFD T0 WHERE T0.\"TableID\" = '{0}' AND T0.\"AliasID\" = '{1}'", TableName, FieldName);
                }
                bool retflag = false;

                if (oRecordSet.RecordCount == 1)
                {
                    retflag = true;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (retflag == true) return true;

                if (oUserFieldsMD == null)
                {
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                }
                oUserFieldsMD.TableName = TableName;
                oUserFieldsMD.Name = FieldName;
                oUserFieldsMD.Description = FieldDesc;
                switch (FieldType)
                {
                    case UDFFieldType.Alpha:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;
                        if (LinkedTable != string.Empty)
                        {
                            oUserFieldsMD.LinkedTable = LinkedTable;
                        }
                        if (LinkedUDO != string.Empty)
                        {
                            oUserFieldsMD.LinkedUDO = LinkedUDO;
                        }
                        break;
                    case UDFFieldType.Text:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo; // alphanumeric type                   
                                                                              // oUserFieldsMD.EditSize = FieldSize;
                        break;
                    case UDFFieldType.Integer:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric; //  Integer type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
                        oUserFieldsMD.EditSize = FieldSize;
                        break;
                    case UDFFieldType.Date:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date; // Date type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
                        break;

                    case UDFFieldType.Time:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date; // Time type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time;
                        break;

                    case UDFFieldType.Amount:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float; // Amount type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum; // Amount type
                        break;
                    case UDFFieldType.Quantity:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float; // Amount type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity; // Amount type
                        break;
                    case UDFFieldType.Percent:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float; // Amount type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage; // Amount type
                        break;
                    case UDFFieldType.UnitTotal:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float; // Amount type
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement; // Amount type
                        break;

                    case UDFFieldType.Rate:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate;
                        break;

                    case UDFFieldType.Price:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
                        break;

                    case UDFFieldType.Link:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo; // alphanumeric type 
                        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link;
                        oUserFieldsMD.EditSize = FieldSize;
                        break;

                    case UDFFieldType.YN:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "Y";
                        oUserFieldsMD.ValidValues.Description = "Yes";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "N";
                        oUserFieldsMD.ValidValues.Description = "No";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.IssueMethod:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "M";
                        oUserFieldsMD.ValidValues.Description = "Manual";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "B";
                        oUserFieldsMD.ValidValues.Description = "Blackflush";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.BOMCompType:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "4";
                        oUserFieldsMD.ValidValues.Description = "Item";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "290";
                        oUserFieldsMD.ValidValues.Description = "Resource";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.Layer:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "Outer Layer";
                        oUserFieldsMD.ValidValues.Description = "Outer Layer";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "Middle Layer";
                        oUserFieldsMD.ValidValues.Description = "Middle Layer";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "Inner Layer";
                        oUserFieldsMD.ValidValues.Description = "Inner Layer";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.Status:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "O";
                        oUserFieldsMD.ValidValues.Description = "Open";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "C";
                        oUserFieldsMD.ValidValues.Description = "Cancel";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "L";
                        oUserFieldsMD.ValidValues.Description = "Close";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.QCStastus:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "P";
                        oUserFieldsMD.ValidValues.Description = "Pass";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "H";
                        oUserFieldsMD.ValidValues.Description = "Hold";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "R";
                        oUserFieldsMD.ValidValues.Description = "Reject";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "V";
                        oUserFieldsMD.ValidValues.Description = "Void";
                        oUserFieldsMD.ValidValues.Add();

                        break;


                    case UDFFieldType.Activity:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "D";
                        oUserFieldsMD.ValidValues.Description = "Deisgn";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "C";
                        oUserFieldsMD.ValidValues.Description = "Checked";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "A";
                        oUserFieldsMD.ValidValues.Description = "Approved";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.ActivityStatus:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "O";
                        oUserFieldsMD.ValidValues.Description = "Open";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "C";
                        oUserFieldsMD.ValidValues.Description = "Closed";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                    case UDFFieldType.RewStatus:
                        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                   
                        oUserFieldsMD.EditSize = FieldSize;

                        oUserFieldsMD.ValidValues.Value = "Completed";
                        oUserFieldsMD.ValidValues.Description = "Completed";
                        oUserFieldsMD.ValidValues.Add();

                        oUserFieldsMD.ValidValues.Value = "Pending";
                        oUserFieldsMD.ValidValues.Description = "Pending";
                        oUserFieldsMD.ValidValues.Add();

                        break;

                }
                if (Mandatory)
                {
                    oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;

                }
                if (DefaultVal != "")
                {
                    oUserFieldsMD.DefaultValue = DefaultVal;

                }

                // Add the field to the table
                IRetCode = oUserFieldsMD.Add();

                if (IRetCode != 0)
                {
                    oCompany.GetLastError(out errCode, out ErrMsg);
                }
                return true;
            }
            finally
            {
                if (oUserFieldsMD != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);

                oUserFieldsMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //return false ;
            }
        }

        private Boolean CreateUserObjectMaster(string CodeID, string Name, string TableName, string Child, string Child1, string Child2, Boolean DefaultForm)
        {
            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            if (oUserObjectMD == null)
                oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

            if (oUserObjectMD.GetByKey(CodeID) == true)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return true;
            }

            oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.Code = CodeID;
            oUserObjectMD.Name = Name;

            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.FindColumns.SetCurrentLine(0);
            oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
            oUserObjectMD.FindColumns.ColumnAlias = "Code";
            oUserObjectMD.FindColumns.ColumnDescription = "Code";
            oUserObjectMD.FindColumns.Add();
            oUserObjectMD.FindColumns.SetCurrentLine(1);
            oUserObjectMD.FindColumns.ColumnAlias = "DocNum";
            oUserObjectMD.FindColumns.ColumnDescription = "Doc Num";
            oUserObjectMD.FindColumns.Add();


            oUserObjectMD.TableName = TableName;

            if (Child != "")
            {
                oUserObjectMD.ChildTables.TableName = Child;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (Child1 != "")
            {
                oUserObjectMD.ChildTables.Add();
                oUserObjectMD.ChildTables.TableName = Child1;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (Child2 != "")
            {
                oUserObjectMD.ChildTables.Add();
                oUserObjectMD.ChildTables.TableName = Child2;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (DefaultForm == true)
            {
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            try
            {
                lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { }
                    else
                    { oCompany.GetLastError(out lRetCode, out sErrMsg); }
                else
                { }
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        private Boolean CreateUserObjectDocuments(string CodeID, string Name, string TableName, string Child, string Child1, string Child2, string Child3, Boolean DefaultForm)
        {
            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            if (oUserObjectMD == null)
                oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

            if (oUserObjectMD.GetByKey(CodeID) == true)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return true;
            }

            oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.Code = CodeID;
            oUserObjectMD.Name = Name;

            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.FindColumns.SetCurrentLine(0);
            oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
            oUserObjectMD.FindColumns.ColumnAlias = "DocEntry";
            oUserObjectMD.FindColumns.ColumnDescription = "DocEntry";
            oUserObjectMD.FindColumns.Add();
            oUserObjectMD.FindColumns.SetCurrentLine(1);
            oUserObjectMD.FindColumns.ColumnAlias = "DocNum";
            oUserObjectMD.FindColumns.ColumnDescription = "Doc Num";
            oUserObjectMD.FindColumns.Add();


            oUserObjectMD.TableName = TableName;

            if (Child != "")
            {
                oUserObjectMD.ChildTables.TableName = Child;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (Child1 != "")
            {
                oUserObjectMD.ChildTables.Add();
                oUserObjectMD.ChildTables.TableName = Child1;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (Child2 != "")
            {
                oUserObjectMD.ChildTables.Add();
                oUserObjectMD.ChildTables.TableName = Child2;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            if (Child3 != "")
            {
                oUserObjectMD.ChildTables.Add();
                oUserObjectMD.ChildTables.TableName = Child3;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }

            if (DefaultForm == true)
            {
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
            }

            lRetCode = oUserObjectMD.Add();

            oCompany.GetLastError(out lRetCode, out sErrMsg);
            oApplication.StatusBar.SetText("Feald : " + CodeID + " ErrorMsg : " + lRetCode + " : " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            // check for errors in the process
            if (lRetCode != 0)
                if (lRetCode == -1)
                { }
                else
                { oCompany.GetLastError(out lRetCode, out sErrMsg); }
            else
            { }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        #endregion

        #region Load Forms

        public void LoadXML(string XMLFile, string BrowseBy, string Title, SAPbouiCOM.BoFormMode boFormMode)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                string sPath = null;
                oXmlDoc = new System.Xml.XmlDocument();

                sPath = System.Windows.Forms.Application.StartupPath + "\\Forms\\" + XMLFile + ".xml";
                oXmlDoc.Load(sPath);

                Random r = new Random();
                r.Next(1000);

                if ((oXmlDoc.SelectSingleNode("//form") != null))
                {
                    oXmlDoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value =
                    oXmlDoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value.ToString() + "_" + r.Next().ToString();
                    string sXML = oXmlDoc.InnerXml.ToString();
                    oApplication.LoadBatchActions(ref sXML);

                    SAPbouiCOM.Form oForm = oApplication.Forms.ActiveForm;
                    if (Title != "")
                    {
                        oForm.Title = Title;
                    }
                    oForm.SupportedModes = -1;

                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FindRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.AddRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.RemoveRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.RestoreRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.CancelRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.CloseRecord), false);

                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FirstRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.LastRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.NextRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.PreviousRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.AddRow), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.DeleteRow), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FilterTable), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.UDFForm), false);

                    if (BrowseBy != "")
                    {
                        oForm.DataBrowser.BrowseBy = BrowseBy;
                    }
                    oForm.Mode = boFormMode;
                }
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + ">" + XMLFile + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + ">" + XMLFile + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public void Loadxml(string XMLFile, string BrowseBy, string Title, SAPbouiCOM.BoFormMode boFormMode)
        {

            try
            {
                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                string file = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".XMLForms." + XMLFile + ".xml";
                System.IO.Stream Streaming = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(file);
                System.IO.StreamReader StreamRead = new System.IO.StreamReader(Streaming, true);
                xmldoc.LoadXml(StreamRead.ReadToEnd());
                StreamRead.Close();
                Random r = new Random();
                r.Next(1000);

                if ((xmldoc.SelectSingleNode("//form") != null))
                {
                    xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value =
                    xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value.ToString() + "_" + r.Next().ToString();
                    string sXML = xmldoc.InnerXml.ToString();
                    oApplication.LoadBatchActions(ref sXML);

                    SAPbouiCOM.Form oForm = oApplication.Forms.ActiveForm;
                    if (Title != "")
                    {
                        oForm.Title = Title;
                    }
                    oForm.SupportedModes = -1;

                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FindRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.AddRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.RemoveRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.RestoreRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.CancelRecord), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.CloseRecord), false);

                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FirstRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.LastRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.NextRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.PreviousRecord), true);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.AddRow), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.DeleteRow), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.FilterTable), false);
                    oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.UDFForm), false);

                    if (BrowseBy != "")
                    {
                        oForm.DataBrowser.BrowseBy = BrowseBy;
                    }
                    oForm.Mode = boFormMode;
                    //if (FormMode == "A")
                    //{
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    //}
                    //else if (FormMode == "U")
                    //{
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    //}
                    //else if (FormMode == "F")
                    //{
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    //}
                    //else if (FormMode == "O")
                    //{
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    //}
                }

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public void LoadFromXML(string FileName, string BrowseBy, out string FormUID, string Title, string FormMode)
        {
            FormUID = "";
            try
            {

                System.Xml.XmlDocument oXmlDoc = null;
                string sPath = null;
                oXmlDoc = new System.Xml.XmlDocument();

                sPath = System.Windows.Forms.Application.StartupPath + "\\Forms\\" + FileName + ".xml";
                oXmlDoc.Load(sPath);

                string sXML = oXmlDoc.InnerXml.ToString();
                oApplication.LoadBatchActions(ref sXML);

                SAPbouiCOM.Form oForm = oApplication.Forms.ActiveForm;
                FormUID = oForm.UniqueID;
                if (Title != "")
                {
                    oForm.Title = Title;
                }
                oForm.SupportedModes = -1;

                oForm.EnableMenu("1281", true);
                oForm.EnableMenu("1282", true);
                oForm.EnableMenu("1288", true);
                oForm.EnableMenu("1289", true);
                oForm.EnableMenu("1290", true);
                oForm.EnableMenu("1291", true);


                oForm.EnableMenu("1292", true); //Activate Add row menu required for transactions
                oForm.EnableMenu("1293", true); //Activate delete row menu required for transactions
                if (BrowseBy != "")
                {
                    oForm.DataBrowser.BrowseBy = BrowseBy;
                }
                if (FormMode == "A")
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                else if (FormMode == "O")
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }

                //return (oXmlDoc.InnerXml);

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }




        public string LoadFromXML(string FileName, string BrowseBy, string FormMode)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                string sPath = null;
                oXmlDoc = new System.Xml.XmlDocument();
                sPath = System.Windows.Forms.Application.StartupPath + "\\Forms\\" + FileName + ".srf";
                oXmlDoc.Load(sPath);
                string sXML = oXmlDoc.InnerXml.ToString();
                oApplication.LoadBatchActions(ref sXML);
                SAPbouiCOM.Form oForm = oApplication.Forms.Item(FileName);
                oForm.SupportedModes = -1;
                oForm.EnableMenu("1281", true);
                oForm.EnableMenu("1282", true);
                oForm.EnableMenu("1288", true);
                oForm.EnableMenu("1289", true);
                oForm.EnableMenu("1290", true);
                oForm.EnableMenu("1291", true);
                oForm.EnableMenu("1292", true); //Activate Add row menu required for transactions
                oForm.EnableMenu("1293", true); //Activate delete row menu required for transactions
                oForm.EnableMenu("6913", false);

                if (BrowseBy != "")
                {
                    oForm.DataBrowser.BrowseBy = BrowseBy;
                }
                if (FormMode == "A")
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                else if (FormMode == "F")
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }
                return (oXmlDoc.InnerXml);

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return "";
        }



        public void LoadDefaultForm(string udoID)
        {
            SAPbouiCOM.MenuItem menu = oApplication.Menus.Item("47616"); // Link to the Default Forms menu
            try
            {
                if (menu.SubMenus.Count > 0)
                {
                    for (int i = 0; i <= menu.SubMenus.Count - 1; i++)
                    {
                        if (menu.SubMenus.Item(i).String.Contains(udoID))
                        {
                            menu.SubMenus.Item(i).Activate();
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        public bool FormAlreadyExist(string FormId, out string FormUID)
        {
            FormUID = string.Empty;
            int x = 0;
            bool foundG = false;
            for (x = 0; x <= oApplication.Forms.Count - 1; x++)
            {
                if (oApplication.Forms.Item(x).TypeEx == FormId)
                {
                    foundG = true;

                    FormUID = oApplication.Forms.Item(x).UniqueID;
                    break; // TODO: might not be correct. Was : Exit For
                }
            }

            return foundG;
        }

        #endregion

        #region AddChooseFromList

        public void AddChooseFromList(SAPbouiCOM.Form oForm, string cflUniqueID, string objType)
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = objType;
                oCFLCreationParams.UniqueID = cflUniqueID;
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public void AddChooseFromList_WithCond(SAPbouiCOM.Form oForm, string cflID, string objectid, string query, string Alias, ArrayList alCond)
        {
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                try
                {
                    SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                    oCFLs = oForm.ChooseFromLists;
                    SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                    oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = objectid;
                    oCFLCreationParams.UniqueID = cflID;
                    oCFL = oCFLs.Add(oCFLCreationParams);
                }
                catch
                {
                    oCFL = oForm.ChooseFromLists.Item(cflID);
                }
                oCFL.SetConditions(null);
                oCons = oCFL.GetConditions();



                #region Filter CFL values using ArrayList
                if (query == string.Empty)
                {
                    for (int i = 0; i < alCond.Count; i++)
                    {
                        if (oCons.Count > 0) //'If there are already user conditions.
                        {
                            oCons.Item(oCons.Count - 1).Relationship = (SAPbouiCOM.BoConditionRelationship)((alCond[i] as ArrayList)[0]);
                        }
                        oCon = oCons.Add();
                        oCon.BracketOpenNum = 1;
                        oCon.Alias = (alCond[i] as ArrayList)[1].ToString();
                        oCon.CondVal = (alCond[i] as ArrayList)[2].ToString();
                        oCon.Operation = (SAPbouiCOM.BoConditionOperation)((alCond[i] as ArrayList)[3]); //SAPbouiCOM.BoConditionOperation.co_EQUAL;                      
                        oCon.BracketCloseNum = 1;
                        oCFL.SetConditions(oCons);
                    }
                }
                #endregion

                #region Filter CFL values using recordset
                else
                {
                    oRs = returnRecord(query);
                    if (oRs.RecordCount == 0)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = Alias;
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "9999999";
                        oCFL.SetConditions(oCons);
                    }
                    else
                    {
                        while (!oRs.EoF)
                        {
                            if (oCons.Count > 0) //'If there are already user conditions.
                            {
                                oCons.Item(oCons.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            }
                            oCon = oCons.Add();
                            oCon.Alias = Alias;
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oRs.Fields.Item(0).Value.ToString();
                            oRs.MoveNext();
                            oCFL.SetConditions(oCons);
                        }
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region AddRow And DeleteRow


        public void AddRow(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDbDataSource, string Value)
        {
            oMatrix.FlushToDataSource();
            if (oMatrix.VisualRowCount == 0)
                oMatrix.AddRow(1, oMatrix.RowCount);
            else
            {

                if (Value != string.Empty)
                {
                    oDbDataSource.InsertRecord(oMatrix.RowCount);
                    oMatrix.LoadFromDataSource();
                }
            }
        }


        #endregion

        #region Allocate Batch
        public void AllocateBatch_41(SAPbouiCOM.Form Base_Form, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix Base_Matrix, SAPbouiCOM.Matrix Batch_Matrix1, SAPbouiCOM.Matrix Batch_Matrix2)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                string Base_ItemCode, Base_Batch, Batch_Managed, Batch_Matrix1_Item;
                int Batch_Row = 1;
                for (int i = 1; i < Base_Matrix.VisualRowCount; i++)
                {
                    oEdit = ((SAPbouiCOM.EditText)(Base_Matrix.GetCellSpecific("1", i)));
                    Base_ItemCode = oEdit.String;
                    oEdit = ((SAPbouiCOM.EditText)(Base_Matrix.GetCellSpecific("U_Batch", i)));
                    Base_Batch = oEdit.String;

                    Batch_Managed = SelectRecord("SELECT T2.[ManBtchNum] FROM  OITM T2  where  T2.Itemcode='" + Base_ItemCode + "'");
                    if (Batch_Managed == "Y")
                    {
                        for (; Batch_Row <= Batch_Matrix1.VisualRowCount;)
                        {
                            oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix1.GetCellSpecific("5", Batch_Row)));
                            Batch_Matrix1_Item = oEdit.String;
                            try
                            {
                                Batch_Matrix1.Columns.Item("5").Cells.Item(Batch_Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                            }
                            catch { }
                            oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix2.GetCellSpecific("2", 1)));
                            if (oEdit.String == "")
                            {
                                oEdit.String = Base_Batch;

                                oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix2.GetCellSpecific("7", 1))); //Batch Att 1
                                                                                                        //try
                                                                                                        //{
                                                                                                        //    oEdit.String = Base_Att1;
                                                                                                        //}
                                                                                                        //catch { }

                                oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix2.GetCellSpecific("8", 1)));//Batch Att 2
                                                                                                       //try
                                                                                                       //{
                                                                                                       //    oEdit.String = Base_Att2;
                                                                                                       //}
                                                                                                       //catch { }

                                oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix2.GetCellSpecific("10", 1)));//Exp Date
                                                                                                        //try
                                                                                                        //{
                                                                                                        //    oEdit.String = Base_Exp_Date;
                                                                                                        //}
                                                                                                        //catch { }

                                oEdit = ((SAPbouiCOM.EditText)(Batch_Matrix2.GetCellSpecific("11", 1)));//Manf Date
                                                                                                        //try
                                                                                                        //{
                                                                                                        //    oEdit.String = Base_Man_Date;
                                                                                                        //}
                                                                                                        //catch { }

                                oItem = oForm.Items.Item("1");
                                oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                if (Batch_Row == Batch_Matrix1.VisualRowCount)
                                {
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oItem = Base_Form.Items.Item("1");
                                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                Batch_Row++;
                                break;
                            }
                            else
                            {
                                Batch_Row++;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Select Records
        public string SelectRecord(string query)
        {
            SAPbobsCOM.Recordset oRs = null;
            string value = "";

            try
            {
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(query);

                if (oRs.EoF == false)
                {
                    value = oRs.Fields.Item(0).Value.ToString();
                }
                return value;
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return value;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                oRs = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public SAPbobsCOM.Recordset returnRecord(string query)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(query);
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return oRs;
        }
        #endregion



        #region FillCombo
        public void FillCombo(SAPbouiCOM.ComboBox oCombo, string Query)
        {
            SAPbobsCOM.Recordset oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));

            try
            {

                try
                {
                    oCombo.ValidValues.Remove("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                catch
                {

                }
                int Count = oCombo.ValidValues.Count;
                for (int i = 0; i < Count; i++)
                {
                    try
                    {
                        oCombo.ValidValues.Remove(oCombo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch { }
                }
                oRs.DoQuery(Query);

                oCombo.ValidValues.Add("", "");
                if (!oRs.EoF)
                {
                    for (int i = 0; i < oRs.RecordCount; i++)
                    {
                        if (oRs.Fields.Item(1).Value.ToString().Length <= 5)
                        {
                            // ADD SPACE TO AVOID VALIDATION ERROR 
                            oCombo.ValidValues.Add(oRs.Fields.Item(0).Value.ToString(), "    " + oRs.Fields.Item(1).Value.ToString());
                            // oCombo.ValidValues.Add(InsertRec.Fields.Item(0).Value.ToString(), "    ");

                        }
                        else if (oRs.Fields.Item(1).Value.ToString().Length > 50)
                        {
                            // REMOVE EXCESS CHARACTER TO AVOID VALIDATION ERROR
                            oCombo.ValidValues.Add(oRs.Fields.Item(0).Value.ToString(), oRs.Fields.Item(1).Value.ToString().Substring(0, 49));
                            //oCombo.ValidValues.Add(InsertRec.Fields.Item(0).Value.ToString(), "    ");
                        }
                        else
                        {
                            //oCombo.ValidValues.Add(InsertRec.Fields.Item(0).Value.ToString(), "    ");
                            oCombo.ValidValues.Add(oRs.Fields.Item(0).Value.ToString(), oRs.Fields.Item(1).Value.ToString());

                        }
                        oRs.MoveNext();

                    }
                }


            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
        #endregion

        #region Bind_Radio  & Checkbox

        public void Bind_Radio(SAPbouiCOM.Form oForm, int NoOfRButton)
        {
            SAPbouiCOM.OptionBtn opt;
            SAPbouiCOM.UserDataSource oUserdatasource;
            SAPbouiCOM.Item oItem;


            try
            {
                if (NoOfRButton == 2)
                {
                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD1");


                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");
                    opt.Selected = true;
                }
                else if (NoOfRButton == 3)
                {
                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt3");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");

                    opt.Selected = true;
                }

                else if (NoOfRButton == 4)
                {
                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt3");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");

                    opt.Selected = true;
                }

                else if (NoOfRButton == 5)
                {
                    oItem = oForm.Items.Item("opt5");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt5");

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt3");
                    opt.Selected = true;
                }



                else if (NoOfRButton == 6)
                {
                    oItem = oForm.Items.Item("opt6");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt5");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt6");

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt5");

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt3");
                    opt.Selected = true;
                }

                else if (NoOfRButton == 7)
                {
                    oItem = oForm.Items.Item("opt7");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt6");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt7");

                    oItem = oForm.Items.Item("opt5");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt6");

                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt5");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt3");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");

                    opt.Selected = true;
                }

                else if (NoOfRButton == 23)
                {
                    oItem = oForm.Items.Item("opt5");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt5");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");
                    opt.Selected = true;

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD1");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");
                    opt.Selected = true;
                }
                else if (NoOfRButton == 22)
                {
                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");
                    opt.Selected = true;

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD1");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");
                    opt.Selected = true;
                }
                else if (NoOfRButton == 232)
                {
                    oItem = oForm.Items.Item("opt7");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD3", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD3");

                    oItem = oForm.Items.Item("opt6");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt7");
                    opt.Selected = true;

                    oItem = oForm.Items.Item("opt5");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);

                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD2");

                    oItem = oForm.Items.Item("opt4");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt5");

                    oItem = oForm.Items.Item("opt3");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt4");
                    opt.Selected = true;

                    oItem = oForm.Items.Item("opt2");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    oUserdatasource = oForm.DataSources.UserDataSources.Add("BD1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 5);
                    opt.DataBind.SetBound(true, "", "BD1");

                    oItem = oForm.Items.Item("opt1");
                    opt = (SAPbouiCOM.OptionBtn)(oItem.Specific);
                    opt.GroupWith("opt2");
                    opt.Selected = true;
                }
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public void Bind_CheckBox(SAPbouiCOM.Form oForm, string ItemUID)
        {
            SAPbouiCOM.CheckBox chk = null;
            SAPbouiCOM.UserDataSource oUserdatasource;
            SAPbouiCOM.Item oItem = null;
            try
            {
                oItem = oForm.Items.Item(ItemUID);
                chk = (SAPbouiCOM.CheckBox)(oItem.Specific);

                oUserdatasource = oForm.DataSources.UserDataSources.Add("CHK", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1);
                chk.DataBind.SetBound(true, "", "CHK");
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Return Proper Date Format

        public string GetDateFormat(string sDate)
        {
            try
            {
                string sDateFormat = sDate.Substring(0, 4) + "/" + sDate.Substring(4, 2) + "/" + sDate.Substring(6, 2);
                return sDateFormat;
            }
            catch
            {
                return "";
            }
        }

        #endregion

        #region GetDefaultSeries

        public int GetDefaultSeries(string ObjectType)
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.SeriesService oSeriesService = null;
            SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = null;
            SAPbobsCOM.Series oSeries = null;

            //'get company service
            oCmpSrv = (SAPbobsCOM.CompanyService)oCompany.GetCompanyService();
            //'get series service
            oSeriesService = (SAPbobsCOM.SeriesService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService);
            //'get new series
            oSeries = (SAPbobsCOM.Series)oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries);
            //'get DocumentTypeParams for filling the document type
            oDocumentTypeParams = (SAPbobsCOM.DocumentTypeParams)oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams);
            //'set the document type (e.g. A/R Invoice=13)
            oDocumentTypeParams.Document = ObjectType;
            //'get the default series of the SaleOrder documentset the document type
            oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams);
            return oSeries.Series;
        }

        #endregion

        #region GetMaxDocNum
        public int GetMaxDocNum(string objCode, int series)
        {
            string QueryString = "";
            int MaxCode = 0;
            SAPbobsCOM.Recordset oRs = null;
            try
            {

                QueryString = "SELECT MAX(NextNumber) \"NextNum\" FROM NNM1 WHERE ObjectCode='" + objCode + "' AND Series= " + series + "";
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(QueryString);
                string test = oRs.Fields.Item("NextNum").Value.ToString();
                if (oRs.EoF == false)
                {
                    MaxCode = Convert.ToInt16(oRs.Fields.Item("NextNum").Value);
                }
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return MaxCode;
        }
        #endregion

        #region GoodsIssue_Production

        #region GoodsReceipt_Production


        public void GoodsReceipt_Production(string FormUID, string Form_DocEntry, string sDocNum, string sDocDate, string sShift, string sRemks, out string GRDocEntry, SAPbouiCOM.Matrix oMtx1, SAPbouiCOM.Matrix oMtx2)
        {
            //Goods Receipt object. 59       
            GRDocEntry = string.Empty;
            int lretcode;
            string ItemCode, WhsCode, RejWhsCode, DocEntry;
            double Quantity, RejQty;
            string type = "";
            string BaseEntry = string.Empty;

            if (Form_DocEntry != string.Empty)
            {
                string IsAlreadyExist = SelectRecord("SELECT DOCENTRY FROM IGN1 WHERE U_BaseEntry='" + BaseEntry + "'");
                if (IsAlreadyExist != string.Empty)
                {
                    oApplication.SetStatusBarMessage("Goods Receipt is already created.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    return;
                }
            }

            SAPbobsCOM.Documents Prod = null;
            Prod = (SAPbobsCOM.Documents)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry));

            string dDate = sDocDate.Substring(0, 4) + "/" + sDocDate.Substring(4, 2) + "/" + sDocDate.Substring(6, 2);
            DateTime d1 = Convert.ToDateTime(dDate);
            string msg2 = sShift.ToString().Trim() + "," + type + sDocNum + ",PDt-" + dDate.ToString();

            Prod.DocDate = d1;
            Prod.JournalMemo = "G.Recpt," + msg2;
            Prod.Comments = sRemks;

            int Row = 0;
            for (int i = 1; i <= oMtx2.VisualRowCount; i++)
            {
                ItemCode = ((SAPbouiCOM.EditText)oMtx2.GetCellSpecific("V_5", i)).String;
                if (ItemCode == "")
                {
                    break;
                }
                WhsCode = ((SAPbouiCOM.EditText)oMtx2.GetCellSpecific("V_1", i)).String;
                Quantity = double.Parse(((SAPbouiCOM.EditText)oMtx2.GetCellSpecific("V_0", i)).String);
                DocEntry = ((SAPbouiCOM.EditText)oMtx1.GetCellSpecific("V_7", i)).String;


                if (Quantity > 0)
                {
                    if (Row > 0)
                    {
                        Prod.Lines.Add();
                    }
                    Prod.Lines.BaseType = 202;
                    Prod.Lines.BaseEntry = int.Parse(DocEntry);
                    Prod.Lines.Quantity = Quantity;
                   // Prod.Lines.BPThouseCode = WhsCode;
                    Prod.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntComplete;
                    Prod.Lines.UserFields.Fields.Item("U_BaseEntry").Value = BaseEntry;
                    Prod.Lines.SetCurrentLine(Row);
                    Row++;
                }


                RejQty = double.Parse(((SAPbouiCOM.EditText)oMtx2.GetCellSpecific("V_2", i)).String);
                RejWhsCode = ((SAPbouiCOM.EditText)oMtx2.GetCellSpecific("V_3", i)).String;
                if (RejQty > 0)
                {
                    if (Row > 0)
                    {
                        Prod.Lines.Add();
                    }
                    Prod.Lines.BaseType = 202;
                    Prod.Lines.BaseEntry = int.Parse(DocEntry);
                    Prod.Lines.Quantity = RejQty;
                    //Prod.Lines.BPThouseCode = RejWhsCode;
                    Prod.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntReject;
                    Prod.Lines.UserFields.Fields.Item("U_BaseEntry").Value = BaseEntry;

                    Prod.Lines.SetCurrentLine(Row);
                    Row++;
                }
            }
            /*Prod.JournalMemo = "Receipt from Production " + BaseDoc;
            Prod.Comments = "Receipt from Production " + BaseDoc;

            oRsGI = returnReocord("SELECT ItemCode,BPTHouse,DocEntry,PlannedQty [Quantity] FROM OWOR WHERE DocEntry=" + BaseDoc + " ");
            int Row = 0;
            while (!oRsGI.EoF)
            {
                ItemCode = oRsGI.Fields.Item("ItemCode").Value.ToString();
                WhsCode = oRsGI.Fields.Item("BPTHouse").Value.ToString();
                Quantity = double.Parse(oRsGI.Fields.Item("Quantity").Value.ToString());


                Prod.Lines.BaseType = 202;
                Prod.Lines.BaseEntry = int.Parse(oRsGI.Fields.Item("DocEntry").Value.ToString());
                //Prod.Lines.BaseLine = int.Parse(oRsGI.Fields.Item("LineNum").Value.ToString());
                // IssueToProd.Lines.ItemCode = ItemCode;
                Prod.Lines.Quantity = Quantity;
                Prod.Lines.BPThouseCode = WhsCode;

                string BatchNo = SelectReocord("SELECT U_BATCh FROM [@TIS_WO] WHERE DOCENTRY=" + WODocEntry + "");
                Prod.Lines.BatchNumbers.BatchNumber = BatchNo;
                Prod.Lines.BatchNumbers.Quantity = Quantity;


                if (Row > 0)
                {
                    Prod.Lines.Add();
                }
                Prod.Lines.SetCurrentLine(Row);
                Row++;
                oRsGI.MoveNext();
            }*/

            lretcode = Prod.Add();

            if (lretcode != 0)
            {
                int Errcode; string ErrMsg;//05007BR3165
                oCompany.GetLastError(out Errcode, out ErrMsg);
                oApplication.MessageBox("Receipt From Production  : " + ErrMsg, 1, "OK", "", "");
                oApplication.StatusBar.SetText("Receipt From Production  : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                GRDocEntry = "";
            }
            else
            {
                //oApplication.StatusBar.SetText("Receipt from Production created successfully.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                GRDocEntry = oCompany.GetNewObjectKey();
            }

        }

        #endregion

        #endregion

        #region GetPeriod
        public void GetPeriod()
        {
            try
            {
                SAPbouiCOM.Company company = oApplication.Company;
                int Period = company.CurrentPeriod;
                SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();
                SAPbobsCOM.FinancePeriodParams diPeriodParams = (SAPbobsCOM.FinancePeriodParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiFinancePeriodParams);

                diPeriodParams.AbsoluteEntry = Period;
                SAPbobsCOM.FinancePeriod finPeriod = oCompanyService.GetFinancePeriod(diPeriodParams);

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion


        #endregion

        public SqlConnection DBConnection()
        {
            SqlConnection conn = null;
            SAPbobsCOM.Recordset oRs = null;
            string QueryString = null;
            try
            {
                //SetConnection
                QueryString = "SELECT U_DBPwd FROM [@TIS_DATABASE]";
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(QueryString);

                if (!oRs.EoF)
                {
                    clsVariables.DbPass = oRs.Fields.Item("U_DBPwd").Value.ToString();
                }
                else
                    oApplication.MessageBox("Database Connection details missing in user defined table!", 1, "OK", "", "");
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


            try
            {
                clsVariables.Server = oCompany.Server;
                clsVariables.DbName = oCompany.CompanyDB;
                clsVariables.DbUser = oCompany.DbUserName;


                string strConn = "SERVER='" + oCompany.Server + "';DATABASE='" + oCompany.CompanyDB + "';USER ID='" + oCompany.DbUserName + "';Password='" + clsVariables.DbPass + "'";
                conn = new SqlConnection(strConn);
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return conn;
        }

        public void FillCombo_UDF(SAPbouiCOM.ComboBox oCombo, string tableName, string fieldName)
        {
            StringBuilder sbQuery = new StringBuilder();
            sbQuery.Append(" SELECT T1.\"FldValue\",T1.\"Descr\" ");
            sbQuery.Append(" FROM \"CUFD\" T0 ");
            sbQuery.Append(" INNER JOIN \"UFD1\" T1 ON T0.\"TableID\" = T1.\"TableID\" AND T0.\"FieldID\" = T1.\"FieldID\" ");
            sbQuery.Append(" WHERE  T0.\"TableID\" = '" + tableName + "' AND T0.\"AliasID\" = '" + fieldName + "' ");
            FillCombo(oCombo, sbQuery.ToString());
        }

        public DateTime ConvertStrToDate(string strDate, string strFormat)
        {
            DateTime oDate = default(DateTime);
            try
            {

                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-GB", false);
                System.Globalization.CultureInfo newCi = (System.Globalization.CultureInfo)ci.Clone();

                System.Threading.Thread.CurrentThread.CurrentCulture = newCi;
                oDate = DateTime.ParseExact(strDate, strFormat, ci.DateTimeFormat);

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return oDate;
        }

        public void RefreshRecord()
        {
            oApplication.ActivateMenuItem(Convert.ToString((int)SAPMenuEnum.PreviousRecord));
            oApplication.ActivateMenuItem(Convert.ToString((int)SAPMenuEnum.NextRecord));
        }

        public void FillCombo_Series_Custom(SAPbouiCOM.Form oForm, string objCode, string seriesUDF, string mode)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("Series").Specific;
                string docDate = ((SAPbouiCOM.EditText)oForm.Items.Item(seriesUDF).Specific).Value;

                if (mode == "Load")
                {
                    StringBuilder sbQuery = new StringBuilder();
                    sbQuery.Append(" SELECT T0.\"Series\", T0.\"SeriesName\" ");
                    sbQuery.Append(" FROM NNM1 T0 ");
                    sbQuery.Append(" INNER JOIN OFPR T1 ON T0.\"Indicator\" = T1.\"Indicator\"  ");
                    sbQuery.Append(" WHERE ObjectCode ='" + objCode + "'  AND  \"Locked\" = 'N'  AND \"F_RefDate\" <='" + docDate + "'  AND \"T_RefDate\" >= '" + docDate + "'  ");
                    sbQuery.Append(" GROUP BY T0.\"Series\", T0.\"SeriesName\"  ");


                    int Count = oCombo.ValidValues.Count;
                    for (int i = 0; i < Count; i++)
                    {
                        try
                        {
                            oCombo.ValidValues.Remove(oCombo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        catch { }
                    }
                    oRs = returnRecord(sbQuery.ToString());
                    while (!oRs.EoF)
                    {
                        try
                        {
                            oCombo.ValidValues.Add(oRs.Fields.Item("SERIES").Value.ToString(), oRs.Fields.Item("SERIESNAME").Value.ToString());
                        }
                        catch { }
                        oRs.MoveNext();
                    }

                    oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                }
                else
                {
                    oCombo.ValidValues.LoadSeries(objCode, SAPbouiCOM.BoSeriesMode.sf_View);
                }
                oCombo.Item.DisplayDesc = true;
            }
            catch (Exception ex)
            {
                oApplication.SetStatusBarMessage(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

      
    }
}