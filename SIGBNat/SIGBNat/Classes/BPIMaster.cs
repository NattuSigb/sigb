using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using General.Classes;
using static General.Classes.@enum; //if this enum.cs is not accessing then on SAPMenuEnum will show you an error
using System.Collections; // by including line 68 array list error is fine
//using SAPbouiCOM;
using General.Extensions;

namespace SIGBNat
{
    class BPIMaster : Connection //connection to be given then only can access oApplication on catch line 23
    {

        #region Variables  


        SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.EditText oEdit;
        SAPbouiCOM.ComboBox oCombo;
        SAPbouiCOM.Item oItem;
        public const string objType = "BPIM";
        SAPbouiCOM.DBDataSource oDbDataSource = null;
        const string formMenuUID = "BPIM";
                                                   //const string formMenuUID1 = "SOF";
        public const string formTypeEx = "BPIM";
        clsCommon objclsComman = new clsCommon();
        StringBuilder sbQuery = new StringBuilder();

        public const string headerTable = "@BPIM";
        public const string rowTable = "@BPIM1";

        public const string mtxtable = "@BPIM1";

        //public const string formTypeEx = "DWHM";
        ////const string formMenuUID = "DWHM";
        //public const string objType = "DWHM";
        const string BPUID = "BPCode";
        const string BPNUID1 = "BPName";
        public const string cflobjtype = "4";
        public const string CFL2objtype = "2";
        const string fromDeptUID = "BPCode";
        const string toDeptUID = "BPName";
        const string frombinUID = "BPT";
        const string fromDocDUID = "BPIDocDate";
        //const string tobinUID = "BPT";
        const string matrixPrimaryColumnUID = "Item No.";
        const string matrixPrimarColumnUID = "INo";
        public const string matrixPrimaryColumnUDF = "U_INo";
        const string matrixUID = "mtx";
        // public const string matrixPrimaryColumnUDF = "U_ino";
        //public const string LineUID = "LineId";

        public const string CFL = "CFL_1";
        public const string CFL2 = "CFL_0";
        //public const string InUDF = "U_ItemNo";
        public const string InUDF = "U_INo";
        public const string DescUDF = "U_Dscription";
        public const string baseUDF = "U_PPqty";
        public const string stockUDF = "U_SPqty";

        #endregion

        public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region BeforeAction == true

                if (pVal.BeforeAction == true)
                {
                    try
                    {
                        if (pVal.MenuUID == formMenuUID || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        //if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                //Record is directly added without validation
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.FindRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            oItem = oForm.Items.Item(CommonFields.DocNum);
                            oItem.EnableinFindMode();
                            oItem = oForm.Items.Item("Series");
                            oItem.Enable();
                            oItem = oForm.Items.Item("BPCode");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPName");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPT");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPIDocDate");
                            oItem.Disable();
                        }
                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.PreviousRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            oItem = oForm.Items.Item(CommonFields.DocNum);
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPCode");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPName");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPT");
                            oItem.Disable();

                            oItem = oForm.Items.Item(CommonFields.Series);
                            oItem.Disable();
                            oItem = oForm.Items.Item("mtx");
                            oItem.Disable();
                        }

                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.FirstRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            oItem = oForm.Items.Item(CommonFields.DocNum);
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPCode");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPName");
                            oItem.Disable();
                            oItem = oForm.Items.Item("BPT");
                            oItem.Disable();

                            oItem = oForm.Items.Item(CommonFields.Series);
                            oItem.Disable();
                            oItem = oForm.Items.Item("mtx");
                            oItem.Disable();

                        }

                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.NextRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            oItem = oForm.Items.Item(CommonFields.DocNum);
                            oItem.Enable();
                            oItem = oForm.Items.Item("BPCode");
                            oItem.Enable();
                            oItem = oForm.Items.Item("BPName");
                            oItem.Enable();
                            oItem = oForm.Items.Item("BPT");
                            oItem.Disable();

                            oItem = oForm.Items.Item(CommonFields.Series);
                            oItem.Disable();
                            oItem = oForm.Items.Item("mtx");
                            oItem.Disable();

                        }
                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.DeleteRow))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            int i = oApplication.MessageBox("Do you really want to remove line?", 1, "Yes", "No", "");
                            if (i == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }

                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.RemoveRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            int i = oApplication.MessageBox("Do you really want to remove record?", 1, "Yes", "No", "");
                            if (i == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = true : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region BeforeAction == false

                if (pVal.BeforeAction == false)
                {
                    try
                    {
                        if (pVal.MenuUID == formMenuUID || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        {
                            LoadForm(pVal.MenuUID);
                        }
                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRow))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            AddRow(matrixUID, rowTable, matrixPrimaryColumnUDF);
                            AddRow(matrixUID, rowTable, InUDF);
                        }

                        else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.DeleteRow))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            DeleteRow(matrixUID, rowTable, matrixPrimaryColumnUDF);
                            DeleteRow(matrixUID, rowTable, InUDF);
                        }

                        //else if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.PreviousRecord) || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.NextRecord)
                        //   || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.LastRecord) || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.FirstRecord))
                        //{
                        //    oForm = oApplication.Forms.ActiveForm;

                            //oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);


                            //oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);


                       // }



                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }


        public void ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region Before_Action == true
                if (pVal.BeforeAction == true)
                {
                    try
                    {
                        #region CHOOSE FROM LIST

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            if (pVal.ItemUID == BPUID)
                            {
                                oForm = oApplication.Forms.Item(pVal.FormUID);
                                oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                ArrayList alCondVal = new ArrayList();
                                ArrayList temp = new ArrayList();
                                string query = "Select * FROM \"OCRD\" ";
                               // objclsComman.AddChooseFromList_WithCond(oForm,CFL,CFL2objtype, query.ToString(),"CardCode",alCondVal);
                                objclsComman.AddChooseFromList_WithCond(oForm, CFL, CFL2objtype, query.ToString(), "CardCode", alCondVal);
                            }


                            else if (pVal.ItemUID == matrixUID)
                            {
                                if (pVal.ColUID == matrixPrimarColumnUID)
                                {
                                    oForm = oApplication.Forms.Item(pVal.FormUID);
                                    SAPbouiCOM.Column oColumn;
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                    oMatrix.FlushToDataSource();
                                    oDbDataSource = oForm.DataSources.DBDataSources.Item(rowTable);

                                    if (pVal.ColUID == matrixPrimarColumnUID)
                                    {
                                        oColumn = oMatrix.Columns.Item(matrixPrimarColumnUID);

                                        ArrayList alCondVal = new ArrayList();
                                        ArrayList temp = new ArrayList();
                                        string query = "Select * FROM \"OITM \" ";
                                        objclsComman.AddChooseFromList_WithCond(oForm, CFL2, cflobjtype, query.ToString(), "ItemCode", alCondVal);
                                    }
                                }
                            }

                        }

                        #endregion

                        #region T_et_ITEM_PRESSED
                        else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                            #region Add Record

                            if (pVal.ItemUID == Convert.ToString((int)SAPButtonEnum.Add))
                            {
                                oForm = oApplication.Forms.Item(pVal.FormUID);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                   
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                                    //oMatrix.FlushToDataSource();

                                   

                                   
                                        oDbDataSource = oForm.DataSources.DBDataSources.Item(rowTable);
                                        string itemcode = oForm.DataSources.DBDataSources.Item(rowTable).GetValue(InUDF, 0).Trim();
                                        

                                        if ((itemcode == string.Empty))
                                        {
                                            BubbleEvent = false;
                                            oApplication.StatusBar.SetText("itemcode is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            return;
                                        }

                                       

                                    oMatrix.LoadFromDataSource();
                                }
                            }

                            #endregion
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region Before_Action == false
                if (pVal.BeforeAction == false)
                {
                    try
                    {

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            SAPbouiCOM.DataTable oDataTable = null;
                            oForm = oApplication.Forms.Item(pVal.FormUID);
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            oDataTable = oCFLEvento.SelectedObjects;

                            if (oDataTable == null || oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                return;
                            }

                            if (oCFLEvento.ChooseFromListUID == CFL)
                            {
                                oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                string Avaistk = oDataTable.GetValue("CardCode", 0).ToString();
                                oDbDataSource.SetValue("U_BPCode", 0, Avaistk);
                                string Avaist = oDataTable.GetValue("CardName", 0).ToString();
                                oDbDataSource.SetValue("U_BPName", 0, Avaist);
                                string Avaisk = oDataTable.GetValue("CardType", 0).ToString();
                                if (Avaisk == "C")
                                {
                                    Avaisk = "Customer";
                                }
                                if (Avaisk == "S")
                                {
                                    Avaisk = "Supplier";
                                }
                                if (Avaisk == "L")
                                {
                                    Avaistk = "Lead";
                                }
                                oDbDataSource.SetValue("U_BPT", 0, Avaisk);
                               
                            }
                            else if (oCFLEvento.ChooseFromListUID == CFL2)
                            {
                               
                                oDbDataSource = oForm.DataSources.DBDataSources.Item(rowTable);
                                string List = oDataTable.GetValue("ItemCode", 0).ToString();
                                string ListDesc = oDataTable.GetValue("ItemName", 0).ToString();
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                                
                                oDbDataSource.SetValue(InUDF, pVal.Row - 1, List);
                                oDbDataSource.SetValue(DescUDF, pVal.Row - 1, ListDesc);
                                //oDbDataSource.SetValue(baseUDF, pVal.Row - 1, BaseUOM);
                                //oDbDataSource.SetValue(stockUDF, pVal.Row - 1, AvailStock);
                                if (pVal.Row == oMatrix.RowCount)
                                {
                                    oDbDataSource.InsertRecord(oMatrix.VisualRowCount);
                                    int RowNo = 1;
                                    for (int i = 0; i < oDbDataSource.Size; i++)
                                    {
                                        oDbDataSource.SetValue(CommonFields.LineId, i, RowNo.ToString());
                                        RowNo = RowNo + 1;
                                    }
                                }

                                oMatrix.LoadFromDataSource();
                                oMatrix.Columns.Item(matrixPrimarColumnUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                clsVariables.boolCFLSelected = true;
                                SAPbouiCOM.ICellPosition oPos = oMatrix.GetCellFocus();
                                clsVariables.ColNo = oPos.ColumnIndex;
                                clsVariables.RowNo = oPos.rowIndex;
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                }

                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        //  
        #region Methods
        public void LoadForm(string MenuID)
        {
            clsVariables.boolCFLSelected = false;

            if (MenuID == formMenuUID)
            {
                string formUID = "";
                objclsComman.LoadXML(MenuID, "", string.Empty, SAPbouiCOM.BoFormMode.fm_ADD_MODE);
                oForm = oApplication.Forms.ActiveForm;
                oForm.DataSources.UserDataSources.Add("Close", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Item("Close").Value = "N";
                //oForm.EnableMenu("5895", true);

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                if (oMatrix.VisualRowCount == 0)
                {
                    oMatrix.AddRow(1, 1);
                    SetLineId(oMatrix);
                }
            }


            oForm = oApplication.Forms.ActiveForm;
            EnableControls(oForm.UniqueID);
            

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;

            if (oMatrix.VisualRowCount == 0)
            {
                oMatrix.AddRow(1, 1);
            }

            //oItem = oForm.Items.Item("DocNum");
            //oItem.EnableinFindMode();

            //oItem = oForm.Items.Item("DocEntry");
            //oItem.EnableinFindMode();

            //oItem = oForm.Items.Item("Series");
            //oItem.Enable();

            //oItem = oForm.Items.Item(fromDeptUID);
            //oItem.Enable();


            //oItem = oForm.Items.Item(BPNUID1);
            //oItem.EnableinAddMode();

            //oItem = oForm.Items.Item(toDeptUID);
            //oItem.EnableinAddMode();

            //oItem = oForm.Items.Item(frombinUID);
            //oItem.EnableinFindMode();

            //oItem = oForm.Items.Item(fromDocDUID);
            //oItem.EnableinAddMode();

            #region Series And DocNum

            try
            {
                oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("BPIDocDate").Specific;
                oEdit.String = "t";

                objclsComman.FillCombo_Series_Custom(oForm, objType, "BPIDocDate", "Load");

                #region Set DocNum

                string defaultseries = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("Series", 0).Trim();
                if (defaultseries == string.Empty)
                {
                    oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("Series").Specific;
                    oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    defaultseries = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("Series", 0).Trim();
                }
                string MaxCode = oForm.BusinessObject.GetNextSerialNumber(defaultseries.ToString(), objType).ToString();
                int inMaxCode = objclsComman.GetMaxDocNum(objType.ToString(), int.Parse(defaultseries));
                oForm.DataSources.DBDataSources.Item(headerTable).SetValue("DocNum", 0, MaxCode.ToString());

                #endregion
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            }

            

            #endregion

        }

    



        private void AddRow(string matrixUID, string tableName, string matrixPrimaryUDF)
        {
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
            if (oMatrix.VisualRowCount == 0)
            {
                oMatrix.AddRow(1, 1);
                return;
            }
            oMatrix.FlushToDataSource();
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(tableName);
            string value = oDBDataSource.GetValue(matrixPrimaryUDF, oMatrix.VisualRowCount - 1).ToString().Trim();
            objclsComman.AddRow(oMatrix, oDBDataSource, value);
        }
        private void DeleteRow(string matrixUID, string tableName, string matrixPrimaryUDF)
        {
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
            oMatrix.FlushToDataSource();
            oDbDataSource = oForm.DataSources.DBDataSources.Item(tableName);
            int RowNo = 1;
            for (int i = 0; i < oDbDataSource.Size; i++)
            {
                string value = oDbDataSource.GetValue(matrixPrimaryUDF, i).ToString().Trim();
                if (value == string.Empty)
                {
                    oDbDataSource.RemoveRecord(i);
                }
                oDbDataSource.SetValue(CommonFields.LineId, i, RowNo.ToString());
                RowNo = RowNo + 1;
            }
            oMatrix.LoadFromDataSource();
        }
        private void EnableControls(string formUID)
        {
            oForm = oApplication.Forms.Item(formUID);
            oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.AddRow), true);
            oForm.EnableMenu(Convert.ToString((int)SAPMenuEnum.DeleteRow), true);
        }

        private void SetLineId(SAPbouiCOM.Matrix oMatrix)
        {
            int rowNo = 1;
            oMatrix.FlushToDataSource();
            oDbDataSource = oForm.DataSources.DBDataSources.Item(rowTable);
            for (int i = 0; i < oDbDataSource.Size; i++)
            {
                oDbDataSource.SetValue(CommonFields.LineId, i, rowNo.ToString());
                rowNo = rowNo + 1;
            }
            oMatrix.LoadFromDataSource();
        }

        #endregion
    }


}