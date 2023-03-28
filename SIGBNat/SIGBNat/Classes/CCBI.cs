using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using General.Classes;
using static General.Classes.@enum;
using General.Extensions;
using SAPbouiCOM;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
//using OfficeOpenXml;
using System.Web.Hosting;
using System.Net.Http;
using System.Data;
namespace SIGBNat
{
    class CCBI:Connection
    {
        #region Variables

        SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.EditText oEdit;
        SAPbouiCOM.ComboBox oCombo;
        SAPbobsCOM.Recordset checkoRs;
        SAPbobsCOM.Recordset recordset;
        SAPbouiCOM.Item oItem1;
        SAPbouiCOM.LinkedButton olinkedButton;
        SAPbouiCOM.Column oColumn;
        // SAPbobsCOM.Recordset oRset;
        SAPbouiCOM.Item oItem;
        public const string objType = "CCBI";
        SAPbouiCOM.DBDataSource oDbDataSource = null;
        const string formMenuUID = "CCBI";
        public const string formTypeEx = "CCBI";
        clsCommon objclsComman = new clsCommon();
        StringBuilder sbQuery = new StringBuilder();
        public const string headerTable = "@CCBI";
        public const string rowTable = "@CCBI1";
        public const string mtxtable = "@CCBI1";
        const string BPUID = "code";
        const string BPNUID1 = "cname";
        SAPbobsCOM.Recordset oRs;
        public const string cflobjtype = "4";
        public const string CFL2objtype = "2";
        const string frombinUID = "BPT";
        const string matrixPrimarColumnUID = "CIno";
        public const string matrixPrimaryColumnUDF = "U_CIno";
        const string matrixUID = "mtx";
        public const string LineUID = "LineId";
        public SAPbobsCOM.Recordset pORs { get; set; }
        public const string CFL = "CFL_1";
        public const string CFL2 = "CFL_0";
        public const string InUDF = "U_CIno";
        public const string DescUDF = "U_CDesc";
        public const string baseUDF = "U_Cqty";
        public const string GroupUDF = "U_group";
        clsCommon objclsCommon = new clsCommon();
       // StringBuilder sbQuery = new StringBuilder();
        #endregion
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
                                objclsComman.AddChooseFromList_WithCond(oForm, CFL2, CFL2objtype, query.ToString(), "CardCode", alCondVal);
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
                                        objclsComman.AddChooseFromList_WithCond(oForm, CFL, cflobjtype, query.ToString(), "ItemCode", alCondVal);
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
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {

                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                                   // oMatrix.FlushToDataSource();
                                    string cardc = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_code", 0).Trim();
                                    string doctyp = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_DocTy", 0);
                                    string docstat = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_DocStd", 0);
                                    if ((cardc == string.Empty))
                                    {
                                        BubbleEvent = false;
                                        oApplication.StatusBar.SetText("Card Code  is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                    if ((doctyp == string.Empty))
                                    {
                                        BubbleEvent = false;
                                        oApplication.StatusBar.SetText("Document Type is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                    if ((docstat == string.Empty))
                                    {
                                        BubbleEvent = false;
                                        oApplication.StatusBar.SetText("Document Status is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                   
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

                                if (oCFLEvento.ChooseFromListUID == CFL2)
                            {
                                oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                string a = oDataTable.GetValue("CardCode", 0).ToString();
                                oDbDataSource.SetValue("U_code", 0, a);
                                string b = oDataTable.GetValue("CardName", 0).ToString();
                                oDbDataSource.SetValue("U_cname", 0, b);
                                string c = oDataTable.GetValue("CardType", 0).ToString();
                                if (c == "C")
                                {
                                    c = "Customer";
                                }
                                if (c == "S")
                                {
                                    c = "Supplier";
                                }
                                if (c == "L")
                                {
                                    c = "Lead";
                                }
                                oDbDataSource.SetValue("U_Type", 0, c);
                                string d = oDataTable.GetValue("Balance", 0).ToString();
                                oDbDataSource.SetValue("U_Balance", 0, d);
                                string e = oDataTable.GetValue("Currency", 0).ToString();
                                oDbDataSource.SetValue("u_typ", 0, e);
                                string f = oDataTable.GetValue("GroupCode", 0).ToString();
                                //oDbDataSource.SetValue(GroupUDF, 0, f);
                                
                                sbQuery.Length = 0;
                                sbQuery.Append(" SELECT \"GroupName\" ");
                                sbQuery.Append(" FROM \"OCRG\" ");
                                sbQuery.Append(" WHERE \"GroupCode\" ='" + f + "' ");
                                string grpname = SelectRecord(sbQuery.ToString());
                                oDbDataSource.SetValue(GroupUDF, 0, grpname);
                            }
                            else if (oCFLEvento.ChooseFromListUID == CFL)
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

                        }
                        #region T_et_ITEM_PRESSED
                        
                            else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                            {
                           
                            #region FIND BUTTON
                            if (pVal.ItemUID == "3")
                            {
                                SAPbouiCOM.DBDataSource oDBDataSource;
                               
                                string ItemTable = "@DOCTYP";
                                oDBDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                string cardcode = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_code", 0);
                                string head = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_DocTy", 0);
                                string docst = oForm.DataSources.DBDataSources.Item(headerTable).GetValue("U_DocStd", 0);
                                if ((head == string.Empty))
                                {
                                    BubbleEvent = false;
                                    oApplication.StatusBar.SetText("Document Type is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return;
                                }
                                if ((docst == string.Empty))
                                {
                                    BubbleEvent = false;
                                    oApplication.StatusBar.SetText("Document Status is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return;
                                }

                                sbQuery.Length = 0;
                                sbQuery.Append(" SELECT T0.\"Code\",T0.\"Name\",T0.\"U_child\",T0.\"U_ObjT\" ");
                                sbQuery.Append(" FROM \"@DOCTYP\" T0 ");
                                sbQuery.Append(" WHERE T0.\"Name\" = '" + head + "' ");
                                pORs = objclsCommon.returnRecord(sbQuery.ToString());
                                string a = pORs.Fields.Item("Code").Value.ToString();
                                string b = pORs.Fields.Item("U_child").Value.ToString();
                                string c = pORs.Fields.Item("U_ObjT").Value.ToString();

                                #region link to docunum based on the doucument
                                oColumn = oMatrix.Columns.Item("Doc");
                                olinkedButton= oColumn.ExtendedObject;
                                olinkedButton.LinkedObjectType = c;
                                #endregion

                                if (docst == "Close") { docst = "C"; }
                                if (docst == "Open") { docst = "O"; }
                                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                if (c == "202")
                                {
                                    string query = "SELECT T0.DocNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.Price,T1.WhsCode,T1.TaxCode,T1.UomCode,T3.ItmsGrpNam,T0.DocDate FROM " + a + " T0 INNER JOIN " + b + " T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OITM T2 ON T1.ItemCode = T2.ItemCode INNER JOIN OITB T3 ON T2.ItmsGrpCod = T3.ItmsGrpCod WHERE T0.CardCode = '" + cardcode + "' and T0.DocStatus = '" + docst + "' and T0.ObjType = '" + c + "' ";
                                    oRs.DoQuery(query);
                                    int rc = oRs.RecordCount;
                                    oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                    oDbDataSource.SetValue("U_Rec", 0, rc.ToString());
                                    if (oRs.RecordCount > 0)
                                    {
                                        for (int i = 0; i <= oRs.RecordCount - 1; i++)
                                        {

                                            oMatrix.AddRow();
                                            string value2 = oRs.Fields.Item("ItemCode").Value.ToString();
                                            value2 = string.IsNullOrEmpty(value2) ? "0" : value2;
                                            string value3 = oRs.Fields.Item("Dscription").Value.ToString();
                                            value3 = string.IsNullOrEmpty(value3) ? "0" : value3;
                                            string value4 = oRs.Fields.Item("Quantity").Value.ToString();
                                            value4 = string.IsNullOrEmpty(value4) ? "0" : value4;
                                            string value5 = oRs.Fields.Item("DocNum").Value.ToString();
                                            value5 = string.IsNullOrEmpty(value5) ? "0" : value5;
                                            string value6 = oRs.Fields.Item("Price").Value.ToString();
                                            value6 = string.IsNullOrEmpty(value6) ? "0" : value6;
                                            string value7 = oRs.Fields.Item("ItmsGrpNam").Value.ToString();
                                            value7 = string.IsNullOrEmpty(value7) ? "0" : value7;
                                            string value8 = oRs.Fields.Item("UomCode").Value.ToString();
                                            value8 = string.IsNullOrEmpty(value8) ? "0" : value8;
                                            string value9 = oRs.Fields.Item("WhsCode").Value.ToString();
                                            value9 = string.IsNullOrEmpty(value9) ? "0" : value9;
                                            string value11 = oRs.Fields.Item("DocDate").Value.ToString();
                                            value11 = string.IsNullOrEmpty(value11) ? "0" : value11;
                                            value11 = value11.Substring(0, 11);
                                            string value12 = oRs.Fields.Item("VatSum").Value.ToString();
                                            value12 = string.IsNullOrEmpty(value12) ? "0" : value12;


                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ID").Cells.Item(i + 1).Specific).Value = (i + 1).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CIno").Cells.Item(i + 1).Specific).Value = value2;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CDesc").Cells.Item(i + 1).Specific).Value = value3;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cqty").Cells.Item(i + 1).Specific).Value = value4;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Doc").Cells.Item(i + 1).Specific).Value = value5;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Price").Cells.Item(i + 1).Specific).Value = value6;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ItemType").Cells.Item(i + 1).Specific).Value = value7;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("UomCode1").Cells.Item(i + 1).Specific).Value = value8;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("whsCode1").Cells.Item(i + 1).Specific).Value = value9;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("DocDate1").Cells.Item(i + 1).Specific).Value = value11;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ta").Cells.Item(i + 1).Specific).Value = value12;
                                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("LineID").Cells.Item(i + 1).Specific).Value = oRs.Fields.Item("LineNum").Value.ToString();

                                            oRs.MoveNext();

                                        }
                                        oApplication.StatusBar.SetText("Data has been fetched now you can ADD", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    }
                                }
                                else
                                {
                                    string query = "SELECT T0.DocNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.Price,T1.WhsCode,T1.VatSum,T1.UomCode,T3.ItmsGrpNam,T0.DocDate FROM " + a + " T0 INNER JOIN " + b + " T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OITM T2 ON T1.ItemCode = T2.ItemCode INNER JOIN OITB T3 ON T2.ItmsGrpCod = T3.ItmsGrpCod WHERE T0.CardCode = '" + cardcode + "' and T0.DocStatus = '" + docst + "' and T0.ObjType = '" + c + "' ";
                                    oRs.DoQuery(query);
                                    int rc = oRs.RecordCount;
                                    oDbDataSource = oForm.DataSources.DBDataSources.Item(headerTable);
                                    oDbDataSource.SetValue("U_Rec", 0, rc.ToString());
                                    if (oRs.RecordCount > 0)
                                    {
                                        for (int i = 0; i <= oRs.RecordCount - 1; i++)
                                        {

                                            oMatrix.AddRow();
                                            string value2 = oRs.Fields.Item("ItemCode").Value.ToString();
                                            value2 = string.IsNullOrEmpty(value2) ? "0" : value2;
                                            string value3 = oRs.Fields.Item("Dscription").Value.ToString();
                                            value3 = string.IsNullOrEmpty(value3) ? "0" : value3;
                                            string value4 = oRs.Fields.Item("Quantity").Value.ToString();
                                            value4 = string.IsNullOrEmpty(value4) ? "0" : value4;
                                            string value5 = oRs.Fields.Item("DocNum").Value.ToString();
                                            value5 = string.IsNullOrEmpty(value5) ? "0" : value5;
                                            string value6 = oRs.Fields.Item("Price").Value.ToString();
                                            value6 = string.IsNullOrEmpty(value6) ? "0" : value6;
                                            string value7 = oRs.Fields.Item("ItmsGrpNam").Value.ToString();
                                            value7 = string.IsNullOrEmpty(value7) ? "0" : value7;
                                            string value8 = oRs.Fields.Item("UomCode").Value.ToString();
                                            value8 = string.IsNullOrEmpty(value8) ? "0" : value8;
                                            string value9 = oRs.Fields.Item("WhsCode").Value.ToString();
                                            value9 = string.IsNullOrEmpty(value9) ? "0" : value9;
                                            string value11 = oRs.Fields.Item("DocDate").Value.ToString();
                                            value11 = string.IsNullOrEmpty(value11) ? "0" : value11;
                                            value11 = value11.Substring(0, 11);
                                            string value12 = oRs.Fields.Item("VatSum").Value.ToString();
                                            value12 = string.IsNullOrEmpty(value12) ? "0" : value12;


                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ID").Cells.Item(i + 1).Specific).Value = (i + 1).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CIno").Cells.Item(i + 1).Specific).Value = value2;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CDesc").Cells.Item(i + 1).Specific).Value = value3;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cqty").Cells.Item(i + 1).Specific).Value = value4;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Doc").Cells.Item(i + 1).Specific).Value = value5;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Price").Cells.Item(i + 1).Specific).Value = value6;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ItemType").Cells.Item(i + 1).Specific).Value = value7;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("UomCode1").Cells.Item(i + 1).Specific).Value = value8;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("whsCode1").Cells.Item(i + 1).Specific).Value = value9;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("DocDate1").Cells.Item(i + 1).Specific).Value = value11;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ta").Cells.Item(i + 1).Specific).Value = value12;
                                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("LineID").Cells.Item(i + 1).Specific).Value = oRs.Fields.Item("LineNum").Value.ToString();

                                            oRs.MoveNext();

                                        }
                                        oApplication.StatusBar.SetText("Complete your detail, Now You Proceed to ADD", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    }
                                    else if (oRs.RecordCount == 0)
                                    {
                                        oApplication.StatusBar.SetText("There is no Item Your selected BP CODE can you check another one", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        oApplication.StatusBar.SetText("Mismatch BpCode and Document Type can you check that.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                }

                            }

                            #endregion

                            #region disable fields in matrix 
                            oColumn = oMatrix.Columns.Item("Doc");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("DocDate1");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("ta");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("whsCode1");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("UomCode1");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("ItemType");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("Price");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("Cqty");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("CDesc");
                            oColumn.Editable = false;
                            oColumn = oMatrix.Columns.Item("CIno");
                            oColumn.Editable = false;
                            #endregion

                            #region Clear Matrix
                            if (pVal.ItemUID == "Clear")
                            {
                                oMatrix.Clear();
                            }
                            #endregion

                            #region T_et_ITEM_PRESSED
                            else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                            {
                                #region Add Record

                                oForm = (SAPbouiCOM.Form)oApplication.Forms.Item(pVal.FormUID);
                                if (pVal.ItemUID == Convert.ToString((int)SAPButtonEnum.Add))
                                {
                                    if (pVal.FormTypeEx == formTypeEx && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        string code = oForm.DataSources.DBDataSources.Item(headerTable).GetValue(CommonFields.DocNum, 0).ToString();
                                        if (code.Trim() == string.Empty)
                                        {
                                            LoadForm(Convert.ToString((int)SAPMenuEnum.AddRecord));
                                            return;
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region Print Method
                            if (pVal.ItemUID == "Print")
                            {
                                PostExcel();
                            }
                            #endregion

                        }
                        #endregion
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
                            //oItem = oForm.Items.Item("Series");
                            //oItem.Disable();
                        }
                        if (pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.PreviousRecord) 
                            || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.NextRecord) 
                            || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.FirstRecord) 
                            || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.LastRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
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

            }

            oForm = oApplication.Forms.ActiveForm;
           // EnableControls(oForm.UniqueID);
            oForm.Freeze(true);
            //below line is for dropdown of DocType
            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTy").Specific;
            objclsCommon.FillCombo(oCombo, "SELECT \"Name\",\"Name\" FROM \"@DOCTYP\"");
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
            oForm.Freeze(false);

            #region Series And DocNum

            try
            {
                oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("SDocDate").Specific;
                oEdit.String = "t";

                objclsComman.FillCombo_Series_Custom(oForm, objType, "SDocDate", "Load");

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

            oItem = oForm.Items.Item("Series");
            oItem.EnableinFindMode();

            oItem = oForm.Items.Item("DocNum");
            oItem.EnableinFindMode();

            #endregion


        }

        #endregion

        #region Print Method
        public void PostExcel()
        {
            // Create a new workbook object
            IWorkbook workbook = new XSSFWorkbook();

            // Create a new sheet in the workbook
            ISheet sheet = workbook.CreateSheet("New sheet");



            IRow headerRow = sheet.CreateRow(0);

            headerRow.CreateCell(0).SetCellValue("LineID");
            headerRow.CreateCell(1).SetCellValue("DocNo");
            headerRow.CreateCell(2).SetCellValue("ItemCode");
            headerRow.CreateCell(3).SetCellValue("dscription");
            headerRow.CreateCell(4).SetCellValue("Quantity");
            headerRow.CreateCell(5).SetCellValue("Price");
            headerRow.CreateCell(6).SetCellValue("TaxCode");
            headerRow.CreateCell(7).SetCellValue("ItemType");
            headerRow.CreateCell(8).SetCellValue("UOM Code");
            headerRow.CreateCell(9).SetCellValue("Whse");
            headerRow.CreateCell(10).SetCellValue("Document Date");
            //headerRow.CreateCell(11).SetCellValue("Bin Location");
           
            oRs.MoveFirst();
            for (int i = 1; i <= oRs.RecordCount; i++)
            {
                string LineID = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("ID", i)).String;
                string DocNo = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("Doc", i)).String;
                string ItemCode = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("CIno", i)).String;
                string dscription = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("CDesc", i)).String;
                string Quantity = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("Cqty", i)).String;
                string Price = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("Price", i)).String;
                string TaxCode = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("ta", i)).String;
                string ItemType = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("ItemType", i)).String;
                string UOMCode = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("UomCode1", i)).String;
                string Whse = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("whsCode1", i)).String;
                string DocumentDate = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("DocDate1", i)).String;
                //string Bin = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("1470000168", i)).String;
              
                // Create some rows of data
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(LineID);
                row.CreateCell(1).SetCellValue(DocNo);
                row.CreateCell(2).SetCellValue(ItemCode);
                row.CreateCell(3).SetCellValue(dscription);
                row.CreateCell(4).SetCellValue(Quantity);
                row.CreateCell(5).SetCellValue(Price);
                row.CreateCell(6).SetCellValue(TaxCode);
                row.CreateCell(7).SetCellValue(ItemType);
                row.CreateCell(8).SetCellValue(UOMCode);
                row.CreateCell(9).SetCellValue(Whse);
                row.CreateCell(10).SetCellValue(DocumentDate);
                //row.CreateCell(11).SetCellValue(Bin);
                oRs.MoveNext();
               
            }

            // Write the workbook to a file
            string path = "Carddetails.ODF";
            var fullpath = Path.Combine("C:\\Users\\Socius IGB\\Desktop", path);
            using (FileStream stream = new FileStream(fullpath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
            oApplication.StatusBar.SetText(fullpath, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }





        #endregion
    }
   
}

