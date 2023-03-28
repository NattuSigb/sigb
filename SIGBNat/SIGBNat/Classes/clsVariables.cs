//using DepartmentWiseWharehouseManagement.Classes;
using SIGBNat;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace General.Classes
{
    class clsVariables
    {

        #region Global Variable
        private static string sServer;
        private static string sDbName;
        private static string sDbUser;
        private static string sDbPass;
        private static string sDocEntry;
        private static string sBaseEntry;
        private static string sBaseLine;
        private static string sBaseOrderNo;
        private static string sBaseItemCode;
        private static string sSpecialMixItemCode;
        private static string sBaseObjectType;
        private static string sBaseBatch;
        private static string sBaseShift;
        private static string sBaseWhsCode;
     

        private static double dblBaseQuantity;

        private static ArrayList alList;
 /*       private static List<clsItemEntity> itemList;
        private static List<clsBOMEntity> bomList;
 */
        private static string sBaseFormUID;
        private static string sBaseFormTypeEx;

        private static bool oboolCFLSelected;
        private static bool oboolNewFormOpen;

        private static int iRowNo, iColNo;
        private static SAPbouiCOM.Form oBaseForm;
        private static SAPbouiCOM.DataTable ostuffDataTable;

        #endregion

        # region Login Information

        public static string Server
        {
            get { return sServer; }
            set { sServer = value; }
        }

        public static string DbName
        {
            get { return sDbName; }
            set { sDbName = value; }
        }

        public static string DbUser
        {
            get { return sDbUser; }
            set { sDbUser = value; }
        }

        public static string DbPass
        {
            get { return sDbPass; }
            set { sDbPass = value; }
        }

        #endregion

        #region Variables

        public static string BaseFormUID
        {
            get { return sBaseFormUID; }
            set { sBaseFormUID = value; }
        }

        public static string BaseFormTypeEx
        {
            get { return sBaseFormTypeEx; }
            set { sBaseFormTypeEx = value; }
        }

        public static string DocEntry
        {
            get { return sDocEntry; }
            set { sDocEntry = value; }
        }

        public static bool boolCFLSelected
        {
            get { return oboolCFLSelected; }
            set { oboolCFLSelected = value; }
        }

        public static bool boolNewFormOpen
        {
            get { return oboolNewFormOpen; }
            set { oboolNewFormOpen = value; }
        }

        public static SAPbouiCOM.DataTable stuffDataTable
        {
            get { return ostuffDataTable; }
            set { ostuffDataTable = value; }
        }

        public static int RowNo
        {
            get { return iRowNo; }
            set { iRowNo = value; }
        }
        public static int ColNo
        {
            get { return iColNo; }
            set { iColNo = value; }
        }

        public static SAPbouiCOM.Form BaseForm
        {
            get { return oBaseForm; }
            set { oBaseForm = value; }
        }
        public static string BaseEntry
        {
            get { return sBaseEntry; }
            set { sBaseEntry = value; }
        }
        public static string BaseLine
        {
            get { return sBaseLine; }
            set { sBaseLine = value; }
        }
        public static string BaseOrderNo
        {
            get { return sBaseOrderNo; }
            set { sBaseOrderNo = value; }
        }
        public static ArrayList AlList
        {
            get { return alList; }
            set { alList = value; }
        }
/*
        public static List<clsItemEntity> ItemList
        {
            get { return itemList; }
            set { itemList = value; }
        }

        public static List<clsBOMEntity> BOMList
        {
            get { return bomList; }
            set { bomList = value; }
        }
 */
        public static string BaseItemCode
        {
            get { return sBaseItemCode; }
            set { sBaseItemCode = value; }
        }
        public static string SpecialMixItemCode
        {
            get { return sSpecialMixItemCode; }
            set { sSpecialMixItemCode = value; }
        }
        public static double BaseQuantity
        {
            get { return dblBaseQuantity; }
            set { dblBaseQuantity = value; }
        }
        public static string BaseWhsCode
        {
            get { return sBaseWhsCode; }
            set { sBaseWhsCode = value; }
        }
        public static string BaseObjectType
        {
            get { return sBaseObjectType; }
            set { sBaseObjectType = value; }
        }
        public static string BaseBatch
        {
            get { return sBaseBatch; }
            set { sBaseBatch = value; }
        }
        public static string BaseShift
        {
            get { return sBaseShift; }
            set { sBaseShift = value; }
        }

        //public static string Clinic
        //{
        //    get { return Clinic; }
        //    set { Clinic = value; }
        //}
        #endregion

    }
}


