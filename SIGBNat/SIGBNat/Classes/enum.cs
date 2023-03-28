using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace General.Classes
{
    class @enum
    {
        public enum YesNoEnum
        {
            [Description("Yes")]
            Y,
            [Description("No")]
            N,
        }

        public enum SAPButtonEnum
        {
            Add = 1,
            Cancel = 2,
        }

        public enum SAPMenuEnum
        {
            FindRecord = 1281,
            AddRecord = 1282,
            RemoveRecord = 1283,
            CancelRecord = 1284,
            RestoreRecord = 1285,
            CloseRecord = 1286,
            LastRecord = 1291,
            FirstRecord = 1290,
            NextRecord = 1288,
            PreviousRecord = 1289,
            RefreshRecord = 1304,
            AddRow = 1292,
            DeleteRow = 1293,
            DuplicateRow = 1294,
            FilterTable = 4870,
            UDFForm = 6913,
            SystemInitialisation = 8192,
            Modules = 43520
        }

        public enum SAPFormUIDEnum
        {
            MainMenu = 169,
            SalesOrder = 139,
            Delivery = 140,
            Return = 180,
            ARInvoice = 133,
            ARCreditMemo = 179,
            BusinessPartnerMasterData = 134,
            ItemMasterData = 150,
            PurchaseOrder = 142,
            GoodsReceipt = 721,
            GoodsIssue = 720,
            InventoryTransfer = 940,
            Production = 65211,
            ReceiptFromProduction = 65214,
            ReceiptFromProductionUDFForm = -65214,
            IssueForProduction = 65213,
            BatchInward = 41,
            BinCodeOutward = 1470000007,
            BusinessPlace = 1320000702
        }

        public enum SAPCustomFormUIDEnum
        {


            [Description("Create BPI UDO")]
           // DWHMUDO,
           BPIUDO,

            [Description("BPI Add-on")]
            //DWHMADO,
            BPIADO,

            [Description("BPI Form")]
            //DWHM,
            //SIGBNattu,
            BPIM,

            [Description("CCBI Form")]
            //DWHM,
            //SIGBNattu,
            //SOF,
            CCBI,

        }

        public enum SAPCommonFieldEnum
        {
            CardCode,
            ItemCode
        }

        public enum UDFFieldType
        {
            Alpha,
            Text,
            Integer,
            Date,
            Time,
            Amount,
            Quantity,
            Percent,
            UnitTotal,
            Rate,
            Price,
            Link,
            YN,
            Status,
            Activity,
            ActivityStatus,
            QCStastus,
            BOMCompType,
            IssueMethod,
            Layer,
            QCType,
            RewStatus,
            Numeric,
        }

        public enum TableType
        {
            StandardTable,
            CustomTable
        }

        public enum SAPCommonMaskModeEnum
        {
            All = -1,
            Ok = 1,
            Add = 2,
            Find = 4,
            View = 8
        }

        public enum SAPCommonColorEnum
        {
            White = 16777215,
            Green = 20221504,
            Red = 70221504
        }
    }
}