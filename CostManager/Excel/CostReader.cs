using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

using NPOI.XSSF.UserModel;

using NPOI.SS.UserModel;
using System.Xml.Linq;
using static ExcelReaderUtility.ProductReader;
using NPOI.POIFS.Crypt.Dsig;

using CostManager;

namespace ExcelReaderUtility
{
    public class CostReader
    {

        public class MaterialData
        {
            /// <summary>
            /// ID
            /// </summary>
            public string id;
            /// <summary>
            /// 分類
            /// </summary>
            public string kind;
            /// <summary>
            /// 名称
            /// </summary>
            public string name;
            /// <summary>
            /// 原価
            /// </summary>
            public float cost;
            /// <summary>
            /// 原価基準容量
            /// </summary>
            public float gram;
        }
        public class WorkerData
        {
            /// <summary>
            /// ID
            /// </summary>
            public string id;
            /// <summary>
            /// 名称
            /// </summary>
            public string name;
            /// <summary>
            /// 時給
            /// </summary>
            public uint hourlyPay;
        }
        public class PackageData
        {
            /// <summary>
            /// ID
            /// </summary>
            public string id;
            /// <summary>
            /// 名称
            /// </summary>
            public string name;
            /// <summary>
            /// 単価(１個辺り）
            /// </summary>
            public float cost;
        }

        /// <summary>
        /// 原価基本データ
        /// </summary>
        private List<MaterialData> lstMaterial = new List<MaterialData>();
        private List<WorkerData> lstWorker = new List<WorkerData>();
        private List<PackageData> lstPackage = new List<PackageData>();

        public int ReadExcel( string excelFilePath)
        {
            if ( !File.Exists(excelFilePath))
            {
                return -1;
            }


            var workbook = ExcelReader.GetWorkbook(excelFilePath, "xlsx");
            if( workbook==null)
            {
                Utility.MessageError($"{excelFilePath}\nを開けません");
                return -1;
            }

            //データクリア
            lstMaterial.Clear();
            lstWorker.Clear();
            lstPackage.Clear();
            int sheetNum = workbook.NumberOfSheets;

            for (int iSheet = 0; iSheet < sheetNum; iSheet++)
            {
                XSSFSheet sheet = (XSSFSheet)((XSSFWorkbook)workbook).GetSheetAt(iSheet);

                string sheetName = sheet.SheetName;

                Dictionary<string, int> dicColmunIndex = new Dictionary<string, int>();

                bool bReadTitle = false;
                int iRow = 0;
                while (true)
                {
                    List<ExcelReader.TextInfo> lstRowData;

                    int rc = ReadRowData(sheet, iRow, out lstRowData);
                    if (rc == -1)
                    {
                        //読み込み終了
                        break;
                    }
                    if (lstRowData != null)
                    {
                        if (!bReadTitle)
                        {
                            //最初の行をタイトル行とする。
                            bReadTitle = true;
                        }
                        else
                        {
                            switch (sheetName)
                            {
                                case "原材料":
                                    {
                                        MaterialData data = new MaterialData();
                                        GetExcelRowData(lstRowData, (int)MaterialItemName.ID, ref data.id);
                                        GetExcelRowData(lstRowData, (int)MaterialItemName.Kind, ref data.kind);
                                        GetExcelRowData(lstRowData, (int)MaterialItemName.Name, ref data.name);
                                        GetExcelRowData(lstRowData, (int)MaterialItemName.Cost, ref data.cost);
                                        GetExcelRowData(lstRowData, (int)MaterialItemName.Gram, ref data.gram);
                                        var wk = lstMaterial.Find(x => x.id == data.id);
                                        if (wk != null)
                                        {
                                            //IDの重複
                                            Utility.MessageError($"原価データベースの原材料に重複したIDがあります。\n({wk.id}) {wk.name}\n({data.id}) {data.name}\n\n重複しないIDを設定してください");
                                            return -1;
                                        }
                                        lstMaterial.Add(data);
                                    }
                                    break;
                                case "作業者":
                                    {
                                        WorkerData data = new WorkerData();
                                        GetExcelRowData(lstRowData, (int)WorkerItemName.ID, ref data.id);
                                        GetExcelRowData(lstRowData, (int)WorkerItemName.Name, ref data.name);
                                        GetExcelRowData(lstRowData, (int)WorkerItemName.HourlyPay, ref data.hourlyPay);
                                        var wk = lstWorker.Find(x => x.id == data.id);
                                        if (wk != null)
                                        {
                                            //IDの重複
                                            Utility.MessageError($"原価データベースの作業者に重複したIDがあります。\n({wk.id}) {wk.name}\n({data.id}) {data.name}\n\n重複しないIDを設定してください");
                                            return -1;
                                        }
                                        lstWorker.Add(data);
                                    }
                                    break;
                                case "包装材":
                                    {
                                        PackageData data = new PackageData();
                                        GetExcelRowData(lstRowData, (int)PackageItemName.ID, ref data.id);
                                        GetExcelRowData(lstRowData, (int)PackageItemName.Name, ref data.name);
                                        GetExcelRowData(lstRowData, (int)PackageItemName.Cost, ref data.cost);
                                        var wk = lstPackage.Find(x => x.id == data.id);
                                        if (wk != null)
                                        {
                                            //IDの重複
                                            Utility.MessageError($"原価データベースの包装材に重複したIDがあります。\n({wk.id}) {wk.name}\n({data.id}) {data.name}\n\n重複しないIDを設定してください");
                                            return -1;
                                        }
                                        lstPackage.Add(data);
                                    }
                                    break;
                            }

                        }
                    }

                    iRow++;
                }
            }

            return 0;

        }



        private int GetExcelRowData(List<ExcelReader.TextInfo> rowData, int itemIndex, ref string value)
        {
            value = rowData[itemIndex].Text;
            return 0;
        }

        private int GetExcelRowData(List<ExcelReader.TextInfo> rowData, int itemIndex, ref int value)
        {
            value = 0;
            if (int.TryParse(rowData[itemIndex].Text, out value) == false)
            {
                return -1;
            }
            return 0;
        }
        private int GetExcelRowData(List<ExcelReader.TextInfo> rowData, int itemIndex, ref uint value)
        {
            value = 0;
            if (uint.TryParse(rowData[itemIndex].Text, out value) == false)
            {
                return -1;
            }
            return 0;
        }
        private int GetExcelRowData(List<ExcelReader.TextInfo> rowData, int itemIndex, ref float value)
        {
            value = 0;
            if (float.TryParse(rowData[itemIndex].Text, out value) == false)
            {
                return -1;
            }
            return 0;
        }



        private int ReadRowData(XSSFSheet sheet, int iRow, out List<ExcelReader.TextInfo> lstRowITems)
        {
            lstRowITems = null;

            int cellNum = ExcelReader.GetCellCount(sheet, iRow);
            if (cellNum < 0) return -1; //読み込み終了


            if (string.IsNullOrEmpty(ExcelReader.CellValue(sheet, iRow, 0)))
            {
                return 0;
            }
            lstRowITems = new List<ExcelReader.TextInfo>();

            for (int iCol = 0; iCol < cellNum; iCol++)
            {
                string text = ExcelReader.CellValue(sheet, iRow, iCol);

                ExcelReader.TextInfo textInfo = new ExcelReader.TextInfo();
                textInfo.col = iCol;
                textInfo.row = iRow;
                textInfo.textData = text.Replace("\n", "\r\n");
                lstRowITems.Add(textInfo);
            }

            return 0;

        }

        /// <summary>
        /// 原材料名一覧取得
        /// </summary>
        /// <returns></returns>
        public List<MaterialData> GetMaterialList(string kind = null)
        {
            List<MaterialData> lst;
            if (kind == null)
            {
                lst = lstMaterial;
            }
            else
            {
                lst = lstMaterial.FindAll(x => x.kind == kind);
            }
            return lst;
        }
        /// <summary>
        /// 原材料IDからその登録データを取得
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public MaterialData GetMaterialtDataByID( string id)
        {
            if( string.IsNullOrEmpty(id)) return null;    
            return lstMaterial.FirstOrDefault(x => x.id == id);
        }


        /// <summary>
        /// 作業者名一覧取得
        /// </summary>
        /// <returns></returns>
        public List<WorkerData> GetWorkerList()
        {
            return lstWorker;
        }
        /// <summary>
        /// 作業者IDからその登録データを取得
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public WorkerData GetWorkerDataByID(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return lstWorker.FirstOrDefault(x => x.id == id);
        }



        /// <summary>
        /// 包装材名一覧取得
        /// </summary>
        /// <returns></returns>
        public List<PackageData> GetPackageList()
        {
            return lstPackage;
        }
        /// <summary>
        /// 作業者IDからその登録データを取得
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public PackageData GetPackageDataByID(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return lstPackage.FirstOrDefault(x => x.id == id);
        }

    }
}
