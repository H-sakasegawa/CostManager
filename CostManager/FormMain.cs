using ExcelReaderUtility;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;

namespace CostManager
{
    public partial class FormMain : Form
    {
        enum COL_INDEX
        {
            CHECK = 0,
            KIND,
            NAME,
            PRODUCT_NUM,    //制作数
            ROW_COST,       //原価率
            MATERIAL_COST,  //材料費
            WORKER_COST,    //人件費
            PACKAGE_COST,   //包装費
            PRICE,          //定価
            PROFIT,         //利益率

        }
        OptionData optionData = new OptionData();

        ProductReader productReader = new ProductReader();

        CostReader costReader = new CostReader();

        string curCostDataFilePath = null;

        public FormMain()
        {
            InitializeComponent();


        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(optionData.DataBasePath))
            {
                Utility.MessageConfirm("データベースのフォルダが未設定です。\nオプション画面からデータベースのフォルダを設定してください。","データベースパス");
                EnalbeControl(false);
                return;
            }
            ReadDataBase(optionData.DataBasePath);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            optionData.SaveOptions();
        }

        int ReadDataBase(string databasePath)
        {
            string productDataBase = Path.Combine(databasePath, Const.ProductFileName);
            string costDataBase = Path.Combine(databasePath, Const.CostFileName);
            if (!File.Exists(productDataBase))
            {
                Utility.MessageError($"{Const.ProductFileName} が見つかりません");
                return -1;
            }
            if (!File.Exists(costDataBase))
            {
                Utility.MessageError($"{Const.CostFileName} が見つかりません");
                return -1;
            }
            //-------------------------------------------
            // 商品データを読み込む
            //-------------------------------------------
            productReader.ReadExcel(productDataBase);
            //-------------------------------------------
            // 原価データを読み込む
            //-------------------------------------------
            costReader.ReadExcel(costDataBase);

            return 0;
        }


        private void cmbMainListKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            var SelectKind = cmbMainListKind.Text;

            if (SelectKind == Const.SelectAll)
            {
                foreach (DataGridViewRow row in gridList.Rows)
                {
                    row.Visible = true;
                }
            }
            else
            {
                foreach (DataGridViewRow row in gridList.Rows)
                {
                    CostData costData = (CostData)row.Tag;
                    var product = productReader.GetProductDataByID(costData.ProductId);
                    if (SelectKind == product.kind)
                    {
                        row.Visible = true;
                    }
                    else
                    {
                        row.Visible = false;
                    }
                }
            }
        }


        private void btnDelete_Click(object sender, EventArgs e)
        {
            List<int> lstCheckedRow = new List<int>();
            //チェックボックスがONの行があるか？
            var rows = gridList.Rows;
            for (int iRow = 0; iRow < rows.Count; iRow++)
            {
                object value = (bool)(rows[iRow].Cells[(int)COL_INDEX.CHECK].Value);

                if (value!=null && (bool)value == true)
                {
                    lstCheckedRow.Add(iRow);
                }
            }
            if (lstCheckedRow.Count > 0)
            {
                if (Utility.MessageConfirm("選択されている項目を削除します。\nよろしいですか？", "削除") != DialogResult.OK)
                {
                    return;
                }
                for (int iRow = lstCheckedRow.Count - 1; iRow >= 0; iRow--)
                {
                    rows.RemoveAt(lstCheckedRow[iRow]);
                }
            }
            else
            {
                //選択行を削除
                foreach (DataGridViewRow row in gridList.SelectedRows)
                {
                    rows.Remove(row);
                }
            }
        }

        /// <summary>
        /// メインリストに商品情報（コストオブジェクトを追加
        /// </summary>
        /// <param name="product"></param>
        public void AddProduct(ProductReader.ProductData product)
        {
            CostData costData = new CostData(product, costReader);
            AddCostDataToRow(costData);
        }
        public void AddCostDataToRow(CostData costData)
        {

            var index = gridList.Rows.Add();
            var row = gridList.Rows[index];

            row.Cells[(int)COL_INDEX.CHECK].Value = false;
            row.Cells[(int)COL_INDEX.KIND].Value = costData.Kind;
            row.Cells[(int)COL_INDEX.NAME].Value = costData.ProductName;
            row.Tag = costData;

            UpdateRowValues(row);


            //メインリストの分類コンボボックスのアイテム更新
            cmbMainListKind.Items.Clear();
            cmbMainListKind.Items.Add(Const.SelectAll);
            foreach (DataGridViewRow row2 in gridList.Rows)
            {
                var value = row2.Cells[(int)COL_INDEX.KIND].Value;

                if( cmbMainListKind.Items.IndexOf( value )<0)
                {
                    cmbMainListKind.Items.Add(value);
                }
            }
            cmbMainListKind.SelectedIndex = 0;
        }


        public void UpdateRowValues(DataGridViewRow row)
        {
            CostData costData = (CostData)row.Tag;

            float costRate;
            float rowCostRate;
            float laborCostRate;
            float packageCostRate;
            float profit;

            costData.Calc(out costRate, out rowCostRate, out laborCostRate, out packageCostRate, out profit);

            row.Cells[(int)COL_INDEX.PRODUCT_NUM].Value     = costData.ProductNum;
            row.Cells[(int)COL_INDEX.ROW_COST].Value        = costRate * 100;
            row.Cells[(int)COL_INDEX.MATERIAL_COST].Value   = rowCostRate * 100;
            row.Cells[(int)COL_INDEX.WORKER_COST].Value     = laborCostRate * 100;
            row.Cells[(int)COL_INDEX.PACKAGE_COST].Value    = packageCostRate * 100;
            row.Cells[(int)COL_INDEX.PRICE].Value           = costData.Price;
            row.Cells[(int)COL_INDEX.PROFIT].Value          = profit;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormItemSelector frm = new FormItemSelector(this, productReader);
            frm.Show();

        }

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach(DataGridViewRow row in gridList.Rows)
            {
                row.Cells[(int)COL_INDEX.CHECK].Value = chkAll.Checked;
            }
        }

        private void gridList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var selectRows = gridList.SelectedRows;
            if (selectRows == null) return;

            var costData = (CostData)selectRows[0].Tag;

            FormEditCost frm = new FormEditCost(costData, costReader);
            if( frm.ShowDialog() == DialogResult.OK)
            {
                //メイングリッドの各種コストを計算して表示　★★
                UpdateRowValues(selectRows[0]);
            }
        }
        private void gridList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                gridList_CellDoubleClick(null, null);
            }
        }


        private void mnuOption_Click(object sender, EventArgs e)
        {
            string orgPath = optionData.DataBasePath;

            FormOption frm = new FormOption(optionData);
            if( frm.ShowDialog()== DialogResult.OK )
            {
                if (string.IsNullOrEmpty(optionData.DataBasePath))
                {
                    Utility.MessageConfirm("データベースのフォルダが未設定です。\nオプション画面からデータベースのフォルダを設定してください。", "データベースパス");
                    EnalbeControl(false);
                    return;
                }
                EnalbeControl(true);

                //パスに変更が発生したら読み込みなおし
                if (orgPath != optionData.DataBasePath)
                {
                    ReadDataBase(optionData.DataBasePath);
                }
            }
        }

        void EnalbeControl(bool nFlg)
        {
            button1.Enabled = nFlg;

        }
        /// <summary>
        /// 保存メニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mnuSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(curCostDataFilePath))
            {
                //新規状態からの初めての保存
                curCostDataFilePath = GetSaveFileName();
                if (string.IsNullOrEmpty(curCostDataFilePath))
                {
                    return;
                }
            }
            SaveData(curCostDataFilePath);
        }
        /// <summary>
        /// 保存ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolBtnSave_Click(object sender, EventArgs e)
        {
            mnuSave_Click(null, null);

        }
        /// <summary>
        /// 名前を付けて保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mnuSaveAs_Click(object sender, EventArgs e)
        {
            string filePath = GetSaveFileName(curCostDataFilePath);
            if( string.IsNullOrEmpty(filePath))
            {
                return;
            }
            SaveData(filePath);
        }
        private string GetSaveFileName(string initFilePath=null)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = initFilePath;
            if (dlg.ShowDialog()== DialogResult.OK)
            {
                return dlg.FileName;
            }
            return null;
        }
        private void SaveData(string filePath)
        {
            CostDataList costDataList = new CostDataList();
            foreach (DataGridViewRow row in gridList.Rows)
            {
                CostData cost = (CostData)row.Tag;

                costDataList.Add(cost);
            }
            using (TextWriter writer = File.CreateText(filePath))
            {
                var serializer = new Serializer();
                serializer.Serialize(writer, costDataList);
            }
        }


        private void mnuLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var loadData = LoadData(dlg.FileName);
                ReadCostDataToRow(loadData);
            }
        }

        private void gridList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void gridList_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {

                // ドラッグ中のファイルやディレクトリの取得
                string[] drags = (string[])e.Data.GetData(DataFormats.FileDrop);
                if( drags.Length >1)
                {
                    return;
                };

                string filePath = drags[0];
                if (!System.IO.File.Exists(filePath))
                {
                    // ファイル以外であればイベント・ハンドラを抜ける
                    return;

                }
                var loadData = LoadData(filePath);
                ReadCostDataToRow(loadData);

                e.Effect = DragDropEffects.Copy;
            }
        }

        private CostDataList LoadData(string filePath)
        {
            curCostDataFilePath = filePath;

            CostDataList costDataList;
            using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
            {
                var deserializer = new Deserializer();
                costDataList = deserializer.Deserialize<CostDataList>(reader);
            }
            return costDataList;
        }
        void ReadCostDataToRow(CostDataList loadData)
        {
            gridList.Rows.Clear();

            foreach (var data in loadData.ListCostDatas)
            {
                //読み込まれたデータにCost管理がひもついていないので、ここで設定
                data.AttachCostReader(costReader);

                AddCostDataToRow(data);
            }


        }

     }
}
