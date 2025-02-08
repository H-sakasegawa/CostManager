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
using System.Reflection;
using NPOI.HSSF.Record;

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
            //------------------------------
            MAX

        }
        //OptionData optionData = new OptionData();
        //ProductReader productReader = new ProductReader();
        //CostReader costReader = new CostReader();
        string curCostDataFilePath = null;
        bool bAltDataFlg = false;

        public FormMain()
        {
            InitializeComponent();

            var assembly = Assembly.GetExecutingAssembly().GetName();
            var ver = assembly.Version;

            // アセンブリ名 1.0.0.0
            this.Text = $"{assembly.Name} - {ver.Major}.{ver.Minor}.{ver.Build}.{ver.Revision}";

        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            LoadUserSetting();

            //現在のサイズでフォームの最小サイズを指定
           // this.MinimumSize = new Size(this.Width, this.Height);

            if (string.IsNullOrEmpty(Global.optionData.DataBasePath))
            {
                Utility.MessageConfirm("データベースのフォルダが未設定です。\nオプション画面からデータベースのフォルダを設定してください。", "データベースパス");
                EnalbeControl(false);
                return;
            }
            ReadDataBase(Global.optionData.DataBasePath);

            UpdateListFont();


            gridList.MouseWheel += OnGridList_MouseWheel;

        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (bAltDataFlg)
            {
                if (Utility.MessageConfirm("データが変更されています。\n\n保存しますか？") == DialogResult.OK)
                {
                    mnuSave_Click(null, null);
                }
            }

            SaveUserSetting();
            Global.optionData.SaveOptions();
        }

        private void LoadUserSetting()
        {
            Utility.LoadUserSetting(this, 
                                    Properties.Settings.Default.FrmMainLocX,
                                    Properties.Settings.Default.FrmMainLocY,
                                    Properties.Settings.Default.FrmMainSizeW,
                                    Properties.Settings.Default.FrmMainSizeH
                                    );
            int iCol = 0;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColChkW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColKindW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColNameW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColNumW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColCostRateW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColMaterialW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColWorkerlW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColPackageW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColPriceW;
            gridList.Columns[iCol++].Width = Properties.Settings.Default.ProductListColProfitW;
        }
        private void SaveUserSetting()
        {
            Properties.Settings.Default.FrmMainLocX = this.Location.X;
            Properties.Settings.Default.FrmMainLocY = this.Location.Y;
            Properties.Settings.Default.FrmMainSizeW = this.Size.Width;
            Properties.Settings.Default.FrmMainSizeH = this.Size.Height;
            int iCol = 0;
            Properties.Settings.Default.ProductListColChkW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColKindW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColNameW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColNumW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColCostRateW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColMaterialW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColWorkerlW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColPackageW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColPriceW = gridList.Columns[iCol++].Width;
            Properties.Settings.Default.ProductListColProfitW = gridList.Columns[iCol++].Width;

            Properties.Settings.Default.Save();
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
            Global.productReader.ReadExcel(productDataBase);
            //-------------------------------------------
            // 原価データを読み込む
            //-------------------------------------------
            Global.costReader.ReadExcel(costDataBase);

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
                    var product = Global.productReader.GetProductDataByID(costData.ProductId);
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

        void OnGridList_MouseWheel(object sender, MouseEventArgs e)
        {

            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (e.Delta > 0)
                {
                    Properties.Settings.Default.ProductListFontSize -= Const.prodListFontSizeInc;
                }
                else
                {
                    Properties.Settings.Default.ProductListFontSize += Const.prodListFontSizeInc;
                }

                UpdateListFont();
            }
        }
        /// <summary>
        /// グリッドのフォント設定
        /// </summary>
        private void UpdateListFont()
        {
            gridList.Font = new Font(gridList.Font.Name, Properties.Settings.Default.ProductListFontSize);
            int intRowHeight = (int)(float.Parse(gridList.Font.Size.ToString()) + 12);
            for (int i = 0; i < gridList.Rows.Count; i++)
            {
                gridList.Rows[i].Height = intRowHeight;
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

                if (value != null && (bool)value == true)
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
            //編集ありフラグセット
            bAltDataFlg = true;

        }

        /// <summary>
        /// メインリストに商品情報（コストオブジェクトを追加
        /// </summary>
        /// <param name="product"></param>
        public void AddProduct(ProductReader.ProductData product)
        {
            CostData costData = new CostData(product, Global.costReader);
            AddCostDataToRow(costData);
        }
        public void AddCostDataToRow(CostData costData)
        {
            //編集ありフラグセット
            bAltDataFlg = true;

            var index = gridList.Rows.Add();
            var row = gridList.Rows[index];

            row.Cells[(int)COL_INDEX.CHECK].Value = false;
            row.Cells[(int)COL_INDEX.KIND].Value = costData.Kind;
            row.Cells[(int)COL_INDEX.NAME].Value = costData.ID_Name(Global.optionData.DispIDtoList);
            row.Tag = costData;

            UpdateRowValues(row);


            //メインリストの分類コンボボックスのアイテム更新
            cmbMainListKind.Items.Clear();
            cmbMainListKind.Items.Add(Const.SelectAll);
            foreach (DataGridViewRow row2 in gridList.Rows)
            {
                var value = row2.Cells[(int)COL_INDEX.KIND].Value;

                if (cmbMainListKind.Items.IndexOf(value) < 0)
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

            int rc = costData.Calc(out costRate, out rowCostRate, out laborCostRate, out packageCostRate, out profit);

            row.Cells[(int)COL_INDEX.PRODUCT_NUM].Value = costData.ProductNum;
            row.Cells[(int)COL_INDEX.ROW_COST].Value = costRate * 100;
            row.Cells[(int)COL_INDEX.MATERIAL_COST].Value = rowCostRate * 100;
            row.Cells[(int)COL_INDEX.WORKER_COST].Value = laborCostRate * 100;
            row.Cells[(int)COL_INDEX.PACKAGE_COST].Value = packageCostRate * 100;
            if (costData.ProductNum != 0)
            {
                row.Cells[(int)COL_INDEX.PRICE].Value = costData.Price / costData.ProductNum; //1個辺りの定価
                row.Cells[(int)COL_INDEX.PROFIT].Value = profit / costData.ProductNum; //１個辺りの利益
            }else
            {
                row.Cells[(int)COL_INDEX.PRICE].Value =0; 
                row.Cells[(int)COL_INDEX.PROFIT].Value = 0; 
            }

            Color forColor = Color.Black;
            if (rc != 0)
            {   //文字の色を赤色にする
                forColor = Color.Red;
            }
            for (int i = 0; i < (int)COL_INDEX.MAX; i++)
            {
                row.Cells[i].Style.ForeColor = forColor;
            }
        }
        FormItemSelector frm = null;
        private void button1_Click(object sender, EventArgs e)
        {

            if (frm == null || frm.IsDisposed)
            {
                frm = new FormItemSelector(this, Global.productReader);
            }
            frm.Show();
            frm.TopMost = true;

        }

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in gridList.Rows)
            {
                row.Cells[(int)COL_INDEX.CHECK].Value = chkAll.Checked;
            }
        }

        private void gridList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var selectRows = gridList.SelectedRows;
            if (selectRows == null) return;

            var costData = (CostData)selectRows[0].Tag;

            FormEditCost frm = new FormEditCost(costData, Global.costReader, Global.productReader);
            if (frm.ShowDialog() == DialogResult.OK)
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
            string orgPath = Global.optionData.DataBasePath;

            bool orgDispIDtoList = Global.optionData.DispIDtoList;

            FormOption frm = new FormOption();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(Global.optionData.DataBasePath))
                {
                    Utility.MessageConfirm("データベースのフォルダが未設定です。\nオプション画面からデータベースのフォルダを設定してください。", "データベースパス");
                    EnalbeControl(false);
                    return;
                }
                EnalbeControl(true);

                //パスに変更が発生したら読み込みなおし
                if (orgPath != Global.optionData.DataBasePath)
                {
                    ReadDataBase(Global.optionData.DataBasePath);
                }
                //名称表示のID表示設定に変更があったか？
                if(orgDispIDtoList != Global.optionData.DispIDtoList)
                {
                    UpdateListName();
                }
            }
        }
        /// <summary>
        /// 一覧に表示されている商品名の名称表示更新
        /// </summary>
        void UpdateListName()
        {
            foreach (DataGridViewRow row in gridList.Rows)
            {
                CostData cost = (CostData)row.Tag;

                row.Cells[(int)COL_INDEX.NAME].Value = cost.ID_Name(Global.optionData.DispIDtoList);
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

            //編集フラグをリセット
            bAltDataFlg = false;
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
            curCostDataFilePath = GetSaveFileName(curCostDataFilePath);
            if (string.IsNullOrEmpty(curCostDataFilePath))
            {
                return;
            }
            SaveData(curCostDataFilePath);
        }
        private string GetSaveFileName(string initFilePath = null)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = initFilePath;
            dlg.AddExtension = true;
            dlg.DefaultExt = "yaml";
            dlg.Filter = "原価管理データ|*.yaml|全て|*.*";

            if (dlg.ShowDialog() == DialogResult.OK)
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
            dlg.FileName = curCostDataFilePath;
            dlg.AddExtension = true;
            dlg.DefaultExt = "yaml";
            dlg.Filter = "原価管理データ|*.yaml|全て|*.*";
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
                if (drags.Length > 1)
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
                data.AttachCostReader(Global.costReader);

                AddCostDataToRow(data);
            }
            //編集ありフラグリセット
            bAltDataFlg = false;


        }

        /// <summary>
        /// 印刷ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolBtnPrint_Click(object sender, EventArgs e)
        {
            CreatePrintData();

            PrintDocumentEx pd = CreatePrintDocument();
            pd.ResetPageIndex();

            FormPrintPreview frm = new FormPrintPreview(this, pd);
            frm.ShowDialog();
        }
    }
}
