using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelReaderUtility;

using static ExcelReaderUtility.CostReader;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Eval;

namespace CostManager
{
    public partial class FormEditCost : Form
    {
        enum COL_MATERIAL
        {
            CHECK = 0,
            KIND,
            NAME,
            AMOUNT,
            COST,
        }
        enum COL_WORKER
        {
            CHECK = 0,
            NAME,
            TIME,
            COST,
        }
        enum COL_PACKAGE
        {
            CHECK = 0,
            NAME,
            NUM,
            COST,
        }

        ProductReader productReader;
        CostReader costReader;
        CostData costData;
        CostData costDataBk = new CostData();

        FormItemSelector frmMaterial = null;
        FormItemSelector frmWorker = null;
        FormItemSelector frmPackage = null;


        public FormEditCost(CostData costData, CostReader costReader, ProductReader productReader)
        {
            InitializeComponent();
            this.costData = costData;
            costData.CopyTo( costDataBk );

            this.costReader = costReader;
            this.productReader = productReader;
        }

        private void FormEditCost_Load(object sender, EventArgs e)
        {
            LoadUserSetting();

            splitContainer1.Dock = DockStyle.Fill;
            splitContainer2.Dock = DockStyle.Fill;
            splitContainer3.Dock = DockStyle.Fill;

            grdRowMaterial.Dock = DockStyle.Fill;
            grdWorker.Dock = DockStyle.Fill;
            grdPackage.Dock = DockStyle.Fill;

            //編集列に色付け
            grdRowMaterial.Columns[(int)COL_MATERIAL.AMOUNT].DefaultCellStyle.BackColor = Color.LightBlue;
            grdWorker.Columns[(int)COL_WORKER.TIME].DefaultCellStyle.BackColor = Color.LightBlue;

            grdRowMaterial.CellValueChanged -= grdRowMaterial_CellValueChanged;
            grdWorker.CellValueChanged -= grdWorker_CellValueChanged;
            DispCostData();
            grdRowMaterial.CellValueChanged += grdRowMaterial_CellValueChanged;
            grdWorker.CellValueChanged += grdWorker_CellValueChanged;

            //現在のサイズでフォームの最小サイズを指定
            this.MinimumSize = new Size(this.Width, this.Height);

        }

        private void FormEditCost_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.Cancel)
            {
                if (!costData.IsSame(costDataBk))
                {
                    if (Utility.MessageConfirm("データが変更されています。\n\n編集内容を設定しますか？") == DialogResult.OK)
                    {
                        DialogResult = DialogResult.OK;
                    }else
                    {
                        //編集前の状態に戻す
                        costDataBk.CopyTo(costData);
                    }
                }
            }
            
            SaveUserSetting();
        }


        private void LoadUserSetting()
        {
            Utility.LoadUserSetting(this,
                                   Properties.Settings.Default.FrmEditCostLocX,
                                   Properties.Settings.Default.FrmEditCostLocY,
                                   Properties.Settings.Default.FrmEditCostSizeW,
                                   Properties.Settings.Default.FrmEditCostSizeH
                                   );
        }
        private void SaveUserSetting()
        {
            Properties.Settings.Default.FrmEditCostLocX = this.Location.X;
            Properties.Settings.Default.FrmEditCostLocY = this.Location.Y;
            Properties.Settings.Default.FrmEditCostSizeW = this.Size.Width;
            Properties.Settings.Default.FrmEditCostSizeH = this.Size.Height;
            Properties.Settings.Default.Save();
        }
        void DispCostData()
        {
            bool bDispID = Global.optionData.DispIDtoList;

            lblProductName.Text = Utility.RemoveCRLF(costData.ProductName);
            txtMakeNum.Text     = costData.ProductNum.ToString();
            lblPriceSum.Text = costData.Price.ToString();
            if (costData.ProductNum > 0)
            {
                txtPrice.Text = (costData.Price / costData.ProductNum).ToString();
            }else
            {
                txtPrice.Text = (0).ToString();
            }
            //原材料
            foreach (var val in costData.LstMaterialCost)
            {
                //原材料IDから原材料情報を取得
                var material = costReader.GetMaterialtDataByID(val.ID);
                int index = grdRowMaterial.Rows.Add();
                var row = grdRowMaterial.Rows[index];
                grdRowMaterial.Columns[(int)COL_MATERIAL.AMOUNT].DefaultCellStyle.BackColor = Color.LightBlue;

                row.Tag = val;
                row.Cells[(int)COL_MATERIAL.CHECK].Value = false;
                row.Cells[(int)COL_MATERIAL.KIND].Value = material.kind;
                row.Cells[(int)COL_MATERIAL.NAME].Value = material.ToString(bDispID);
                row.Cells[(int)COL_MATERIAL.AMOUNT].Value = val.AmountUsed; //使用量
                row.Cells[(int)COL_MATERIAL.COST].Value = val.CalcCost(); //原価 × val.amountUsed
            }
            //作業者
            foreach (var val in costData.LstWorkerCost)
            {
                //作業者名から作業者情報を取得
                var worker = costReader.GetWorkerDataByID(val.ID);
                int index = grdWorker.Rows.Add();
                var row = grdWorker.Rows[index];
                grdWorker.Columns[(int)COL_WORKER.TIME].DefaultCellStyle.BackColor = Color.LightBlue;

                row.Tag = val;
                row.Cells[(int)COL_WORKER.CHECK].Value = false;
                row.Cells[(int)COL_WORKER.NAME].Value = worker.ToString(bDispID);
                row.Cells[(int)COL_WORKER.TIME].Value = val.WorkingTime; //時間（分)
                row.Cells[(int)COL_WORKER.COST].Value = (float)(worker.hourlyPay / 60.0) * val.WorkingTime; //時給 / 60 × val.workingTime
            }

            //包装材
            foreach (var val in costData.LstPackageCost)
            {
                //包装材から包装材情報を取得
                var package = costReader.GetPackageDataByID(val.ID);
                int index = grdPackage.Rows.Add();
                var row = grdPackage.Rows[index];

                row.Tag = val;
                row.Cells[(int)COL_PACKAGE.CHECK].Value = false;
                row.Cells[(int)COL_PACKAGE.NAME].Value = package.ToString(bDispID);
                row.Cells[(int)COL_PACKAGE.NUM].Value = val.RequiredNum; //必要数
                row.Cells[(int)COL_PACKAGE.COST].Value = package.cost * val.RequiredNum; //単価 × val.requiredNum
            }
            //ラベル情報更新
            UpdateLabelData();

        }



        void UpdateLabelData()
        {
            //原材料費
            lblCostMaterial.Text = costData.GetMateralCost().ToString(Const.strFmtC);
            //人件費
            lblCostWorker.Text = costData.GetWorkerCost().ToString(Const.strFmtC);
            //パッケージ費
            lblCostPackage.Text = costData.GetPackageCost().ToString(Const.strFmtC);

            var costAll = costData.GetAllCost();
            //総原価
            lblCostAll.Text = costData.GetAllCost().ToString(Const.strFmtC);
            //製品単価
            lblCostOne.Text = costData.GetCostOne().ToString(Const.strFmtC);

            var allCost = costData.GetAllCost();
            if (allCost < costData.Price && costData.Price!=0)
            {
                lblCostRate.ForeColor = Color.Black;
                lblProfitRate.ForeColor = Color.Black;
            }else
            {
                lblCostRate.ForeColor = Color.Red;
                lblProfitRate.ForeColor = Color.Red;
            }
            //原価率
            lblCostRate.Text = ((allCost / costData.Price)).ToString("P2");// "%"
                                                                           //利益率
            lblProfitRate.Text = (costData.GetProfitRate()).ToString("P2");// "%"
            var product = productReader.GetProductDataByID(costData.ProductId);

            //栄養成分
            lblCalorie.Text = product.Calorie;
            lblProtein.Text = product.Protein;
            lblLipids.Text = product.Lipids;
            lblCarbohydrates.Text = product.Carbohydrates;
            lblSalt.Text = product.Salt;
        }

        //============================================================
        //  原材料
        //============================================================
        /// <summary>
        /// 原材料追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddMaterial_Click(object sender, EventArgs e)
        {
            if (frmMaterial== null || frmMaterial.IsDisposed)
            {
                frmMaterial = new FormItemSelector(this, costReader.GetMaterialList());
            }
            frmMaterial.Show();
            frmMaterial.TopMost = true;
        }
        /// <summary>
        /// 原材料削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelMaterial_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdRowMaterial);
            foreach( var row in rows)
            {
                costData.LstMaterialCost.Remove((MaterialCost)row.Tag);
                grdRowMaterial.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }


        /// <summary>
        /// 原材料一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public bool AddMaterial(MaterialData data)
        {
            //重複チェック
            if(costData.LstMaterialCost.Find(x=> x.ID == data.id)!=null)
            {
                Utility.MessageError($"同じ原材料が登録済みです。\n\n{data.ToString(true)}");
                return false;
            }
            MaterialCost cost = new MaterialCost(data.id, 0, costReader);
            costData.LstMaterialCost.Add(cost);

            var index = grdRowMaterial.Rows.Add();
            var row = grdRowMaterial.Rows[index];

            row.Tag = cost;
            row.Cells[(int)COL_MATERIAL.CHECK].Value = false;
            row.Cells[(int)COL_MATERIAL.KIND].Value = data.kind;
            row.Cells[(int)COL_MATERIAL.NAME].Value = data.ToString(Global.optionData.DispIDtoList);
            row.Cells[(int)COL_MATERIAL.AMOUNT].Value = 0;

            //ラベル情報更新
            UpdateLabelData();

            return true;


        }
        /// <summary>
        /// 原材料 資料量変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void grdRowMaterial_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex == (int)COL_MATERIAL.AMOUNT)
            {
                var row = grdRowMaterial.Rows[e.RowIndex];
                MaterialCost data = (MaterialCost)row.Tag;

                data.AmountUsed = Convert.ToSingle(row.Cells[e.ColumnIndex].Value);

                row.Cells[(int)COL_MATERIAL.COST].Value = data.CalcCost();
            }
            //ラベル情報更新
            UpdateLabelData();
        }
        private void grdRowMaterial_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == (int)COL_MATERIAL.AMOUNT)
            {
                float value = 0;
                if (float.TryParse(e.FormattedValue.ToString(), out value) == false)
                {
                    Utility.MessageError("使用量の入力値が正しくありません");
                    e.Cancel = true;
                }
            }

        }
        //============================================================
        //  作業者
        //============================================================
        /// <summary>
        /// 作業者追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddWorker_Click(object sender, EventArgs e)
        {
            if (frmWorker == null || frmWorker.IsDisposed)
            {
                frmWorker = new FormItemSelector(this, costReader.GetWorkerList());
            }
            frmWorker.Show();
            frmWorker.TopMost = true;
        }
        /// <summary>
        /// 作業者削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelWorker_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdWorker);
            foreach (var row in rows)
            {
                costData.LstWorkerCost.Remove((WorkerCost)row.Tag);
                grdWorker.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }
        /// <summary>
        /// 作業者一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public void AddWorker(WorkerData data)
        {
            WorkerCost cost = new WorkerCost(data.id, 0, costReader);
            costData.LstWorkerCost.Add(cost);

            var index = grdWorker.Rows.Add();
            var row = grdWorker.Rows[index];

            row.Tag = cost;
            row.Cells[(int)COL_WORKER.CHECK].Value = false;
            row.Cells[(int)COL_WORKER.NAME].Value = data.ToString(Global.optionData.DispIDtoList);
            row.Cells[(int)COL_WORKER.TIME].Value = 0;
            //ラベル情報更新
            UpdateLabelData();

        }
        /// <summary>
        /// 作業者 作業時間変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void grdWorker_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (e.ColumnIndex == (int)COL_WORKER.TIME)
            {
                var row = grdWorker.Rows[e.RowIndex];
                WorkerCost data = (WorkerCost)row.Tag;
                data.WorkingTime = Convert.ToUInt32(row.Cells[e.ColumnIndex].Value);

                row.Cells[(int)COL_WORKER.COST].Value = data.CalcCost();
            }
            //ラベル情報更新
            UpdateLabelData();
        }
        private void grdWorker_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == (int)COL_WORKER.TIME)
            {
                uint value = 0;
                if (uint.TryParse(e.FormattedValue.ToString(), out value) == false)
                {
                    Utility.MessageError("作業時間の入力値は正の整数値で入力してください");
                    e.Cancel = true;
                }
            }
        }
        //============================================================
        //  包装材
        //============================================================
        /// <summary>
        /// 包装材追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddPackage_Click(object sender, EventArgs e)
        {

            if (frmPackage == null || frmPackage.IsDisposed)
            {
                frmPackage = new FormItemSelector(this, costReader.GetPackageList());
            }
            frmPackage.Show();
            frmPackage.TopMost = true;

        }
        /// <summary>
        /// 包装材削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelPackage_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdPackage);
            foreach (var row in rows)
            {
                costData.LstPackageCost.Remove((PackageCost)row.Tag);
                grdPackage.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }

        /// <summary>
        /// 包装材一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public bool AddPackage(PackageData data)
        {
            //重複チェック
            if (costData.LstPackageCost.Find(x => x.ID == data.id) != null)
            {
                Utility.MessageError($"同じ包装材が登録済みです。\n\n{data.ToString(true)}");
                return false;
            }


            PackageCost packageCost = new PackageCost(data.id, costData.ProductNum, costReader);
            costData.LstPackageCost.Add(packageCost);

            var index = grdPackage.Rows.Add();
            var row = grdPackage.Rows[index];

            row.Tag = packageCost;
            row.Cells[(int)COL_PACKAGE.CHECK].Value = false;
            row.Cells[(int)COL_PACKAGE.NAME].Value = data.ToString(Global.optionData.DispIDtoList);
            row.Cells[(int)COL_PACKAGE.NUM].Value = costData.ProductNum;
            row.Cells[(int)COL_PACKAGE.COST].Value = data.cost * costData.ProductNum; //単価 × val.requiredNum;
            //ラベル情報更新
            UpdateLabelData();

            return true;
        }

        //=============================================================
        //=============================================================
        //=============================================================

        /// <summary>
        /// グリッドの選択行一覧取得
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        List<DataGridViewRow> GetSelectRows(DataGridView grid)
        {
            var rows = grid.Rows;
            List<DataGridViewRow> lstRows = new List<DataGridViewRow>();
            for (int iRow = 0; iRow < rows.Count; iRow++)
            {
                object value = (bool)(rows[iRow].Cells[0].Value);

                if (value != null && (bool)value == true)
                {
                    lstRows.Add(rows[iRow]);
                }
            }
            if (lstRows.Count <= 0)
            {
                var selectRows = grid.SelectedRows;
                if (selectRows != null)
                {
                    for (int i = 0; i < selectRows.Count; i++)
                    {
                        lstRows.Add(selectRows[i]);
                    }
                }
            }
            return lstRows;
        }

        /// <summary>
        /// OKボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }
        /// <summary>
        /// キャンセルボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }



        /// <summary>
        /// 制作数変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtMakeNum_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMakeNum.Text)) return;

            uint value;
            if( !uint.TryParse(txtMakeNum.Text, out value))
            {
                Utility.MessageError("制作数に不正な文字が入力されました");
                return;
            }
            //制作数
            costData.ProductNum = value;
            //包装材の必要数を更新
            foreach(DataGridViewRow row in grdPackage.Rows)
            {
                PackageCost data = (PackageCost)row.Tag;
                data.RequiredNum = value;
                row.Cells[(int)COL_PACKAGE.NUM].Value = value;
                row.Cells[(int)COL_PACKAGE.COST].Value = data.CalcCost(value);
            }
            //ラベル情報更新
            UpdateLabelData();
        }

        private void txtPrice_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPrice.Text)) return;

            uint value;
            if (!uint.TryParse(txtPrice.Text, out value))
            {
                Utility.MessageError("定価(1個辺り)に不正な文字が入力されました");
                return;
            }
            //制作数
            costData.Price = value * costData.ProductNum;
            //定価
            lblPriceSum.Text = costData.Price.ToString(Const.strFmtC);
            //ラベル情報更新
            UpdateLabelData();

        }

    }
}
