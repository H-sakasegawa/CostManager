﻿using ExcelReaderUtility;
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
using static ExcelReaderUtility.CostReader;

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

        CostReader costBaseInfo;
        CostData costData;
        CostData costDataBk = new CostData();
        public FormEditCost(CostData costData, CostReader costBaseInfo)
        {
            InitializeComponent();
            this.costData = costData;
            costData.CopyTo( costDataBk );

            this.costBaseInfo = costBaseInfo;
        }

        private void FormEditCost_Load(object sender, EventArgs e)
        {
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

        void DispCostData()
        {
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
                var material = costBaseInfo.GetMaterialtDataByID(val.ID);
                int index = grdRowMaterial.Rows.Add();
                var row = grdRowMaterial.Rows[index];
                grdRowMaterial.Columns[(int)COL_MATERIAL.AMOUNT].DefaultCellStyle.BackColor = Color.LightBlue;

                row.Cells[(int)COL_MATERIAL.CHECK].Value = false;
                row.Cells[(int)COL_MATERIAL.KIND].Value = material.kind;
                row.Cells[(int)COL_MATERIAL.NAME].Value = material.name;
                row.Cells[(int)COL_MATERIAL.AMOUNT].Value = val.AmountUsed; //使用量
                row.Cells[(int)COL_MATERIAL.COST].Value = material.cost * val.AmountUsed; //原価 × val.amountUsed
                row.Tag = val;
            }
            //作業者
            foreach (var val in costData.LstWorkerCost)
            {
                //作業者名から作業者情報を取得
                var worker = costBaseInfo.GetWorkerDataByID(val.ID);
                int index = grdWorker.Rows.Add();
                var row = grdWorker.Rows[index];
                grdWorker.Columns[(int)COL_WORKER.TIME].DefaultCellStyle.BackColor = Color.LightBlue;

                row.Cells[(int)COL_WORKER.CHECK].Value = false;
                row.Cells[(int)COL_WORKER.NAME].Value = worker.name;
                row.Cells[(int)COL_WORKER.TIME].Value = val.WorkingTime; //時間（分)
                row.Cells[(int)COL_WORKER.COST].Value = (float)(worker.hourlyPay / 60.0) * val.WorkingTime; //時給 / 60 × val.workingTime
                row.Tag = val;
            }

            //包装材
            foreach (var val in costData.LstPackageCost)
            {
                //包装材から包装材情報を取得
                var package = costBaseInfo.GetPackageDataByID(val.ID);
                int index = grdPackage.Rows.Add();
                var row = grdPackage.Rows[index];

                row.Cells[(int)COL_PACKAGE.CHECK].Value = false;
                row.Cells[(int)COL_PACKAGE.NAME].Value = package.name;
                row.Cells[(int)COL_PACKAGE.NUM].Value = val.RequiredNum; //必要数
                row.Cells[(int)COL_PACKAGE.COST].Value = package.cost * val.RequiredNum; //単価 × val.requiredNum
                row.Tag = val;
            }
            //ラベル情報更新
            UpdateLabelData();

        }

        void UpdateLabelData()
        {
            //原材料費
            lblCostMaterial.Text = costData.GetMateralCost().ToString("C");
            //人件費
            lblCostWorker.Text = costData.GetWorkerCost().ToString("C");
            //パッケージ費
            lblCostPackage.Text = costData.GetPackageCost().ToString("C");

            var costAll = costData.GetAllCost();
            //総原価
            lblCostAll.Text = costData.GetAllCost().ToString("C");
            //製品単価
            if (costData.ProductNum > 0)
            {
                lblCostOne.Text = (costAll / costData.ProductNum).ToString("C");
            }else
            {
                lblCostOne.Text = (0).ToString();
            }
            var allCost = costData.GetAllCost();
            if (allCost < costData.Price && costData.Price!=0)
            {
                lblCostRate.ForeColor = Color.Black;
                lblProfitRate.ForeColor = Color.Black;
                //原価率
                lblCostRate.Text = ((allCost / costData.Price)).ToString("P2");// "%"
                 //利益率
                lblProfitRate.Text = (costData.GetProfitRate()).ToString("P2");// "%"
            }else
            {
                lblCostRate.ForeColor = Color.Red;
                lblProfitRate.ForeColor = Color.Red;
                //原価率
                lblCostRate.Text = (0).ToString("P2");// "%"
                //利益率
                lblProfitRate.Text = (0).ToString("P2");// "%"
            }
        }

        /// <summary>
        /// 原材料追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddMaterial_Click(object sender, EventArgs e)
        {
            FormItemSelector frm = new FormItemSelector( this, costBaseInfo.GetMaterialList());
            frm.ShowDialog();
        }
        /// <summary>
        /// 原材料削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelMaterial_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdRowMaterial.Rows);
            foreach( var row in rows)
            {
                costData.LstMaterialCost.Remove((MaterialCost)row.Tag);
                grdRowMaterial.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }

        List<DataGridViewRow> GetSelectRows(DataGridViewRowCollection rows)
        {
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
                var selectRows = grdRowMaterial.SelectedRows;
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
        /// 作業者追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddWorker_Click(object sender, EventArgs e)
        {
            FormItemSelector frm = new FormItemSelector(this, costBaseInfo.GetWorkerList());
            frm.ShowDialog();
        }
        /// <summary>
        /// 作業者削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelWorker_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdWorker.Rows);
            foreach (var row in rows)
            {
                costData.LstWorkerCost.Remove((WorkerCost)row.Tag);
                grdWorker.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }
        /// <summary>
        /// 包装材追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddPackage_Click(object sender, EventArgs e)
        {
            FormItemSelector frm = new FormItemSelector(this, costBaseInfo.GetPackageList());
            frm.ShowDialog();
        }
        /// <summary>
        /// 包装材削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelPackage_Click(object sender, EventArgs e)
        {
            var rows = GetSelectRows(grdPackage.Rows);
            foreach (var row in rows)
            {
                costData.LstPackageCost.Remove((PackageCost)row.Tag);
                grdPackage.Rows.Remove(row);
            }
            //ラベル情報更新
            UpdateLabelData();
        }
        /// <summary>
        /// 原材料一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public void AddMaterial(MaterialData data)
        {
            MaterialCost cost = new MaterialCost(data.id, 0, costBaseInfo);
            costData.LstMaterialCost.Add(cost);

            var index = grdRowMaterial.Rows.Add();
            var row = grdRowMaterial.Rows[index];

            row.Cells[(int)COL_MATERIAL.CHECK].Value = false;
            row.Cells[(int)COL_MATERIAL.KIND].Value = data.kind;
            row.Cells[(int)COL_MATERIAL.NAME].Value = data.name;
            row.Tag = cost;

            //ラベル情報更新
            UpdateLabelData();


        }
        /// <summary>
        /// 作業者一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public void AddWorker(WorkerData data)
        {
            WorkerCost cost = new WorkerCost(data.id, 0, costBaseInfo);
            costData.LstWorkerCost.Add(cost);

            var index = grdRowMaterial.Rows.Add();
            var row = grdRowMaterial.Rows[index];

            row.Cells[(int)COL_WORKER.CHECK].Value = false;
            row.Cells[(int)COL_WORKER.NAME].Value = data.name;
            row.Tag = cost;
            //ラベル情報更新
            UpdateLabelData();

        }
        /// <summary>
        /// 包装材一覧からの追加要求処理
        /// </summary>
        /// <param name="data"></param>
        public void AddPackage(PackageData data)
        {
            PackageCost cost = new PackageCost(data.id, 0, costBaseInfo);
            costData.LstPackageCost.Add(cost);

            var index = grdRowMaterial.Rows.Add();
            var row = grdRowMaterial.Rows[index];

            row.Cells[(int)COL_PACKAGE.CHECK].Value = false;
            row.Cells[(int)COL_PACKAGE.NAME].Value = data.name;
            row.Tag = cost;
            //ラベル情報更新
            UpdateLabelData();
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
            costDataBk.CopyTo(costData);

            DialogResult = DialogResult.Cancel;
            this.Close();
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
                data.AmountUsed = Convert.ToUInt32(row.Cells[e.ColumnIndex].Value);

                row.Cells[(int)COL_MATERIAL.COST].Value = data.CalcCost();
            }
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
            lblPriceSum.Text = costData.Price.ToString("C");
            //ラベル情報更新
            UpdateLabelData();

        }

    }
}
