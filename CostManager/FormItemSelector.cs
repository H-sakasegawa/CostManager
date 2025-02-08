using ExcelReaderUtility;
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

using static ExcelReaderUtility.CostReader;

namespace CostManager
{
    public partial class FormItemSelector : Form
    {
        enum GridVieCol
        {
            GVC_CHECK = 0,
            GVC_VALUE1,
            GVC_VALUE2,
        }
         enum enmSelectType
        {
            PRODUCT = 0,
            MATERIAL,
            WORKER,
            PACKAGE
        }
        FormMain frmMain=null;
        FormEditCost frmEditCost = null;
        ProductReader productBaseInfo = null;
        List<MaterialData> lstMaterial = null;
        List<WorkerData> lstWorker = null;
        List<PackageData> lstPackage = null;

        enmSelectType selectType;


        public FormItemSelector(FormMain frmMain, ProductReader productBaseInfo)
        {
            InitializeComponent();
            this.Text = "商品選択";
            this.frmMain = frmMain;
            this.productBaseInfo = productBaseInfo;

            selectType = enmSelectType.PRODUCT;

            //グリッド設定
            var columns = grdProductNames.Columns;
            columns.Clear();
            DataGridViewCheckBoxColumn colChk = new DataGridViewCheckBoxColumn();
            colChk.Name = "select";
            colChk.HeaderText = "選択";
            colChk.Width = 40;
            columns.Add(colChk);

            DataGridViewTextBoxColumn colText = new DataGridViewTextBoxColumn();
            colText.Name = "kind";
            colText.HeaderText = "分類";
            colText.ReadOnly = true;
            colText.Width = 70;
            columns.Add(colText);

            colText = new DataGridViewTextBoxColumn();
            colText.Name = "name";
            colText.HeaderText = "商品名";
            colText.ReadOnly = true;
            colText.Width = 300;
            columns.Add(colText);

            bool bDispID = Global.optionData.DispIDtoList;

            var lstProduct = productBaseInfo.GetProductList();
            foreach (var product in lstProduct)
            {
                var index = grdProductNames.Rows.Add(product.kind);
                var row = grdProductNames.Rows[index];
                row.Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                row.Cells[(int)GridVieCol.GVC_VALUE1].Value = product.kind;
                row.Cells[(int)GridVieCol.GVC_VALUE2].Value = product.ToString(bDispID);
                row.Tag = product;
            }
            int iCol = 0;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColChkW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColKindW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColNameW;

            //-------------------------------------------
            // コンボボックスアイテム設定
            //-------------------------------------------
            //種別コンボボックス
            var lstKind = productBaseInfo.GetKindList();
            cmbKind.Items.Add(Const.SelectAll);
            foreach (var s in lstKind)
            {
                cmbKind.Items.Add(s);
            }
            cmbKind.SelectedIndex = 0;
        }
        //原材料選択
        public FormItemSelector(FormEditCost frmEditCost, List<MaterialData> lstMaterial)
        {
            InitializeComponent();
            this.Text = "原材料選択";
            this.frmEditCost = frmEditCost;
            this.lstMaterial = lstMaterial;
            selectType = enmSelectType.MATERIAL;

            //グリッド設定
            var columns = grdProductNames.Columns;
            columns.Clear();
            DataGridViewCheckBoxColumn colChk = new DataGridViewCheckBoxColumn();
            colChk.Name = "select";
            colChk.HeaderText = "選択";
            colChk.Width = 40;
            columns.Add(colChk);

            DataGridViewTextBoxColumn colText = new DataGridViewTextBoxColumn();
            colText.Name = "kind";
            colText.HeaderText = "分類";
            colText.ReadOnly = true;
            colText.Width = 70;
            columns.Add(colText);

            colText = new DataGridViewTextBoxColumn();
            colText.Name = "name";
            colText.HeaderText = "原材料";
            colText.ReadOnly = true;
            colText.Width = 300;
            columns.Add(colText);

            bool bDispID = Global.optionData.DispIDtoList;
            foreach (var material in lstMaterial)
            {
                var index = grdProductNames.Rows.Add(material.kind);
                var row = grdProductNames.Rows[index];
                row.Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                row.Cells[(int)GridVieCol.GVC_VALUE1].Value = material.kind;
                row.Cells[(int)GridVieCol.GVC_VALUE2].Value = material.ToString(bDispID);
                row.Tag = material;
            }
            int iCol = 0;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColChkW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColKindW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColNameW;

            //-------------------------------------------
            // コンボボックスアイテム設定
            //-------------------------------------------
            //種別コンボボックス
            cmbKind.Items.Add(Const.SelectAll);
            foreach (var s in lstMaterial)
            {
                bool bSame = false;
                for(int i= 0; i< cmbKind.Items.Count; i++)
                {
                    if((string)cmbKind.Items[i] == s.kind)
                    {
                        bSame = true;
                        break;
                    }
                }
                if( !bSame) cmbKind.Items.Add(s.kind);
            }
            cmbKind.SelectedIndex = 0;
        }
        //作業者選択
        public FormItemSelector(FormEditCost frmEditCost, List<WorkerData> lstWorker)
        {
            InitializeComponent();
            this.Text = "作業者選択";
            this.frmEditCost = frmEditCost;
            this.lstWorker = lstWorker;
            cmbKind.Visible = false;
            lblKind.Visible = false;
            selectType = enmSelectType.WORKER;

            //グリッド設定
            var columns = grdProductNames.Columns;
            columns.Clear();
            DataGridViewCheckBoxColumn colChk = new DataGridViewCheckBoxColumn();
            colChk.Name = "select";
            colChk.HeaderText = "選択";
            colChk.Width = 40;
            columns.Add(colChk);

            DataGridViewTextBoxColumn colText = new DataGridViewTextBoxColumn();
            colText = new DataGridViewTextBoxColumn();
            colText.Name = "name";
            colText.HeaderText = "作業者";
            colText.ReadOnly = true;
            colText.Width = 200;
            columns.Add(colText);

            bool bDispID = Global.optionData.DispIDtoList;
            foreach (var worker in lstWorker)
            {
                var index = grdProductNames.Rows.Add();
                var row = grdProductNames.Rows[index];
                row.Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                row.Cells[(int)GridVieCol.GVC_VALUE1].Value = worker.ToString(bDispID);
                row.Tag = worker;
            }

            int iCol = 0;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColChkW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColNameW;

        }
        //包装材選択
        public FormItemSelector(FormEditCost frmEditCost, List<PackageData> lstPackage)
        {
            InitializeComponent();
            this.Text = "包装材選択";
            this.frmEditCost = frmEditCost;
            this.lstPackage = lstPackage;
            cmbKind.Visible = false;
            lblKind.Visible = false;
            selectType = enmSelectType.PACKAGE;
            //グリッド設定
            var columns = grdProductNames.Columns;
            columns.Clear();
            DataGridViewCheckBoxColumn colChk = new DataGridViewCheckBoxColumn();
            colChk.Name = "select";
            colChk.HeaderText = "選択";
            colChk.Width = 40;
            columns.Add(colChk);

            DataGridViewTextBoxColumn colText = new DataGridViewTextBoxColumn();
            colText = new DataGridViewTextBoxColumn();
            colText.Name = "name";
            colText.HeaderText = "包装材";
            colText.ReadOnly = true;
            colText.Width = 200;
            columns.Add(colText);

            bool bDispID = Global.optionData.DispIDtoList;
            foreach (var package in lstPackage)
            {
                var index = grdProductNames.Rows.Add();
                var row = grdProductNames.Rows[index];
                row.Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                row.Cells[(int)GridVieCol.GVC_VALUE1].Value = package.ToString(bDispID);
                row.Tag = package;
            }
            int iCol = 0;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColChkW;
            grdProductNames.Columns[iCol++].Width = Properties.Settings.Default.ItemSelectColNameW;
        }


        private void FormProductList_Load(object sender, EventArgs e)
        {
            //現在のサイズでフォームの最小サイズを指定
            this.MinimumSize = new Size(this.Width, this.Height);

            LoadUserSetting();
            UpdateListFont();

            grdProductNames.MouseWheel += LstProductNames_MouseWheel;

        }
        private void FormItemSelector_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveUserSetting();

        }

        private void LoadUserSetting()
        {
            Utility.LoadUserSetting(this,
                                   Properties.Settings.Default.FrmItemSelectLocX,
                                   Properties.Settings.Default.FrmItemSelectLocY,
                                   Properties.Settings.Default.FrmItemSelectSizeW,
                                   Properties.Settings.Default.FrmItemSelectSizeH
                                   );

        }
        private void SaveUserSetting()
        {
            Properties.Settings.Default.FrmItemSelectLocX = this.Location.X;
            Properties.Settings.Default.FrmItemSelectLocY = this.Location.Y;
            Properties.Settings.Default.FrmItemSelectSizeW = this.Size.Width;
            Properties.Settings.Default.FrmItemSelectSizeH = this.Size.Height;

            Properties.Settings.Default.ItemSelectColChkW = grdProductNames.Columns[(int)GridVieCol.GVC_CHECK].Width;
            switch(selectType)
            {
                case enmSelectType.PRODUCT:
                case enmSelectType.MATERIAL:
                    Properties.Settings.Default.ItemSelectColKindW = grdProductNames.Columns[(int)GridVieCol.GVC_VALUE1].Width;
                    Properties.Settings.Default.ItemSelectColNameW = grdProductNames.Columns[(int)GridVieCol.GVC_VALUE2].Width;
                    break;
                case enmSelectType.WORKER:
                case enmSelectType.PACKAGE:
                    Properties.Settings.Default.ItemSelectColNameW = grdProductNames.Columns[(int)GridVieCol.GVC_VALUE1].Width;
                    break;

            }
            Properties.Settings.Default.Save();
        }

        private void LstProductNames_MouseWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (e.Delta > 0)
                {
                    Properties.Settings.Default.ItemSelectFontSize -= Const.WheelInc;
                    if(Properties.Settings.Default.ItemSelectFontSize<0.5f)
                    {
                        Properties.Settings.Default.ItemSelectFontSize = 0.5f;
                    }
                }
                else
                {
                    Properties.Settings.Default.ItemSelectFontSize += Const.WheelInc;
                }

                UpdateListFont();
            }
        }
        /// <summary>
        /// グリッドのフォント設定
        /// </summary>
        private void UpdateListFont()
        {
            grdProductNames.Font = new Font(grdProductNames.Font.Name, Properties.Settings.Default.ItemSelectFontSize);
            int intRowHeight = (int)(float.Parse(grdProductNames.Font.Size.ToString()) + 12);
            for (int i = 0; i < grdProductNames.Rows.Count; i++)
            {
                grdProductNames.Rows[i].Height = intRowHeight;
            }
        }

        /// <summary>
        /// 分類コンボボックス
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            DoFilter();
        }

        /// <summary>
        /// 商品名フィルタ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            DoFilter();
        }

        private void DoFilter()
        { 
            foreach (DataGridViewRow row in grdProductNames.Rows)
            {
                string kind = null;
                string name = null;
                switch (selectType)
                {
                    case enmSelectType.PRODUCT:
                    case enmSelectType.MATERIAL:
                        kind = (string)row.Cells[(int)GridVieCol.GVC_VALUE1].Value;
                        name = (string)row.Cells[(int)GridVieCol.GVC_VALUE2].Value;
                        break;
                    case enmSelectType.WORKER:
                    case enmSelectType.PACKAGE:
                        name = (string)row.Cells[(int)GridVieCol.GVC_VALUE1].Value;
                        break;

                }

                row.Visible = true;
                //種別によるフィルタ
                if ( kind!=null)
                {
                    if (cmbKind.Text != Const.SelectAll)
                    {
                        if(cmbKind.Text != kind)
                        {
                            row.Visible = false;
                        }
                    }
                }
                //名称フィルタ
                if (!string.IsNullOrEmpty(txtFilter.Text))
                {
                    if (name.IndexOf(txtFilter.Text) < 0)
                    {
                        row.Visible = false;
                    }
                }
            }
        }


        private void btnAdd_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> lstCheckedRow = new List<DataGridViewRow>();
            //チェックボックスがONの行があるか？
            var rows = grdProductNames.Rows;
            for (int iRow = 0; iRow < rows.Count; iRow++)
            {
                object value = (bool)(rows[iRow].Cells[(int)GridVieCol.GVC_CHECK].Value);

                if (value != null && (bool)value == true)
                {
                    lstCheckedRow.Add(rows[iRow]);
                }
            }
            if (lstCheckedRow.Count <=0)
            {
                var selectRows = grdProductNames.SelectedRows;
                if (selectRows == null) return;

                for (int i = 0; i < selectRows.Count; i++)
                {
                    lstCheckedRow.Add(selectRows[i]);
                }
            }

            bool bRc = true;;
            foreach(var row in lstCheckedRow)
            {
                switch (selectType)
                {
                    case enmSelectType.PRODUCT:
                        {
                            var value = (ProductReader.ProductData)row.Tag;
                            frmMain.AddProduct(value);
                        }
                        break;
                    case enmSelectType.MATERIAL:
                        {
                            var value = (MaterialData)row.Tag;
                            bRc = frmEditCost.AddMaterial(value);
                        }
                        break;
                    case enmSelectType.WORKER:
                        {
                            var value = (WorkerData)row.Tag;
                            frmEditCost.AddWorker(value);
                        }
                        break;
                    case enmSelectType.PACKAGE:
                        {
                            var value = (PackageData)row.Tag;
                            frmEditCost.AddPackage(value);
                        }
                        break;
                }
                if( !bRc)
                {
                    break;
                }
            }


        }



        //商品名 スペースキー選択
        private void grdProductNames_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                btnAdd_Click(null, null);
            }
        }

        private void grdProductNames_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            btnAdd_Click(null, null);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in grdProductNames.Rows)
            {
                row.Cells[(int)GridVieCol.GVC_CHECK].Value = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
