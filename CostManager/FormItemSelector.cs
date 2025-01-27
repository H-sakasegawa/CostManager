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

            foreach (var worker in lstWorker)
            {
                var index = grdProductNames.Rows.Add();
                grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE1].Value = worker.name;
                grdProductNames.Rows[index].Tag = worker;
            }


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

            foreach (var package in lstPackage)
            {
                var index = grdProductNames.Rows.Add();
                grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE1].Value = package.name;
                grdProductNames.Rows[index].Tag = package;
            }
        }


        private void FormProductList_Load(object sender, EventArgs e)
        {
            //現在のサイズでフォームの最小サイズを指定
            this.MinimumSize = new Size(this.Width, this.Height);


        }
        private void cmbKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            //選択された種別の商品名をグリッドに設定
            grdProductNames.Rows.Clear();

            if (selectType == enmSelectType.PRODUCT)
            {
                var lstProduct = productBaseInfo.GetProductList(cmbKind.Text);
                foreach (var product in lstProduct)
                {
                    var index = grdProductNames.Rows.Add(product.kind);
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE1].Value = product.kind;
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE2].Value = product.name;
                    grdProductNames.Rows[index].Tag = product;
                }
            }else if(selectType == enmSelectType.MATERIAL)
            {
                foreach (var material in lstMaterial)
                {
                    if(cmbKind.Text != Const.SelectAll && material.kind != cmbKind.Text)
                    {
                        continue;
                    }
                    var index = grdProductNames.Rows.Add(material.kind);
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_CHECK].Value = false;
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE1].Value = material.kind;
                    grdProductNames.Rows[index].Cells[(int)GridVieCol.GVC_VALUE2].Value = material.name;
                    grdProductNames.Rows[index].Tag = material;
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
                            frmEditCost.AddMaterial(value);
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
