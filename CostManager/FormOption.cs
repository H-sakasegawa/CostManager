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
using ExcelReaderUtility;

namespace CostManager
{
    public partial class FormOption : Form
    {
        public FormOption()
        {
            InitializeComponent();
        }

        private void FormOption_Load(object sender, EventArgs e)
        {
            txtDataBasePath.Text = Global.optionData.DataBasePath;
            chkDispIDtoList.Checked = Global.optionData.DispIDtoList;
        }
        private void FormOption_FormClosing(object sender, FormClosingEventArgs e)
        {
            Global.optionData.SaveOptions();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(txtDataBasePath.Text))
            {
                Utility.MessageError("指定されたデータベースフォルダは存在しません");
                return;
            }
            Global.optionData.DataBasePath = txtDataBasePath.Text;

            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderSelectDialog dlg = new FolderSelectDialog();
            dlg.Path = Global.optionData.DataBasePath;
            dlg.Title = "商品データベース格納フォルダを選択";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtDataBasePath.Text = dlg.Path;
            }

        }
        private void chkDispIDtoList_CheckedChanged(object sender, EventArgs e)
        {
            Global.optionData.DispIDtoList = chkDispIDtoList.Checked;
        }

    }
}
