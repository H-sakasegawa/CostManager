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
        OptionData optionData;
        public FormOption(OptionData optionData)
        {
            InitializeComponent();
            this.optionData = optionData;
        }

        private void FormOption_Load(object sender, EventArgs e)
        {
            txtDataBasePath.Text = optionData.DataBasePath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(! Directory.Exists(txtDataBasePath.Text))
            {
                Utility.MessageError("指定されたデータベースフォルダは存在しません");
                return;
            }
            optionData.DataBasePath = txtDataBasePath.Text;

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
            dlg.Path = optionData.DataBasePath;
            dlg.Title = "商品データベース格納フォルダを選択";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtDataBasePath.Text = dlg.Path;
            }

        }
    }
}
