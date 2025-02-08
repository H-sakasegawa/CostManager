using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using NPOI.SS.Formula.Functions;
using ExcelReaderUtility;
using NPOI.XSSF.Streaming.Values;
using static NPOI.HSSF.Util.HSSFColor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Runtime.ConstrainedExecution;
using System.Data;
using static CostManager.DrawUtil2;
using static System.Windows.Forms.AxHost;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.HPSF;
using System.Runtime.InteropServices;
using YamlDotNet.Core.Tokens;

namespace CostManager
{
    public partial class FormMain : Form
    {

        public enum PrintDataType
        {
            GRID_LIST,  //グリッド一覧印刷
            PRODUCT_COST_INFO   //商品毎のコスト情報
        }

        /// <summary>
        /// 印刷ドキュメント拡張クラス
        /// </summary>
        public class PrintDocumentEx : System.Drawing.Printing.PrintDocument
        {
            public PrintDocumentEx()
                : base()
            {
            }

            public void ResetPageIndex()
            {
                printGridDataIndex = 0;
                curPrintPageNo = 0;
            }

            public int printGridDataIndex;
            public int curPrintPageNo = 0;

            public PrintDataType printDataType;

        }

        List<CostData> lstPrintData = new List<CostData>();

        PaperKind paperKind = System.Drawing.Printing.PaperKind.A4;

        float A4HeightMM = 0;
        float A4WidthMM = 0;
        float PrintGapLeft = 8; 
        float PrintGapTop = 10;
        float rowHeight = 4; //表のセルの高さ
        float FooterHeight = 15;
        /// <summary>
        /// 原材料のグリッドの横領域サイズ
        /// </summary>
        float materialGridW = 0;

        int pageNum;

        /// <summary>
        /// 印刷ドキュメント情報作成
        /// </summary>
        /// <returns></returns>
        private PrintDocumentEx CreatePrintDocument()
        {

            PrintDocumentEx pd = new PrintDocumentEx();

            var ps = new System.Drawing.Printing.PrinterSettings();


            //A4用紙
            foreach (System.Drawing.Printing.PaperSize psize in pd.PrinterSettings.PaperSizes)
            {
                if (psize.Kind == paperKind)
                {
                    pd.DefaultPageSettings.PaperSize = psize;
                    break;
                }
            }
            pd.DefaultPageSettings.Landscape = false; //縦

            A4HeightMM = (float)DrawUtil2.inch_to_mm(pd.DefaultPageSettings.PaperSize.Height / 100.0); //mm
            A4WidthMM = (float)DrawUtil2.inch_to_mm(pd.DefaultPageSettings.PaperSize.Width / 100.0); //mm


            pageNum = GetPageNum();


            //PrintPageイベントハンドラの追加
            pd.PrintPage +=
                new System.Drawing.Printing.PrintPageEventHandler(pd_PrintPage);

            return pd;
        }




        private void CreatePrintData()
        {
            lstPrintData.Clear();
            //printDataIndex = 0; //印刷ラベルIndex初期化

            //印刷データ収集
            foreach (DataGridViewRow row in gridList.Rows)
            {
                CostData data = (CostData)row.Tag;

                lstPrintData.Add(data);
            }
        }


        private void pd_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            PrintDocumentEx pd = (PrintDocumentEx)sender;
            pd.curPrintPageNo++;



            if (pd.printDataType == PrintDataType.GRID_LIST)
            {
                //========================================
                // 原価管理グリッドデータの印刷
                //========================================

                e.HasMorePages = DrawGridList(pd, e.Graphics);

            }
            else if (pd.printDataType == PrintDataType.PRODUCT_COST_INFO)
            {
                //========================================
                // 各商品毎の原価の印刷
                //========================================
                e.HasMorePages = DrawProductCost(pd, e.Graphics);
            }

            DrawFooter(e.Graphics, pd.curPrintPageNo, pageNum );

        }

        bool DrawGridList(PrintDocumentEx pd, Graphics gPreview)
        {

            List<DrawGridInf> lstDrawGridInf = new List<DrawGridInf>()
            {
                new DrawGridInf(15, "分類","", StringAlignment.Near),
                new DrawGridInf(60, "商品名","", StringAlignment.Near),
                new DrawGridInf(15, "制作数","", StringAlignment.Far),
                new DrawGridInf(15, "原価率","", StringAlignment.Far),
                new DrawGridInf(15, "原材料","", StringAlignment.Far),
                new DrawGridInf(15, "人件費","", StringAlignment.Far),
                new DrawGridInf(15, "包装材","", StringAlignment.Far),
                new DrawGridInf(15, "定価","", StringAlignment.Far),
                new DrawGridInf(15, "利益額","", StringAlignment.Far)
            };

            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);


            float y = util.DrawGridTitle(PrintGapLeft, PrintGapTop, rowHeight, lstDrawGridInf);

            for (int targetIndex = pd.printGridDataIndex; targetIndex < lstPrintData.Count; targetIndex++)
            {

                int iItem = 0;
                var item = lstPrintData[targetIndex];

                float costRate;
                float rowCostRate;
                float laborCostRate;
                float packageCostRate;
                float profit;

                lstDrawGridInf[iItem++].value = item.Kind;
                lstDrawGridInf[iItem++].value = item.ProductName;
                lstDrawGridInf[iItem++].value = item.ProductNum.ToString();

                int rc = item.Calc(out costRate, out rowCostRate, out laborCostRate, out packageCostRate, out profit);
                Color color = Color.Black;
                if (rc != 0)
                {
                    color = Color.Red;
                    lstDrawGridInf[iItem++].value = 0.ToString("P1");
                    lstDrawGridInf[iItem++].value = 0.ToString("P1");
                    lstDrawGridInf[iItem++].value = 0.ToString("P1");
                    lstDrawGridInf[iItem++].value = 0.ToString("P1");
                    lstDrawGridInf[iItem++].value = item.Price.ToString("C");
                    lstDrawGridInf[iItem++].value = item.GetProfit().ToString("C");

                }
                else
                {
                    lstDrawGridInf[iItem++].value = costRate.ToString("P1");
                    lstDrawGridInf[iItem++].value = rowCostRate.ToString("P1");
                    lstDrawGridInf[iItem++].value = laborCostRate.ToString("P1");
                    lstDrawGridInf[iItem++].value = packageCostRate.ToString("P1");
                    lstDrawGridInf[iItem++].value = item.Price.ToString("C");
                    lstDrawGridInf[iItem++].value = profit.ToString("C");
                }
                y = util.DrawGridValue(PrintGapLeft, y, rowHeight, lstDrawGridInf, color);

                if (y > A4HeightMM - FooterHeight)
                {
                    //次のページの印刷開始データインデックス設定
                    pd.printGridDataIndex = targetIndex++;
                    //改行
                    return true;
                }
            }

            //一覧印刷がおわったので、次の商品別の原価情報印刷に移行
            pd.printDataType = PrintDataType.PRODUCT_COST_INFO;
            pd.printGridDataIndex = 0;

            return true; //改ページ
        }
        bool DrawProductCost(PrintDocumentEx pd, Graphics gPreview)
        {
            var costData = lstPrintData[pd.printGridDataIndex];
            float x = 0;
            float y = 0;

            //商品名
            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, 8, PrintGapLeft, PrintGapTop);
            util.DrawTitle(x, y- rowHeight, rowHeight, $"(ID:{costData.ProductId}) {costData.ProductName}", Color.Blue);

            //原材料
            DrawMaterial(gPreview, ref x, ref y, costData);
            //作業者
            y += 2;
            DrawWorker(gPreview, ref x, ref y, costData);
            //梱包材
            y += 2;
            DrawPackage(gPreview, ref x, ref y, costData);
            //コスト情報
            y += 2;
            DrawCostInfo(gPreview, ref x, ref y, costData);

            pd.printGridDataIndex++;

            if (pd.printGridDataIndex < lstPrintData.Count - 1)
            {
                return true; //改ページ
            }
            return false;
        }

        private void DrawMaterial(Graphics gPreview, ref float x, ref float y, CostData costData )
        {
            //原材料
            List<DrawGridInf> lstDrawGridInf = new List<DrawGridInf>()
            {
                new DrawGridInf(15, "分類","", StringAlignment.Near),
                new DrawGridInf(30, "原材料","", StringAlignment.Near),
                new DrawGridInf(10, "使用料","", StringAlignment.Far),
                new DrawGridInf(15, "費用","", StringAlignment.Far),
            };
            //原材料のグリッド描画幅
            materialGridW = lstDrawGridInf.Sum(inf => inf.width);

            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);
            y = util.DrawTitle(x, y, rowHeight, "■原材料", Color.Black);
            y = util.DrawGridTitle(x, y, rowHeight, lstDrawGridInf);

            foreach (var item in costData.LstMaterialCost)
            {
                int iItem = 0;
                lstDrawGridInf[iItem++].value = item.GetKind();
                lstDrawGridInf[iItem++].value = item.GetName();
                lstDrawGridInf[iItem++].value = item.AmountUsed.ToString("0g");
                lstDrawGridInf[iItem++].value = item.CalcCost().ToString("C");
                y = util.DrawGridValue(x, y, rowHeight, lstDrawGridInf, Color.Black);
            }
        }
        private void DrawWorker(Graphics gPreview, ref float x, ref float y, CostData costData)
        {
            //全行数分の縦幅
            float allHeight = costData.LstWorkerCost.Count * rowHeight;
            //タイトル
            allHeight += rowHeight;

            if (y + allHeight > A4HeightMM - FooterHeight)
            {
                x = materialGridW + 10;
                y = 0;
            }

            //作業者
            List<DrawGridInf> lstDrawGridInf = new List<DrawGridInf>()
            {
                new DrawGridInf(30, "作業者","", StringAlignment.Near),
                new DrawGridInf(10, "時間(分)","", StringAlignment.Far),
                new DrawGridInf(15, "費用","", StringAlignment.Far),
            };

            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);
            y = util.DrawTitle(x,y, rowHeight, "■作業者", Color.Black);
            y = util.DrawGridTitle(x, y, rowHeight, lstDrawGridInf);

            foreach (var item in costData.LstWorkerCost)
            {
                int iItem = 0;
                lstDrawGridInf[iItem++].value = item.GetName();
                lstDrawGridInf[iItem++].value = item.WorkingTime.ToString("0分");
                lstDrawGridInf[iItem++].value = item.CalcCost().ToString("C");
                y = util.DrawGridValue(x, y, rowHeight, lstDrawGridInf, Color.Black);
            }
        }

        private void DrawPackage(Graphics gPreview, ref float x, ref float y, CostData costData)
        {
            //全行数分の縦幅
            float allHeight = costData.LstPackageCost.Count * rowHeight;
            //タイトル
            allHeight += rowHeight;

            if (y + allHeight > A4HeightMM - PrintGapTop - FooterHeight)
            {
                x = materialGridW+10;
                y = 0;
            }

            //作業者
            List<DrawGridInf> lstDrawGridInf = new List<DrawGridInf>()
            {
                new DrawGridInf(30, "包装材","", StringAlignment.Near),
                new DrawGridInf(10, "必要数","", StringAlignment.Far),
                new DrawGridInf(15, "費用","", StringAlignment.Far),
            };

            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);
            y = util.DrawTitle(x, y, rowHeight, "■包装材", Color.Black);
            y = util.DrawGridTitle(x, y, rowHeight, lstDrawGridInf);

            foreach (var item in costData.LstPackageCost)
            {
                int iItem = 0;
                lstDrawGridInf[iItem++].value = item.GetName();
                lstDrawGridInf[iItem++].value = item.RequiredNum.ToString();
                lstDrawGridInf[iItem++].value = item.CalcCost().ToString("C");
                y = util.DrawGridValue(x, y, rowHeight, lstDrawGridInf, Color.Black);
            }
        }

        private void DrawCostInfo(Graphics gPreview, ref float x, ref float y, CostData costData)
        {
            //x値が右列になっていなければ右列最上段に変更
            if( x <= PrintGapLeft)
            {
                x = materialGridW + 10;
                y = 0;
            }

            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);
            //原材料
            y = util.DrawTitle(x, y, rowHeight, "■原価/単価", Color.Black);

            DrawCostInfo_Sub(util, ref x, ref y, "原材料", costData.GetMateralCost().ToString("C"));
            DrawCostInfo_Sub(util, ref x, ref y, "人件費", costData.GetWorkerCost().ToString("C"));
            DrawCostInfo_Sub(util, ref x, ref y, "パケージ費", costData.GetPackageCost().ToString("C"));
            y += 1;
            DrawCostInfo_Sub(util, ref x, ref y, "総原価", costData.GetAllCost().ToString("C"));
            DrawCostInfo_Sub(util, ref x, ref y, "製品単価", costData.GetCostOne().ToString("C"));

            //原価率/利益率
            y += 3;
            y = util.DrawTitle(x, y, rowHeight, "■原価率/利益率", Color.Black);

            DrawCostInfo_Sub(util, ref x, ref y, "定価", costData.Price.ToString("C"));
            DrawCostInfo_Sub(util, ref x, ref y, "原価率", costData.GetCostRate().ToString("P1"));
            DrawCostInfo_Sub(util, ref x, ref y, "利益率", costData.GetProfitRate().ToString("P1"));

            //栄養成分
            y += 3;
            var product = Global.productReader.GetProductDataByID(costData.ProductId);
            y = util.DrawTitle(x, y, rowHeight, "■栄養成分", Color.Black);
            DrawCostInfo_Sub(util, ref x, ref y, "熱量", product.Calorie);
            DrawCostInfo_Sub(util, ref x, ref y, "タンパク質", product.Protein);
            DrawCostInfo_Sub(util, ref x, ref y, "脂質", product.Lipids);
            DrawCostInfo_Sub(util, ref x, ref y, "炭水化物", product.Carbohydrates);
            DrawCostInfo_Sub(util, ref x, ref y, "食塩", product.Salt);

        }

        private void DrawCostInfo_Sub(DrawUtil2 util, ref float x, ref float y, string rowTitle, string value)
        {
            DrawGridInf2 item = new DrawGridInf2(rowTitle, value, 30, 15);
            y = util.DrawGridRowTitle(x, y, rowHeight, item);

        }

        public int GetPageNum()
        {
            //グリッド印刷ページ数
            int num = GetGridPageNum();

            //商品数分のページを加算
            num += lstPrintData.Count();

            return num;
        }

        /// <summary>
        /// グリッド印刷ページ数カウント
        /// </summary>
        /// <returns></returns>
        public int GetGridPageNum()
        {
            int pageNum = 1;
            float y = PrintGapTop;

            foreach (var item in lstPrintData)
            {
                y += rowHeight;

                if (y > A4HeightMM)
                {
                    pageNum++;
                    y = PrintGapTop;
                }
            }
            return pageNum;
        }
        void DrawFooter(Graphics gPreview, int curPageNo, int pageNum)
        {
            DrawUtil2 util = new DrawUtil2(gPreview, Const.defaultFontName, Const.defaultFontSize, PrintGapLeft, PrintGapTop);
            util.DrawFooter(curPageNo, pageNum, A4WidthMM, A4HeightMM);
        }

    }
}