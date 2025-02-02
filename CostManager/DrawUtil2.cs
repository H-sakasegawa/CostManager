using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using static System.Windows.Forms.AxHost;

namespace CostManager
{
    internal class DrawUtil2
    {

        Graphics graphics = null;

        string fontName = Const.defaultFontName;
        float fontSize = 6;
        float fontSizeItemHeader = 6;

        //印刷ページの左上の余白（コンストラクタで確定）
        float PageGapTopMM = 0;
        float PageGapLeftMM = 0;
        //グリッド内の余白
        float GridGapX = 1;
        float GridGapY = 1;

        Font font;

        ////mm → 座
        //float dpiX = 0;
        //float dpiY = 0;
        //private const float MillimetersPerInch = 25.4f;
        public DrawUtil2(Graphics g,  string fontName)
        {
            this.graphics = g;
            this.graphics.PageUnit = GraphicsUnit.Millimeter;
            this.fontName = fontName;

            Init();
        }

        public DrawUtil2(Graphics graphics,  string fontName,  float fontSize, float PageGapTopMM, float PageGapLeftMM)
        {

            this.graphics = graphics;
            this.graphics.PageUnit = GraphicsUnit.Millimeter;
            this.fontName = fontName;
            this.fontSize = fontSize;
            this.PageGapTopMM = PageGapTopMM;
            this.PageGapLeftMM = PageGapLeftMM;

            Init();

        }

        private void Init()
        {
            font = new Font(fontName, fontSize, FontStyle.Regular);

        }


        float _X(float x) { return PageGapLeftMM + x; }
        float _Y(float y) { return PageGapTopMM + y; }

        public class DrawGridInf
        {
            public DrawGridInf( float width, string title, string value, StringAlignment alignment)
            {
                this.width = width;
                this.title = title;
                this.value = value;
                this.alignment = alignment;
            }
            public float width;
            public string title;
            public string value;
            public StringAlignment alignment;
        }

        public class DrawGridInf2
        {
            public DrawGridInf2(string title, string value, float titleWidth, float valueWidth)
            {
                this.title = title;
                this.value = value;
                this.titleWidth = titleWidth;
                this.valueWidth = valueWidth;
            }
            public string title;
            public string value;
            public float titleWidth;
            public float valueWidth;
        }
        public class DrawCostDetailInf
        {

        }

        public float DrawTitle(float x, float y, float h,string s, Color foreColor)
        {
            Brush brs = new SolidBrush(foreColor);
            StringFormat sf = new StringFormat();
            graphics.DrawString(s, font, brs, new PointF(_X(x), _Y(y)));
            //次のグリッドセルの描画開始Y座標を返す
            return y + h;
        }

        public float DrawGridTitle(float x, float y, float h, List<DrawGridInf> lstItem)
        {
            foreach (var item in lstItem)
            {
                DrawCell(ref x, y,  h, item.width, item.title, StringAlignment.Near, Color.Black, Color.LightBlue);
            }
            //次のグリッドセルの描画開始Y座標を返す
            return y + h;
        }
        public float DrawGridValue(float x, float y, float h, List<DrawGridInf> lstItem, Color color)
        {
            foreach (var item in lstItem)
            {
                DrawCell(ref x, y,h, item.width, item.value, item.alignment, color);
            }
            //次のグリッドセルの描画開始Y座標を返す
            return y + h;
        }

        public float DrawGridRowTitle(float x, float y, float h, DrawGridInf2 item)
        {
            DrawCell(ref x, y, h, item.titleWidth, item.title, StringAlignment.Near, Color.Black, Color.LightBlue);
            DrawCell(ref x , y, h, item.valueWidth, item.value, StringAlignment.Far, Color.Black);

            //次のグリッドセルの描画開始Y座標を返す
            return y + h;
        }


        private void DrawCell(ref float x, float y, float height, float width, string s, StringAlignment alignment, Color foreColor, Color backColor = default)
        {
            if(backColor != default)
            {
                Brush backBrs = new SolidBrush(backColor);
                graphics.FillRectangle(backBrs, _X(x),_Y(y), width, height);
            }

            Pen pen = new Pen(Color.Black, (float)0.1);
            graphics.DrawRectangle(pen, _X(x), _Y(y), width, height);

            Brush brs = new SolidBrush(foreColor);
            StringFormat sf = new StringFormat();
            sf.Alignment = alignment;
            graphics.DrawString(s, font, brs, 
                                new RectangleF(_X(x) + GridGapX, _Y(y) + GridGapY, width- GridGapX, height - GridGapY),
                                sf);

            x += width;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="pen"></param>
        /// <param name="p1">ライン座標</param>
        /// <param name="p2">ライン座標</param>
        public void DrawLine(Pen pen, PointF p1, PointF p2)
        {
            graphics.DrawLine(pen, p1, p2);
        }
        public int DrawFooter(int curPageNo, int pageNum, float pageWidth, float pageHeight)
        {
            Font font = new Font(fontName, fontSizeItemHeader, FontStyle.Regular);
            Brush brs = new SolidBrush(Color.FromArgb(255, 0, 0, 0));

            string pageStr = $"{curPageNo} / {pageNum}";
            SizeF size = graphics.MeasureString(pageStr, font);


            float x = (pageWidth - size.Width)/ 2;
            float y = pageHeight - 10;
            graphics.DrawString(pageStr, font, brs, x, y);
            return 0;
        }

        // インチ(inch)をミリメートル(mm)に変換
        public static double inch_to_mm(double inch)
        {
            return inch * 25.4;
        }

    }
}
