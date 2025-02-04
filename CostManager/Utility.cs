using NPOI.SS.Formula.Functions;
using NPOI.Util.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelReaderUtility.ProductReader;
using System.Drawing;

internal class Utility
{
    /// <summary>
    /// 今日から指定された日数のDateTimeを返す
    /// </summary>
    /// <param name="days"></param>
    /// <returns></returns>
    public static DateTime GetValidDate(int days)
    {
        var today = DateTime.Now;

        return today.Add(TimeSpan.FromDays(days - 1));

    }

    public static void MessageError(string msg)
    {
        MessageBox.Show(msg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
    public static void MessageInfo(string msg, string title = "情報")
    {
        MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    public static DialogResult MessageConfirm(string msg, string title= "確認")
    {
        return MessageBox.Show(msg, title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
    }
    public static int MILLI2POINT(float milli)
    {
        return (int)(milli / 0.352777);
    }
    public static float POINT2MILLI(int point)
    {
        return (float)(point * 0.352777);
    }

    public static float ToFloat(string value)
    {
        return float.Parse(value.Trim());
    }
    public static int ToInt(string value)
    {
        return int.Parse(value.Trim());
    }
    public static bool ToBoolean(string value)
    {
        if (!string.IsNullOrEmpty(value.Trim()))
        {
            int tmpValue = 0;
            int.TryParse(value, out tmpValue);
            return tmpValue == 0 ? false : true;
        }
        return false;
    }

    public static string RemoveCRLF(string s)
    {
        //改行コードを削除
        s = s.Replace("\n", "");
        return s.Replace("\r", "");
    }

    public static void LoadUserSetting(Form frm,
                                int settingLocX, int settingLocY,
                                int settingSizeW, int settingSizeH
                                )
    {
        //setting情報
        int WinX = settingLocX;
        int WinY = settingLocY;

        if (WinX < 0)
        {   //小さすぎたら補正
            WinX = 0;
        }
        if (WinY < 0)
        {   //小さすぎたら補正
            WinY = 0;
        }
        frm.Location = new Point(WinX, WinY);

        int SizeW = settingSizeW;
        int SizeH = settingSizeH;
        if (SizeW < 200)
        {   //小さすぎたら補正
            SizeW = frm.Size.Width;
        }
        if (SizeH < 200)
        {   //小さすぎたら補正
            SizeH = frm.Size.Height;
        }
        frm.Size = new Size(SizeW, SizeH);

    }

}

