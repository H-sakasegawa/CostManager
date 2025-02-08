using ExcelReaderUtility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostManager
{
    public class Global
    {
        public static OptionData optionData = new OptionData();
        public static ProductReader productReader = new ProductReader();
        public static CostReader costReader = new CostReader();

    }
}
