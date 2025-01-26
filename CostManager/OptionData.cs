using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CostManager
{
    public class OptionData
    {

        public OptionData() {
            DataBasePath = Properties.Settings.Default.DataBasePath;

        }
        public void SaveOptions()
        {
            Properties.Settings.Default.DataBasePath = DataBasePath;
            Properties.Settings.Default.Save();

        }

        public string DataBasePath;
    }
}
