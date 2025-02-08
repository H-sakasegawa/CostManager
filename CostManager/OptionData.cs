using MathNet.Numerics;
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
        }
        public void SaveOptions()
        {
            Properties.Settings.Default.Save();

        }

        public string DataBasePath
        {
            get
            {
               return Properties.Settings.Default.DataBasePath;
            }
            set
            {
                Properties.Settings.Default.DataBasePath = value;
            }
        }
        public bool DispIDtoList
        {
            get
            {
                return Properties.Settings.Default.DispIDtoList;
            }
            set
            {
                Properties.Settings.Default.DispIDtoList = value;
            }
        }

    }
}
