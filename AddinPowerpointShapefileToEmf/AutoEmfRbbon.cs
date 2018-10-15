using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace AddinPowerpointShapefileToEmf
{
    public partial class AutoEmfRbbon
    {
        private void AutoEmfRbbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonLaunch_Click(object sender, RibbonControlEventArgs e)
        {
            new SettingForm().ShowDialog();
         
        }
    }
}
