using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using Newtonsoft.Json.Linq;


namespace ExcelAddIn1
{

    public partial class Ribbon1
    {
        



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //搜尋取代 Task Pane
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //show broken operation task pane
            UserControl1 myUserControl1 = new UserControl1();
            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane = (Microsoft.Office.Tools.CustomTaskPane)Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "北研專案");
            myCustomTaskPane.Visible = true;

        }

    }
}
