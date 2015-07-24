﻿using System;
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
        public string brokenTaskPaneTitle = "北研專案 1.0";
        Microsoft.Office.Tools.CustomTaskPaneCollection thisCustomTaskPanes;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //搜尋取代 Task Pane
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl1 myUserControl1 = null;
            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane = null;
            Excel.Window activeWin = Globals.ThisAddIn.Application.ActiveWindow;
            thisCustomTaskPanes = Globals.ThisAddIn.CustomTaskPanes;

            //already had panes
            int matchPanes = 0;
            for (int v1 = 0; v1 < thisCustomTaskPanes.Count; v1++)
            {
                Microsoft.Office.Tools.CustomTaskPane ctp = (Microsoft.Office.Tools.CustomTaskPane)thisCustomTaskPanes[v1];
                try
                {
                    Excel.Window ctpWin = (Excel.Window)ctp.Window;
                    if (ctp.Title == brokenTaskPaneTitle && ctpWin.Index==activeWin.Index)
                    {
                        //MessageBox.Show(ctpWin.Index.ToString());
                        //MessageBox.Show(activeWin.Index.ToString());
                        matchPanes += 1;
                        if (ctp.Visible == false) { ctp.Visible = true; } else { ctp.Visible = false; }
                    }
                }
                catch (Exception ex) {
                    MessageBox.Show(ex.ToString());
                }
                
            }
            //no any taskpane exist
            if (activeWin.Panes.Count==0 || matchPanes==0)
            {
                myUserControl1 = new UserControl1();
                myCustomTaskPane = (Microsoft.Office.Tools.CustomTaskPane)thisCustomTaskPanes.Add(myUserControl1, brokenTaskPaneTitle,activeWin);
                myCustomTaskPane.Visible = true;
                myCustomTaskPane.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);
            }

        }

        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane taskPane = sender as Microsoft.Office.Tools.CustomTaskPane;
            thisCustomTaskPanes.Remove(taskPane);
        }

    }
}
