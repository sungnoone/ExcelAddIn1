using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    public partial class UserControl2 : UserControl
    {
        //Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;

        public UserControl2()
        {
            InitializeComponent();
        }

        //去除標記
        private void button1_Click(object sender, EventArgs e)
        {
            txtMessage.Text = "";//clear message
            //setting ranges
            oWB = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            Excel.Range usedRange = null;
            //all cells
            usedRange = oSheet.UsedRange;
            txtMessage.Text += "Used Range:"+usedRange.Cells.Count+Environment.NewLine;
            styleMarkRevert(usedRange);//InDesign樣式標記復原
            txtMessage.Text += "Used Range:" + usedRange.Cells.Count + Environment.NewLine;
            sectionMarkRevert(usedRange);//分段標記
            MessageBox.Show("OK!");
        }

        //分段標記復原
        private void sectionMarkRevert(Excel.Range usedRange)
        {
            //all contain match result cell collections
            Dictionary<string, ArrayList> findResultDict = new Dictionary<string, ArrayList>();
            foreach (Excel.Range singleCell in usedRange.Cells)
            {
                string textVal = singleCell.Text;
                string pat = @"(\/\/)";
                Regex r = new Regex(pat, RegexOptions.IgnoreCase);
                ArrayList cellAryList = new ArrayList();//single cell,all match result
                Match m = r.Match(textVal);
                while (m.Success)
                {
                    ArrayList matchAryList = new ArrayList();//single cell,one match result
                    txtMessage.Text += "index: " + m.Index + " Char: " + textVal[m.Index] + Environment.NewLine;
                    //match first position
                    matchAryList.Add(m.Index);
                    for (int i = 1; i <= m.Groups.Count; i++)
                    {
                        Group g = m.Groups[i];
                        txtMessage.Text += "Group[" + i + "]: " + g + Environment.NewLine;
                        /*CaptureCollection cc = g.Captures;
                        for (int j=0; j < cc.Count;j++ )
                        {
                            Capture c = cc[j];
                            txtMessage.Text += "Capture[" + j + "]: " + c + Environment.NewLine;
                        }*/
                        matchAryList.Add(g.Length);
                    }
                    cellAryList.Add(matchAryList);
                    m = m.NextMatch();
                }
                if (cellAryList.Count != 0)
                {
                    findResultDict.Add(singleCell.Address, cellAryList);
                }
            }
            //replace from dictionary
            foreach (KeyValuePair<string, ArrayList> items in findResultDict)
            {
                txtMessage.Text += items.Key + Environment.NewLine;
                Excel.Range locateCell = oSheet.get_Range(items.Key);
                //locateCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                ArrayList aryList1 = items.Value;
                txtMessage.Text += items.Key + Environment.NewLine;
                for (int poi = aryList1.Count - 1; poi >= 0; poi--)
                {
                    ArrayList aryList2 = (ArrayList)aryList1[poi];
                    txtMessage.Text += "poistion:" + aryList2[0] + "  Length:" + aryList2[1] + Environment.NewLine;
                    Excel.Characters g1 = locateCell.Characters[(int)aryList2[0] + 1, (int)aryList2[1]];
                    g1.Text = "\r\n";
                }
                //Excel.Characters getChars = locateCell.Characters[aryList1[0],aryList1[1]];
                //getChars.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
        }

        //樣式標記復原
        private void styleMarkRevert(Excel.Range usedRange) {
            //all contain match result cell collections
            Dictionary<string, ArrayList> findResultDict = new Dictionary<string, ArrayList>();
            foreach (Excel.Range singleCell in usedRange.Cells)
            {
                string textVal = singleCell.Text;
                //string pat = @"(~@)(.*)(@~)";
                string pat = @"(~@C\w{3}-)(.*?)(@~)";
                Regex r = new Regex(pat, RegexOptions.IgnoreCase);
                //string replaced = r.Replace(textVal, "$2");
                //txtMessage.Text += replaced + Environment.NewLine;
                //singleCell.Value = replaced;
                ArrayList cellAryList = new ArrayList();//single cell,all match result
                Match m = r.Match(textVal);
                while (m.Success)
                {
                    ArrayList matchAryList = new ArrayList();//single cell,one match result
                    txtMessage.Text += "index: " + m.Index + " Char: " + textVal[m.Index] + Environment.NewLine;
                    //match first position
                    matchAryList.Add(m.Index);
                    for (int i = 1; i <= m.Groups.Count; i++)
                    {
                        Group g = m.Groups[i];
                        txtMessage.Text += "Group[" + i + "]: " + g + Environment.NewLine;
                        /*CaptureCollection cc = g.Captures;
                        for (int j=0; j < cc.Count;j++ )
                        {
                            Capture c = cc[j];
                            txtMessage.Text += "Capture[" + j + "]: " + c + Environment.NewLine;
                        }*/
                        matchAryList.Add(g.Length);
                    }
                    cellAryList.Add(matchAryList);
                    m = m.NextMatch();
                }
                if (cellAryList.Count != 0)
                {
                    findResultDict.Add(singleCell.Address, cellAryList);
                }
            }
            //replace from dictionary
            foreach (KeyValuePair<string, ArrayList> items in findResultDict)
            {
                txtMessage.Text += items.Key + Environment.NewLine;

                Excel.Range locateCell = oSheet.get_Range(items.Key);
                //locateCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                ArrayList aryList1 = items.Value;
                txtMessage.Text += items.Key + Environment.NewLine;
                for (int poi = aryList1.Count - 1; poi >= 0; poi--)
                {
                    ArrayList aryList2 = (ArrayList)aryList1[poi];
                    txtMessage.Text += "poistion:" + aryList2[0] + "  Length:" + aryList2[1] + Environment.NewLine;
                    //Excel.Characters getChars = locateCell.Characters[(int)aryList2[0]+1,(int)aryList2[1]];
                    Excel.Characters g1 = locateCell.Characters[(int)aryList2[0] + 1, (int)aryList2[1]];
                    Excel.Characters g2 = locateCell.Characters[(int)aryList2[0] + 1 + (int)aryList2[1], (int)aryList2[2]];
                    Excel.Characters g3 = locateCell.Characters[(int)aryList2[0] + 1 + (int)aryList2[1] + (int)aryList2[2], (int)aryList2[3]];
                    //g1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    //g2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Tomato);
                    //g3.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
                    g3.Text = "";
                    g1.Text = "";
                }
                //Excel.Characters getChars = locateCell.Characters[aryList1[0],aryList1[1]];
                //getChars.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
        }

        //=========================== sub function ========================================.

        #region sub function

        //search character and apply color
        private void searchReplace2(Excel.Range operateRange, string beforeStr, string midStr, string afterStr, string fontName)
        {
            //txtMessage.Text += "搜尋字串: "+searchString+"\r\n";
            object missing = System.Reflection.Missing.Value;//null parameter
            //Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            string firstAddress = "";
            string searchString = beforeStr + midStr + afterStr;
            Excel.Range findRng = operateRange.Find(searchString, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);

            if (findRng == null)
            {//no search result
                return;
            }
            if (firstFind == null)
            {
                firstFind = findRng;//assign first range as replace range
            }
            firstAddress = firstFind.get_Address(Excel.XlReferenceStyle.xlA1);// get first range address

        }

        #endregion


    }
}
