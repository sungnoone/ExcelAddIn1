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


namespace ExcelAddIn1
{

    public partial class Ribbon1
    {
        
        public string REST_HOST = "http://192.168.1.229:8088";
        public string REST_URL_CONNECT_SERVICE = "/srv/"; //測試連接服務
        public string REST_URL_CN_WORDS_ALL = "/srv/cn/get/all/"; //查詢全部生字
        public string REST_URL_CN_WORDS_INSERT_ONE = "/srv/cn/insert/one/"; //新增生字到資料庫

        public string SEARCH_KEYWORD = "解釋"; //搜尋關鍵詞
        public string FILE_CACHE_DB_DATA = @"C:\Users\Public\WriteText.txt";

        public string FILE_LOG_CELL_OPERATION = @"C:\Users\Public\FILE_LOG_CELL_OPERATION.txt";


        //fields name map
        string CN_WORDS_ID = "words_id";
        string CN_WORDS_CONTENT = "words_content";


        //Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //搜尋取代
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            object missing = System.Reflection.Missing.Value;
            //Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            string firstAddress = "";
            
            oWB = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            Excel.Range usedRange = oSheet.UsedRange;
            Excel.Range findRng = usedRange.Find(SEARCH_KEYWORD, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);
            if (firstFind == null){
                    firstFind = findRng;
            }
            firstAddress = firstFind.get_Address(Excel.XlReferenceStyle.xlA1);
            bool firstDo = false;
            string str = "";
            while(findRng!=null){
                string findAddress = findRng.get_Address(Excel.XlReferenceStyle.xlA1);
                if (findAddress == firstAddress && firstDo == false){
                    //第一個還沒做
                     writeToFile(findAddress + ":\r\n" + findRng.Value2 + "\r\n", FILE_CACHE_DB_DATA, false);
                    change_cell_content_attribute(findRng, SEARCH_KEYWORD);
                    //MessageBox.Show("first:"+findAddress);
                    firstDo = true;
                }
                else if (findAddress == firstAddress && firstDo == true)
                {
                    //已重複回車了
                    //MessageBox.Show("已重複回車了");
                    break;
                }
                else {
                    //str += str + findRng + ":\r\n" + findRng.Value2 + "\r\n";
                    writeToFile(findAddress + ":\r\n" + findRng.Value2 + "\r\n", FILE_CACHE_DB_DATA, true);
                    change_cell_content_attribute(findRng, SEARCH_KEYWORD);
                    //MessageBox.Show(findAddress);
                }
                findRng = usedRange.FindNext(findRng);
            }
            //string s = readFromFile();
            //MessageBox.Show(s);
        }


        //hello world 測試
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            WebRequest req = WebRequest.Create(REST_HOST+REST_URL_CONNECT_SERVICE);
            req.Method = "GET";

            HttpWebResponse resp = req.GetResponse() as HttpWebResponse;
            if (resp.StatusCode == HttpStatusCode.OK)
            {
                using (Stream respStream = resp.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(respStream, Encoding.UTF8);
                    MessageBox.Show(reader.ReadToEnd());
                }
            }
            else {
                MessageBox.Show(string.Format("Status Code:{0}, Status Description:{1}",resp.StatusCode, resp.StatusDescription));
            }

        }


        //查詢全部生字
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            WebRequest req = WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_ALL);
            req.Method = "GET";

            HttpWebResponse resp = req.GetResponse() as HttpWebResponse;
            if (resp.StatusCode == HttpStatusCode.OK)
            {
                using (Stream respStream = resp.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(respStream, Encoding.UTF8);
                    MessageBox.Show(reader.ReadToEnd());
                }
            }
            else
            {
                MessageBox.Show(string.Format("Status Code:{0}, Status Description:{1}", resp.StatusCode, resp.StatusDescription));
            }
        }


        //新增生字到資料庫
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string url = REST_HOST + REST_URL_CN_WORDS_INSERT_ONE;
            string word_id = "cn0000001";
            string word_content = "一統天下";

            string json = "{\""+CN_WORDS_ID+"\":\""+word_id+"\"," + "\""+CN_WORDS_CONTENT+"\":\""+word_content+"\"}";
            MessageBox.Show(json.ToString());
            
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_INSERT_ONE);
            req.Method = "POST";
            req.ContentType = "text/plain; charset=utf-8";
            using(var streamWriter = new StreamWriter(req.GetRequestStream())){
                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();

                var resp = (HttpWebResponse)req.GetResponse();
                if (resp.StatusCode == HttpStatusCode.OK)
                {
                    using (Stream respStream = resp.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(respStream, Encoding.UTF8);
                        MessageBox.Show(reader.ReadToEnd());
                    }
                }
                else
                {
                    MessageBox.Show(string.Format("Status Code:{0}, Status Description:{1}", resp.StatusCode, resp.StatusDescription));
                }
            }

        }


        //=========================== sub function ========================================.

        #region sub function


        //cell content search & replace
        private bool change_cell_content_attribute(Excel.Range locateCell, string searchStr) 
        {
            
            string pattern = searchStr; //defined regular search pattern

            try {
                //using cell address as log title
                writeToFile(locateCell.get_Address(Excel.XlReferenceStyle.xlR1C1) + "\r\n", FILE_LOG_CELL_OPERATION, true);
            }
            catch (Exception e) {
                MessageBox.Show("write cell operation log into file failed!!"+e.ToString());
                return false;
            }

            //matching in cell string
            string cellValue = Convert.ToString(locateCell.Value2.ToString());
            foreach (Match match in Regex.Matches(cellValue, pattern))
            {
                writeToFile("\t\tvalue:" + match.Value + " index:" + match.Index +" length:" + match.Length + "\r\n", FILE_LOG_CELL_OPERATION, true);
                locateCell.Characters[match.Index+1, match.Length].Font.Color = Excel.XlRgbColor.rgbLightGreen;
                //for (int v1 = match.Index; v1 <= match.Length; v1++ )
                //{
                //    locateCell.Characters[v1, 1].Font.Color = Excel.XlRgbColor.rgbLightGreen;
                //    locateCell.Characters[v1, 1].Font.Name = "新細明體";
                //    //locateCell.Characters[v1, 1].Font.Size = 12;
                //    locateCell.Characters[v1, 1].Font.Bold = true;
                //}
            }
            return true;
        }


        #endregion


        //=========================== common function ========================================


        #region common function


        //
        private Excel.Range GetSpecifiedRange(string matchStr, Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            currentFind = objWs.get_Range("A1", missing).Find(matchStr, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);
            return currentFind;
        }


        //read from file
        private string readFromFile() 
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Text Files (.txt)|*.txt|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = true;

            // Call the ShowDialog method to show the dialog box.
            //bool? userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            string s = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Open the selected file to read.
                System.IO.StreamReader sr =new  System.IO.StreamReader(openFileDialog1.FileName);
                s = sr.ReadToEnd();
                //MessageBox.Show(s);
                sr.Close();
            }
            return s;
        }


        //write to file
        private bool writeToFile(string contentStr, string fileName, bool append) 
        {
            try {
                //System.IO.File.WriteAllText(fileName, contentStr, true);
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fileName, append);
                
                sw.Write(contentStr);
                sw.Close();
                return true;
            }
            catch(Exception e)  {
                MessageBox.Show("write file error:"+e.ToString());
                return false;
            }
        }


        #endregion




    }
}
