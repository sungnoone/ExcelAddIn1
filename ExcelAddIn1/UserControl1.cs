using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using Newtonsoft.Json.Linq;


namespace ExcelAddIn1
{
    public partial class UserControl1 : UserControl
    {
        public string REST_HOST = "http://192.168.1.229:8088";//開發主機
        //public string REST_HOST = "http://192.168.1.230:8088";//正式主機
        public string REST_URL_CONNECT_SERVICE = "/srv/"; //測試連接服務
        public string REST_URL_CN_WORDS_BROKEN = "/vsto/cn/get/broken"; //查詢全部破音字
        public string REST_URL_CN_WORDS_BROKEN_FONT = "/vsto/cn/get/broken/font"; //查詢破音字體庫
        //public string REST_URL_CN_WORDS_INSERT_ONE = "/srv/cn/insert/one/"; //新增生字到資料庫

        //public string SEARCH_KEYWORD = "解釋"; //搜尋關鍵詞
        public string FILE_CACHE_DB_DATA = @"C:\Users\Public\WriteText.txt";

        public string FILE_LOG_CELL_OPERATION = @"C:\Users\Public\FILE_LOG_CELL_OPERATION.txt";
        public string PATH_TEMP = Path.GetTempPath();
        public string FILE_API_RESPONSE = "file_api_response.log";

        //words_broken fields mappings
        public string FIELDS_BROKEN_ID = "編號";
        public string FIELDS_BROKEN_WORD = "字";

        //words_broken_font fields mappings
        public string FIELDS_BROKEN_FONT_FIELD1 = "id";
        public string FIELDS_BROKEN_FONT_FIELD2 = "before_word";
        public string FIELDS_BROKEN_FONT_FIELD3 = "word";
        public string FIELDS_BROKEN_FONT_FIELD4 = "after_word";
        public string FIELDS_BROKEN_FONT_FIELD5 = "apply_font_name";
        
        //fields name map
        string CN_WORDS_ID = "words_id";
        string CN_WORDS_CONTENT = "words_content";

        //Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;

        //color picker
        int colorValue = 0;

        public UserControl1()
        {
            InitializeComponent();
        }


        //選色
        private void button3_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                //txtMessage.Text = colorDialog1.Color.ToArgb().ToString()+"\r\n";
                colorValue = colorDialog1.Color.ToArgb();
                button3.BackColor = Color.FromArgb(colorValue);
            }
        }


        //執行標注作業
        private void button1_Click(object sender, EventArgs e)
        {
            txtMessage.Text = "";//clear message
            //query all broken words(newton require)
            JArray words = (JArray)queryBrokenAll(REST_HOST,REST_URL_CN_WORDS_BROKEN);
            //writeToFile(words.ToString(), PATH_TEMP + FILE_API_RESPONSE, false);//log query result
            progressBar1.Value = 0;
            progressBar1.Maximum = words.Count;
            foreach (dynamic word in words)
            {
                var jObj = (JObject)word;
                progressBar1.Value += 1;
                foreach (JToken token in jObj.Children())
                {
                    if (token is JProperty)
                    {
                        //writeToFile(word[FIELDS_BROKEN_ID] + ":" + word[FIELDS_BROKEN_WORD] + "\r\n", PATH_TEMP + "words.log", true);
                        var prop = token as JProperty;
                        if (prop.Name == FIELDS_BROKEN_WORD)
                        {
                            //setting ranges
                            oWB = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
                            oSheet = (Excel.Worksheet)oWB.ActiveSheet;
                            Excel.Range usedRange = null;
                            //writeToFile(prop.Value + "\r\n", PATH_TEMP + "words.log", true);
                            //all doc or selection range
                            txtMessage.Text += prop.Value + "\r\n";
                            if (radioButton1.Checked == true){
                                //all cells
                                usedRange = oSheet.UsedRange;
                                //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                                searchReplace1(usedRange, prop.Value.ToString() );//all doc - search custom function
                            }else{
                                //user selection range
                                usedRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                                //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                                if (usedRange == null || usedRange.Cells.Count <= 1)
                                {
                                    txtMessage.Text += "沒有選取範圍" + "\r\n";
                                    return;
                                }
                                searchReplace1(usedRange,prop.Value.ToString());//selection range - search custom function
                            }
                        }
                    }
                }
            }
        }

        //執行多音字體套用
        private void button2_Click(object sender, EventArgs e)
        {
            txtMessage.Text = "";//clear message
            //query all broken words(newton require)
            //try to connect RestFul service
            JArray words;
            try {
                words = (JArray)queryBrokenAll(REST_HOST, REST_URL_CN_WORDS_BROKEN_FONT);
            }catch(Exception ex){
                txtMessage.Text += ex.ToString()+"服務連線失敗!!\r\n";
                return;
            }
            if(words==null){
                txtMessage.Text += e.ToString() + "服務連線結果為空!!\r\n";
                return;
            }
            progressBar2.Value = 0;
            progressBar2.Maximum = words.Count;
            foreach (dynamic word in words)
            {
                var jObj = (JObject)word;
                progressBar2.Value += 1;

                string before_word = "";
                string body_word = "";
                string after_word = "";
                string applyFontName = "";

                foreach (JToken token in jObj.Children())
                {
                    //compound search string = before + body + after
                    if (token is JProperty)
                    {
                        var prop = token as JProperty;
                        //compound search string = before + body + after
                        if(prop.Name==FIELDS_BROKEN_FONT_FIELD2){
                            //before word
                            before_word = prop.Value.ToString();
                        }else if (prop.Name == FIELDS_BROKEN_FONT_FIELD3) {
                            //word body
                            body_word = prop.Value.ToString();
                        }else if (prop.Name == FIELDS_BROKEN_FONT_FIELD4) {
                            //after word
                            after_word = prop.Value.ToString();
                        }else if(prop.Name == FIELDS_BROKEN_FONT_FIELD5){
                            //apply font name
                            applyFontName = prop.Value.ToString();
                        }
                    }
                }
                string myFindText = before_word + body_word + after_word;
                //if search text empty
                if (myFindText != "" && myFindText != null && applyFontName!="" && applyFontName!=null)
                {
                    //setting ranges
                    oWB = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
                    oSheet = (Excel.Worksheet)oWB.ActiveSheet;
                    Excel.Range usedRange = null;
                    //writeToFile(prop.Value + "\r\n", PATH_TEMP + "words.log", true);
                    //all doc or selection range
                    txtMessage.Text += myFindText + " ==> " + applyFontName + "\r\n";
                    if (radioButton3.Checked == true)
                    {
                        //all cells
                        usedRange = oSheet.UsedRange;
                        //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                        searchReplace2(usedRange, before_word, body_word, after_word, applyFontName);//all doc - search custom function
                    }
                    else
                    {
                        //user selection range
                        usedRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                        //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                        if (usedRange == null || usedRange.Cells.Count <= 1)
                        {
                            txtMessage.Text += "沒有選取範圍" + "\r\n";
                            return;
                        }
                        searchReplace2(usedRange, before_word, body_word, after_word, applyFontName);//selection range - search custom function
                    }
                }
                else
                {
                    txtMessage.Text += "搜尋字串為空!\r\n";
                }
            }
        }

 
        //=========================== sub function ========================================.

        #region sub function

        //search character and apply color
        private void searchReplace1(Excel.Range operateRange, string searchString)
        {
            //MessageBox.Show(searchString);
            object missing = System.Reflection.Missing.Value;//null parameter
            //Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            string firstAddress = "";

            Excel.Range findRng = operateRange.Find(searchString, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);

            if (findRng == null){//no search result
                return; 
            }
            if (firstFind == null)
            {
                firstFind = findRng;//assign first range as replace range
            }
            firstAddress = firstFind.get_Address(Excel.XlReferenceStyle.xlA1);// get first range address
            bool firstDo = false;//flag of first excution
            while (findRng != null)
            {
                string findAddress = findRng.get_Address(Excel.XlReferenceStyle.xlA1);
                if (findAddress == firstAddress && firstDo == false)
                {
                    //第一個還沒做
                    writeToFile(findAddress + ":\r\n" + findRng.Value2 + "\r\n", FILE_CACHE_DB_DATA, false);
                    changeCellAttribute1(findRng, searchString);
                    firstDo = true;//change first excute flag. first range already be excuted
                }
                else if (findAddress == firstAddress && firstDo == true)
                {
                    //已重複回車了
                    break;
                }
                else
                {
                    //normal cell operation
                    changeCellAttribute1(findRng, searchString);
                }
                findRng = operateRange.FindNext(findRng);
            }
        }

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
            bool firstDo = false;//flag of first excution
            while (findRng != null)
            {
                string findAddress = findRng.get_Address(Excel.XlReferenceStyle.xlA1);
                if (findAddress == firstAddress && firstDo == false)
                {
                    //第一個還沒做
                    //writeToFile(findAddress + ":\r\n" + findRng.Value2 + "\r\n", FILE_CACHE_DB_DATA, false);
                    changeCellAttribute2(findRng, beforeStr, midStr, afterStr, fontName);
                    firstDo = true;//change first excute flag. first range already be excuted
                }
                else if (findAddress == firstAddress && firstDo == true)
                {
                    //已重複回車了
                    break;
                }
                else
                {
                    //normal cell operation
                    changeCellAttribute2(findRng, beforeStr, midStr, afterStr, fontName);
                }
                findRng = operateRange.FindNext(findRng);
            }

        }

        #endregion

        //=========================== common function ========================================


        #region common function

        private JArray queryBrokenAll(string restHost, string restUrl)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(restHost + restUrl);
            req.ContentType = "application/json";
            req.Method = "GET";

            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            if (resp.StatusCode == HttpStatusCode.OK)
            {
                using (Stream respStream = resp.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(respStream);
                    System.Text.Encoding.GetEncoding("utf-8");//encoding response stream
                    JsonTextReader jsonReader = new JsonTextReader(reader);//convert streamReader to newtonsoft jsonTextReader
                    var serializer = new JsonSerializer();//declare JsonSerial Object
                    dynamic serDes = serializer.Deserialize(jsonReader);//Deserialize jsonTextReader
                    JArray wordsJson = (JArray)serDes.rows;
                    return wordsJson;
                }
            }
            else
            {
                MessageBox.Show(string.Format("Status Code:{0}, Status Description:{1}", resp.StatusCode, resp.StatusDescription));
                return null;
            }
        }

        //cell content search & replace
        private bool changeCellAttribute1(Excel.Range locateCell, string searchStr)
        {
            string pattern = searchStr; //defined regular search pattern
            try
            {
                //using cell address as log title
                writeToFile(locateCell.get_Address(Excel.XlReferenceStyle.xlR1C1) + "\r\n", FILE_LOG_CELL_OPERATION, true);
            }
            catch (Exception e)
            {
                MessageBox.Show("write cell operation log into file failed!!" + e.ToString());
                return false;
            }

            //matching in cell string
            string cellValue = Convert.ToString(locateCell.Value2.ToString());
            foreach (Match match in Regex.Matches(cellValue, pattern))
            {
                writeToFile("\t\tvalue:" + match.Value + " index:" + match.Index + " length:" + match.Length + "\r\n", FILE_LOG_CELL_OPERATION, true);
                locateCell.Characters[match.Index + 1, match.Length].Font.Color = Color.FromArgb(colorValue);
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

        //cell content search & change character attributes - color, font name
        private bool changeCellAttribute2(Excel.Range locateCell, string beforeStr, string midStr, string afterStr, string fontName)
        {
            string pattern = beforeStr+midStr+afterStr; //defined regular search pattern
            try
            {
                //using cell address as log title
                writeToFile(locateCell.get_Address(Excel.XlReferenceStyle.xlR1C1) + "\r\n", FILE_LOG_CELL_OPERATION, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("write cell operation log into file failed!!" + ex.ToString());
                return false;
            }

            //matching in cell string
            if (locateCell.Value2 == null) { return false; }//cell content null
            string cellValue = Convert.ToString(locateCell.Value2.ToString());
            foreach (Match match in Regex.Matches(cellValue, pattern))
            {
                //apply color
                //locateCell.Characters[match.Index+beforeStr.Length+1, midStr.Length].Font.Color = Excel.XlRgbColor.rgbPurple;
                locateCell.Characters[match.Index + beforeStr.Length + 1, midStr.Length].Font.Color = Color.FromArgb(colorValue);
                //apply font
                locateCell.Characters[match.Index+beforeStr.Length+1, midStr.Length].Font.Name = fontName;
            }
            return true;
        }

        //
        private Excel.Range GetSpecifiedRange(string matchStr, Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Range currentFind = null;
            //Excel.Range firstFind = null;
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
                System.IO.StreamReader sr = new System.IO.StreamReader(openFileDialog1.FileName);
                s = sr.ReadToEnd();
                //MessageBox.Show(s);
                sr.Close();
            }
            return s;
        }


        //write to file
        private bool writeToFile(string contentStr, string fileName, bool append)
        {
            try
            {
                //System.IO.File.WriteAllText(fileName, contentStr, true);
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fileName, append);

                sw.Write(contentStr);
                sw.Close();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("write file error:" + e.ToString());
                return false;
            }
        }

        //read from servcice
        private void readFromApi()
        {
            WebRequest req = WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_BROKEN);
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

        #endregion








    }
}
