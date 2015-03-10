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
        public string REST_HOST = "http://192.168.1.229:8088";
        public string REST_URL_CONNECT_SERVICE = "/srv/"; //測試連接服務
        public string REST_URL_CN_WORDS_BROKEN = "/vsto/cn/get/broken/"; //查詢全部破音字
        public string REST_URL_CN_WORDS_INSERT_ONE = "/srv/cn/insert/one/"; //新增生字到資料庫

        //public string SEARCH_KEYWORD = "解釋"; //搜尋關鍵詞
        public string FILE_CACHE_DB_DATA = @"C:\Users\Public\WriteText.txt";

        public string FILE_LOG_CELL_OPERATION = @"C:\Users\Public\FILE_LOG_CELL_OPERATION.txt";
        public string PATH_TEMP = Path.GetTempPath();
        public string FILE_API_RESPONSE = "file_api_response.log";

        //破音字欄位名稱定義
        public string FIELDS_BROKEN_ID = "編號";
        public string FIELDS_BROKEN_WORD = "字";

        //fields name map
        string CN_WORDS_ID = "words_id";
        string CN_WORDS_CONTENT = "words_content";

        //Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;


        public UserControl1()
        {
            InitializeComponent();
        }


        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }


        //執行標注作業
        private void button1_Click(object sender, EventArgs e)
        {
            txtMessage.Text = "";//clear message
            //query all broken words(newton require)
            JArray words = (JArray)queryBrokenAll();
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
                            if (radioButton1.Checked == true){
                                //all cells
                                usedRange = oSheet.UsedRange;
                                //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                                searchReplace1(prop.Value.ToString(), usedRange);//all doc - search custom function
                            }else{
                                //user selection range
                                usedRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                                //txtMessage.Text += usedRange.Cells.Count.ToString() + " 個儲存格作業範圍" + "\r\n";
                                if (usedRange == null || usedRange.Cells.Count <= 1)
                                {
                                    txtMessage.Text += "沒有選取範圍" + "\r\n";
                                    return;
                                }
                                searchReplace1(prop.Value.ToString(), usedRange);//selection range - search custom function
                            }
                        }
                    }
                }
            }
        }


        //連線測試
        private void button3_Click(object sender, EventArgs e)
        {
            WebRequest req = WebRequest.Create(REST_HOST + REST_URL_CONNECT_SERVICE);
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


        //查詢全部生字
        private void button2_Click(object sender, EventArgs e)
        {
            //WebRequest req = WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_BROKEN);
            //req.Method = "GET";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_BROKEN);
            req.ContentType = "application/json";
            req.Method = "GET";

            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            if (resp.StatusCode == HttpStatusCode.OK)
            {
                using (Stream respStream = resp.GetResponseStream())
                {
                    //web response to stream reader
                    //StreamReader reader = new StreamReader(respStream, Encoding.UTF8);
                    StreamReader reader = new StreamReader(respStream);
                    System.Text.Encoding.GetEncoding("utf-8");//encoding response stream
                    JsonTextReader jsonReader = new JsonTextReader(reader);//convert streamReader to newtonsoft jsonTextReader
                    var serializer = new JsonSerializer();//declare JsonSerial Object
                    dynamic serDes = serializer.Deserialize(jsonReader);//Deserialize jsonTextReader
                    JArray wordsJson = (JArray)serDes.rows;
                    foreach (var word in wordsJson)
                    {
                        writeToFile(word[FIELDS_BROKEN_ID] + ":" + word[FIELDS_BROKEN_WORD] + "\r\n", PATH_TEMP + "words.log", true);
                    }

                    //Console.Write(serDes);
                    //MessageBox.Show(serDes.ToString());                    
                    writeToFile(wordsJson.ToString(), PATH_TEMP + FILE_API_RESPONSE, false);
                }
            }
            else
            {
                MessageBox.Show(string.Format("Status Code:{0}, Status Description:{1}", resp.StatusCode, resp.StatusDescription));
            }
        }


        //新增生字到資料庫
        private void button4_Click(object sender, EventArgs e)
        {
            string url = REST_HOST + REST_URL_CN_WORDS_INSERT_ONE;
            string word_id = "cn0000001";
            string word_content = "一統天下";

            string json = "{\"" + CN_WORDS_ID + "\":\"" + word_id + "\"," + "\"" + CN_WORDS_CONTENT + "\":\"" + word_content + "\"}";
            MessageBox.Show(json.ToString());

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_INSERT_ONE);
            req.Method = "POST";
            req.ContentType = "text/plain; charset=utf-8";
            using (var streamWriter = new StreamWriter(req.GetRequestStream()))
            {
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

        //search character and apply color
        private void searchReplace1(string searchString, Excel.Range operateRange)
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
                    change_cell_content_attribute(findRng, searchString);
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
                    change_cell_content_attribute(findRng, searchString);
                }
                findRng = operateRange.FindNext(findRng);
            }
        }


        private JArray queryBrokenAll()
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(REST_HOST + REST_URL_CN_WORDS_BROKEN);
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


        #endregion


        //=========================== common function ========================================


        #region common function

        //cell content search & replace
        private bool change_cell_content_attribute(Excel.Range locateCell, string searchStr)
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
                locateCell.Characters[match.Index + 1, match.Length].Font.Color = Excel.XlRgbColor.rgbLightGreen;
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
