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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    public partial class UserControl3 : UserControl
    {
        //user difinition variables
        ArrayList aryWords = new ArrayList();//words collection
        ArrayList aryFilePath = new ArrayList();//data source excel files list collection
        ArrayList aryFields = new ArrayList();//data fields name collection
        ArrayList aryKeyFieldName = new ArrayList();//Search key fields name collection
        JObject joFilterFieldsKeys = new JObject();//fields filter keys collection
        JObject joFilterFieldsExcludeKeys = new JObject();//fields filter Exclude keys collection
        //Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;
        object missing = System.Reflection.Missing.Value;//null parameter
        //default input value file
        string defaultValFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "ExcedlAddIn1.conf");

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (File.Exists(defaultValFile) == true)
            {
                try
                {
                    JObject jo = JObject.Parse(File.ReadAllText(defaultValFile));
                    txtPath1.Text = jo["filePath1"].ToString();
                    txtPath2.Text = jo["filePath2"].ToString();
                    txtPath3.Text = jo["filePath3"].ToString();
                    txtPath4.Text = jo["filePath4"].ToString();
                }
                catch { }
            }
        }

        public UserControl3()
        {
            InitializeComponent();
        }

        #region "Checkbox Function"

        //欄位條件排除定義文字檔是否作用
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                button5.Enabled = true;
                txtPath4.Enabled = true;
            }
            else
            {
                button5.Enabled = false;
                txtPath4.Enabled = false;
            }
        }

        //生字候選清單文字檔 作用與否
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                button2.Enabled = true;
                txtPath1.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
                txtPath1.Enabled = false;
            }
        }

        #endregion

        #region "Button - User Choose Files"

        //open 生字候選清單
        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt file|*.txt";
            openFileDialog1.Title = "選擇生字候選清單文字檔";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath1.Text = "";
                foreach(string strFileName in openFileDialog1.FileNames){
                    if (txtPath1.Text == ""){
                        txtPath1.Text += strFileName;
                    }else {
                        txtPath1.Text += ";" + strFileName;
                    }                    
                }                
            }
        }
        //open 生字資料表路徑定義
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt file|*.txt";
            openFileDialog1.Title = "選擇生字資料表路徑定義文字檔";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath2.Text = openFileDialog1.FileName;
            }
        }
        //open 擷取來源欄位定義
        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt file|*.txt";
            openFileDialog1.Title = "選擇擷取來源欄位定義文字檔";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath3.Text = openFileDialog1.FileName;
            }
        }
        //欄位篩選排除條件定義
        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt file|*.txt";
            openFileDialog1.Title = "欄位篩選排除條件定義檔";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath4.Text = "";
                foreach (string strFileName in openFileDialog1.FileNames)
                {
                    if (txtPath4.Text == "")
                    {
                        txtPath4.Text += strFileName;
                    }
                    else
                    {
                        txtPath4.Text += ";" + strFileName;
                    }
                }  
            }
        }

        #endregion

        #region "Button - Run"

        //執行擷取生字資料作業
        private void button1_Click(object sender, EventArgs e)
        {
            txtMessage.Text = "";//clear message
            //Source Excel Files
            if (checkSourceFiles(txtPath2.Text) == false) { return; }

            //open fields definition list
            if (checkFieldsList(txtPath3.Text) == false) { return; }

            //check and get key words lists content(include condition)
            if(checkBox1.Checked==true){if (checkKeyWords(txtPath1.Text) == false) { return; }}

            //get exclude condition
            if (checkBox2.Checked == true){if (checkExcludeFieldsVal(txtPath4.Text) == false){return;}}

            //txtMessage.Text += "Words: "+aryWords.Count + Environment.NewLine;
            txtMessage.Text += "Filess: " + aryFilePath.Count + Environment.NewLine;
            txtMessage.Text += "Fields: " + aryFields.Count + Environment.NewLine;

            //=============Run================

            //new worksheet
            Excel.Workbook newWB = (Excel.Workbook)Globals.ThisAddIn.Application.Workbooks.Add();
            Excel.Worksheet newSheet = (Excel.Worksheet)newWB.Worksheets[1];
            newSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            //loop in every source file
            //fill new title on new sheet
            Excel.Range titleCell = newSheet.get_Range("A1", Type.Missing);//first cell of title row in new sheet
            foreach (string fn in aryFields)
            {
                titleCell.Value = fn;
                titleCell = titleCell.Offset[0, 1];//next cell on the same row
            }

            //ready copy to new cell by row & cell
            int nn = 1;//row 1 fill title, row number start from nn+1=2
            for (int v1 = 0; v1 < aryFilePath.Count; v1++)
            {
                //open excel
                if (File.Exists((string)aryFilePath[v1]) == false)
                {
                    txtMessage.Text += aryFilePath[v1] + " 不存在" + Environment.NewLine;
                    continue;
                }
                txtMessage.Text += aryFilePath[v1] + Environment.NewLine;
                oWB = (Excel.Workbook)Globals.ThisAddIn.Application.Workbooks.Open((string)aryFilePath[v1]);
                oWB.Windows[1].Visible = false;
                oSheet = (Excel.Worksheet)oWB.Sheets[1];
                //Excel.Range usedRange = null;
                //all cells

                //usedRange = oSheet.UsedRange;
                Excel.Range headRow = oSheet.Rows[1];//header row
                Excel.Range realDataRange = getUsedRange(oSheet);//no empty range
                //txtMessage.Text += realDataRange.Rows[1].Row.ToString() + " ~ " + realDataRange.Rows[realDataRange.Rows.Count].Row.ToString() + Environment.NewLine;

                //columns collection
                JObject aryCols = collectCols(headRow, aryFields);

                //rows collection
                //JObject aryRows = collectRows(usedRange, aryWords, (string)aryKeyFieldName[0]);
                ArrayList aryRows = new ArrayList();
                if(checkBox1.Checked==true && checkBox2.Checked==true){
                    aryRows = collectRows1(realDataRange, joFilterFieldsKeys, joFilterFieldsExcludeKeys);
                }else if(checkBox1.Checked==true && checkBox2.Checked==false){
                    aryRows = collectRows2(realDataRange, joFilterFieldsKeys);
                }else if(checkBox1.Checked==false && checkBox2.Checked==true){
                    aryRows = collectRows3(realDataRange, joFilterFieldsExcludeKeys);
                }
                
                //no any row had to be catch, next source file
                if (aryRows == null) {
                    oWB.Close(false);
                    txtMessage.Text+="Rows null."+Environment.NewLine;
                    continue; 
                }

                int rowCount = 0;
                foreach (var rowNum in aryRows)
                {
                    if (rowCount != aryRows.Count - 1)
                    {
                        txtMessage.Text += rowNum.ToString() + " , ";
                    }
                    else
                    {
                        txtMessage.Text += rowNum.ToString() + Environment.NewLine;
                    }
                    rowCount += 1;
                }

                //Remove first row (head row)
                aryRows = removeHeadRow(aryRows);

                /*
                int rowCount = 0;
                foreach(var rowNum in aryRows){
                    if (rowCount != aryRows.Count-1){
                        txtMessage.Text += rowNum.ToString() + " , ";
                    }else {
                        txtMessage.Text += rowNum.ToString()+Environment.NewLine;
                    }
                    rowCount += 1;
                }*/


                 
                //loop for every key word for getting fields
                //warning for duplication field name on the same sheet source
                /*
                foreach (string field in aryFields)
                {
                    if ((aryCols[field] as JArray).Count > 1)
                    {
                        txtMessage.Text += "請注意...在[" + aryFilePath[v1] + "]" + Environment.NewLine;
                        txtMessage.Text += "來源中有重複的欄位:[" + field + "]重複的欄位資料將被捨棄。" + Environment.NewLine;
                    }
                    else if ((aryCols[field] as JArray).Count == 0)
                    {
                        txtMessage.Text += "請注意...在[" + aryFilePath[v1] + "]" + Environment.NewLine;
                        txtMessage.Text += "來源中沒有欄位:[" + field + "]存在。" + Environment.NewLine;
                    }
                }*/

                //after row & column collection ready
                foreach(var rowPoi in aryRows){
                    nn += 1;
                    //txtMessage.Text += "Row: " + catchRowNum.ToString()+" Type:"+catchRowNum.GetType().ToString()+ Environment.NewLine;
                    Excel.Range newCell = newSheet.get_Range("A" + nn.ToString(), Type.Missing);
                    foreach(string fn in aryFields){
                        JArray colPoiCollect = (JArray)aryCols[fn];
                        if(colPoiCollect==null){//source no field 
                            newCell = newCell.Offset[0, 1];
                            continue;
                        }
                        if (colPoiCollect.Count==0)//source no field 
                        {//source no field 
                            newCell = newCell.Offset[0, 1];
                            continue;
                        }
                        //source field name exist in souce sheet
                        int colPoi = (int)colPoiCollect[0];
                        Excel.Range sourceCell = oSheet.Rows[rowPoi].Columns[colPoi];
                        sourceCell.Copy(newCell);
                        newCell = newCell.Offset[0, 1];
                    }
                }

                /*old -- delete
                foreach (string wd in aryWords) 
                {
                    if (aryRows[wd] == null) { continue; }
                    foreach (int r in (aryRows[wd] as JArray))
                    {
                        Excel.Range wRow = usedRange.Rows[r];
                        nn += 1;
                        Excel.Range newCell = newSheet.get_Range("A" + nn.ToString(), Type.Missing);
                        foreach (string field in aryFields)
                        {
                            if ((aryCols[field] as JArray).Count == 0)
                            {
                                newCell = newCell.Offset[0, 1];
                                continue;
                            }
                            //txtMessage.Text += field + Environment.NewLine;
                            //txtMessage.Text += "column count: "+wRow.Columns.Count + Environment.NewLine;
                            Excel.Range wCell = null;
                            JArray aryColNum = aryCols[field] as JArray;
                            //copy the first column address into new sheet cell,ignore second duplication field name column
                            wCell = wRow.Columns[aryColNum[0]];
                            wCell.Copy(newCell);
                            newCell = newCell.Offset[0, 1];
                        }
                    }
                }*/
                oWB.Close(false);
            }

            //save user input
            JObject jo = new JObject(
                new JProperty("filePath1", txtPath1.Text),
                new JProperty("filePath2", txtPath2.Text),
                new JProperty("filePath3", txtPath3.Text),
                new JProperty("filePath4", txtPath4.Text));
            File.WriteAllText(defaultValFile, jo.ToString());
            MessageBox.Show("OK!");

            return;
        }

        #endregion

        #region "Row & Column Position Collection"

        //rows collection
        private JObject collectRows(Excel.Range operateRange, ArrayList dictAry, string colName) 
        {
            //ArrayList collectAddress = new ArrayList();
            JObject collectAddress = new JObject();
            //match title column (only one match)
            int keyColNum = 0;
            Excel.Range titleRow = operateRange.Rows[1];
            foreach(Excel.Range titleCell in titleRow.Cells){
                if(titleCell.Text==colName){
                    keyColNum=titleCell.Column;
                    break;
                }
            }
            //no match col name
            if(keyColNum==0){
                return null;
            }
            Excel.Range cols=operateRange.Columns[keyColNum];
            foreach (string keyWord in dictAry)
            {
                JArray rowNum = new JArray();
                object missing = System.Reflection.Missing.Value;//null parameter
                Excel.Range firstFind = null;
                string firstAddress = "";
                Excel.Range findRng = cols.Find(keyWord, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);
                if (findRng == null)
                {
                    continue;//no search result
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
                        //collectAddress.Add(keyWord,findRng.Row);
                        rowNum.Add(findRng.Row);
                        //txtMessage.Text += keyWord+": "+findRng.Address+ Environment.NewLine;
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
                        //collectAddress.Add(findRng.Row.ToString(),keyWord);
                        rowNum.Add(findRng.Row);
                        //txtMessage.Text += keyWord + ": " + findRng.Address + Environment.NewLine;
                    }
                    findRng = cols.FindNext(findRng);
                }
                collectAddress.Add(keyWord, rowNum);
            }
            //txtMessage.Text += "Rows collectAddress: " +collectAddress.Count + Environment.NewLine;
            return collectAddress;
        }

        //include & exclude
        private ArrayList collectRows1(Excel.Range operateRange, JObject includeJSON , JObject excludeJSON) 
        {
            //filter condition
            //both include and exclude

            //include condition
            ArrayList arylIncludeRows = new ArrayList();
            JObject joAllFilterIncludeRows = new JObject();
            foreach (var j_in_obj in includeJSON)
            {
                JObject joSingleFieldRows = getRowsCollection(operateRange, j_in_obj.Key, (JArray)j_in_obj.Value);
                //Intersection all fields match rows
                JArray jaVal = (JArray)joSingleFieldRows[j_in_obj.Key];
                joAllFilterIncludeRows.Add(j_in_obj.Key, jaVal);
            }
            arylIncludeRows = intersectionFieldsRows(joAllFilterIncludeRows);

            //exclude condition
            ArrayList arylExcludeRows = new ArrayList();
            JObject joAllFilterExcludeRows = new JObject();
            foreach (var j_ex_obj in excludeJSON)
            {
                JObject joSingleFieldRows = getRowsCollection(operateRange, j_ex_obj.Key, (JArray)j_ex_obj.Value);
                JArray jaVal = (JArray)joSingleFieldRows[j_ex_obj.Key];
                joAllFilterExcludeRows.Add(j_ex_obj.Key, jaVal);
            }
            arylExcludeRows = unionFieldsRows(joAllFilterExcludeRows);

            //Include & Exclude mixed operation(Removed Exclude Rows from Include Rows)
            ArrayList aryMixedRows = mixedInExFieldsRows(arylIncludeRows, arylExcludeRows);
            return aryMixedRows;
        }

        //Include only
        private ArrayList collectRows2(Excel.Range operateRange, JObject includeJSON)
        {
            //include condition
            ArrayList arylIncludeRows = new ArrayList();
            JObject joAllFilterIncludeRows = new JObject();
            foreach (var j_in_obj in includeJSON)
            {
                JObject joSingleFieldRows = getRowsCollection(operateRange, j_in_obj.Key, (JArray)j_in_obj.Value);
                //Intersection all fields match rows
                JArray jaVal = (JArray)joSingleFieldRows[j_in_obj.Key];
                joAllFilterIncludeRows.Add(j_in_obj.Key, jaVal);
            }
            arylIncludeRows = intersectionFieldsRows(joAllFilterIncludeRows);
            return arylIncludeRows;
        }

        //Exclude only(all rows are the Include rows)
        private ArrayList collectRows3(Excel.Range operateRange, JObject excludeJSON) 
        {
            if(operateRange.Equals(null)){
                return null;
            }
            if(operateRange.Cells.Count==0){
                return null;
            }
            //filter condition
            //include condition
            ArrayList arylIncludeRows = new ArrayList();
            txtMessage.Text += operateRange.Rows[1].Address.ToString()+Environment.NewLine;
            foreach(Excel.Range r1 in operateRange.Rows){
                //bool rowEmpty = ifRowEmpty(r1);
                //if (rowEmpty == false) {
                    arylIncludeRows.Add(r1.Row);
                //}
            }

            //exclude condition
            ArrayList arylExcludeRows = new ArrayList();
            JObject joAllFilterExcludeRows = new JObject();
            foreach (var j_ex_obj in excludeJSON)
            {
                JObject joSingleFieldRows = getRowsCollection(operateRange, j_ex_obj.Key, (JArray)j_ex_obj.Value);
                //txtMessage.Text += "joSingleFieldRows: " + joSingleFieldRows.ToString() + Environment.NewLine;
                JArray jaVal = (JArray)joSingleFieldRows[j_ex_obj.Key];
                joAllFilterExcludeRows.Add(j_ex_obj.Key, jaVal);
            }
            arylExcludeRows = unionFieldsRows(joAllFilterExcludeRows);


            //Include & Exclude mixed operation(Removed Exclude Rows from Include Rows)
            ArrayList aryMixedRows = mixedInExFieldsRows(arylIncludeRows, arylExcludeRows);
            return aryMixedRows;
        }

        //find match value in column for getting rows number collection
        private JObject getRowsCollection(Excel.Range searchRange,string fieldName,JArray aryVal) 
        {
            JObject collect_rows_match_key = new JObject();
            ArrayList memoAry = new ArrayList(); //for contains compare used only

            int colNum = 0;
            Excel.Range titleRow = searchRange.Rows[1];//The first must be column title name
            //search match name title cell, getting column number
            foreach (Excel.Range titleCell in titleRow.Cells)
            {
                if (titleCell.Text == fieldName)
                {
                    colNum = titleCell.Column;//get column number
                    break;
                }
            }
            if (colNum == 0){return null;}//no match col name
            //Excel.Range noheadSearchRange = removeHeadRow(searchRange);//remove head line, if range include 
            Excel.Range cols = searchRange.Columns[colNum];
            foreach (string keyWord in aryVal)
            {
                JArray rowsNumCollect = new JArray();//collect one key string find match result rows number in one field
                memoAry = new ArrayList();
                object missing = System.Reflection.Missing.Value;//null parameter
                Excel.Range firstFind = null;
                string firstAddress = "";
                Excel.Range findRng = cols.Find(keyWord, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);
                if (findRng == null)
                {
                    continue;//no search result
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
                        //exclude row number already exist
                        if (memoAry.Contains(findRng.Row) == false)//deduplication row number
                        {                            
                            rowsNumCollect.Add(findRng.Row);
                            memoAry.Add(findRng.Row);
                        }
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
                        if (memoAry.Contains(findRng.Row) == false)//deduplication row number
                        {                            
                            rowsNumCollect.Add(findRng.Row);
                            memoAry.Add(findRng.Row);
                        }
                    }
                    findRng = cols.FindNext(findRng);
                }
                collect_rows_match_key.Add(keyWord, rowsNumCollect);
            }
            
            // {field name:[row_number]}.
            JObject joFieldRows = new JObject();
            JArray jaFieldRowsNum = new JArray();
            memoAry = new ArrayList(); //for contains compare used only
            foreach(var joobj in collect_rows_match_key){
                JArray ja = (JArray)joobj.Value;
                foreach(var rowNum in ja) {
                    if(memoAry.Contains(rowNum)==false){
                        jaFieldRowsNum.Add(rowNum);
                        memoAry.Add(rowNum);
                    }
                }
            }
            joFieldRows.Add(fieldName,sortJArray(jaFieldRowsNum));
            //txtMessage.Text += "joFieldRows: " + joFieldRows + Environment.NewLine;
            return joFieldRows;
        }

        //columns collection
        private JObject collectCols(Excel.Range operateRange, ArrayList dictAry)
        {
            //ArrayList collectAddress = new ArrayList();
            JObject collectAddress = new JObject();
            //loop for every item in find array
            foreach(string keyWord in dictAry){
                JArray colNum = new JArray();
                foreach (Excel.Range thisCell in operateRange.Cells)
                {
                    if (thisCell.Text == keyWord)
                    {
                        colNum.Add(thisCell.Column);
                    }
                }
                collectAddress.Add(keyWord,colNum);
            }
            return collectAddress; 
        }

        #endregion

        #region "User Input Parameters Collection"

        //check & get Include fields value(multi files)
        private bool checkKeyWords(string strPath)
        {
            joFilterFieldsKeys = new JObject();
            //seprate files name
            if (strPath == "")
            {
                txtMessage.Text += "沒有輸入欄位篩選條件定義檔路徑" + Environment.NewLine;
                return false;
            }
            string[] keyWords_filesPaths = strPath.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string keyWord_filePath in keyWords_filesPaths)
            {
                //file not exist
                if (File.Exists(keyWord_filePath) == false)
                {
                    txtMessage.Text += keyWord_filePath + " 檔案不存在" + Environment.NewLine;
                    continue;
                }
                //file exist
                //open & read field filter definition
                JArray jContent = new JArray();
                ArrayList aryContent = new ArrayList();//for deduplication compare usage
                string jName = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(keyWord_filePath))
                {
                    int wCount = 0;
                    while (sr.Peek() >= 0)
                    {
                        string[] strValAry = jContent.ToObject<string[]>();
                        string val = sr.ReadLine() as string;
                        if (wCount == 0)
                        {
                            jName = val;//Text file first line must be field title
                        }
                        else if (aryContent.Contains(val) == false)//remove duplication
                        {
                            //txtMessage.Text += "val:" +val+ Environment.NewLine;
                            jContent.Add(val);
                            aryContent.Add(val);
                        }
                        wCount += 1;
                    }
                }
                joFilterFieldsKeys.Add(jName, jContent);
            }
            if ((joFilterFieldsKeys == null) || (joFilterFieldsKeys.Count == 0))
            {
                txtMessage.Text += "沒有定義任何關鍵詞" + Environment.NewLine;
                return false;
            }
            return true;
        }

        //check fields list(single path)
        private bool checkFieldsList(string strPath)
        {
            aryFields = new ArrayList();
            if (strPath == "")
            {
                txtMessage.Text += "沒有輸入擷取欄位名稱定義" + Environment.NewLine;
                return false;
            }
            //check file exist or not
            if (File.Exists(strPath) == false)
            {
                txtMessage.Text += strPath + "檔案路徑不存在" + Environment.NewLine;
                return false;
            }
            //open words list
            using (System.IO.StreamReader sr = new System.IO.StreamReader(strPath))
            {
                int wCount = 0;
                while (sr.Peek() >= 0)
                {
                    string val = sr.ReadLine() as string;
                    //remove duplication
                    if (!aryFields.Contains(val))
                    {
                        wCount += 1;
                        aryFields.Add(val);
                    }
                }
            }
            if (aryFields == null)
            {
                txtMessage.Text += "沒有定義任何來源欄位名稱." + Environment.NewLine;
                return false;
            }
            return true;
        }

        //check source excel files definition(single path)
        private bool checkSourceFiles(string strPath)
        {
            aryFilePath = new ArrayList();
            if (strPath == "")
            {
                txtMessage.Text += "沒有輸入生字資料表來源路徑定義" + Environment.NewLine;
                return false;
            }
            //check file exist or not
            if (File.Exists(strPath) == false)
            {
                txtMessage.Text += strPath + "檔案路徑不存在" + Environment.NewLine;
                return false;
            }
            //open source excel files list
            using (System.IO.StreamReader sr = new System.IO.StreamReader(strPath))
            {
                while (sr.Peek() >= 0)
                {
                    //aryFilePath.Add(sr.ReadLine());
                    aryFilePath.Add(sr.ReadLine());
                }
            }
            //no any file path definition
            if ((aryFilePath.Count == 0) || (aryFilePath == null))
            {
                txtMessage.Text += strPath + "沒有任何檔案路徑定義" + Environment.NewLine;
                return false;
            }
            return true;
        }

        //check & get Exclude fields value(multi files)
        private bool checkExcludeFieldsVal(string strPath)
        {
            joFilterFieldsExcludeKeys = new JObject();
            //seprate files name
            if (strPath == "")
            {
                txtMessage.Text += "沒有輸入欄位篩選排除條件定義檔路徑" + Environment.NewLine;
                return false;
            }
            string[] keyWords_filesPaths = strPath.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string keyWord_filePath in keyWords_filesPaths)
            {
                //file not exist
                if (File.Exists(keyWord_filePath) == false)
                {
                    txtMessage.Text += keyWord_filePath + " 檔案不存在" + Environment.NewLine;
                    continue;
                }
                //file exist
                //open & read field filter definition
                JArray jContent = new JArray();
                ArrayList aryContent = new ArrayList();//for deduplication compare usage
                string jName = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(keyWord_filePath))
                {
                    int wCount = 0;
                    while (sr.Peek() >= 0)
                    {
                        string val = sr.ReadLine() as string;
                        if (wCount == 0)
                        {
                            jName = val;//Text file first line must be field title
                        }
                        else if (aryContent.Contains(val)==false)
                        {
                            //remove duplication
                            jContent.Add(val);
                            aryContent.Add(val);
                        }
                        wCount += 1;
                    }
                }
                joFilterFieldsExcludeKeys.Add(jName, jContent);
            }
            if ((joFilterFieldsExcludeKeys == null) || (joFilterFieldsExcludeKeys.Count == 0))
            {
                txtMessage.Text += "排除條件沒有定義任何關鍵詞" + Environment.NewLine;
                return false;
            }
            return true;
        }

        #endregion

        #region "Tools Function"

        //Sort JArray
        private JArray sortJArray(JArray jarray) { 
            //Convert to ArrayList
            ArrayList aryList = new ArrayList();
            foreach(var jaObj in jarray){
                aryList.Add(jaObj);
            }
            //sort by arraylist method
            aryList.Sort();
            //Convert to JArray
            JArray retJArray = new JArray();
            foreach(var alObj in aryList){
                retJArray.Add(alObj);
            }
            return retJArray;
        }

        //Intersection all fields match rows in one array
        //jobject sample: {"生字":[2,3,4],"部首":[木,一,虫]}
        private ArrayList intersectionFieldsRows(JObject jobject) {
            if(jobject==null){
                return null;
            }
            if(jobject.Count==0){
                return null;
            }

            //Choose the min amount json pair
            int rowMin = 1048576;
            string fName = "";
            int iCount = 0;
            foreach(var joObj in jobject){
                string joName = (string)joObj.Key;
                JArray jaVal = (JArray)joObj.Value;
                if (jaVal.Count < rowMin) {
                    rowMin=jaVal.Count;// min rows count
                    fName = joName;//min field name
                }
                    iCount+=1;
            }
            //
            //txtMessage.Text += "Min Field Name: "+fName+" Min Number:"+rowMin+Environment.NewLine;
            //ArrayList basicAry = new ArrayList();
            //JArray pAry = (JArray)jobject[fName];
            //basicAry = (ArrayList)jobject[fName].ToObject(typeof(ArrayList));
            //Intersection(fill min amount array for compare)
            ArrayList intersection = (ArrayList)jobject[fName].ToObject(typeof(ArrayList));
            foreach (var joObj in jobject){
                string joName = (string)joObj.Key;
                JArray jaVal = (JArray)joObj.Value;
                //txtMessage.Text += joName + Environment.NewLine+jaVal+ Environment.NewLine;
                if (joName == fName) { continue; }//compare target no need to be loop
                foreach(var poi in jaVal){//loop for every poi on jobject every jarray value
                    bool rowContain = false;
                    foreach (var basicPoi in intersection)
                    {//loop for every poi jarray on basic array
                        // poi type:jvalue , basicPoi type:int64 , we need convert to the same type for comparing
                        if (poi.ToString() == basicPoi.ToString()) { rowContain = true; continue; }
                    }
                    if(rowContain==true){intersection.Add(poi); }
                }
            }
            //txtMessage.Text += "Intersection Count:" + intersection.Count + Environment.NewLine;
            return intersection;
        }

        //Union all fields match rows in one array
        //jobject sample: {"字音辨別":[不製作],"成語教學":[不製作]}
        private ArrayList unionFieldsRows(JObject jobject)
        {
            if (jobject == null)
            {
                return null;
            }
            if (jobject.Count == 0)
            {
                return null;
            }

            //Choose the min amount json pair
            int rowMax = 0;
            string fName = "";
            int iCount = 0;
            foreach (var joObj in jobject)
            {
                string joName = (string)joObj.Key;
                JArray jaVal = (JArray)joObj.Value;
                if (jaVal.Count > rowMax)
                {
                    rowMax = jaVal.Count;// min rows count
                    fName = joName;//min field name
                }
                iCount += 1;
            }
            //ArrayList basicAry = new ArrayList();
            //basicAry = (ArrayList)jobject[fName].ToObject(typeof(ArrayList));
            //Union
            ArrayList union = new ArrayList();
            union = (ArrayList)jobject[fName].ToObject(typeof(ArrayList));
            foreach (var joObj in jobject)
            {
                string joName = (string)joObj.Key;
                JArray jaVal = (JArray)joObj.Value;
                //txtMessage.Text += "joName: " + joName+Environment.NewLine+" jaVal:"+jaVal.ToString()+ Environment.NewLine;
                if (joName == fName) { continue; }//compare target no need to be loop
                foreach (var poi in jaVal)
                {//loop for every poi on jobject every jarray value
                    bool rowContain = false;
                    foreach (var basicPoi in union)
                    {//loop for every poi jarray on basic array
                        // poi type:jvalue , basicPoi type:int64 , we need convert to the same type for comparing
                        if (poi.ToString() == basicPoi.ToString()) { rowContain = true; continue; }
                    }
                    if(rowContain==false){union.Add(poi);}
                }
            }
            return union;
        }

        //Include & Exclude mixed operation(Removed Exclude Rows from Include Rows)
        private ArrayList mixedInExFieldsRows(ArrayList includeAryl, ArrayList excludeAryl){
            //no include condition
            if (includeAryl == null){return null;}
            if (includeAryl.Count == 0){return null;}
            //had include condition, no exclude condition
            if (excludeAryl == null) { return includeAryl; }
            if (excludeAryl.Count == 0) { return includeAryl; }
            //mixed
            ArrayList mixedAryl = new ArrayList();
            foreach(var incl in includeAryl){
                bool ifContain = false;
                foreach(var excl in excludeAryl){
                    if(incl.ToString()==excl.ToString()){
                        ifContain = true;
                        break;
                    }
                }
                if(ifContain==false){
                    mixedAryl.Add(incl);
                }
            }
            return mixedAryl;
        }

        //Check if row empty
        private bool checkEmptyRow(Excel.Range singleRow) {
            bool ret = true;
            foreach(Excel.Range inCell in singleRow.Cells){
                if (inCell == null) {
                    continue; 
                }
                if ((string)inCell.Text.Trim() == "") 
                {
                    continue; 
                }
                ret = false;
                break;
            }
            return ret;
        }

        //Get worksheet used range(trim empty row at the end)
        private Excel.Range getUsedRange(Excel.Worksheet usedSheet) {
            Excel.Range usedRange = usedSheet.UsedRange;
            Excel.Range startCell = usedSheet.Rows[1].Cells[1];
            //Excel.Range startCell = usedSheet.Rows[2].Cells[1];
            Excel.Range endCell = usedSheet.Rows[usedRange.Rows.Count].Cells[usedRange.Rows[usedRange.Rows.Count].Cells.Count];

            //pass header row
            int emptyRowLimit = 10;
            int emptyRowCount = 0;
            int lastRow = endCell.Row;
            for (int v1 = 2; v1 <= usedRange.Rows.Count;v1++ )
            {
                bool emptyRow = ifRowEmpty(usedRange.Rows[v1]);
                //txtMessage.Text += "emptyRow: " + v1 + " is " + emptyRow.ToString() + Environment.NewLine;
                if (emptyRow == false) { 
                    //row is not empty, empty row number zero
                    lastRow = v1; 
                    emptyRowCount=0;
                    endCell = usedRange.Rows[v1].Cells[usedRange.Rows[v1].Cells.Count];
                    continue;
                }
                else if (emptyRow == true && emptyRowCount <= emptyRowLimit){
                    //empty row not reach the max limit
                    emptyRowCount += 1;
                }
                else if (emptyRow == true && emptyRowCount > emptyRowLimit) { 
                    //empty row reach the max limit
                    break;
                }
            }
            usedRange = usedSheet.get_Range(startCell, endCell);
            return usedRange;
        }

        //remove head line(title)
        /*
        private Excel.Range removeHeadRow(Excel.Range operateRange) {
            if (operateRange.Equals(null)){return null;}
            if (operateRange.Rows.Count==0) { return null; }
            Excel.Range startCell = operateRange.Rows[1].Cells[1];
            Excel.Range endCell = operateRange.Rows[operateRange.Rows.Count].Cells[operateRange.Rows[operateRange.Rows.Count].Cells.Count];
            if(startCell.Row==1){
                startCell = operateRange.Rows[2].Cells[1];
                operateRange = operateRange.get_Range(startCell,endCell);//if include head line, reset start range
            }
            return operateRange;
        }
        */
        private ArrayList removeHeadRow(ArrayList aryl) {
            ArrayList newAryl = new ArrayList();
            if(aryl.Equals(null)){return aryl;}
            if (aryl.Count==0){return aryl;}
            foreach(var obj in aryl){
                if (obj.ToString() != "1") {
                    newAryl.Add(obj);
                }
            }
            return newAryl;
        }


        //Check if row is empty(true:empty row , false:not empty row)
        private bool ifRowEmpty(Excel.Range singleRow){
            //txtMessage.Text += "singleRow: " + singleRow.Address.ToString() + Environment.NewLine;
            bool ret = true;
            int checkCellMaxLimit = 26;
            foreach(Excel.Range singleCell in singleRow.Cells){
                if (singleCell.Column > checkCellMaxLimit) { break; }//to the compare cell max,no need to be compare
                if (singleCell.Text.ToString()==""){//cell emptt
                    if(singleCell.Column <= checkCellMaxLimit){
                        ret = true; continue;//cell empty, and not reach the column limit
                    }else{
                        ret = true; break;//cell empty, and reach the column limit
                    }
                }else if ((string)singleCell.Text.Trim() == ""){//cell is empty after trim
                    if (singleCell.Column <= checkCellMaxLimit){
                        ret = true; continue;//cell empty, and not reach the column limit
                    }else{
                        ret = true; break;//cell empty, and reach the column limit
                    }
                }
                //If any cell is not empty, row is not empty
                ret = false;
                break;//row had data, no needed to continue loop
            }
            //txtMessage.Text += "Row: " + singleRow.Row + " is " + ret.ToString() + Environment.NewLine;
            return ret;
        }


        #endregion


    }
}
