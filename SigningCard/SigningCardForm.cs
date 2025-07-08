//using DocumentFormat.OpenXml.Spreadsheet;
//using Moq;
using Aspose.Cells;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
//using DocumentFormat.OpenXml.Wordprocessing;

namespace SigningCard
{
    public partial class SigningCardForm : Form
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern void OutputDebugString(string message);
        const int dataViewColums = 7;
        int totalDays = 0;
        int weekDayFirstDay = 0;//本月第一天是星期几
        bool[] holidayData = new bool[31];//最多31天
        string execlPath = null;
        string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        class Reason
        {
            public string Reason1 { get; set; }
            public string Reason2 { get; set; }
        }
        BindingList<Reason> ReasonList = new BindingList<Reason>();

        public SigningCardForm()
        {
            InitializeComponent();
        }

        private void SelectMonth(int year, int month)
        {
            for (int i = 0; i < 31; i++)
            {
                holidayData[i] = false;
            }
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    dataGridViewHoliday.Rows[j].Cells[i].Style.BackColor = System.Drawing.Color.White;
                    dataGridViewHoliday.Rows[j].Cells[i].Value = "";
                }
            }

            totalDays = DateTime.DaysInMonth(year, month);
            weekDayFirstDay = (int)DateTime.Parse(
                string.Format("{0:D}/{1:D}/{2:D}", year, month, 1)).DayOfWeek;

            for (int i = 0; i < totalDays; i++)
            {
                int row = (i + weekDayFirstDay) / dataViewColums;
                int colum = (i + weekDayFirstDay) % dataViewColums;
                dataGridViewHoliday.Rows[row].Cells[colum].Value = i + 1;
                if (colum == 0 || colum == 6)
                {
                    holidayData[i] = true;
                    dataGridViewHoliday.Rows[row].Cells[colum].Style.BackColor = System.Drawing.Color.Green;
                }
                else
                {
                    holidayData[i] = false;
                }
            }
        }
        private void ButtonImport_Click(object sender, EventArgs e)
        {
            openFileDialogPunchRecord.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";
            DialogResult dr = openFileDialogPunchRecord.ShowDialog();
            if (dr == DialogResult.OK)
            {
                execlPath = openFileDialogPunchRecord.FileName;
                AnalyzeExcel(execlPath);
            }
        }

        private void DataGridViewHoliday_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void SigningCardForm_Load(object sender, EventArgs e)
        {
            dataGridViewHoliday.AllowUserToResizeRows = false;
            dataGridViewHoliday.Rows.Add(5);
            dataGridViewHoliday.AllowUserToResizeColumns = false;
            for (int i = 0; i < this.dataGridViewHoliday.Columns.Count; i++)
            {
                dataGridViewHoliday.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridViewHoliday.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridViewHoliday.Columns[i].ReadOnly = true;
            }

            DateTime date = DateTime.Now.AddMonths(0);
            SelectMonth(date.Year, date.Month);

            try
            {
                using (System.IO.StreamReader file = System.IO.File.OpenText(System.AppDomain.CurrentDomain.BaseDirectory + "reason.json"))
                {
                    string json = file.ReadToEnd();
                    ReasonList = JsonConvert.DeserializeObject<BindingList<Reason>>(json);
                }
            }
            catch (Exception)
            {

            }
            dataGridView1.AutoGenerateColumns = false;                    // 防止自由生成所有数据列

            dataGridView1.DataSource = new BindingSource(ReasonList, null); 
        }

        private void DataGridViewHoliday_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void DataGridViewHoliday_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int index = e.RowIndex * dataViewColums + e.ColumnIndex - weekDayFirstDay;
            if (index >= totalDays || index < 0)
            {
                return;
            }
            if (holidayData[index])
            {
                dataGridViewHoliday.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.White;
                holidayData[index] = false;
            }
            else
            {
                dataGridViewHoliday.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.Green;
                holidayData[index] = true;
            }
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            SelectMonth(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month);
        }

        private void AnalyzeExcel(string fileName)
        {
            List<DateTime> importDateList = new List<DateTime>();
            string strNameNO = "";
            if (fileName != null)
            {
                string excelName = fileName;
                Aspose.Cells.Workbook excel = new Aspose.Cells.Workbook(excelName);
                importDateList = GetImportExcelRoute(excel);
                strNameNO = excel.Worksheets[0].Cells[1, 1].StringValue + "_" + excel.Worksheets[0].Cells[1, 0].StringValue;
            }
            else
            {
                if (Clipboard.ContainsText())
                {
                    //dateTimePicker1.Value.Year, dateTimePicker1.Value.Month
                    string clipboardText = Clipboard.GetText();

                    // 3. 使用正则表达式匹配时间戳
                    string timePattern = @"(\d{2}:\d{2}:\d{2})"; // 匹配时间戳
                    Regex timeRegex = new Regex(timePattern);

                    // 4. 使用正则表达式匹配日期
                    string dayPattern = @"^\d{1,2}$"; // 匹配日期，假设日期为1-2位数字
                    Regex dayRegex = new Regex(dayPattern);

                    // 5. 解析数据并提取每天的打卡时间
                    int currentDay = 0; // 当前处理的日期
                    bool hasPunchTime = false; // 标记当前日期是否有打卡记录

                    var year = Convert.ToInt32(Regex.Match(clipboardText, @"年份:\s*(\d{4})").Groups[1].Value);
                    var month = Convert.ToInt32(Regex.Match(clipboardText, @"月份:\s*(\d{1,2})").Groups[1].Value);
                    // 正则表达式匹配名字_ID
                    string idPattern = @"\w+_\d+";

                    // 查找所有匹配的结果
                    MatchCollection matches = Regex.Matches(clipboardText, idPattern);
                    if (matches.Count>0)
                    {
                        strNameNO = Regex.Replace(matches[0].Value, @"^\d+", ""); ;
                    }

                    foreach (var line in clipboardText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string trimmedLine = line.Trim();

                        // 5.1 检查当前行是否是日期
                        if (dayRegex.IsMatch(trimmedLine))
                        {
                            // 提取日期
                            if (currentDay != 0 && hasPunchTime)
                            {
                                // 如果上一天有打卡时间，继续处理下一天
                                currentDay = 0; // 处理完一天后重置
                            }

                            currentDay = int.Parse(trimmedLine);
                            hasPunchTime = false; // 重置标志

                            continue; // 跳过这一行，处理下一行的时间
                        }

                        // 5.2 如果当前行有打卡时间，匹配并添加打卡记录
                        var timeMatches = timeRegex.Matches(trimmedLine);
                        if (timeMatches.Count > 0)
                        {
                            hasPunchTime = true; // 表示该天有打卡记录
                            foreach (Match match in timeMatches)
                            {
                                string timeStr = match.Value;
                                DateTime punchTime = DateTime.ParseExact(timeStr, "HH:mm:ss", null);

                                // 将打卡时间和日期合并
                                if (year!=0 && month!=0)
                                {
                                    DateTime punchDateTime = new DateTime(year, month, currentDay, punchTime.Hour, punchTime.Minute, punchTime.Second);
                                    importDateList.Add(punchDateTime);
                                }
                                else
                                {
                                    DateTime punchDateTime = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, currentDay, punchTime.Hour, punchTime.Minute, punchTime.Second);
                                    importDateList.Add(punchDateTime);
                                }
                            }
                        }
                    }

                }
            }
             
            if (importDateList.Count <= 0)
            {
                return;
            }

            List<DateTime> singingCardList = new List<DateTime>();//签卡
            List<DateTime> overtimeList = new List<DateTime>();//加班
            for (int i = 0; i < totalDays; i++)
            {
                //一个一个找签卡
                //当天的签卡情况
                List<DateTime> curDayList = new List<DateTime>();
                foreach (var item in importDateList)
                {
                    if (item.Day == i + 1)
                    {
                        curDayList.Add(item);
                    }
                }

                if (true == holidayData[i])
                {
                    if (curDayList.Count > 0)
                    {
                        //休息日 有加班 暂时不处理 获取加班后手动添加
                        for (int k = 0; k < curDayList.Count; k++)
                        {
                            overtimeList.Add(curDayList[k]);
                        }
                    }
                }
                else
                {
                    DateTime[] dtModel = new DateTime[6]{new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 8, 30, 0),
                    new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 12, 0, 0),
                    new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 13, 30, 0),
                    new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 0, 0),
                    new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 30, 0),
                    new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 20, 30, 0)};
                    //各时段上下班判断有效区间
                    DateTime dtAMU1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 7, 30, 0);
                    DateTime dtAMU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 8, 35, 59);

                    DateTime dtAMD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 11, 55, 0);
                    DateTime dtAMD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 12, 45, 59);

                    DateTime dtPMU1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 12, 46, 0);
                    DateTime dtPMU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 13, 35, 59);

                    DateTime dtPMD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 17, 55, 0);
                    DateTime dtPMD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 15, 59);

                    //加班有效区间
                    DateTime dtOTU1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 16, 0);
                    DateTime dtOTU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 59, 59);

                    DateTime dtOTD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 19, 01, 0);
                    DateTime dtOTD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 23, 59, 59);

                    //夜班
                    DateTime dtND1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 0, 0, 0);
                    DateTime dtND2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 7, 29, 59);

                    //正常上班
                    bool[] szbSingingCardSts = new bool[6] { false, false, false, false, false, false };

                    for (int j = 0; j < curDayList.Count; j++)
                    {
                        if (curDayList[j] >= dtAMU1 && curDayList[j] <= dtAMU2)
                        {
                            //早上签卡存在
                            szbSingingCardSts[0] = true;
                        }
                        else if (curDayList[j] >= dtAMD1 && curDayList[j] <= dtAMD2)
                        {
                            szbSingingCardSts[1] = true;
                        }
                        else if (curDayList[j] >= dtPMU1 && curDayList[j] <= dtPMU2)
                        {
                            szbSingingCardSts[2] = true;
                        }
                        else if (curDayList[j] >= dtPMD1 && curDayList[j] <= dtPMD2)
                        {
                            szbSingingCardSts[3] = true;
                        }
                        else if (curDayList[j] >= dtOTU1 && curDayList[j] <= dtOTU2)
                        {
                            szbSingingCardSts[4] = true;
                            overtimeList.Add(curDayList[j]);
                        }
                        else if (curDayList[j] >= dtOTD1 && curDayList[j] <= dtOTD2)
                        {
                            szbSingingCardSts[5] = true;
                            overtimeList.Add(curDayList[j]);
                        }
                        else if (curDayList[j] >= dtND1 && curDayList[j] <= dtND2)
                        {
                            overtimeList.Add(curDayList[j]);
                        }
                    }

                    //每天需要的签卡 有加班缺卡则另外添加
                    for (int k = 0; k < 3; k++) //加班连班情况第四个卡需要考虑是否存在只有一个加班下班卡情况
                    {
                        if (false == szbSingingCardSts[k])
                        {
                            singingCardList.Add(dtModel[k]);
                        }
                    }
                    if (true == szbSingingCardSts[3])
                    {
                        if (true == szbSingingCardSts[4] && false == szbSingingCardSts[5])
                        {
                            //加班下班缺卡
                            singingCardList.Add(dtModel[5]);
                            overtimeList.Add(dtModel[5]);
                        }
                        else if (false == szbSingingCardSts[4] && true == szbSingingCardSts[5])
                        {
                            //加班上班缺卡
                            singingCardList.Add(dtModel[4]);
                            overtimeList.Add(dtModel[4]);
                        }
                    }
                    else if (false == szbSingingCardSts[3])
                    {
                        if (true == szbSingingCardSts[4] && false == szbSingingCardSts[5])
                        {
                            //下班缺卡  及 加班下班缺卡
                            singingCardList.Add(dtModel[3]);
                            singingCardList.Add(dtModel[5]);
                            overtimeList.Add(dtModel[5]);
                        }
                        else if (false == szbSingingCardSts[4] && true == szbSingingCardSts[5])
                        {
                            if (checkBox1.Checked)
                            {
                                //连班 不打下班卡 及加班上班卡
                                overtimeList.Add(dtModel[3]);
                            }
                            else
                            {
                                singingCardList.Add(dtModel[3]);
                                singingCardList.Add(dtModel[4]);
                                overtimeList.Add(dtModel[4]);
                            }
                        }
                        else if (false == szbSingingCardSts[4] && false == szbSingingCardSts[5])
                        {
                            //下班缺卡无加班
                            singingCardList.Add(dtModel[3]);
                        }
                    }

                }
            }
            singingCardList.Sort();
            overtimeList.Sort();

            //写回excel中
            //签卡
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(System.IO.Stream.Null);
                Aspose.Cells.Worksheet sheet = wb.Worksheets[0];
                //设置样式
                Aspose.Cells.Style style = wb.CreateStyle();
                style.ForegroundColor = System.Drawing.Color.FromArgb(128, 128, 128);
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Center;
                style.Pattern = BackgroundType.Solid;
                sheet.Name = "明细1";
                //绑定数据
                sheet.Cells[0, 0].PutValue("序号");
                sheet.Cells[0, 1].PutValue("员工姓名");
                sheet.Cells[0, 2].PutValue("签卡日期");
                sheet.Cells[0, 3].PutValue("时间");
                sheet.Cells[0, 4].PutValue("类型");
                sheet.Cells[0, 5].PutValue("事由");
                //绑定样式
                sheet.Cells[0, 0].SetStyle(style);
                sheet.Cells[0, 1].SetStyle(style);
                sheet.Cells[0, 2].SetStyle(style);
                sheet.Cells[0, 3].SetStyle(style);
                sheet.Cells[0, 4].SetStyle(style);
                sheet.Cells[0, 5].SetStyle(style);


                Aspose.Cells.Workbook excelTmp = new Aspose.Cells.Workbook(System.AppDomain.CurrentDomain.BaseDirectory + "测试文件/签卡申请流程模板.xls");
                Aspose.Cells.Style styleTmp1 = excelTmp.Worksheets[0].Cells[1,2].GetStyle();
                Aspose.Cells.Style styleTmp2 = excelTmp.Worksheets[0].Cells[1,3].GetStyle();


                int i = 1;
                foreach (var item in singingCardList)
                {
                    sheet.Cells[i, 0].PutValue(i);
                    sheet.Cells[i, 1].PutValue(strNameNO);
                    
                    sheet.Cells[i, 2].PutValue(item.ToString("yyyy-MM-dd"));
                    sheet.Cells[i, 3].PutValue(item.ToString("HH:mm"));
                    sheet.Cells[i, 4].PutValue("正常签卡");
                    //sheet.Cells[i, 5].PutValue("正常上班");
                    try
                    {
                        if (ReasonList[i-1] != null && ReasonList[i-1].Reason1 != null)
                        {
                            sheet.Cells[i, 5].PutValue(ReasonList[i-1].Reason1);
                        }
                        else
                        {
                            sheet.Cells[i, 5].PutValue("正常上班");
                        }
                    }
                    catch (Exception)
                    {
                        sheet.Cells[i, 5].PutValue("正常上班");
                    }
                    //绑定样式
                    sheet.Cells[i, 2].SetStyle(styleTmp1);
                    sheet.Cells[i, 3].SetStyle(styleTmp2);
                    i++;
                }

                wb.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\签卡测试.xls");
            }
            XLDeleteSheet(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\签卡测试.xls", "Evaluation Warning");

            //写回excel中
            //加班
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                Aspose.Cells.Worksheet sheet = wb.Worksheets[0];
                //设置样式
                Aspose.Cells.Style style = wb.CreateStyle();
                style.ForegroundColor = System.Drawing.Color.FromArgb(128, 128, 128);
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Center;
                style.Pattern = BackgroundType.Solid;
                sheet.Name = "明细1";
                //绑定数据
                sheet.Cells[0, 0].PutValue("序号");
                sheet.Cells[0, 1].PutValue("员工姓名");
                sheet.Cells[0, 2].PutValue("加班类型");
                sheet.Cells[0, 3].PutValue("加班事由");
                sheet.Cells[0, 4].PutValue("起始日期");
                sheet.Cells[0, 5].PutValue("加上1");
                sheet.Cells[0, 6].PutValue("加下1");
                sheet.Cells[0, 7].PutValue("加上2");
                sheet.Cells[0, 8].PutValue("加下2");
                sheet.Cells[0, 9].PutValue("加上3");
                sheet.Cells[0, 10].PutValue("加下3");
                sheet.Cells[0, 11].PutValue("加上4");
                sheet.Cells[0, 12].PutValue("加下4");
                sheet.Cells[0, 13].PutValue("关闭同步时间");
                sheet.Cells[0, 14].PutValue("人员类型");
                sheet.Cells[0, 15].PutValue("人员安全级别");
                sheet.Cells[0, 16].PutValue("入职日期");
                //绑定样式
                for (int m = 0; m < 17; m++)
                {
                    sheet.Cells[0, m].SetStyle(style);
                }



                Aspose.Cells.Workbook excelTmp = new Aspose.Cells.Workbook(System.AppDomain.CurrentDomain.BaseDirectory + "测试文件/加班申请单模板.xls");
                Aspose.Cells.Style styleTmp1 = excelTmp.Worksheets[0].Cells[1, 4].GetStyle();
                Aspose.Cells.Style styleTmp2 = excelTmp.Worksheets[0].Cells[1, 5].GetStyle();



                int i = 0;
                int j = 0;
                int nLastDay = 0;
                foreach (var item in overtimeList)
                {
                    if (nLastDay != item.Day)
                    {
                        i++;
                        nLastDay = item.Day;
                        j = 0;

                        sheet.Cells[i, 0].PutValue(i);
                        sheet.Cells[i, 1].PutValue(strNameNO);

                        if (holidayData[item.Day-1])
                        {
                            sheet.Cells[i, 2].PutValue("周六日加班");
                        }
                        else
                        {
                            sheet.Cells[i, 2].PutValue("平时加班");
                        }

                        try
                        {
                            if (ReasonList[i-1] != null && ReasonList[i - 1].Reason2 != null)
                            {
                                sheet.Cells[i, 3].PutValue(ReasonList[i - 1].Reason2);
                            }
                            else
                            {
                                sheet.Cells[i, 3].PutValue("测试");
                            }
                        }
                        catch (Exception)
                        {
                            sheet.Cells[i, 3].PutValue("测试");
                        }
                        
                        sheet.Cells[i, 4].PutValue(item.ToString("yyyy-MM-dd"));
                        //绑定样式
                        sheet.Cells[i, 4].SetStyle(styleTmp1);
                    }

                    sheet.Cells[i, j + 5].SetStyle(styleTmp2);
                    sheet.Cells[i, j + 5].PutValue(item.ToString("HH:mm"));
                    j++;

                }

                wb.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\加班测试.xls");
            }

            XLDeleteSheet(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\加班测试.xls", "Evaluation Warning");
        }

        public bool XLDeleteSheet(string fileName, string sheetToDelete)
        {
            bool returnValue = true;
            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fs);
                workbook.RemoveSheetAt(1);
                using (var fsw = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    workbook.Write(fsw);
                    fsw.Close();
                }
                    
            }
            return returnValue;
        }
        //循环遍历获取excel的中每行每列的值  
        public List<DateTime> GetImportExcelRoute(Aspose.Cells.Workbook excel)
        {
            int icount = excel.Worksheets.Count;

            List<DateTime> routList = new List<DateTime>();
            for (int i = 0; i < 1/*icount*/; i++)
            {
                Aspose.Cells.Worksheet sheet = excel.Worksheets[i];
                Cells cells = sheet.Cells;
                int rowcount = cells.MaxRow;//行数需要+1
                int columncount = cells.MaxColumn;

                //获取标题所在的列
                for (int r  = 1; r <= rowcount; r++)
                {
                    //string[] szstrRow = new string[columncount + 1];
                    //for (int c = 0; c <= columncount; c++)
                    //{
                    //    string strVal = cells[r, c].StringValue.Trim();
                    //    Debug.Write(strVal);
                    //    Debug.Write("\t");
                    //    szstrRow[c] = strVal;
                    //}
                    //Debug.WriteLine("");
                    //DateTime dtt = cells[r, 2].DateTimeValue;
                    //DateTime dtd = cells[r, 1].DateTimeValue;
                    //dtd = dtd.AddHours(dtt.Hour);
                    // dtd = dtd.AddMinutes(dtt.Minute);
                    //dtd = dtd.AddSeconds(dtt.Second);

                    DateTime dtt = Convert.ToDateTime(cells[r, 2].StringValue);
                    DateTime selDate = DateTime.Now.AddMonths(0);
                    if (dtt.Month == dateTimePicker1.Value.Month)
                    {
                        routList.Add(dtt);
                    }
                }
                //break;//只读一个worksheet
            }
            return routList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ReasonList
            string json = JsonConvert.SerializeObject(ReasonList);
            //BindingList<Reason> m = JsonConvert.DeserializeObject<BindingList<Reason>>(json);
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory+"reason.json"))
            {
                file.Write(json);
            }

        }

        private void buttonClipBoard_Click(object sender, EventArgs e)
        {
            AnalyzeExcel(null);
        }
    }
}
