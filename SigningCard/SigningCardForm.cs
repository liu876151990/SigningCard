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
using System.Net.Http;
using System.Net;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using SigningCard.Properties;
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
        int adjustedFirstDay = 0;//调整为星期一为第一列后的索引
        bool[] holidayData = new bool[31];//最多31天
        string execlPath = null;
        string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        class Reason
        {
            public string Reason1 { get; set; }
            public string Reason2 { get; set; }
        }
        BindingList<Reason> ReasonList = new BindingList<Reason>();
        ToolTip toolTipOa;

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
            // 获取本月第一天是星期几（Sunday=0, Monday=1, ..., Saturday=6）
            DateTime firstDay = DateTime.Parse(string.Format("{0:D}/{1:D}/{2:D}", year, month, 1));
            weekDayFirstDay = (int)firstDay.DayOfWeek;
            
            // 转换为星期一为第一列的索引（Monday=0, Tuesday=1, ..., Sunday=6）
            adjustedFirstDay = weekDayFirstDay == 0 ? 6 : weekDayFirstDay - 1;

            for (int i = 0; i < totalDays; i++)
            {
                int row = (i + adjustedFirstDay) / dataViewColums;
                int colum = (i + adjustedFirstDay) % dataViewColums;
                dataGridViewHoliday.Rows[row].Cells[colum].Value = i + 1;
                
                // 判断当前日期是星期几
                DateTime currentDate = new DateTime(year, month, i + 1);
                DayOfWeek dayOfWeek = currentDate.DayOfWeek;
                
                // 星期日（Sunday=0）或星期六（Saturday=6）为节假日
                if (dayOfWeek == DayOfWeek.Sunday || dayOfWeek == DayOfWeek.Saturday)
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

            toolTipOa = new ToolTip();
            toolTipOa.SetToolTip(buttonOaAutoImport, "单击：导入考勤；Shift+单击：设置 OA 账号并保存到本机");
        }

        private void DataGridViewHoliday_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void DataGridViewHoliday_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int index = e.RowIndex * dataViewColums + e.ColumnIndex - adjustedFirstDay;
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
             
            AnalyzeImportedDateList(importDateList, strNameNO);
        }

        private void AnalyzeImportedDateList(List<DateTime> importDateList, string strNameNO)
        {
            if (importDateList == null || importDateList.Count <= 0)
            {
                MessageBox.Show("没有可导入的考勤打卡记录。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            MessageBox.Show("已生成签卡与加班导出文件。", "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void ShowOaAccountSettings()
        {
            string loginId;
            string password;
            TryLoadOaCredentials(out loginId, out password);
            bool remember;
            if (!ShowOaCredentialDialog("OA 账号设置", out loginId, out password, out remember, true, loginId ?? "", password ?? ""))
            {
                return;
            }
            if (remember && !string.IsNullOrWhiteSpace(loginId))
            {
                SaveOaCredentials(loginId, password ?? "");
                MessageBox.Show("已保存 OA 账号（密码已用本机 Windows 用户级加密存储）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                ClearOaCredentials();
                MessageBox.Show("已清除本机保存的 OA 登录信息。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void buttonOaAutoImport_Click(object sender, EventArgs e)
        {
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                ShowOaAccountSettings();
                return;
            }

            string loginId;
            string password;
            bool usedSaved = TryLoadOaCredentials(out loginId, out password);
            if (!usedSaved)
            {
                bool remember;
                if (!ShowOaCredentialDialog("OA 自动导入 - 登录", out loginId, out password, out remember, false))
                {
                    return;
                }
                if (string.IsNullOrWhiteSpace(loginId) || password == null)
                {
                    return;
                }
                if (remember)
                {
                    SaveOaCredentials(loginId, password);
                }
            }

            buttonOaAutoImport.Enabled = false;
            buttonOaAutoImport.Text = "导入中...";
            try
            {
                List<DateTime> dateTimes = await FetchAttendanceFromOaAsync(loginId, password, dateTimePicker1.Value.Year, dateTimePicker1.Value.Month);
                string nameNo = loginId;
                AnalyzeImportedDateList(dateTimes, nameNo);
            }
            catch (Exception ex)
            {
                if (usedSaved)
                {
                    ClearOaCredentials();
                    MessageBox.Show("已保存的登录已失效，已清除本地记录。\r\n请 Shift+单击「OA自动导入」重新设置账号，或再次导入时输入账号密码。\r\n\r\n" + ex.Message, "OA 登录失效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("OA自动导入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                buttonOaAutoImport.Enabled = true;
                buttonOaAutoImport.Text = "OA自动导入";
            }
        }

        private static bool TryLoadOaCredentials(out string loginId, out string password)
        {
            loginId = "";
            password = "";
            try
            {
                if (string.IsNullOrWhiteSpace(Settings.Default.OaLoginId) || string.IsNullOrWhiteSpace(Settings.Default.OaPasswordProtected))
                {
                    return false;
                }
                loginId = Settings.Default.OaLoginId.Trim();
                byte[] cipher = Convert.FromBase64String(Settings.Default.OaPasswordProtected);
                byte[] plain = ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser);
                password = Encoding.UTF8.GetString(plain);
                return !string.IsNullOrEmpty(loginId);
            }
            catch
            {
                return false;
            }
        }

        private static void SaveOaCredentials(string loginId, string password)
        {
            Settings.Default.OaLoginId = (loginId ?? "").Trim();
            byte[] plain = Encoding.UTF8.GetBytes(password ?? "");
            byte[] cipher = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
            Settings.Default.OaPasswordProtected = Convert.ToBase64String(cipher);
            Settings.Default.Save();
        }

        private static void ClearOaCredentials()
        {
            Settings.Default.OaLoginId = "";
            Settings.Default.OaPasswordProtected = "";
            Settings.Default.Save();
        }

        private bool ShowOaCredentialDialog(string title, out string loginId, out string password, out bool remember, bool settingsMode, string initialLogin = "", string initialPassword = "")
        {
            loginId = "";
            password = "";
            remember = false;

            using (var form = new Form())
            using (var lblUser = new Label())
            using (var txtUser = new TextBox())
            using (var lblPwd = new Label())
            using (var txtPwd = new TextBox())
            using (var chkRemember = new CheckBox())
            using (var btnOk = new Button())
            using (var btnCancel = new Button())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.Width = 400;
                form.Height = 220;

                lblUser.Text = "账号：";
                lblUser.Left = 12;
                lblUser.Top = 14;
                lblUser.AutoSize = true;

                txtUser.Left = 72;
                txtUser.Top = 10;
                txtUser.Width = 300;
                txtUser.Text = initialLogin;

                lblPwd.Text = "密码：";
                lblPwd.Left = 12;
                lblPwd.Top = 44;
                lblPwd.AutoSize = true;

                txtPwd.Left = 72;
                txtPwd.Top = 40;
                txtPwd.Width = 300;
                txtPwd.UseSystemPasswordChar = true;
                txtPwd.Text = initialPassword;

                chkRemember.Text = "记住账号密码（本机加密保存，下次无需输入）";
                chkRemember.Left = 12;
                chkRemember.Top = 76;
                chkRemember.Width = 360;
                chkRemember.Checked = true;

                btnOk.Text = "确定";
                btnOk.Left = 216;
                btnOk.Top = 110;
                btnOk.Width = 75;
                btnOk.DialogResult = DialogResult.OK;

                btnCancel.Text = "取消";
                btnCancel.Left = 297;
                btnCancel.Top = 110;
                btnCancel.Width = 75;
                btnCancel.DialogResult = DialogResult.Cancel;

                form.Controls.Add(lblUser);
                form.Controls.Add(txtUser);
                form.Controls.Add(lblPwd);
                form.Controls.Add(txtPwd);
                form.Controls.Add(chkRemember);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog(this) != DialogResult.OK)
                {
                    return false;
                }
                loginId = txtUser.Text.Trim();
                password = txtPwd.Text;
                remember = chkRemember.Checked;
                if (string.IsNullOrEmpty(loginId))
                {
                    // 账号设置里取消「记住」时，可不填账号，仅用于清除本机已保存的凭据
                    if (!(settingsMode && !remember))
                    {
                        MessageBox.Show("请输入账号。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                }
                if (settingsMode)
                {
                    if (remember && string.IsNullOrEmpty(password))
                    {
                        MessageBox.Show("勾选「记住账号密码」时，密码不能为空。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(password))
                    {
                        MessageBox.Show("请输入密码。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                }
                return true;
            }
        }

        private async System.Threading.Tasks.Task<List<DateTime>> FetchAttendanceFromOaAsync(string loginId, string password, int year, int month)
        {
            var cookieContainer = new CookieContainer();
            var handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            using (var client = new HttpClient(handler))
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                client.DefaultRequestHeaders.Referrer = new Uri("https://oa.hanslaser.com/wui/index.html");
                client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

                var rsaInfo = await GetRsaInfoAsync(client);
                string encLogin = RsaEncryptForOa(loginId, rsaInfo.PublicKey, rsaInfo.RsaCode, rsaInfo.RsaFlag);
                string encPwd = RsaEncryptForOa(password, rsaInfo.PublicKey, rsaInfo.RsaCode, rsaInfo.RsaFlag);

                var loginPayload = new Dictionary<string, string>
                {
                    {"islanguid", "7"},
                    {"loginid", encLogin},
                    {"userpassword", encPwd},
                    {"dynamicPassword", ""},
                    {"tokenAuthKey", ""},
                    {"validatecode", ""},
                    {"validateCodeKey", ""},
                    {"logintype", "1"},
                    {"messages", ""},
                    {"isie", "false"},
                    {"appid", ""},
                    {"service", ""},
                    {"isRememberPassword", "true"}
                };
                var loginResp = await client.PostAsync("https://oa.hanslaser.com/api/hrm/login/checkLogin", new FormUrlEncodedContent(loginPayload));
                loginResp.EnsureSuccessStatusCode();
                string loginText = await loginResp.Content.ReadAsStringAsync();
                dynamic loginObj = JsonConvert.DeserializeObject(loginText);
                string msgcode = loginObj?.msgcode?.ToString();
                string loginstatus = loginObj?.loginstatus?.ToString();
                if (msgcode != "0" && loginstatus != "true")
                {
                    throw new Exception((string)(loginObj?.msg ?? "登录失败"));
                }

                // 一次请求即可返回当月考勤（与 HAR 中接口一致，date 传当月任意一天即可；这里用当月 1 日）
                string dateParam = string.Format("{0}-{1}-{2}", year, month, 1);
                string url = "https://oa.hanslaser.com/api/hans/Administrative/getinfo?date=" + Uri.EscapeDataString(dateParam) + "&workcode=&_=" + DateTimeOffset.Now.ToUnixTimeMilliseconds();
                string text = await client.GetStringAsync(url);
                var result = ParseHansAdministrativeResponse(text, year, month);
                result.Sort();
                return result;
            }
        }

        private List<DateTime> ParseHansAdministrativeResponse(string text, int year, int month)
        {
            var list = new List<DateTime>();
            if (string.IsNullOrWhiteSpace(text))
            {
                return list;
            }

            dynamic arr = JsonConvert.DeserializeObject(text);
            if (arr == null)
            {
                return list;
            }

            foreach (var dayObj in arr)
            {
                var am = dayObj?.AM;
                if (am == null)
                {
                    continue;
                }
                foreach (var item in am)
                {
                    DateTime dt;
                    if (DateTime.TryParse(item.ToString(), out dt) && dt.Year == year && dt.Month == month)
                    {
                        list.Add(dt);
                    }
                }
            }
            return list;
        }

        private class RsaInfo
        {
            public string PublicKey { get; set; }
            public string RsaCode { get; set; }
            public string RsaFlag { get; set; }
        }

        private async System.Threading.Tasks.Task<RsaInfo> GetRsaInfoAsync(HttpClient client)
        {
            string url = "https://oa.hanslaser.com/rsa/weaver.rsa.GetRsaInfo?ts=" + DateTimeOffset.Now.ToUnixTimeMilliseconds();
            string text = await client.GetStringAsync(url);
            dynamic obj = JsonConvert.DeserializeObject(text);
            return new RsaInfo
            {
                PublicKey = obj?.rsa_pub?.ToString(),
                RsaCode = obj?.rsa_code?.ToString() ?? "",
                RsaFlag = obj?.rsa_flag?.ToString() ?? "``RSA``"
            };
        }

        /// <summary>
        /// 与前端 rsa.js 一致：明文分段(240) + rsa_code 后 RSA 加密，每段密文后接 rsa_flag；不做 URL 编码（由 FormUrlEncodedContent 处理）。
        /// </summary>
        private string RsaEncryptForOa(string value, string base64PublicKey, string rsaCode, string rsaFlag)
        {
            const int groupLength = 240;
            var blocks = new List<string>();
            for (int i = 0; i < value.Length; i += groupLength)
            {
                int len = Math.Min(groupLength, value.Length - i);
                blocks.Add(value.Substring(i, len));
            }
            if (blocks.Count == 0)
            {
                blocks.Add("");
            }

            using (var rsa = DecodeX509PublicKey(Convert.FromBase64String(base64PublicKey)))
            {
                var sb = new StringBuilder();
                foreach (var block in blocks)
                {
                    byte[] plain = Encoding.UTF8.GetBytes(block + rsaCode);
                    byte[] encrypted = rsa.Encrypt(plain, false);
                    sb.Append(Convert.ToBase64String(encrypted));
                    sb.Append(rsaFlag);
                }
                return sb.ToString();
            }
        }

        private RSACryptoServiceProvider DecodeX509PublicKey(byte[] x509Key)
        {
            using (BinaryReader reader = new BinaryReader(new MemoryStream(x509Key)))
            {
                ushort twobytes = reader.ReadUInt16();
                if (twobytes == 0x8130) reader.ReadByte();
                else if (twobytes == 0x8230) reader.ReadInt16();
                else throw new Exception("无效RSA公钥格式");

                byte[] seq = reader.ReadBytes(15);
                byte[] seqOid = new byte[] { 0x30, 0x0D, 0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x01, 0x01, 0x05, 0x00 };
                for (int i = 0; i < seqOid.Length; i++)
                {
                    if (seq[i] != seqOid[i]) throw new Exception("无效RSA公钥OID");
                }

                twobytes = reader.ReadUInt16();
                if (twobytes == 0x8103) reader.ReadByte();
                else if (twobytes == 0x8203) reader.ReadInt16();
                else throw new Exception("无效RSA公钥BIT STRING");

                byte bt = reader.ReadByte();
                if (bt != 0x00) throw new Exception("无效RSA公钥填充");

                twobytes = reader.ReadUInt16();
                if (twobytes == 0x8130) reader.ReadByte();
                else if (twobytes == 0x8230) reader.ReadInt16();
                else throw new Exception("无效RSA公钥SEQUENCE");

                twobytes = reader.ReadUInt16();
                int modsize;
                if (twobytes == 0x8102) modsize = reader.ReadByte();
                else if (twobytes == 0x8202)
                {
                    byte high = reader.ReadByte();
                    byte low = reader.ReadByte();
                    modsize = BitConverter.ToUInt16(new byte[] { low, high }, 0);
                }
                else
                {
                    throw new Exception("无效RSA公钥模数");
                }

                byte firstModByte = reader.ReadByte();
                reader.BaseStream.Seek(-1, SeekOrigin.Current);
                if (firstModByte == 0x00) 
                {
                    reader.ReadByte();
                    modsize -= 1;
                }
                byte[] modulus = reader.ReadBytes(modsize);

                if (reader.ReadByte() != 0x02) throw new Exception("无效RSA公钥指数");
                int expbytes = reader.ReadByte();
                byte[] exponent = reader.ReadBytes(expbytes);

                var rsa = new RSACryptoServiceProvider();
                rsa.ImportParameters(new RSAParameters { Modulus = modulus, Exponent = exponent });
                return rsa;
            }
        }

        private string PromptInput(string prompt, string title, string defaultValue = "", bool isPassword = false)
        {
            using (var form = new Form())
            using (var label = new Label())
            using (var textBox = new TextBox())
            using (var buttonOk = new Button())
            using (var buttonCancel = new Button())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.Width = 420;
                form.Height = 170;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                label.Left = 12;
                label.Top = 16;
                label.Width = 380;
                label.Text = prompt;

                textBox.Left = 12;
                textBox.Top = 44;
                textBox.Width = 380;
                textBox.Text = defaultValue;
                textBox.UseSystemPasswordChar = isPassword;

                buttonOk.Text = "确定";
                buttonOk.Left = 236;
                buttonOk.Top = 82;
                buttonOk.Width = 75;
                buttonOk.DialogResult = DialogResult.OK;

                buttonCancel.Text = "取消";
                buttonCancel.Left = 317;
                buttonCancel.Top = 82;
                buttonCancel.Width = 75;
                buttonCancel.DialogResult = DialogResult.Cancel;

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(buttonOk);
                form.Controls.Add(buttonCancel);
                form.AcceptButton = buttonOk;
                form.CancelButton = buttonCancel;

                return form.ShowDialog(this) == DialogResult.OK ? textBox.Text.Trim() : "";
            }
        }

        private async void ButtonGetHolidays_Click(object sender, EventArgs e)
        {
            int year = dateTimePicker1.Value.Year;
            buttonGetHolidays.Enabled = false;
            buttonGetHolidays.Text = "获取中...";

            try
            {
                await GetHolidaysFromWeb(year);
                MessageBox.Show($"已成功获取 {year} 年节假日数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取节假日失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                buttonGetHolidays.Enabled = true;
                buttonGetHolidays.Text = "获取节假日";
            }
        }

        private async System.Threading.Tasks.Task GetHolidaysFromWeb(int year)
        {
            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(10);
                string url = $"http://timor.tech/api/holiday/year/{year}";

                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string json = await response.Content.ReadAsStringAsync();
                dynamic result = JsonConvert.DeserializeObject(json);

                // 首先重置当前月份的所有节假日数据
                for (int i = 0; i < totalDays; i++)
                {
                    // 重置为默认状态：周六周日是节假日，其他不是
                    int dayOfMonth = i + 1;
                    DateTime currentDate = new DateTime(year, dateTimePicker1.Value.Month, dayOfMonth);
                    DayOfWeek dayOfWeek = currentDate.DayOfWeek;
                    
                    if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
                    {
                        holidayData[i] = true;
                    }
                    else
                    {
                        holidayData[i] = false;
                    }
                    
                    // 更新日历显示颜色
                    int row = (i + adjustedFirstDay) / dataViewColums;
                    int col = (i + adjustedFirstDay) % dataViewColums;
                    
                    if (row < dataGridViewHoliday.RowCount && col < dataGridViewHoliday.ColumnCount)
                    {
                        dataGridViewHoliday.Rows[row].Cells[col].Style.BackColor = holidayData[i] ? System.Drawing.Color.Green : System.Drawing.Color.White;
                    }
                }

                // 然后根据API数据更新节假日状态
                if (result.holiday != null)
                {
                    // holiday 是一个字典，键是日期字符串（如 "01-01"），值是节假日信息
                    foreach (var kvp in result.holiday)
                    {
                        // kvp 是一个键值对，需要获取其 Value
                        var holidayInfo = kvp.Value;
                        string dateStr = holidayInfo.date.ToString();
                        
                        if (!string.IsNullOrEmpty(dateStr))
                        {
                            DateTime holidayDate = DateTime.Parse(dateStr);
                            if (holidayDate.Year == year && holidayDate.Month == dateTimePicker1.Value.Month)
                            {
                                int dayIndex = holidayDate.Day - 1;
                                if (dayIndex >= 0 && dayIndex < totalDays)
                                {
                                    int row = (dayIndex + adjustedFirstDay) / dataViewColums;
                                    int col = (dayIndex + adjustedFirstDay) % dataViewColums;
                                    
                                    if (row < dataGridViewHoliday.RowCount && col < dataGridViewHoliday.ColumnCount)
                                    {
                                        // 检查是否是调休日（holiday: false）
                                        bool isWorkDay = holidayInfo.holiday != null && holidayInfo.holiday.ToString() == "False";
                                        
                                        if (isWorkDay)
                                        {
                                            // 调休工作日，标记为非节假日（白色）
                                            holidayData[dayIndex] = false;
                                            dataGridViewHoliday.Rows[row].Cells[col].Style.BackColor = System.Drawing.Color.White;
                                        }
                                        else
                                        {
                                            // 正常节假日，标记为绿色
                                            holidayData[dayIndex] = true;
                                            dataGridViewHoliday.Rows[row].Cells[col].Style.BackColor = System.Drawing.Color.Green;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
