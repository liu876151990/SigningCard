﻿//using DocumentFormat.OpenXml.Spreadsheet;
//using Moq;
using System.Diagnostics;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                    dataGridViewHoliday.Rows[j].Cells[i].Style.BackColor = Color.White;
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
                    dataGridViewHoliday.Rows[row].Cells[colum].Style.BackColor = Color.Green;
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
                dataGridViewHoliday.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
                holidayData[index] = false;
            }
            else
            {
                dataGridViewHoliday.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Green;
                holidayData[index] = true;
            }
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            SelectMonth(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month);
        }

        private void AnalyzeExcel(string fileName)
        {
            string excelName = fileName;
            Workbook excel = new Workbook(excelName);
            List<DateTime> importDateList = GetImportExcelRoute(excel);

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
                    DateTime dtAMU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 8, 35, 0);

                    DateTime dtAMD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 11, 55, 0);
                    DateTime dtAMD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 12, 45, 0);

                    DateTime dtPMU1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 12, 46, 0);
                    DateTime dtPMU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 13, 35, 0);

                    DateTime dtPMD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 17, 55, 0);
                    DateTime dtPMD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 15, 0);

                    //加班有效区间
                    DateTime dtOTU1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 18, 16, 0);
                    DateTime dtOTU2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 19, 30, 0);

                    DateTime dtOTD1 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 19, 31, 0);
                    DateTime dtOTD2 = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, i + 1, 23, 59, 0);

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
                    }

                    //每天需要的签卡 有加班缺卡则另外添加
                    for (int k = 0; k < 4; k++)
                    {
                        if (false == szbSingingCardSts[k])
                        {
                            singingCardList.Add(dtModel[k]);
                        }
                    }
                    if (true == szbSingingCardSts[4] && false == szbSingingCardSts[5])
                    {
                        singingCardList.Add(dtModel[5]);
                        overtimeList.Add(dtModel[5]);
                    }
                    else if (false == szbSingingCardSts[4] && true == szbSingingCardSts[5])
                    {
                        singingCardList.Add(dtModel[4]);
                        overtimeList.Add(dtModel[4]);
                    }
                }
            }
            singingCardList.Sort();
            overtimeList.Sort();

            //写回excel中
            //签卡
            {
                Workbook wb = new Workbook(System.IO.Stream.Null);
                Worksheet sheet = wb.Worksheets[0];
                //设置样式
                Style style = wb.CreateStyle();
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
                

                Workbook excelTmp = new Workbook(System.AppDomain.CurrentDomain.BaseDirectory + "测试文件/签卡申请流程模板.xls");
                Style styleTmp1 = excelTmp.Worksheets[0].Cells[1,2].GetStyle();
                Style styleTmp2 = excelTmp.Worksheets[0].Cells[1,3].GetStyle();

                string strNameNO = excel.Worksheets[0].Cells[1,1].StringValue + "_" + excel.Worksheets[0].Cells[1, 0].StringValue;

                int i = 1;
                foreach (var item in singingCardList)
                {
                    sheet.Cells[i, 0].PutValue(i);
                    sheet.Cells[i, 1].PutValue(strNameNO);
                    
                    sheet.Cells[i, 2].PutValue(item.ToShortDateString().Replace('/', '-'));
                    sheet.Cells[i, 3].PutValue(item.ToShortTimeString());
                    sheet.Cells[i, 4].PutValue("正常签卡");
                    sheet.Cells[i, 5].PutValue("正常上班");
                    //绑定样式
                    sheet.Cells[i, 2].SetStyle(styleTmp1);
                    sheet.Cells[i, 3].SetStyle(styleTmp2);
                    i++;
                }

                wb.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\签卡测试.xls");
            }

            //写回excel中
            //加班
            {
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                //设置样式
                Style style = wb.CreateStyle();
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



                Workbook excelTmp = new Workbook(System.AppDomain.CurrentDomain.BaseDirectory + "测试文件/加班申请单模板.xls");
                Style styleTmp1 = excelTmp.Worksheets[0].Cells[1, 4].GetStyle();
                Style styleTmp2 = excelTmp.Worksheets[0].Cells[1, 5].GetStyle();



                int i = 0;
                int j = 0;
                int nLastDay = 0;
                string strNameNO = excel.Worksheets[0].Cells[1, 1].StringValue + "_" + excel.Worksheets[0].Cells[1, 0].StringValue;
                foreach (var item in overtimeList)
                {
                    if (nLastDay != item.Day)
                    {
                        i++;
                        nLastDay = item.Day;
                        j = 0;

                        sheet.Cells[i, 0].PutValue(i);
                        sheet.Cells[i, 1].PutValue(strNameNO);

                        sheet.Cells[i, 2].PutValue("平时加班");
                        sheet.Cells[i, 3].PutValue("测试");
                        sheet.Cells[i, 4].PutValue(item.ToShortDateString().Replace('/','-'));
                        //绑定样式
                        sheet.Cells[i, 4].SetStyle(styleTmp1);
                    }

                    sheet.Cells[i, j + 5].SetStyle(styleTmp2);
                    sheet.Cells[i, j + 5].PutValue(item.ToShortTimeString());
                    j++;

                }

                wb.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"输出文件\加班测试.xls");
            }
        }

        //循环遍历获取excel的中每行每列的值  
        public List<DateTime> GetImportExcelRoute(Workbook excel)
        {
            int icount = excel.Worksheets.Count;

            List<DateTime> routList = new List<DateTime>();
            for (int i = 0; i < 1/*icount*/; i++)
            {
                Worksheet sheet = excel.Worksheets[i];
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

    }
}