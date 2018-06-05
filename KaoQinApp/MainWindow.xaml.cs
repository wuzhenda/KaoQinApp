using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KaoQinApp
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".dat";
            dlg.Filter = "Text documents (.dat)|*.dat";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                FileNameTextBox.Text = filename;

                try
                {
                    var records = ReadRecorder(filename);
                    string tablename = "2018六月考勤";
                    if (records.Count > 0)
                    {
                        var ondate = records[0].OnDate;
                        var year = ondate.Year;
                        var month = ondate.Month;
                        tablename = "" + year + "" + Constant.MonthDictionary[month] + "月考勤";
                    }

                    ExportExcel(tablename,records);

                }
                catch(Exception e2)
                {
                    MessageBox.Show(""+e2.Message);
                }

            }
        }

        public static List<Record> ReadRecorder(string filePath)
        {
            List<Record> listRecords = new List<Record>();

            System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open,
                System.IO.FileAccess.Read);

            System.IO.StreamReader sr = new System.IO.StreamReader(fs, Encoding.ASCII);
            
            //记录每次读取的一行记录  
            string strLine = "";

            //逐行读取CSV中的数据  
            while ((strLine = sr.ReadLine()) != null)
            {                
                
                var aryLine = strLine.Split('\t');

                var record = new Record();
                record.Bianhao = aryLine[0].Trim(' ');
                record.OnDate = Convert.ToDateTime(aryLine[1]);
                record.One1 = aryLine[2];
                record.Zero1 = aryLine[3];
                record.One2 = aryLine[4];
                record.Zero2 = aryLine[5];

                listRecords.Add(record);
            }


            sr.Close();
            fs.Close();
            return listRecords;
        }

        public static void CheckIfValid(List<Record> listRecords)
        {

            
        }

        bool isLateOrEarlyOffWork(DateTime dt)
        {
            //判断当前时间是否在工作时间段内
            string _strWorkingDayAM = "08:30";//工作时间上午08:30
            string _strWorkingDayPM = "17:30";
            TimeSpan dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            TimeSpan dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;
            
            TimeSpan dspNow = dt.TimeOfDay;
            if (dspNow > dspWorkingDayAM && dspNow < dspWorkingDayPM)
            {
                return true;
            }

            return false;
        }

        public void ExportExcel(string tableName, List<Record> listRecords)
        {
            try
            {
                //创建一个工作簿
                IWorkbook workbook = new HSSFWorkbook();

                //创建一个 sheet 表
                ISheet sheet = workbook.CreateSheet(tableName);

                //创建一行
                IRow rowH = sheet.CreateRow(0);

                //创建一个单元格
                ICell cell = null;

                //创建单元格样式
                ICellStyle cellStyle = workbook.CreateCellStyle();

                //创建格式
                IDataFormat dataFormat = workbook.CreateDataFormat();

                //设置为文本格式，也可以为 text，即 dataFormat.GetFormat("text");
                cellStyle.DataFormat = dataFormat.GetFormat("@");

                //设置列名
                //foreach (DataColumn col in dt.Columns)
                //{
                //    //创建单元格并设置单元格内容
                //    rowH.CreateCell(col.Ordinal).SetCellValue(col.Caption);

                //    //设置单元格格式
                //    rowH.Cells[col.Ordinal].CellStyle = cellStyle;
                //}

                //设置第一行列名
                // (DataColumn col in dt.Columns)
                {
                    int start = 0;
                    //创建单元格并设置单元格内容
                    rowH.CreateCell(start).SetCellValue("姓名");
                    //设置单元格格式
                    rowH.Cells[start].CellStyle = cellStyle;

                    rowH.CreateCell(start+1).SetCellValue("工号");
                    //设置单元格格式
                    rowH.Cells[start+1].CellStyle = cellStyle;

                    rowH.CreateCell(start+2).SetCellValue("刷卡记录");
                    //设置单元格格式
                    rowH.Cells[start+2].CellStyle = cellStyle;

                    rowH.CreateCell(start+3).SetCellValue("备注");
                    //设置单元格格式
                    rowH.Cells[start+3].CellStyle = cellStyle;
                }

                //{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
                int lastDay = -1;
                var OndayDic = new Dictionary<string, int>{
                                        {"1003920",0 },
                                        {"1004984",0 },
                                        {"1004699",0 },
                                        {"1006056",0 },
                                        {"1006177",0 },
                                    };
                List<ExcelItem> listExcelItems = new List<ExcelItem>();
                for (int i = 0; i < listRecords.Count; i++)
                {
                    var record = listRecords[i];
                    var excelItem = new ExcelItem();
                    excelItem.Bianhao = record.Bianhao;
                    excelItem.Name = Constant.NameDictionary[record.Bianhao];
                    excelItem.RecordDate = record.OnDate.ToString();
                    if (isLateOrEarlyOffWork(record.OnDate))
                    {
                        excelItem.Note = "迟到或早退";
                    }


                    if (-1 == lastDay)
                    {
                        lastDay = record.OnDate.Day;
                    }
                    else if (lastDay != record.OnDate.Day)
                    {
                        //统计当前缺勤
                        foreach (var item in OndayDic)
                        {
                            if (item.Value == 0)
                            {
                                ExcelItem absent = new ExcelItem();
                                absent.Bianhao = item.Key;
                                absent.Name= Constant.NameDictionary[item.Key];
                                absent.Note = "缺席或请假";
                                listExcelItems.Add(absent);
                            }
                            else if (item.Value < 2)
                            {
                                ExcelItem absent = new ExcelItem();
                                absent.Bianhao = item.Key;
                                absent.Name = Constant.NameDictionary[item.Key];
                                absent.Note = "漏打卡一次";
                                listExcelItems.Add(absent);
                            }
                        }

                        //隔天,插入空记录
                        ExcelItem tmp = new ExcelItem();
                        listExcelItems.Add(tmp);

                        //更新天
                        lastDay = record.OnDate.Day;
                        //新的一天设置初始状态
                        var keys = new List<String>();
                        foreach (var itemkey in OndayDic.Keys)
                        {
                            keys.Add(itemkey);
                        }
                        foreach (var itemkey in keys)
                        {
                            OndayDic[itemkey] = 0;
                        }
                    }
                    else
                    {//同一天,非第一次

                    }

                    OndayDic[record.Bianhao]++;

                    listExcelItems.Add(excelItem);
                }
                //最后一天记录统计
                //统计当前缺勤
                foreach (var item in OndayDic)
                {
                    if (item.Value == 0)
                    {
                        ExcelItem absent = new ExcelItem();
                        absent.Bianhao = item.Key;
                        absent.Name = Constant.NameDictionary[item.Key];
                        absent.Note = "缺席或请假";
                        listExcelItems.Add(absent);
                    }
                    else if (item.Value <2)
                    {
                        ExcelItem absent = new ExcelItem();
                        absent.Bianhao = item.Key;
                        absent.Name = Constant.NameDictionary[item.Key];
                        absent.Note = "漏打卡一次";
                        listExcelItems.Add(absent);
                    }
                }

                //}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}

                //写入数据
                for (int i = 0; i < listExcelItems.Count; i++)
                {
                    //跳过第一行，第一行为列名
                    IRow row = sheet.CreateRow(i + 1);

                    var excelItem = listExcelItems[i];

                    //for (int j = 0; j < 4; j++)
                    {
                        int j = 0;
                        cell = row.CreateCell(j);
                        cell.SetCellValue(excelItem.Name);
                        cell = row.CreateCell(j+1);
                        cell.SetCellValue(excelItem.Bianhao);
                        cell = row.CreateCell(j+2);
                        cell.SetCellValue(excelItem.RecordDate);
                        cell = row.CreateCell(j+3);
                        cell.SetCellValue(excelItem.Note);
                        if (string.IsNullOrEmpty(excelItem.Note))
                        {
                            
                            //cell.CellStyle = cellStyle;
                            //cellStyle.DataFormat = dataFormat.GetFormat("@");
                        }
                    }
                }

                //设置导出文件路径
                string path = "D:\\"+tableName+"-智能部件苏州分部";

                //设置新建文件路径及名称
                string savePath = path + ".xls";

                //创建文件
                FileStream file = new FileStream(savePath, FileMode.CreateNew, System.IO.FileAccess.Write);

                //创建一个 IO 流
                MemoryStream ms = new MemoryStream();

                //写入到流
                workbook.Write(ms);

                //转换为字节数组
                byte[] bytes = ms.ToArray();

                file.Write(bytes, 0, bytes.Length);
                file.Flush();

                //还可以调用下面的方法，把流输出到浏览器下载
                //OutputClient(bytes);

                //释放资源
                bytes = null;

                ms.Close();
                ms.Dispose();

                file.Close();
                file.Dispose();

                workbook.Close();
                sheet = null;
                workbook = null;

                MessageBox.Show("Done OK,file at D:\\");
            }
            catch (Exception ex)
            {
                MessageBox.Show(""+ex);
            }
        }

    }
}
