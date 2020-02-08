using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, List<string>> dic = GetSheetData("E:/work/ConsoleApp1/ConsoleApp1/File/TMS.xlsx");
            WriteTxt(dic);
        }

        /// <summary>
        /// 根据路径读取excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetSheetData(string filePath)
        {
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                //2007版本
                var xssfworkbook = new XSSFWorkbook(file);
                int count = xssfworkbook.NumberOfSheets;
                for (int i = 0; i < count; i++)
                {
                    ISheet sheet = xssfworkbook.GetSheetAt(i);
                    for (int j = 1; j < 400; j++)
                    {
                        XSSFRow cRow = (XSSFRow)sheet.GetRow(j);
                        if (cRow == null)
                        {
                            continue;
                        }
                        ICell rowFirstCell = cRow.GetCell(0);
                        if (rowFirstCell == null)
                        {
                            continue;
                        }
                        
                        if (rowFirstCell.IsMergedCell)
                        {
                            for (int ii = 0; ii < sheet.NumMergedRegions; ii++)
                            {
                                var cellrange = sheet.GetMergedRegion(ii);
                                if (rowFirstCell.ColumnIndex >= cellrange.FirstColumn && rowFirstCell.ColumnIndex <= cellrange.LastColumn
                                    && rowFirstCell.RowIndex >= cellrange.FirstRow && rowFirstCell.RowIndex <= cellrange.LastRow)
                                {
                                    XSSFRow firstRow = (XSSFRow)sheet.GetRow(cellrange.FirstRow);
                                    string tableName = Convert.ToString(firstRow.GetCell(0));
                                    string key = dic.Keys.FirstOrDefault(a => a == tableName);
                                    if (string.IsNullOrWhiteSpace(key))
                                    {
                                        List<string> list = new List<string>();
                                        list.Add("sys_serialnumber");
                                        list.Add("sys_data_date");
                                        list.Add("sys_data_status");
                                        list.Add(cRow.GetCell(2).ToString());
                                        dic.Add(tableName, list);
                                    }
                                    else
                                    {
                                        dic[key].Add(cRow.GetCell(2).ToString());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return dic;
        }

        public static void WriteTxt(Dictionary<string, List<string>> dic) {
            
            foreach (var item in dic.Keys)
            {
                FileStream fs = new FileStream("E:/txt/" + item + "_L1.txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("DROP TABLE IF EXISTS "+item+"_L1;");
                sw.WriteLine("CREATE EXTERNAL TABLE " + item + "_L1");
                sw.Write("(");
                for (int i = 0; i < dic[item].Count; i++)
                {
                    if (i == dic[item].Count - 1)
                    {
                        sw.WriteLine(dic[item][i] + " string)");
                    }
                    else
                    {
                        sw.WriteLine(dic[item][i] + " string,");
                    }
                }
                sw.WriteLine("ROW FORMAT DELIMITED FIELDS TERMINATED BY ','");
                sw.WriteLine("STORED AS TEXTFILE LOCATION 'abfs://tstdatalake@nchntsdep003sta.dfs.core.chinacloudapi.cn/test/inbound/thirdParty/CN/Sales/SalesOrder/TMS/Order/"+item
                    +"'");
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
            }
        }
    }
}
