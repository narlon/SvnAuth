using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace SvnAuth
{
    class Program
    {
        private static bool hasWriteHead = false;

        static void Main(string[] args)
        {
            FileInfo file = new FileInfo("./SvnAuth.xlsx");
            var ep = new ExcelPackage(file);
            ExcelWorkbook workbookIn = ep.Workbook;

            StreamWriter sw = new StreamWriter("./dav_svn_st.authz");
            foreach (var sheet in workbookIn.Worksheets)
            {
                CheckSheet(sw, sheet);
            }

            sw.Close();
        }

        private static void CheckSheet(StreamWriter sw, ExcelWorksheet sheet)
        {
            int collumCount = sheet.Dimension.End.Column;
            int rowCount = sheet.Dimension.End.Row;
            if (!hasWriteHead)
            {
                sw.WriteLine("[groups]");
                string groupName = "";
                List<string> groupMember = new List<string>();
                for (int i = 2; i <= collumCount; i++) //写组信息
                {
                    string val = sheet.GetValue(1, i).ToString();
                    if (val.StartsWith("@")) //是一个组
                    {
                        if (groupName != "") //把上一个组信息写入文件
                        {
                            sw.WriteLine(string.Format("{0}={1}", groupName.Substring(1), string.Join(",", groupMember.ToArray())));
                        }
                        groupName = val;
                        groupMember.Clear();
                    }
                    else
                    {
                        groupMember.Add(val);
                    }
                }
                if (groupName != "") //把最后一个组信息写入文件
                {
                    sw.WriteLine(string.Format("{0}={1}", groupName.Substring(1), string.Join(",", groupMember.ToArray())));
                }
                hasWriteHead = true;
            }

            sw.WriteLine();
            for (int j = 2; j <= rowCount; j++) //遍历其他所有行
            {
                string dirName = sheet.GetValue(j, 1).ToString(); //读取路径名字
                sw.WriteLine("[{0}]", dirName);
                sw.WriteLine("* =");
                for (int i = 2; i <= collumCount; i++) //写组信息
                {
                    var valObj = sheet.GetValue(j, i);
                    if (valObj != null) //说明有某种权限
                    {
                        string opName = sheet.GetValue(1, i).ToString();
                        sw.WriteLine("{0} = {1}", opName, valObj.ToString());
                    }
                }
                sw.WriteLine();
            }
        }
    }
}
