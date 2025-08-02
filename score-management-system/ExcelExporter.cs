using System;
using System.Collections.Generic;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ScoreManagementSystem
{
    public static class ExcelExporter
    {
        public static void ExportToExcel(List<Student> students, string filePath)
        {
            if (students == null)
                throw new ArgumentNullException(nameof(students));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("文件路径不能为空", nameof(filePath));

            // 创建工作簿
            XSSFWorkbook workbook = new XSSFWorkbook();

            // 创建工作表
            ISheet sheet = workbook.CreateSheet("学生成绩");

            // 创建表头行
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("学号");
            headerRow.CreateCell(1).SetCellValue("姓名");
            headerRow.CreateCell(2).SetCellValue("班级");
            headerRow.CreateCell(3).SetCellValue("语文");
            headerRow.CreateCell(4).SetCellValue("数学");
            headerRow.CreateCell(5).SetCellValue("英语");
            headerRow.CreateCell(6).SetCellValue("总分");
            headerRow.CreateCell(7).SetCellValue("平均分");

            // 设置列宽
            for (int i = 0; i < 8; i++)
            {
                sheet.SetColumnWidth(i, 20 * 256); // 20个字符宽
            }

            // 填充数据
            for (int i = 0; i < students.Count; i++)
            {
                Student student = students[i];
                IRow row = sheet.CreateRow(i + 1);

                row.CreateCell(0).SetCellValue(student.Id);
                row.CreateCell(1).SetCellValue(student.Name);
                row.CreateCell(2).SetCellValue(student.ClassName);
                row.CreateCell(3).SetCellValue(student.Chinese);
                row.CreateCell(4).SetCellValue(student.Math);
                row.CreateCell(5).SetCellValue(student.English);
                row.CreateCell(6).SetCellValue(student.TotalScore);
                row.CreateCell(7).SetCellValue(student.AverageScore);
            }

            // 创建文件并写入内容
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
}
