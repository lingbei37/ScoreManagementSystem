using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace ScoreManagementSystem
{
    class Program
    {
        static StudentManager studentManager = new StudentManager();

        static void Main(string[] args)
        {
            bool running = true;

            while (running)
            {
                Console.Clear();
                Console.WriteLine("===== 成绩管理系统 =====");
                Console.WriteLine("1. 添加学生成绩");
                Console.WriteLine("2. 显示所有学生成绩");
                Console.WriteLine("3. 按成绩排序显示");
                Console.WriteLine("4. 查找学生成绩");
                Console.WriteLine("5. 多条件复合查询");
                Console.WriteLine("6. 修改学生成绩");
                Console.WriteLine("7. 删除学生信息");
                Console.WriteLine("8. 导出为Excel文件");
                Console.WriteLine("9. 保存数据");
                Console.WriteLine("10. 加载数据");
                Console.WriteLine("0. 退出系统");
                Console.WriteLine("========================");
                Console.Write("请选择操作: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        AddStudent();
                        break;
                    case "2":
                        ShowAllStudents();
                        break;
                    case "3":
                        SortAndShowStudents();
                        break;
                    case "4":
                        SearchStudent();
                        break;
                    case "5":
                        AdvancedSearch();
                        break;
                    case "6":
                        ModifyStudent();
                        break;
                    case "7":
                        DeleteStudent();
                        break;
                    case "8":
                        ExportToExcel();
                        break;
                    case "9":
                        SaveData();
                        break;
                    case "10":
                        LoadData();
                        break;
                    case "0":
                        running = false;
                        break;
                    default:
                        Console.WriteLine("无效的选择，请重试");
                        break;
                }

                if (running && choice != "0")
                {
                    Console.WriteLine("\n按任意键继续...");
                    Console.ReadKey();
                }
            }

            Console.WriteLine("感谢使用成绩管理系统，再见！");
        }

        static void AddStudent()
        {
            Console.WriteLine("\n----- 添加学生成绩 -----");

            try
            {
                Console.Write("请输入学号: ");
                string id = Console.ReadLine();

                if (studentManager.GetStudentById(id) != null)
                {
                    Console.WriteLine("该学号已存在！");
                    return;
                }

                Console.Write("请输入姓名: ");
                string name = Console.ReadLine();

                Console.Write("请输入班级: ");
                string className = Console.ReadLine();

                Console.Write("请输入语文成绩: ");
                double chinese = double.Parse(Console.ReadLine());

                Console.Write("请输入数学成绩: ");
                double math = double.Parse(Console.ReadLine());

                Console.Write("请输入英语成绩: ");
                double english = double.Parse(Console.ReadLine());

                Student student = new Student(id, name, className, chinese, math, english);
                studentManager.AddStudent(student);

                Console.WriteLine("学生信息添加成功！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加失败: {ex.Message}");
            }
        }

        static void ShowAllStudents()
        {
            Console.WriteLine("\n----- 所有学生成绩 -----");
            studentManager.ShowAllStudents();
        }

        static void SortAndShowStudents()
        {
            Console.WriteLine("\n----- 按成绩排序 -----");
            Console.WriteLine("1. 按语文成绩排序");
            Console.WriteLine("2. 按数学成绩排序");
            Console.WriteLine("3. 按英语成绩排序");
            Console.WriteLine("4. 按总成绩排序");
            Console.Write("请选择排序方式: ");

            string choice = Console.ReadLine();
            Console.WriteLine("1. 升序");
            Console.WriteLine("2. 降序");
            Console.Write("请选择排序顺序: ");
            string order = Console.ReadLine();

            bool ascending = order == "1";

            var sortedStudents = studentManager.SortStudents(choice, ascending);
            studentManager.ShowStudents(sortedStudents);
        }

        static void SearchStudent()
        {
            Console.WriteLine("\n----- 查找学生 -----");
            Console.WriteLine("1. 按学号查找");
            Console.WriteLine("2. 按姓名查找");
            Console.WriteLine("3. 按班级查找");
            Console.Write("请选择查找方式: ");

            string choice = Console.ReadLine();
            Console.Write("请输入查找内容: ");
            string keyword = Console.ReadLine();

            var result = studentManager.SearchStudents(choice, keyword);
            studentManager.ShowStudents(result);
        }

        static void AdvancedSearch()
        {
            Console.WriteLine("\n----- 多条件复合查询 -----");

            Console.Write("请输入班级(不限制请留空): ");
            string className = Console.ReadLine();

            Console.Write("请输入最低总分(不限制请留空): ");
            string minTotalStr = Console.ReadLine();
            double? minTotal = string.IsNullOrEmpty(minTotalStr) ? null : (double?)double.Parse(minTotalStr);

            Console.Write("请输入最高总分(不限制请留空): ");
            string maxTotalStr = Console.ReadLine();
            double? maxTotal = string.IsNullOrEmpty(maxTotalStr) ? null : (double?)double.Parse(maxTotalStr);

            var result = studentManager.AdvancedSearch(className, minTotal, maxTotal);
            studentManager.ShowStudents(result);
        }

        static void ModifyStudent()
        {
            Console.WriteLine("\n----- 修改学生成绩 -----");
            Console.Write("请输入要修改的学生学号: ");
            string id = Console.ReadLine();

            var student = studentManager.GetStudentById(id);
            if (student == null)
            {
                Console.WriteLine("未找到该学生！");
                return;
            }

            try
            {
                Console.WriteLine($"当前信息: {student}");

                Console.Write("请输入新姓名(按回车不修改): ");
                string name = Console.ReadLine();
                if (!string.IsNullOrEmpty(name))
                    student.Name = name;

                Console.Write("请输入新班级(按回车不修改): ");
                string className = Console.ReadLine();
                if (!string.IsNullOrEmpty(className))
                    student.ClassName = className;

                Console.Write("请输入新语文成绩(按回车不修改): ");
                string chineseStr = Console.ReadLine();
                if (!string.IsNullOrEmpty(chineseStr))
                    student.Chinese = double.Parse(chineseStr);

                Console.Write("请输入新数学成绩(按回车不修改): ");
                string mathStr = Console.ReadLine();
                if (!string.IsNullOrEmpty(mathStr))
                    student.Math = double.Parse(mathStr);

                Console.Write("请输入新英语成绩(按回车不修改): ");
                string englishStr = Console.ReadLine();
                if (!string.IsNullOrEmpty(englishStr))
                    student.English = double.Parse(englishStr);

                Console.WriteLine("学生信息修改成功！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"修改失败: {ex.Message}");
            }
        }

        static void DeleteStudent()
        {
            Console.WriteLine("\n----- 删除学生信息 -----");
            Console.Write("请输入要删除的学生学号: ");
            string id = Console.ReadLine();

            bool success = studentManager.DeleteStudent(id);
            if (success)
                Console.WriteLine("学生信息删除成功！");
            else
                Console.WriteLine("未找到该学生，删除失败！");
        }

        static void ExportToExcel()
        {
            Console.WriteLine("\n----- 导出为Excel -----");
            Console.Write("请输入导出文件路径(如: scores.xlsx): ");
            string filePath = Console.ReadLine();

            try
            {
                // 检查是否安装了NPOI库
                bool npoiAvailable = CheckNPOIAvailable();

                if (npoiAvailable)
                {
                    ExcelExporter.ExportToExcel(studentManager.GetAllStudents(), filePath);
                    Console.WriteLine($"成功导出到 {filePath}");
                }
                else
                {
                    Console.WriteLine("NPOI库未找到，无法导出为Excel。");
                    Console.WriteLine("请先通过NuGet安装NPOI库: Install-Package NPOI");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出失败: {ex.Message}");
            }
        }

        static bool CheckNPOIAvailable()
        {
            try
            {
                // 尝试加载NPOI相关的类型
                Type.GetType("NPOI.XSSF.UserModel.XSSFWorkbook, NPOI.XSSF");
                return true;
            }
            catch
            {
                return false;
            }
        }

        static void SaveData()
        {
            Console.WriteLine("\n----- 保存数据 -----");
            Console.Write("请输入保存文件路径(如: data.txt): ");
            string filePath = Console.ReadLine();

            try
            {
                studentManager.SaveToFile(filePath);
                Console.WriteLine($"数据已成功保存到 {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存失败: {ex.Message}");
            }
        }

        static void LoadData()
        {
            Console.WriteLine("\n----- 加载数据 -----");
            Console.Write("请输入加载文件路径: ");
            string filePath = Console.ReadLine();

            try
            {
                studentManager.LoadFromFile(filePath);
                Console.WriteLine($"数据已成功从 {filePath} 加载");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载失败: {ex.Message}");
            }
        }
    }
}
