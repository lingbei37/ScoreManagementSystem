using System;
using System.Collections.Generic;
using System.Linq;

namespace ScoreManagementSystem
{
    public class StudentManager
    {
        private List<Student> students = new List<Student>();

        public void AddStudent(Student student)
        {
            if (student == null)
                throw new ArgumentNullException(nameof(student));

            if (students.Any(s => s.Id == student.Id))
                throw new InvalidOperationException($"学号 {student.Id} 已存在");

            students.Add(student);
        }

        public Student GetStudentById(string id)
        {
            return students.FirstOrDefault(s => s.Id == id);
        }

        public List<Student> GetAllStudents()
        {
            return new List<Student>(students);
        }

        public void ShowAllStudents()
        {
            if (students.Count == 0)
            {
                Console.WriteLine("没有学生信息");
                return;
            }

            Console.WriteLine($"{"学号",-10} {"姓名",-10} {"班级",-10} {"语文",-6} {"数学",-6} {"英语",-6} {"总分",-6} {"平均分",-8}");
            Console.WriteLine("--------------------------------------------------------------------------");

            foreach (var student in students)
            {
                Console.WriteLine($"{student.Id,-10} {student.Name,-10} {student.ClassName,-10} " +
                                  $"{student.Chinese,-6:F1} {student.Math,-6:F1} {student.English,-6:F1} " +
                                  $"{student.TotalScore,-6:F1} {student.AverageScore,-8:F2}");
            }
        }

        public void ShowStudents(List<Student> studentsToShow)
        {
            if (studentsToShow == null || studentsToShow.Count == 0)
            {
                Console.WriteLine("没有找到匹配的学生信息");
                return;
            }

            Console.WriteLine($"{"学号",-10} {"姓名",-10} {"班级",-10} {"语文",-6} {"数学",-6} {"英语",-6} {"总分",-6} {"平均分",-8}");
            Console.WriteLine("--------------------------------------------------------------------------");

            foreach (var student in studentsToShow)
            {
                Console.WriteLine($"{student.Id,-10} {student.Name,-10} {student.ClassName,-10} " +
                                  $"{student.Chinese,-6:F1} {student.Math,-6:F1} {student.English,-6:F1} " +
                                  $"{student.TotalScore,-6:F1} {student.AverageScore,-8:F2}");
            }
        }

        public List<Student> SortStudents(string sortBy, bool ascending)
        {
            var sorted = new List<Student>(students);

            switch (sortBy)
            {
                case "1": // 语文
                    sorted.Sort((s1, s2) => s1.Chinese.CompareTo(s2.Chinese));
                    break;
                case "2": // 数学
                    sorted.Sort((s1, s2) => s1.Math.CompareTo(s2.Math));
                    break;
                case "3": // 英语
                    sorted.Sort((s1, s2) => s1.English.CompareTo(s2.English));
                    break;
                case "4": // 总分
                    sorted.Sort((s1, s2) => s1.TotalScore.CompareTo(s2.TotalScore));
                    break;
                default:
                    throw new ArgumentException("无效的排序方式");
            }

            if (!ascending)
                sorted.Reverse();

            return sorted;
        }

        public List<Student> SearchStudents(string searchBy, string keyword)
        {
            if (string.IsNullOrEmpty(keyword))
                return new List<Student>(students);

            switch (searchBy)
            {
                case "1": // 学号
                    return students.Where(s => s.Id.Contains(keyword)).ToList();
                case "2": // 姓名
                    return students.Where(s => s.Name.Contains(keyword)).ToList();
                case "3": // 班级
                    return students.Where(s => s.ClassName.Contains(keyword)).ToList();
                default:
                    throw new ArgumentException("无效的查找方式");
            }
        }

        public List<Student> AdvancedSearch(string className, double? minTotal, double? maxTotal)
        {
            var result = students.AsQueryable();

            if (!string.IsNullOrEmpty(className))
                result = result.Where(s => s.ClassName == className);

            if (minTotal.HasValue)
                result = result.Where(s => s.TotalScore >= minTotal.Value);

            if (maxTotal.HasValue)
                result = result.Where(s => s.TotalScore <= maxTotal.Value);

            return result.ToList();
        }

        public bool DeleteStudent(string id)
        {
            var student = GetStudentById(id);
            if (student != null)
            {
                students.Remove(student);
                return true;
            }
            return false;
        }

        public void SaveToFile(string filePath)
        {
            using (var writer = new System.IO.StreamWriter(filePath))
            {
                foreach (var student in students)
                {
                    writer.WriteLine(student.ToFileString());
                }
            }
        }

        public void LoadFromFile(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
                throw new System.IO.FileNotFoundException("文件不存在", filePath);

            students.Clear();

            foreach (var line in System.IO.File.ReadAllLines(filePath))
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    try
                    {
                        var student = Student.FromFileString(line);
                        students.Add(student);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"跳过无效数据行: {line}, 错误: {ex.Message}");
                    }
                }
            }
        }
    }
}
