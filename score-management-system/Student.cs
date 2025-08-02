using System;

namespace ScoreManagementSystem
{
    public class Student
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ClassName { get; set; }
        public double Chinese { get; set; }
        public double Math { get; set; }
        public double English { get; set; }

        public double TotalScore => Chinese + Math + English;
        public double AverageScore => TotalScore / 3;

        public Student(string id, string name, string className, double chinese, double math, double english)
        {
            Id = id;
            Name = name;
            ClassName = className;
            Chinese = chinese;
            Math = math;
            English = english;
        }

        public override string ToString()
        {
            return $"学号: {Id}, 姓名: {Name}, 班级: {ClassName}, 语文: {Chinese}, 数学: {Math}, 英语: {English}, 总分: {TotalScore}, 平均分: {AverageScore:F2}";
        }

        // 用于文件存储的格式
        public string ToFileString()
        {
            return $"{Id}|{Name}|{ClassName}|{Chinese}|{Math}|{English}";
        }

        // 从文件字符串解析学生对象
        public static Student FromFileString(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length != 6)
                throw new FormatException("无效的学生数据格式");

            return new Student(
                parts[0],
                parts[1],
                parts[2],
                double.Parse(parts[3]),
                double.Parse(parts[4]),
                double.Parse(parts[5])
            );
        }
    }
}
