using OfficeOpenXml;
using System;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        var schedule = new Schedule
        {
            Faculties = new Dictionary<string, Faculty>
            {
                {"Факультет економічних наук", new Faculty
                    {
                        Specializations = new Dictionary<string, Specialization>
                        {
                            {"Економіка", new Specialization
                                {

                                }
                            },
                            {"Фінанси", new Specialization
                                {

                                }
                            },
                            {"Менеджмент", new Specialization
                                {

                                }
                            },
                            {"Маркетинг", new Specialization
                                {

                                }
                            }
                        }
                    }
                }
            }
        };
        var daysOfWeek = new HashSet<string>() { "Понеділок", "Вівторок", "Середа", "Четвер", "П`ятниця", "Субота", "Неділя" };
        var timeSlots = new HashSet<string>() { "8.30-9.50", "10.00-11.20", "11.40-13.00", "13.30-14.50", "15.00-16.20", "16.30-17.50" };
        var specialization = new HashSet<string>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string filePath = "../../../fen.xlsx";

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            Console.OutputEncoding = UTF8Encoding.UTF8;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            var a = worksheet.Cells;
            var b = a.Where(cell => cell.Value is double || cell.Value is string && !string.IsNullOrEmpty((string?)cell.Value)).Skip(11).ToArray();

            var currentDay = "Понеділок";
            var count = b.Length;
            var flag = false;
            var time = "8.30-9.50";

            for (var i = 0; i < count-4; i++)
            {
                if (daysOfWeek.Contains(b[i].Value.ToString()))
                {
                    currentDay = b[i].Value.ToString();
                }
                if (flag || timeSlots.Contains(b[i].Value.ToString()) && !daysOfWeek.Contains(b[i+1].Value.ToString()) && !timeSlots.Contains(b[i+1].Value.ToString()))
                {
                    if (!flag)
                    {
                        time = b[i].Value.ToString();
                    }
                    flag = false;
                    string subjectName = b[i + 1].Value.ToString();
                    //Console.WriteLine(subjectName + " " + currentDay);
                    if (subjectName.Contains('('))
                    {
                        int c = subjectName.LastIndexOf('(');
                        int d = subjectName.LastIndexOf(')');
                        string specKey = subjectName.Substring(c + 1, d - c - 1);
                        //Console.WriteLine(specKey);
                        if (specKey.Contains("ек.") || specKey.Contains("екон.") || specKey.Contains("ек"))
                        {
                            specialization.Add("Економіка");
                        }
                        if (specKey.Contains("фін.") || specKey.Contains("фінанси")) 
                        {
                            specialization.Add("Фінанси");
                        }
                        if (specKey.Contains("маркетинг") || specKey.Contains("марк.") || specKey.Contains("мар.") || specKey.Contains("марк,"))
                        {
                            specialization.Add("Маркетинг");
                        }
                        if (specKey.Contains("мен.") || specKey.Contains("мен,") || specKey.Contains("марк,мен"))
                        {
                            specialization.Add("Менеджмент");
                        }
                        if(specialization.Count == 0)
                        {
                            Console.WriteLine(specKey);
                            specialization.Add("Менеджмент");
                        }
                    }
                    Group group = new Group()
                    {
                        Name = b[i + 2].Value.ToString(),
                        Classroom = b[i + 4].Value.ToString(),
                        Weeks = b[i + 3].Value.ToString(),
                        Time = time,
                        DayOfWeek = currentDay,
                    };
                    //Console.WriteLine(b[i + 1].Value.ToString());
                    

                    if (schedule.Faculties["Факультет економічних наук"].Specializations["Економіка"].Subjects.Keys.Contains(subjectName))
                    {
                        schedule.Faculties["Факультет економічних наук"].Specializations["Економіка"].Subjects[subjectName].Groups.Add(group);
                    }
                    else
                    {
                        schedule.Faculties["Факультет економічних наук"].Specializations["Економіка"].Subjects.Add(subjectName, new Subject(group));
                    }
                    
                    if(i + 5 < count && !timeSlots.Contains(b[i + 5].Value.ToString()) && !daysOfWeek.Contains(b[i + 5].Value.ToString()) && !daysOfWeek.Contains(b[i + 4].Value.ToString())) {
                        
                        i += 3;
                        flag = true;
                    }

                    specialization.Clear();
                }
            }



        }

        

        /*
         var date = DateTime.Now;
        
        makeTable();

        Console.WriteLine((DateTime.Now - date).TotalSeconds);
        Console.ReadKey();
         
        */
    }
    static void makeTable()
    {
        var schedule = new Schedule
        {
            Faculties = new Dictionary<string, Faculty>
            {
                {"Факультет інформатики", new Faculty
                    {
                        Specializations = new Dictionary<string, Specialization>
                        {
                            {"Інженерія програмного забезпечення", new Specialization
                                {

                                }
                            }
                        }
                    }
                }
            }
        };

        var daysOfWeek = new HashSet<string>() { "Понеділок", "Вівторок", "Середа", "Четвер", "П`ятниця", "Субота", "Неділя" };

        //        schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects.Add("Предмет", new Subject(new Group() { Name="123"}) );

        //Console.WriteLine($"Faculty: {schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects["Предмет"].Groups[0].Name}");


        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string filePath = "../../../fi.xlsx";

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            Console.OutputEncoding = UTF8Encoding.UTF8;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            var a = worksheet.Cells;
            var b = a.Where(cell => cell.Value is double || cell.Value is string && !string.IsNullOrEmpty((string?)cell.Value)).Skip(12).ToArray();

            var currentDay = "Понеділок";
            var count = b.Length;
            var flag = false;
            var time = "8:30-9:50";

            for (var i = 0; i < count; i++)
            {
                if (daysOfWeek.Contains(b[i].Value.ToString()))
                {
                    currentDay = b[i].Value.ToString();
                }
                if (flag || (b[i].Value.ToString().Contains(':') && b[i + 1].Value.ToString().Contains(',')))
                {
                    if (!flag)
                    {
                        time = b[i].Value.ToString();
                    }
                    flag = false;
                    Group group = new Group()
                    {
                        Name = b[i + 2].Value.ToString(),
                        Classroom = b[i + 4].Value.ToString(),
                        Weeks = b[i + 3].Value.ToString(),
                        Time = time,
                        DayOfWeek = currentDay,
                    };


                    string subjectName = b[i + 1].Value.ToString();


                    if (subjectName.LastIndexOf(',') != subjectName.IndexOf(','))
                    {
                        int index = subjectName.LastIndexOf(',');
                        subjectName = subjectName.Substring(0, index) + "." + subjectName.Substring(index + 1);
                    }

                    if (schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects.Keys.Contains(subjectName))
                    {
                        schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects[subjectName].Groups.Add(group);
                    }
                    else
                    {
                        schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects.Add(subjectName, new Subject(group));
                    }

                    if (i + 5 < count && !b[i + 5].Value.ToString().Contains(':') && !daysOfWeek.Contains(b[i + 5].Value.ToString()) && !daysOfWeek.Contains(b[i + 4].Value.ToString()))
                    {
                        //Console.WriteLine(b[i+5].Value.ToString());
                        flag = true;
                        i += 3;
                    }
                }
            }

            foreach (var item in schedule.Faculties["Факультет інформатики"].Specializations["Інженерія програмного забезпечення"].Subjects)
            {
                Console.WriteLine(item.Key);
                foreach (var item2 in item.Value.Groups)
                {
                    Console.WriteLine(item2.DayOfWeek + " " + item2.Time + " " + item2.Name + " " + item2.Classroom + " " + item2.Weeks);
                }
            }
        }
    }
}
