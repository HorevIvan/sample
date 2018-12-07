using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter1
{
    static class Program
    {
        private static string _FileName;

        static void Main(string[] args)
        {
            _FileName = "log_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt";

            $"Icons8 LLC {DateTime.Now.Year}".Log();

            $"Log file: {_FileName}".Log();

            try
            {
                Console.Write("Please enter target directory: ");

                var directory = Console.ReadLine();

                $"Selected directory: {directory}".Log();

                ProcessDirectory(directory);
            }
            catch (Exception exception)
            {
                $"EXCEPTION: {exception.Message}".Log();
            }

            $"End".Log();
        }

        private static void ProcessDirectory(string directory)
        {
            var files = Directory.GetFiles(directory, "*.xlsx");

            $"Finded {files.Length} files".Log();

            foreach (var file in files)
            {
                try
                {
                    $"----------------------------------------".Log();

                    var fileInfo = new FileInfo(file);

                    $"Processing: {fileInfo.Name}".Log();

                    ProcessFile(fileInfo);

                }
                catch (Exception exception)
                {
                    $"EXCEPTION: {exception.Message}".Log();
                }
            }
        }

        private static void ProcessFile(FileInfo fileInfo)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var workbook = excelPackage.Workbook;

                var worksheet = workbook.Worksheets.First();

                try
                {
                    var cell = worksheet.Cells[7, 5];

                    var data = cell.Value;

                    $"Cell: {cell}, data: {cell}".Log();

                    var parts = data.ToString().Split('-').ToArray();

                    if (parts.Length < 2)
                    {
                        $"ERROR: wrong value {data} in {cell}".Log();

                        return;
                    }

                    if (int.TryParse(parts[0], out int month))
                    {
                        var nextMonth = month == 12 ? 1 : month + 1;

                        var newValue = nextMonth + "-" + data.ToString().Split('-')[1];

                        $"New value: {newValue}".Log();

                        cell.Value = newValue;
                    }
                    else $"ERROR: failed to get first value from {cell}".Log();
                }
                catch (Exception exception)
                {
                    ($"EXCEPTION: " + exception.Message).Log();

                    return;
                }

                try
                {
                    var cell = worksheet.Cells[8, 5];

                    var data = cell.Value;

                    $"Cell: {cell}, data: {cell}".Log();

                    var parts = data.ToString().Split(' ').ToArray();

                    if (parts.Length < 3)
                    {
                        $"ERROR: wrong value {data} in {cell}".Log();

                        return;
                    }

                    var month = parts[1].ToLower();

                    var nextMonthStr = GetNextMonth(month);

                    var nextMonthInt = GetMonthNumber(nextMonthStr);

                    var year = int.Parse(parts[2]);

                    if (nextMonthInt == 1) year++;

                    var daysInNextMonth = DateTime.DaysInMonth(year, nextMonthInt);

                    var newValue = daysInNextMonth + " " + nextMonthStr + " " + year;

                    $"New value: {newValue}".Log();

                    cell.Value = newValue;
                }
                catch (Exception exception)
                {
                    ($"EXCEPTION: " + exception.Message).Log();

                    return;
                }

                excelPackage.Save();

                $"Saved".Log();
            }
        }

        private static Int32 GetMonthNumber(String month)
        {
            month = month.ToLower();

            switch (month)
            {
                case "jan": return 1;
                case "feb": return 2;
                case "mar": return 3;
                case "apr": return 4;
                case "may": return 5;
                case "jun": return 6;
                case "jul": return 7;
                case "aug": return 8;
                case "sep": return 9;
                case "oct": return 10;
                case "nov": return 11;
                case "dec": return 12;

                default: throw new Exception("Wrong month " + month);
            }
        }

        private static String GetNextMonth(String month)
        {
            month = month.ToLower();

            switch (month)
            {
                case "jan": return "Feb";
                case "feb": return "Mar";
                case "mar": return "Apr";
                case "apr": return "May";
                case "may": return "Jun";
                case "jun": return "Jul";
                case "jul": return "Aug";
                case "aug": return "Sep";
                case "sep": return "Oct";
                case "oct": return "Nov";
                case "nov": return "Dec";
                case "dec": return "Jan";

                default: throw new Exception("Wrong month " + month);
            }
        }

        public static void Log(this string str)
        {
            str = $"[{DateTime.Now.ToString("HH:mm:ss.ff")}] {str}";

            Console.WriteLine(str);

            File.AppendAllText(_FileName, str + "\n");

            var strHash = str.GetHashCode();

            Console.Beep(37 + Math.Abs(strHash % 10000), 50 + Math.Abs(strHash % 50));
        }
    }
}
