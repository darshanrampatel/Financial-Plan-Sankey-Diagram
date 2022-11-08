﻿using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace FinancialPlanSankey
{
    class Program
    {
        static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var excelFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Dropbox", "Finances", "Financial Plan.xlsm");
            var copiedFile = Path.GetRandomFileName();
            Console.WriteLine($"Copying {excelFile}...");
            File.Copy(excelFile, copiedFile);

            using (var stream = File.Open(copiedFile, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    // Gets or sets a value indicating whether to set the DataColumn.DataType 
                    // property in a second pass.
                    UseColumnDataType = true,

                    // Gets or sets a callback to determine whether to include the current sheet
                    // in the DataSet. Called once per sheet before ConfigureDataTable.
                    FilterSheet = (tableReader, sheetIndex) => tableReader.Name == "Transactions" /* || tableReader.Name == "Expenses" */,

                    // Gets or sets a callback to obtain configuration options for a DataTable. 
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        // Gets or sets a value indicating the prefix of generated column names.
                        EmptyColumnNamePrefix = "Column",

                        // Gets or sets a value indicating whether to use a row from the 
                        // data as column names.
                        UseHeaderRow = true,

                        // Gets or sets a callback to determine whether to include the 
                        // current row in the DataTable.
                        FilterRow = (rowReader) =>
                        {
                            int progress = (int)Math.Ceiling(rowReader.Depth / (decimal)rowReader.RowCount * 100); // progress is in the range 0..100
                            Console.WriteLine($"Reading {rowReader.Name} {progress}%");
                            return true;
                        },

                        // Gets or sets a callback to determine whether to include the specific
                        // column in the DataTable. Called once per column after reading the 
                        // headers.
                        FilterColumn = (rowReader, columnIndex) => true
                    }
                });
                if (result.Tables.Count > 0)
                {
                    Console.Clear();
                    var sb = new StringBuilder();
                    foreach (DataTable table in result.Tables)
                    {
                        if (table.Columns.Count > 0 && table.Rows.Count > 0)
                        {
                            Console.WriteLine($"{table.TableName}: {table.Columns.Count} columns, {table.Rows.Count} rows");
                            Console.WriteLine();
                        }
                        switch (table.TableName)
                        {
                            case "Expenses":
                                var expensesList = table.ToList<Expenses>();
                                var groupedExpenses = expensesList
                                    .GroupBy(g => g.Category.Split(":").First())
                                    .Select(g => new GroupedExpenses { GroupedCategory = g.Key, SubExpenses = g.ToList() });
                                var groupIndex = 0;
                                foreach (var groupedExpense in groupedExpenses)
                                {
                                    if (groupedExpense.GroupTotal < 0)
                                    {
                                        var groupColour = GetColourFromNumber(groupIndex).BackgroundColour;
                                        sb.AppendLine($":{groupedExpense.GroupedCategory} {groupColour}");
                                        sb.AppendLine($"Net [{-1 * groupedExpense.GroupTotal * 12:#.00}] {groupedExpense.GroupedCategory} {groupColour}");
                                        foreach (var subCategory in groupedExpense.SubExpenses)
                                        {
                                            if (subCategory.TwelveMonths < 0)
                                            {
                                                if (subCategory.Category != groupedExpense.GroupedCategory)
                                                {
                                                    sb.AppendLine($":{subCategory.Category} {groupColour}");
                                                }
                                                sb.AppendLine($"{groupedExpense.GroupedCategory} [{-1 * subCategory.TwelveMonths * 12:#.00}] {(subCategory.Category == groupedExpense.GroupedCategory ? $"{groupedExpense.GroupedCategory}: Unassigned" : subCategory.Category)} {groupColour}");
                                            }
                                        }
                                        groupIndex++;
                                    }
                                }
                                break;
                            case "Transactions":
                                var transactionsList = table.ToList<Transactions>();
                                var charityTransactions = transactionsList
                                    .Where(t => t.Category == "Charitable Donations")
                                    .GroupBy(t => t.Date.Year)
                                    .Select(g => new { Year = g.Key, Total = Math.Abs(g.Sum(t => t.Amount)) });
                                var incomeTransactions = transactionsList
                                    .Where(t => t.Category == "Wages & Salary" ||
                                       (t.Category == "Taxes" && t.Subcategory == "Income Tax") ||
                                       (t.Category == "Taxes" && t.Subcategory == "National Insurance")
                                     )
                                    .GroupBy(t => t.Date.Year)
                                    .Select(g => new { Year = g.Key, Total = g.Sum(t => t.Amount) });
                                var charityPercentages = incomeTransactions
                                     .Select(i => new CharityPercentage
                                     {
                                         Year = i.Year,
                                         IncomeTotal = i.Total,
                                         CharityTotal = charityTransactions.FirstOrDefault(c => c.Year == i.Year)?.Total ?? 0
                                     })
                                    .ToList();
                                Console.WriteLine($"     --Charitable Donations By Year--    ");
                                Console.WriteLine($"Year: {"Charity",11} / {"Income",11}  = {"%",6}");
                                foreach (var year in charityPercentages)
                                {
                                    Console.WriteLine(year);
                                }
                                Console.WriteLine($" *The current year's income is estimated");
                                Console.WriteLine();
                                var lockdownStart = new DateTime(2020, 3, 23);
                                lockdownStart = new DateTime(2022, 1, 1); // set to 2022
                                var lockdownEnd = new DateTime(2022, 12, 31);
                                var lockdownTransactions = transactionsList
                                    .Where(t => t.Date >= lockdownStart && t.Date <= lockdownEnd)
                                    // .Where(t => (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))                                    
                                    .Where(t => t.Category != "Transfer") //|| (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))
                                    .GroupBy(t => t.Category)
                                    .Select(g => new GroupedTransactions
                                    {
                                        Category = g.Key,//g.Key == "Transfer" ? "Net" : g.Key,
                                        Transactions = g.ToList(),
                                    });
                                var transactionIndex = 0;
                                foreach (var group in lockdownTransactions)
                                {
                                    var groupColour = GetColourFromNumber(transactionIndex).BackgroundColour;
                                    sb.AppendLine($":{group.Category} {groupColour}");
                                    if (group.GroupTotal > 0)
                                    {
                                        sb.AppendLine($"{group.Category} [{group.GroupTotal:#.00}] Net {groupColour}");
                                        foreach (var sub in group.Sub)
                                        {
                                            if (sub.Category != group.Category)
                                            {
                                                sb.AppendLine($":{sub.Category} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                sb.AppendLine($"{sub.Category} [{sub.GroupTotal:#.00}] {group.Category} {groupColour}");
                                            }
                                            else
                                            {
                                                sb.AppendLine($"Net [{-1 * sub.GroupTotal:#.00}] {sub.Category}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        sb.AppendLine($"Net [{-1 * group.GroupTotal:#.00}] {group.Category} {groupColour}");
                                        foreach (var sub in group.Sub)
                                        {
                                            if (sub.Category != group.Category)
                                            {
                                                sb.AppendLine($":{sub.Category} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                sb.AppendLine($"{sub.Category} [{sub.GroupTotal:#.00}] WhereTo");
                                            }
                                            else
                                            {
                                                sb.AppendLine($"{group.Category} [{-1 * sub.GroupTotal:#.00}] {(sub.Category == group.Category ? $"{group.Category}: Unassigned" : sub.Category)} {groupColour}");
                                            }
                                        }
                                    }


                                    transactionIndex++;
                                }
                                break;
                            default:
                                Console.WriteLine($"Unknown sheet {table.TableName}");
                                break;
                        }

                    }
                    var output = sb.ToString();
                    // Console.Write(output);
                    WindowsClipboard.SetText(output);
                }
            }
            File.Delete(copiedFile);
            var url = "https://sankeymatic.com/build/";
            Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true }); // https://stackoverflow.com/a/43232486
        }

        public static (string BackgroundColour, string ForegoundColour) GetColourFromNumber(int n)
        {
            if (n > 0)
            {
                n--;
            }
            return Colours[Math.Abs(n % Colours.Length)];
        }

        /// <summary>
        /// Gets a stable six-character hash from a string
        /// </summary>
        /// <param name="s"></param>
        /// <returns>The six-character hash</returns>
        private static string AnonymiseShortHash(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return string.Empty;
            }
            byte[] textData = Encoding.UTF8.GetBytes(s);
            byte[] hash = System.Security.Cryptography.SHA256.HashData(textData);
            return BitConverter.ToString(hash).Replace("-", string.Empty)[..6].ToLowerInvariant();
        }

        /// <summary>
        /// The predefined list of colours
        /// </summary>
        private static readonly (string BackgroundColour, string ForegoundColour)[] Colours = {
            ("#be0032","#ffffff"),
            ("#377eb8","#ffffff"),
            ("#66a61e","#000000"),
            ("#984ea3","#ffffff"),
            ("#e6ab02","#000000"),
            ("#20b2aa","#000000"),
            ("#ff7f00","#000000"),
            ("#7f80cd","#ffffff"),
            ("#b3e900","#000000"),
            ("#c42e60","#ffffff"),
            ("#f781bf","#000000"),
            ("#8dd3c7","#000000"),
            ("#bebada","#000000"),
            ("#fb8072","#000000"),
            ("#80b1d3","#000000"),
            ("#fdb462","#000000"),
            ("#fccde5","#000000"),
            ("#bc80bd","#ffffff"),
            ("#ffed6f","#000000"),
            ("#c4eaff","#000000"),
            ("#1b9e77","#ffffff"),
            ("#d95f02","#ffffff"),
            ("#e7298a","#ffffff"),
            ("#0097ff","#ffffff"),
            ("#00a935","#ffffff"),
            ("#d3060e","#ffffff"),
            ("#191970","#ffffff"),
            ("#848482","#ffffff"),
            ("#ba55d3","#ffffff"),
            ("#3cb371","#000000"),
            ("#c71585","#ffffff"),
        };

        public class Expenses
        {
            public string Category { get; set; }
            public double TwelveMonths { get; set; }
        }

        public class GroupedExpenses
        {
            public string GroupedCategory { get; set; }
            public List<Expenses> SubExpenses { get; set; }
            public double GroupTotal => SubExpenses?.Sum(s => s.TwelveMonths) ?? 0;
        }

        public class GroupedTransactions
        {
            public string Category { get; set; }
            public List<Transactions> Transactions { get; set; }
            public double GroupTotal => Transactions?.Sum(s => s.Amount) ?? 0;
            public List<GroupedTransactions> Sub => Transactions?.GroupBy(t => string.IsNullOrWhiteSpace(t.Subcategory) ? $"{Category}: Unassigned" : $"{Category}: {t.Subcategory}").Select(g => new GroupedTransactions { Category = g.Key, Transactions = g.ToList() }).ToList();
        }

        public class Transactions
        {
            public DateTime Date { get; set; }
            public string Account { get; set; }
            public string Payee { get; set; }
            public double Amount { get; set; }
            public string Category { get; set; }
            public string Subcategory { get; set; }
            public string Memo { get; set; }
            public string Type { get; set; }
        }

        public class CharityPercentage
        {
            public int Year { get; set; }
            public double IncomeTotal { private get; set; }
            public double CharityTotal { get; set; }
            private double ActualOrEstimatedIncomeTotal => Year == DateTime.UtcNow.Year
                                         ? (IncomeTotal / ((double)DateTime.UtcNow.DayOfYear / Helper.GetDaysInYear(DateTime.UtcNow.Year)))
                                         : IncomeTotal;
            public double Percentage => (CharityTotal / ActualOrEstimatedIncomeTotal) * 100;
            public override string ToString() => $"{Year}: £{CharityTotal,10:N2} / £{ActualOrEstimatedIncomeTotal,10:N2}{(Year == DateTime.UtcNow.Year ? "*" : " ")} = {Percentage,5:N2}%";
        }
    }

    // https://codereview.stackexchange.com/a/101827/213992
    public static class Helper
    {
        private static readonly IDictionary<Type, ICollection<PropertyInfo>> _Properties =
            new Dictionary<Type, ICollection<PropertyInfo>>();

        /// <summary>
        /// Converts a DataTable to a list with generic objects
        /// </summary>
        /// <typeparam name="T">Generic object</typeparam>
        /// <param name="table">DataTable</param>
        /// <returns>List with generic objects</returns>
        public static IEnumerable<T> ToList<T>(this DataTable table) where T : class, new()
        {
            try
            {
                var objType = typeof(T);
                ICollection<PropertyInfo> properties;

                lock (_Properties)
                {
                    if (!_Properties.TryGetValue(objType, out properties))
                    {
                        properties = objType.GetProperties().Where(property => property.CanWrite).ToList();
                        _Properties.Add(objType, properties);
                    }
                }

                var list = new List<T>(table.Rows.Count);

                foreach (var row in table.AsEnumerable().Skip(1))
                {
                    var obj = new T();

                    foreach (var prop in properties)
                    {
                        try
                        {
                            var propName = prop.Name switch
                            {
                                nameof(Program.Expenses.TwelveMonths) => "12 Months",
                                _ => prop.Name,
                            };
                            var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                            var safeValue = row[propName] == null ? null : Convert.ChangeType(row[propName], propType);

                            prop.SetValue(obj, safeValue, null);
                        }
                        catch
                        {
                            // ignored
                        }
                    }

                    list.Add(obj);
                }

                return list;
            }
            catch
            {
                return Enumerable.Empty<T>();
            }
        }

        public static int GetDaysInYear(int year) => Enumerable.Range(1, 12).Sum(month => DateTime.DaysInMonth(year, month));
    }
}
