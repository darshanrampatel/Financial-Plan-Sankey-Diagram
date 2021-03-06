﻿using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

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
                                        Console.WriteLine($":{groupedExpense.GroupedCategory} {groupColour}");
                                        Console.WriteLine($"Net [{-1 * groupedExpense.GroupTotal * 12:#.00}] {groupedExpense.GroupedCategory} {groupColour}");
                                        foreach (var subCategory in groupedExpense.SubExpenses)
                                        {
                                            if (subCategory.TwelveMonths < 0)
                                            {
                                                if (subCategory.Category != groupedExpense.GroupedCategory)
                                                {
                                                    Console.WriteLine($":{subCategory.Category} {groupColour}");
                                                }
                                                Console.WriteLine($"{groupedExpense.GroupedCategory} [{-1 * subCategory.TwelveMonths * 12:#.00}] {(subCategory.Category == groupedExpense.GroupedCategory ? $"{groupedExpense.GroupedCategory}: Unassigned" : subCategory.Category)} {groupColour}");
                                            }
                                        }
                                        groupIndex++;
                                    }
                                }
                                break;
                            case "Transactions":
                                var transactionsList = table.ToList<Transactions>();
                                var lockdownStart = new DateTime(2020, 3, 23);
                                var lockdownTransactions = transactionsList
                                    .Where(t => t.Date >= lockdownStart)
                                    // .Where(t => (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))
                                    .Where(t => t.Category != "Transfer" || (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))
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
                                    Console.WriteLine($":{group.Category} {groupColour}");
                                    if (group.GroupTotal > 0)
                                    {
                                        Console.WriteLine($"{group.Category} [{group.GroupTotal:#.00}] Net {groupColour}");
                                        foreach (var sub in group.Sub)
                                        {
                                            if (sub.Category != group.Category)
                                            {
                                                Console.WriteLine($":{sub.Category} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                Console.WriteLine($"{sub.Category} [{sub.GroupTotal:#.00}] {group.Category} {groupColour}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Net [{-1 * sub.GroupTotal:#.00}] {sub.Category}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Net [{-1 * group.GroupTotal:#.00}] {group.Category} {groupColour}");
                                        foreach (var sub in group.Sub)
                                        {
                                            if (sub.Category != group.Category)
                                            {
                                                Console.WriteLine($":{sub.Category} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                Console.WriteLine($"{sub.Category} [{sub.GroupTotal:#.00}] WhereTo");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"{group.Category} [{-1 * sub.GroupTotal:#.00}] {(sub.Category == group.Category ? $"{group.Category}: Unassigned" : sub.Category)} {groupColour}");
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
                }
            }
            File.Delete(copiedFile);
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
    }
}
