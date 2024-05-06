using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace FinancialPlanSankey
{
    class Config
    {
        public string OriginalLoanPrincipalAccount { get; set; }
        public string InvestmentProvider { get; set; }
        public string PensionProvider { get; set; }
        public string DifferenceAddress { get; set; }
    }

    partial class Program
    {
        static Config config;
        static void Main()
        {
            var anonymise = false;
            ConsoleKey response;
            do
            {
                // https://stackoverflow.com/questions/2642585/read-a-variable-in-bash-with-a-default-value
                Console.WriteLine("Anonymise? (y/n) [Default = n]");
                response = Console.ReadKey(true).Key; // Don't show
                if (response == ConsoleKey.Enter)
                {
                    anonymise = false;
                    break;
                }
                anonymise = response == ConsoleKey.Y;

            } while (response != ConsoleKey.Y && response != ConsoleKey.N);

            var monthlyAverage = false;
            do
            {
                // https://stackoverflow.com/questions/2642585/read-a-variable-in-bash-with-a-default-value
                Console.WriteLine("Monthly Average? (y/n) [Default = n]");
                response = Console.ReadKey(true).Key; // Don't show
                if (response == ConsoleKey.Enter)
                {
                    monthlyAverage = false;
                    break;
                }
                monthlyAverage = response == ConsoleKey.Y;

            } while (response != ConsoleKey.Y && response != ConsoleKey.N);

            config = new ConfigurationBuilder().AddUserSecrets<Config>().Build().GetSection(nameof(Config)).Get<Config>();

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
                            Console.Write($"\rReading {rowReader.Name} {progress}%");
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
                                    Console.WriteLine(anonymise ? year.AnonymisedString : year);
                                }
                                Console.WriteLine($" *The current year's income is estimated");
                                Console.WriteLine();
                                var startDate = new DateTime(DateTime.Today.Year, 1, 1);
                                var endDate = DateTime.Today;
                                var proportionOfYear = endDate.DayOfYear / (double)(DateTime.IsLeapYear(endDate.Year) ? 366 : 365);
                                var monthFactor = monthlyAverage ? 12 * proportionOfYear : 1;
                                var transactions = transactionsList
                                    .Where(t => t.Date >= startDate && t.Date <= endDate)
                                    // .Where(t => (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))                                    
                                    .Where(t => t.Category != "Transfer") //|| (t.Category == "Transfer" && !t.Subcategory.Contains("Credit Card")))                                    
                                    .GroupBy(t => t.Category)
                                    .Select(g => new GroupedTransactions
                                    {
                                        Category = g.Key,//g.Key == "Transfer" ? "Net" : g.Key,
                                        MonthFactor = monthFactor,
                                        Transactions = g.ToList(),
                                    });
                                var transactionIndex = 0;
                                foreach (var group in transactions.OrderBy(g => g.GroupTotal)) // Order expenses by largest first
                                {
                                    var groupColour = GetColourFromNumber(transactionIndex).BackgroundColour;
                                    var groupCategory = anonymise ? AnonymiseShortHash(group.Category) : group.Category;
                                    sb.AppendLine($":{groupCategory} {groupColour}");
                                    if (group.GroupTotal > 0)
                                    {
                                        sb.AppendLine($"{groupCategory} [{group.GroupTotal:#.00}] Net {groupColour}");
                                        foreach (var sub in group.Sub.OrderByDescending(s => s.GroupTotal)) // Order income by largest first
                                        {
                                            var subCategory = anonymise ? AnonymiseShortHash(sub.Category) : sub.Category;
                                            if (subCategory != groupCategory)
                                            {
                                                sb.AppendLine($":{subCategory} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                sb.AppendLine($"{subCategory} [{sub.GroupTotal:#.00}] {groupCategory} {groupColour}");
                                            }
                                            else
                                            {
                                                sb.AppendLine($"Net [{-1 * sub.GroupTotal:#.00}] {subCategory}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        sb.AppendLine($"Net [{-1 * group.GroupTotal:#.00}] {groupCategory} {groupColour}");
                                        foreach (var sub in group.Sub.OrderBy(s => s.GroupTotal)) // Order expenses by largest first
                                        {
                                            var subCategory = anonymise ? AnonymiseShortHash(sub.Category) : sub.Category;
                                            if (subCategory != groupCategory)
                                            {
                                                sb.AppendLine($":{subCategory} {groupColour}");
                                            }
                                            if (sub.GroupTotal > 0)
                                            {
                                                sb.AppendLine($"{subCategory} [{sub.GroupTotal:#.00}] Net");
                                            }
                                            else
                                            {
                                                sb.AppendLine($"{groupCategory} [{-1 * sub.GroupTotal:#.00}] {(subCategory == groupCategory ? $"{groupCategory}: Unassigned" : subCategory)} {groupColour}");
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

                    var settingsFileContents = $@"
// === Nodes and Flows ===

{sb}

// === Settings ===

size w 2560
  h 2560
margin l 12
  r 12
  t 18
  b 20
bg color #ffffff
  transparent N
node w 9
  h 50
  spacing 85
  border 0
  theme a
  color #888888
  opacity 1
flow curvature 0.5
  inheritfrom outside-in
  color #999999
  opacity 0.45
layout order automatic
  justifyorigins N
  justifyends N
  reversegraph N
  attachincompletesto nearest
labels color #000000
  highlight 0.55
  fontface sans-serif
labelname appears Y
  size 16
  weight 400
labelvalue appears {(anonymise ? "N" : "Y")}
  fullprecision Y
labelposition first before
  breakpoint 6
value format ',.'
  prefix '£'
  suffix ''
themeoffset a 9
  b 0
  c 0
  d 0
meta mentionsankeymatic N
  listimbalances Y
";
                    WindowsClipboard.SetText(settingsFileContents);
                }
            }
            File.Delete(copiedFile);
            var url = "https://sankeymatic.com/build/";
            Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true }); // https://stackoverflow.com/a/43232486
            Console.ReadLine();
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

        [GeneratedRegex("[\\d-]")]
        private static partial Regex ReplaceDigitsRegex();
        private static string ReplaceDigits(string s) => ReplaceDigitsRegex().Replace(s, "?");
        private static string AnonymiseAmount(double amount, bool padToFixedLength = false) => padToFixedLength ? "???,???.??" : ReplaceDigits($"{amount,10:N2}");
        private static string AnonymisePercentage(double percentage) => ReplaceDigits($"{percentage,5:N2}");

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
            public double MonthFactor { get; init; }
            public double GroupTotal => (Transactions?.Sum(s => s.Amount) ?? 0) / MonthFactor;
            public List<GroupedTransactions> Sub => Transactions?.GroupBy(t => string.IsNullOrWhiteSpace(t.Subcategory) ? $"{Category}: Unassigned" : $"{Category}: {t.Subcategory}").Select(g => new GroupedTransactions { Category = g.Key, Transactions = g.ToList(), MonthFactor = MonthFactor }).ToList();
        }

        public class Transactions
        {
            public DateTime Date { get; set; } // Column B
            public string Account { get; set; } // Column C
            public string Payee { get; set; } // Column D
            public double Amount { get; set; } // Column F
            public string Category { get; set; } // Column G
            public string Subcategory { get; set; } // Column H
            public string Memo { get; set; } // Column I

            /// <summary>
            /// Based on this (gnarly) Excel formula:<br/>
            /// <code>
            /// =IF(ISNUMBER(SEARCH("ISA",$C2)),"ISA",IF(ISNUMBER(SEARCH("Bond",$C2)),"Bond",IF(ISNUMBER(SEARCH([InvestmentProvider],$C2)),IF($G2="Buy Investment","Buy Investment","Investment"),IF(ISNUMBER(SEARCH("Pension",$C2)),IF(ISNUMBER(SEARCH([PensionProvider],$C2)),[PensionProvider],IF($G2="Buy Investment","Pension","Pension")),IF(ISNUMBER(SEARCH("Saver",$C2)),"Saver",IF(ISNUMBER(SEARCH("Loan principal received",$I2)),IF(ISNUMBER(SEARCH([OriginalLoanPrincipalAccount],$C2)),"Cash","House"),IF(AND(ISNUMBER(SEARCH("Difference to move.",$I2)),ISNUMBER(SEARCH([DifferenceAddress],$C2))),"Difference","Cash")))))))
            /// </code>
            /// </summary>
            public string Type // Column K
            {
                get
                {
                    if (string.IsNullOrWhiteSpace(Account))
                    {
                        return "Cash";
                    }
                    if (Account.Contains("ISA", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return "ISA";
                    }
                    else if (Account.Contains("Bond", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return "Bond";
                    }
                    else if (Account.Contains(config.InvestmentProvider, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return Category == "Buy Investment" ? "Buy Investment" : "Investment";
                    }
                    else if (Account.Contains("Pension", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return Account.Contains(config.PensionProvider, StringComparison.InvariantCultureIgnoreCase) ? config.PensionProvider : "Pension";
                    }
                    else if (Account.Contains("Saver", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return "Saver";
                    }
                    else if (Memo?.Contains("Loan principal received", StringComparison.InvariantCultureIgnoreCase) == true)
                    {
                        return Account.Contains(config.OriginalLoanPrincipalAccount, StringComparison.InvariantCultureIgnoreCase) ? "Cash" : "House";
                    }
                    else if (Memo?.Contains("Difference to move.", StringComparison.InvariantCultureIgnoreCase) == true && Account.Contains(config.DifferenceAddress, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return "Difference";
                    }
                    return "Cash";
                }
            }
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
            public string AnonymisedString => $"{Year}: £{AnonymiseAmount(CharityTotal)} / £{AnonymiseAmount(ActualOrEstimatedIncomeTotal, padToFixedLength: true)}{(Year == DateTime.UtcNow.Year ? "*" : " ")} = {Percentage,5:N2}%";
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
