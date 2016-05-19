using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace XLOC.Utility.Extensions
{
    internal static class Extensions
    {
        #region string
        //=================================================
        internal static bool rIsMatch(this string local, string pattern, bool IgnoreCase = true) => Regex.IsMatch(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        internal static string rMatch(this string local, string pattern, bool IgnoreCase = true) => Regex.Match(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None).Value;
        internal static IEnumerable<string> rMatches(this string local, string pattern, bool IgnoreCase = true) => Regex.Matches(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None).Cast<Match>().Select(x => x.Value);
        internal static string rReplace(this string local, string pattern, string replacement, bool IgnoreCase = true) => Regex.Replace(local, pattern, replacement, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        internal static bool HasValue(this string value) => !string.IsNullOrWhiteSpace(value);
        //=================================================
        #endregion

        public static void ForEach<T>(this IEnumerable<T> local, Action<T> action)
        {
            var e = local.GetEnumerator();
            while (e.MoveNext())
                action(e.Current);
        }

        internal static IEnumerable<Cell> GetCaptionCells(this WorksheetPart sheet, int skipCount = 0) => sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().Skip(skipCount).First().Descendants<Cell>().Where(x => x.CellValue != null);
    }
}
