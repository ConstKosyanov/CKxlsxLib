using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CKxlsxLib
{
    public static class Extensions
    {
        public static bool rIsMatch(this string local, string pattern, bool IgnoreCase = true)
        {
            return Regex.IsMatch(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }

        public static string rMatch(this string local, string pattern, bool IgnoreCase = true)
        {
            return Regex.Match(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None).Value;
        }

        public static IEnumerable<string> rMatches(this string local, string pattern, bool IgnoreCase = true)
        {
            return Regex.Matches(local, pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None).Cast<Match>().Select(x => x.Value);
        }

        public static string rReplace(this string local, string pattern, string replacement, bool IgnoreCase = true)
        {
            return Regex.Replace(local, pattern, replacement, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }

        public static bool HasValue(this string value)
        {
            return !string.IsNullOrWhiteSpace(value);
        }

        public static void ForEach<T>(this IEnumerable<T> local, Action<T> action)
        {
            var e = local.GetEnumerator();
            while (e.MoveNext())
                action(e.Current);
        }
    }
}
