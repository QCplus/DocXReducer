using System;

namespace DocxReducer.Helpers
{
    internal static class IntExtensions
    {
        public static string GetLastNDigits(this int number, int n)
        {
            if (n <= 0)
                return "";

            var digits = Math.Abs(number).ToString();

            return digits.Substring(Math.Max(digits.Length - n, 0));
        }
    }
}
