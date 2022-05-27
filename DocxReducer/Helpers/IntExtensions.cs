using System;

namespace DocxReducer.Helpers
{
    internal static class IntExtensions
    {
        public static string GetLastNDigits(this int number, int n)
        {
            if (number == 0)
                return "0";

            string digits = "";
            number = Math.Abs(number);

            for (int i = 0; i < n && number > 0; i++, number /= 10)
                digits = (number % 10).ToString() + digits;

            return digits;
        }
    }
}
