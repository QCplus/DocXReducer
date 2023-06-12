using System;
using System.Linq;

namespace DocxReducer.Helpers
{
    internal static class IntExtensions
    {
        public const string CHARS_FOR_BASE_CONVERSION = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GetLastNDigits(this int number, int n)
        {
            if (n <= 0)
                return "";

            var digits = Math.Abs(number).ToString();

            return digits.Substring(Math.Max(digits.Length - n, 0));
        }

        public static string ToBase(this int number, int targetBase)
        {
            if (targetBase <= 1 || targetBase > CHARS_FOR_BASE_CONVERSION.Length)
                throw new Exception($"Invalid base: {targetBase}");

            var result = "";
            int divisionResult = Math.Abs(number);
            do
            {
                var mod = divisionResult % targetBase;

                result = CHARS_FOR_BASE_CONVERSION[mod] + result;

                divisionResult /= targetBase;

            } while (divisionResult > 0);

            return (number < 0 ? "-" : "") + result;
        }

        public static int TakeFirstBits(this int number, int bitsToTake)
        {
            int mask = Convert.ToInt32(string.Join("", new byte[Math.Min(bitsToTake, 32)].Select(t => 1)), 2);

            return number & mask;
        }
    }
}
