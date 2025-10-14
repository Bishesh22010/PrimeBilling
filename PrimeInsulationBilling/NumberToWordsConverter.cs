using System;

namespace PrimeInsulationBilling.Services
{
    public static class NumberToWordsConverter
    {
        private static readonly string[] Ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
        private static readonly string[] Tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

        public static string ToIndianCurrencyWords(decimal number)
        {
            if (number == 0)
                return "Zero Rupees Only";

            long integerPart = (long)number;
            int decimalPart = (int)((number - integerPart) * 100);

            string words = "";
            words += ConvertToWords(integerPart) + " Rupees";

            if (decimalPart > 0)
            {
                words += " and " + ConvertToWords(decimalPart) + " Paise";
            }

            return words + " Only";
        }

        private static string ConvertToWords(long number)
        {
            if (number == 0) return "";

            string words = "";

            if ((number / 10000000) > 0)
            {
                words += ConvertToWords(number / 10000000) + " Crore ";
                number %= 10000000;
            }

            if ((number / 100000) > 0)
            {
                words += ConvertToWords(number / 100000) + " Lakh ";
                number %= 100000;
            }

            if ((number / 1000) > 0)
            {
                words += ConvertToWords(number / 1000) + " Thousand ";
                number %= 1000;
            }

            if ((number / 100) > 0)
            {
                words += ConvertToWords(number / 100) + " Hundred ";
                number %= 100;
            }

            if (number > 0)
            {
                if (words != "")
                    words += "and ";

                if (number < 20)
                    words += Ones[number];
                else
                {
                    words += Tens[number / 10];
                    if ((number % 10) > 0)
                        words += " " + Ones[number % 10];
                }
            }

            return words.Trim();
        }
    }
}
