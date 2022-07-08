using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LMSBot
{
    public static class ExtensionMethods
    {
        public static string Right(this string input, int count)
        {
            return input.Substring(Math.Max(input.Length - count, 0), Math.Min(count, input.Length));
        }

        public static string Left(this string input, int count)
        {
            return input.Substring(0, Math.Min(input.Length, count));
        }
    }
}
