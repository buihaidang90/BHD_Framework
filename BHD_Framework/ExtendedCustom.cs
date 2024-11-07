using System;
using System.Text.RegularExpressions;

namespace BHD_Framework
{
    public static class StringExtension // Extension methods must be defined in a static class.
    {
        // This is the extension method.
        // The first parameter takes the "this" modifier
        // and specifies the type for which the method is defined.
        public static string TrimAndReduce(this string str)
        {
            return ConvertWhiteSpacesToSingleSpace(str).Trim();
        }

        public static string ConvertWhiteSpacesToSingleSpace(this string value)
        {
            return Regex.Replace(value, @"\s+", " ");
        }

        public static string IsNullOrEmptyReturn(this string value, string ExpectedValue)
        {
            if (String.IsNullOrEmpty(value.Trim())) return ExpectedValue;
            return value;
        }


    }
}

/* Usage example:
using BHD_Framework; // Import the extension method namespace.
namespace Extension_Methods_Simple
{    
    class Program
    {
        static void Main(string[] args)
        {
            string text = "     I'm    wearing the   cheese.  It   isn't   wearing     me!      ";
            text = text.TrimAndReduce();
            System.Console.WriteLine("Word count of s is {0}", i);
            // output: [I'm wearing the cheese. It isn't wearing me!]
            
            System.Console.WriteLine("The value null or empty will return is [" + "".IsNullOrEmptyReturn("%") + "]");
        }
    }
}
*/
