using System.Collections.Generic;
using System.Linq;

namespace Spectrum.SpLoader
{
    static class ArrayExtends
    {
        public static string[] CopyFilled(this string[] array) => (from item in array
                                                                  where !string.IsNullOrEmpty(item)
                                                                  select item).ToArray();
        public static Stack<string> ToStack(this string[] array)
        {
            Stack<string> result = new Stack<string>();
            for(int i = array.GetUpperBound(0); i > -1; i--)
                result.Push(array[i]);
            return result;
        }
    }
}
