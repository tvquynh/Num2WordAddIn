using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn2
{
    using Humanizer;
    using System.Globalization;

    public static class NumberToWordsHelper
    {
        // Hàm đọc số thành chữ tiếng Anh
        public static string ToEnglishWords(Int32 number)
        {
            return number.ToWords().Transform(To.SentenceCase);
        }

        // Hàm đọc số thành chữ tiếng Việt
        public static string ToVietnameseWords(Int32 number)
        {
            return number.ToWords(new CultureInfo("vi")).Transform(To.SentenceCase);
        }
    }

}
