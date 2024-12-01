using Humanizer;
using System.Globalization;
using ExcelDna.Integration;
using System.Linq;
public static class ExcelFunctions
{
    [ExcelFunction(
        Name = "USD",
        Description = "Chuyển số thành chữ tiếng Anh, bao gồm cả phần thập phân.",
        Category = "Custom Functions"
    )]
    public static string ToEnglishWords(
        [ExcelArgument(Name = "Number", Description = "Số cần chuyển đổi sang chữ (tiếng Anh).")] double number)
    {
        return ConvertNumberToWords(number, "en");
    }

    [ExcelFunction(
        Name = "VND",
        Description = "Chuyển số thành chữ tiếng Việt, bao gồm cả phần thập phân.",
        Category = "Custom Functions"
    )]
    public static string ToVietnameseWords(
        [ExcelArgument(Name = "Number", Description = "Số cần chuyển đổi sang chữ (tiếng Việt).")] double number)
    {
        return ConvertNumberToWords(number, "vi");
    }

    private static string ConvertNumberToWords(double number, string cultureCode)
    {
        long integerPart = (long)number; // Lấy phần nguyên
        string integerWords = integerPart.ToWords(new CultureInfo(cultureCode)).Transform(To.SentenceCase);

        string decimalSeparator = cultureCode == "vi" ? " phẩy " : " point ";

        // Lấy phần thập phân và chuyển từng chữ số thành số nguyên trước khi đọc
        string decimalPart = number.ToString("F99").Split('.')[1].TrimEnd('0');
        string decimalWords = string.Join(" ", decimalPart.Select(digit => int.Parse(digit.ToString()).ToWords(new CultureInfo(cultureCode))));

        return decimalWords.Length > 0
            ? integerWords + decimalSeparator + decimalWords
            : integerWords;
    }
}
