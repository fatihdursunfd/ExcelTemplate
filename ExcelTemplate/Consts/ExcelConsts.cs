namespace ExcelTemplate.Consts
{
    public static class ExcelConsts
    {
        // Excel'de izin verilen en büyük pozitif ve negatif sayılar.
        // Referans : https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3

        public const double LargestAllowedPositiveNumber = 9.99999999999999E+307;
        public const double LargestAllowedNegativeNumber = -9.99999999999999E+307;


        public const int MaxFormulaLength = 255;


        // Hücrelere biçimlendirme ekleme. Default formatlar:

        // Tam sayı biçimlendirme   : #,##0
        public const string IntegerFormat = "#,##0";

        // Sayı biçimlendirme       : #,##0.00
        public const string DecimalFormat = "#,##0.00";

        // Tarih biçimlendirme      : dd/MM/yyyy
        public const string DateFormat = "dd/MM/yyyy";

        // Metin biçimlendirme      : @
        public const string TextFormat = "@";

        // Para biçimlendirme       : $#,##0.00
        public const string CurrencyFormat = "$#,##0.00";
    }
}