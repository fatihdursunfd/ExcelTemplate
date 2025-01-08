namespace ExcelTemplate.Helpers
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Belirtilen sütun sayısına göre excel tarzı sütun adlarının bir listesini oluşturur.
        /// Sütun adları, excel'deki formatla benzer şekilde temsil edilir (A, B, C, ..., Z, AA, AB, ..., AZ, BA, ...).
        /// </summary>
        /// <param name="numberOfColumns">Oluşturulacak toplam sütun adı sayısı.</param>
        /// <returns>Oluşturulan sütun adlarının bir listesini döndürür.</returns>
        public static IEnumerable<string> GenerateExcelColumnNames(int numberOfColumns)
        {
            for (int i = 0; i < numberOfColumns; i++)
            {
                int dividend = i;
                string columnName = string.Empty;
                while (dividend >= 0)
                {
                    int modulo = dividend % 26;
                    columnName = (char)(modulo + 65) + columnName;
                    dividend = (dividend / 26) - 1;
                }
                yield return columnName;
            }
        }
    }
}