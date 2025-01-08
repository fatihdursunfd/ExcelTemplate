using ExcelTemplate.Enums;
using ExcelTemplate.Helpers;
using ExcelTemplate.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelTemplate.Services
{
    public class ExcelUtilities : IDisposable, IExcelUtilities
    {
        public ExcelWorksheet WorkSheet;
        public ExcelPackage Excel;

        // Dispose metodunun birden fazla kez çağrılmasını önlemek için
        private bool disposed = false;


        /// <summary>
        /// Varsayılan bir çalışma sayfası ("sheet1") ile yeni bir ExcelUtilities nesnesi oluşturur.
        /// </summary>
        public ExcelUtilities() : this("sheet1") { }



        /// <summary>
        /// Belirtilen isimle yeni bir ExcelUtilities nesnesi oluşturur.
        /// </summary>
        /// <param name="sheetName">Oluşturulacak çalışma sayfasının adı.</param>
        public ExcelUtilities(string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Excel = new();

            WorkSheet = Excel.Workbook.Worksheets.Add(sheetName);
        }



        /// <summary>
        /// ExcelUtilities sınıfının kaynaklarını serbest bırakmak için kullanılır.
        /// Bu metod, ExcelWorksheet ve ExcelPackage nesnelerini düzgün bir şekilde dispose eder.
        /// Dispose pattern uygulanmıştır. bknz : https://docs.microsoft.com/en-us/dotnet/standard/garbage-collection/implementing-dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }



        /// <summary>
        /// ExcelUtilities sınıfının kaynaklarını serbest bırakmak için kullanılır.
        /// Bu metod, yönetilen ve yönetilmeyen kaynakları düzgün bir şekilde dispose eder.
        /// </summary>
        ~ExcelUtilities()
        {
            Dispose(false);
        }



        /// <summary>
        /// Oluşturulan excel dosyasını belirtilen dosya adı ve yol ile kaydetmek için kullanılır.
        /// </summary>
        /// <param name="filename">Kaydedilecek dosyanın adı.</param>
        /// <param name="path">Dosyanın kaydedileceği yol.</param>
        public void Save(string filename, string path)
        {
            var fullPath = @$"{path}{filename}.xlsx";

            var file = File.Create(fullPath);

            file.Close();

            File.WriteAllBytes(fullPath, Excel.GetAsByteArray());
        }



        /// <summary>
        /// Excel dosyasını byte dizisi olarak döndürmek için kullanılır.
        /// Bu metod, oluşturulan Excel dosyasının içeriğini byte dizisi formatında almanızı sağlar.
        /// </summary>
        /// <returns>Excel dosyasının byte dizisi.</returns>
        public byte[] GetExcelAsByteArray()
        {
            return Excel.GetAsByteArray();
        }


        /// <summary>
        /// Yeni bir sayfa eklemek için kullanılır.
        /// Sayfanın adı belirtilir ve isteğe bağlı olarak sayfanın görünürlüğü ayarlanabilir.
        /// </summary>
        /// <param name="sheetName">Eklenecek sayfanın adı.</param>
        /// <param name="isHidden">Sayfanın gizli olup olmayacağı (varsayılan: false).</param>
        /// <returns>Eklenen ExcelWorksheet nesnesi.</returns>
        public ExcelWorksheet AddSheet(string sheetName, bool isHidden = false)
        {
            var sheet = Excel.Workbook.Worksheets.Add(sheetName);

            sheet.Hidden = isHidden ? eWorkSheetHidden.Hidden : eWorkSheetHidden.Visible;

            return sheet;
        }


        /// <summary>
        /// Excel sayfasının temel stil ayarlarını yapmak için kullanılır.
        /// Tab rengi, varsayılan satır yüksekliği, ilk satır yüksekliği,
        /// yatay ve dikey hizalama ile yazı tipi kalınlığı gibi özellikleri ayarlamak için kullanılır.
        /// <param name="WorkSheet">Stil ayarlarının uygulanacağı Excel çalışma sayfası.</param>
        /// <param name="color">Tabın rengi (varsayılan: siyah).</param>
        /// <param name="defaultRowHeight">Varsayılan satır yüksekliği (varsayılan: 12).</param>
        /// <param name="height">İlk satırın yüksekliği (varsayılan: 20).</param>
        /// <param name="horizontalAlignment">İlk satırın yatay hizalaması (varsayılan: Center).</param>
        /// <param name="verticalAlignment">İlk satırın dikey hizalaması (varsayılan: Center).</param>
        /// <param name="isBold">Yazı tipinin kalın olup olmayacağı (varsayılan: false).</param>
        /// </summary>
        public void ApplyDefaultStyling(ExcelWorksheet WorkSheet,
                               Color? color = null,
                               int defaultRowHeight = 12,
                               int height = 20,
                               ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center,
                               ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Center,
                               bool isBold = false)
        {
            WorkSheet.TabColor = color ?? Color.Black;
            WorkSheet.DefaultRowHeight = defaultRowHeight;
            WorkSheet.Row(1).Height = height;
            WorkSheet.Row(1).Style.HorizontalAlignment = horizontalAlignment;
            WorkSheet.Row(1).Style.VerticalAlignment = verticalAlignment;
            WorkSheet.Row(1).Style.Font.Bold = isBold;
        }



        /// <summary>
        /// Belirtilen hücreye stil uygulamak için kullanılır.
        /// Hücreye yatay ve dikey hizalama, yazı tipi kalınlığı ve arka plan rengi gibi stil özellikleri ayarlanır.
        /// </summary>
        /// <param name="workSheet">Stilin uygulanacağı Excel çalışma sayfası.</param>
        /// <param name="row">Stilin uygulanacağı hücrenin satır numarası.</param>
        /// <param name="column">Stilin uygulanacağı hücrenin sütun numarası.</param>
        /// <param name="horizontalAlignment">Hücrenin yatay hizalaması.</param>
        /// <param name="verticalAlignment">Hücrenin dikey hizalaması.</param>
        /// <param name="isBold">Yazı tipinin kalın olup olmayacağı.</param>
        /// <param name="backgroundColor">Hücrenin arka plan rengi (isteğe bağlı).</param>
        public void ApplyStyleToCell(ExcelWorksheet workSheet, 
                                    int row, int column, 
                                    ExcelHorizontalAlignment horizontalAlignment, 
                                    ExcelVerticalAlignment verticalAlignment, 
                                    bool isBold, 
                                    Color? backgroundColor = null)
        {
            var cell = workSheet.Cells[row, column];

            cell.Style.HorizontalAlignment = horizontalAlignment;
            cell.Style.VerticalAlignment = verticalAlignment;
            cell.Style.Font.Bold = isBold;

            if (backgroundColor != null)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor((Color)backgroundColor);
            }
        }



        /// <summary>
        /// Belirtilen excel çalışma sayfasına, verilen sütun adlarını eklemek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Sütun adlarının ekleneceği Excel çalışma sayfası.</param>
        /// <param name="columnNames">Eklenecek sütun adlarını içeren dizi.</param>
        public void AddColumns(ExcelWorksheet workSheet, string[] columnNames)
        {
            for (int i = 0; i < columnNames.Length; i++)
            {
                workSheet.Cells[1, i + 1].Value = columnNames[i];
            }
        }



        /// <summary>
        /// Belirtilen excel çalışma sayfasına, verilen sütun adlarını eklemek için kullanılır.
        /// Sütun adları, bir Dictionary kullanılarak belirtilir; anahtar sütun numarasını, değer ise sütun adını temsil eder.
        /// </summary>
        /// <param name="workSheet">Sütun adlarının ekleneceği Excel çalışma sayfası.</param>
        /// <param name="columns">Eklenecek sütun adlarını içeren sözlük.</param>
        public void AddColumns(ExcelWorksheet workSheet, Dictionary<int, string> columns)
        {
            foreach (var column in columns)
            {
                workSheet.Cells[1, column.Key].Value = column.Value;
            }
        }



        /// <summary>
        /// Belirtilen Excel çalışma sayfasına, verilen hücre aralığında belirli bir doğrulama türünü uygulamak için kullanılır.
        /// Kullanıcıdan alınan verilerin geçerliliğini kontrol etmek amacıyla e-posta, tam sayı, tarih, ondalık, zaman ve metin uzunluğu gibi doğrulama türleri desteklenmektedir.
        /// </summary>
        /// <param name="worksheet">Doğrulamanın uygulanacağı Excel çalışma sayfası.</param>
        /// <param name="range">Doğrulamanın uygulanacağı hücre aralığı.</param>
        /// <param name="validationType">Uygulanacak doğrulama türü.</param>
        /// <param name="values">(isteğe bağlı) Doğrulama türü için kullanılacak değerler dizisi.</param>
        /// <param name="customFormula">(isteğe bağlı) Özel doğrulama için kullanılacak formül.</param>
        public void ApplyValidation(ExcelWorksheet worksheet, string range, ValidationTypes validationType, string[]? values = null, string? customFormula = null)
        {
            switch (validationType)
            {
                case ValidationTypes.Email:
                    ValidationHelper.CreateEmailValidation(worksheet, range);
                    break;

                case ValidationTypes.Integer:
                    ValidationHelper.CreateIntegerValidation(worksheet, range);
                    break;

                case ValidationTypes.Date:
                    ValidationHelper.CreateDateValidation(worksheet, range);
                    break;

                case ValidationTypes.Decimal:
                    ValidationHelper.CreateDecimalValidation(worksheet, range);
                    break;

                case ValidationTypes.Time:
                    ValidationHelper.CreateTimeValidation(worksheet, range);
                    break;

                case ValidationTypes.TextLength:
                    ValidationHelper.CreateTextLengthValidation(worksheet, range);
                    break;

                case ValidationTypes.List:
                    ValidationHelper.CreateListValidation(worksheet, range, values);
                    break;

                case ValidationTypes.Custom:
                    ValidationHelper.CreateCustomValidation(worksheet, range, customFormula);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(validationType), validationType, null);

            }
        }



        /// <summary>
        /// Bir hücreye yorum eklemek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Yorumun ekleneceği çalışma sayfası.</param>
        /// <param name="row">Yorumun ekleneceği hücrenin satır numarası.</param>
        /// <param name="column">Yorumun ekleneceği hücrenin sütun numarası.</param>
        /// <param name="comment">Eklenecek yorum metni.</param>
        /// <param name="author">Yorumun yazarı.</param>
        public void AddComment(ExcelWorksheet workSheet, int row, int column, string comment, string author)
        {
            var cell = workSheet.Cells[row, column];

            cell.AddComment(comment, author);
        }



        /// <summary>
        /// Bir sütuna açılır liste(dropdown) eklemek için kullanılır.
        /// Açılır liste, verilen değerler dizisinden oluşturulur ve belirtilen hücre aralığına uygulanır.
        /// Excel'deki formül karakter uzunluğu kısıtı sebebiyle, eğer formül 255 karakterden uzunsa, değerler yeni bir sayfada tanımlanır.
        /// </summary>
        /// <param name="WorkSheet">Açılır listenin ekleneceği çalışma sayfası.</param>
        /// <param name="values">Açılır listede gösterilecek değerler.</param>
        /// <param name="cell">Açılır listenin uygulanacağı hücre adresi.</param>
        /// <param name="minRow">Açılır listenin uygulanacağı minimum satır numarası.</param>
        /// <param name="maxRow">Açılır listenin uygulanacağı maksimum satır numarası.</param>
        public void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow)
        {
            var formula = "\"" + string.Join(",", values) + "\"";

            if (formula.Length > 255)
            {
                AddDropdownList(WorkSheet, values, cell, minRow, maxRow, $"{cell}_values");
                return;
            }

            var range = WorkSheet.Cells[$"{cell}{minRow}:{cell}{maxRow}"];

            var validation = range.DataValidation.AddListDataValidation();

            validation.Formula.ExcelFormula = formula;
        }



        /// <summary>
        /// Bir sütuna açılır liste(dropdown) eklemek için kullanılır.
        /// Açılır liste, verilen değerler dizisinden oluşturulur ve belirtilen hücre aralığına uygulanır.
        /// Ayrıca, açılır listenin değerlerinin yer alacağı yeni bir sayfa oluşturulur.
        /// </summary>
        /// <param name="WorkSheet">Açılır listenin ekleneceği çalışma sayfası.</param>
        /// <param name="values">Açılır listede gösterilecek değerler.</param>
        /// <param name="cell">Açılır listenin uygulanacağı hücre adresi.</param>
        /// <param name="minRow">Açılır listenin uygulanacağı minimum satır numarası.</param>
        /// <param name="maxRow">Açılır listenin uygulanacağı maksimum satır numarası.</param>
        /// <param name="sheetName">Değerlerin yer alacağı yeni sayfanın ismi.</param>
        public void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow, string sheetName)
        {
            var worksheet = AddSheet(sheetName);

            for (int i = 0; i < values.Length; i++)
            {
                worksheet.Cells[i + 1, 1].Value = values[i];
            }

            var range = WorkSheet.Cells[$"{cell}{minRow}:{cell}{maxRow}"];

            var validation = range.DataValidation.AddListDataValidation();

            validation.Formula.ExcelFormula = $"={sheetName}!$A$1:$A${values.Length}";
        }



        /// <summary>
        /// Birbirine bağımlı dropdownlar oluşturmak için kullanılır.
        /// Bu metod, belirtilen çalışma sayfasında iki hücre arasında bağımlı açılır listeler oluşturur.
        /// İlk hücrede seçilen değere göre ikinci hücredeki seçenekler dinamik olarak güncellenir.
        /// <param name="worksheet">Açılır listelerin oluşturulacağı çalışma sayfası.</param>
        /// <param name="data">İlk hücredeki ve ona bağımlı ikinci hücredeki seçime bağlı olarak gösterilecek değerlerin listesi.</param>
        /// <param name="firstCell">İlk açılır listenin hücre adresi.</param>
        /// <param name="secondCell">İkinci açılır listenin hücre adresi.</param>
        /// <param name="minRow">Açılır listelerin uygulanacağı minimum satır numarası.</param>
        /// <param name="maxRow">Açılır listelerin uygulanacağı maksimum satır numarası.</param>
        /// </summary>
        public void CreateDependentDropdowns(ExcelWorksheet worksheet, Dictionary<string, List<string>> data, string firstCell, string secondCell, int minRow, int maxRow)
        {
            var sourceSheet = AddSheet(nameof(data), true);

            int columnIndex = 1;

            var excelColumnNames = ExcelHelper.GenerateExcelColumnNames(data.Count).ToList();

            foreach (var entry in data)
            {
                sourceSheet.Cells[1, columnIndex].Value = entry.Key;

                DefineNamedRange(sourceSheet, entry.Key, $"{excelColumnNames[columnIndex - 1]}{2}:{excelColumnNames[columnIndex - 1]}{entry.Value.Count + 1}");

                for (int rowIndex = 0; rowIndex < entry.Value.Count; rowIndex++)
                {
                    sourceSheet.Cells[rowIndex + 2, columnIndex].Value = entry.Value[rowIndex];
                }
                columnIndex++;
            }

            AddDropdownList(worksheet, nameof(data), data.Count, firstCell, secondCell, minRow, maxRow);
        }



        /// <summary>
        /// Belirtilen çalışma sayfasında bir hücre aralığına ad tanımlamak için kullanılır.
        /// <param name="workSheet">Adın tanımlanacağı çalışma sayfası.</param>
        /// <param name="rangeName">Tanımlanacak adın ismi.</param>
        /// <param name="cellRange">Adın uygulanacağı hücre aralığı.</param>
        /// </summary>
        public void DefineNamedRange(ExcelWorksheet workSheet, string rangeName, string cellRange)
        {
            Excel.Workbook.Names.Add(rangeName, workSheet.Cells[cellRange]);
        }



        /// <summary>
        /// Otomatik olarak sütunların genişliğini ayarlamak için kullanılır.
        /// Minimum genişlik 10, maksimum genişlik 100 olarak ayarlanmıştır.
        /// <param name="workSheet">AutoFit'in uygulanacağı excel çalışma sayfası.</param>
        /// </summary>
        public void SetAutoFit(ExcelWorksheet WorkSheet)
        {
            if (WorkSheet.Dimension is null)
            {
                return;
            }

            int columnCount = WorkSheet.Dimension.End.Column;

            for (int i = 1; i <= columnCount; i++)
            {
                WorkSheet.Column(i).AutoFit(10, 100);
            }
        }



        /// <summary>
        /// Bütün excelin background rengini setlemek için kullanılır.
        /// <param name="workSheet">Çalışılacak olan excel çalışma sayfası.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        /// </summary>
        public void SetBackGroundColor(ExcelWorksheet workSheet, Color color)
        {
            if (workSheet.Dimension is null)
            {
                return;
            }

            int rowCount = workSheet.Dimension.End.Row;

            for (int i = 0; i <= rowCount; i++)
            {
                workSheet.Row(i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Row(i).Style.Fill.BackgroundColor.SetColor(color);
            }
        }


        /// <summary>
        /// Excelde istenilen satır ve/veya sütunun background rengini setlemek için kullanılır
        /// Eğer bir satır numarası sağlanırsa, o satırın arka plan rengi ayarlanır.
        /// Eğer bir sütun numarası sağlanırsa, o sütunun arka plan rengi ayarlanır.
        /// Her ikisi de sağlanırsa, her ikisi de ayarlanır. Hiçbiri sağlanmazsa, hiçbir işlem yapılmaz.
        /// <param name="workSheet">Çalışılacak olan excel çalışma sayfası.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        /// <param name="rowNumber">Arka plan renginin ayarlanacağı satır numarası (isteğe bağlı).</param>
        /// <param name="columnNumber">Arka plan renginin ayarlanacağı sütun numarası (isteğe bağlı).</param>
        /// </summary>
        public void SetBackGroundColor(ExcelWorksheet workSheet, Color color, int? rowNumber = null, int? columnNumber = null)
        {

            if (workSheet.Dimension is null)
            {
                return;
            }

            int columnCount = workSheet.Dimension.End.Column;
            int rowCount = workSheet.Dimension.End.Row;

            if (rowNumber != null && rowNumber <= rowCount)
            {
                workSheet.Row((int)rowNumber).Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Row((int)rowNumber).Style.Fill.BackgroundColor.SetColor(color);
            }

            if (columnNumber != null && columnNumber <= columnCount)
            {
                workSheet.Column((int)columnNumber).Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Column((int)columnNumber).Style.Fill.BackgroundColor.SetColor(color);
            }
        }



        /// <summary>
        /// Bir hücreye bir değer yazmak için kullanılır.
        /// <param name="workSheet">Değerin yazılacağı çalışma sayfası.</param>
        /// <param name="row">Hücrenin satır numarası.</param>
        /// <param name="column">Hücrenin sütun numarası.</param>
        /// </summary>
        public void WriteCell(ExcelWorksheet workSheet, int row, int column, object value)
        {
            workSheet.Cells[row, column].Value = value;
        }



        /// <summary>
        /// Bir hücredeki değeri temizlemek için kullanılır.
        /// <param name="workSheet">Değerin temizleneceği çalışma sayfası.</param>
        /// <param name="row">Hücrenin satır numarası.</param>
        /// <param name="column">Hücrenin sütun numarası.</param>
        /// </summary>
        public void ClearCell(ExcelWorksheet workSheet, int row, int column)
        {
            workSheet.Cells[row, column].Clear();
        }



        /// <summary>
        /// Belirtilen hücredeki değeri okumak için kullanılır.
        /// </summary>
        /// <param name="workSheet">Değerin okunacağı çalışma sayfası.</param>
        /// <param name="row">Hücrenin satır numarası.</param>
        /// <param name="column">Hücrenin sütun numarası.</param>
        /// <returns>Okunan hücre değeri.</returns>
        public object ReadCell(ExcelWorksheet workSheet, int row, int column)
        {
            return workSheet.Cells[row, column].Value;
        }

        /// <summary>
        /// Belirtilen dosya yolundaki Excel dosyasını okuyup, çalışma sayfasını döndürmek için kullanılır.
        /// </summary>
        /// <param name="filePath">Okunacak Excel dosyasının tam yolu.</param>
        /// <param name="sheetIndex">Dönmek istenen çalışma sayfasının sırası (0 tabanlı). Null ise ilk sayfa döner.</param>
        /// <returns>Okunan ExcelWorksheet nesnesi.</returns>
        /// <exception cref="FileNotFoundException">Dosya bulunamadığında fırlatılır.</exception>
        /// <exception cref="InvalidOperationException">Geçersiz bir çalışma sayfası indeksine erişilmeye çalışıldığında fırlatılır.</exception>
        public ExcelWorksheet ReadWorksheet(string filePath, int? sheetIndex = null)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Belirtilen dosya bulunamadı.", filePath);
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[sheetIndex ?? 0];
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Çalışma sayfası bulunamadı. İndeks: {sheetIndex}");
                }
                return worksheet; // Belirtilen sayfayı döndür veya ilk sayfayı döndür
            }
        }


        /// <summary>
        /// Belirtilen hücre aralığını birleştirmek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Birleştirilecek hücrelerin bulunduğu çalışma sayfası.</param>
        /// <param name="fromRow">Birleştirmenin başlayacağı satır numarası.</param>
        /// <param name="fromColumn">Birleştirmenin başlayacağı sütun numarası.</param>
        /// <param name="toRow">Birleştirmenin biteceği satır numarası.</param>
        /// <param name="toColumn">Birleştirmenin biteceği sütun numarası.</param>
        public void MergeCells(ExcelWorksheet workSheet, int fromRow, int fromColumn, int toRow, int toColumn)
        {
            workSheet.Cells[fromRow, fromColumn, toRow, toColumn].Merge = true;
        }



        /// <summary>
        /// Belirtilen çalışma sayfasını bir şifre ile korur, yetkisiz değişiklikleri engeller.
        /// </summary>
        /// <param name="workSheet">Korunacak çalışma sayfası.</param>
        /// <param name="password">Koruma için ayarlanacak şifre.</param>
        public void ProtectSheet(ExcelWorksheet workSheet, string password)
        {
            workSheet.Protection.IsProtected = true;
            workSheet.Protection.SetPassword(password);
        }



        /// <summary>
        /// Çalışma sayfasındaki belirli bir hücreye bir formül eklemek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Formülün ekleneceği çalışma sayfası.</param>
        /// <param name="row">Hücrenin satır numarası.</param>
        /// <param name="column">Hücrenin sütun numarası.</param>
        /// <param name="formula">Hücreye eklenecek formül.</param>
        public void AddFormula(ExcelWorksheet workSheet, int row, int column, string formula)
        {
            workSheet.Cells[row, column].Formula = formula;
        }



        /// <summary>
        /// Belirli bir hücre aralığına formüle dayalı koşullu biçimlendirme eklemek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Koşullu biçimlendirmenin uygulanacağı çalışma sayfası.</param>
        /// <param name="address">Biçimlendirilecek hücre aralığının adresi.</param>
        /// <param name="formula">Biçimlendirme koşulunu belirleyen formül.</param>
        /// <param name="color">Koşul karşılandığında uygulanacak arka plan rengi.</param>
        public void AddConditionalFormatting(ExcelWorksheet workSheet, string address, string formula, Color color)
        {
            var conditionalFormattingRule = workSheet.ConditionalFormatting.AddExpression(workSheet.Cells[address]);
            conditionalFormattingRule.Formula = formula;
            conditionalFormattingRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormattingRule.Style.Fill.BackgroundColor.Color = color;
        }



        /// <summary>
        /// Belirtilen çalışma sayfasındaki belirli bir hücreyi, belirtilen formatla biçimlendirmek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Biçimlendirilecek hücreyi içeren çalışma sayfası.</param>
        /// <param name="row">Biçimlendirilecek hücrenin satır numarası.</param>
        /// <param name="column">Biçimlendirilecek hücrenin sütun numarası.</param>
        /// <param name="format">Hücreye uygulanacak format.</param>
        public void FormatCell(ExcelWorksheet workSheet, int row, int column, string format)
        {
            var cell = workSheet.Cells[row, column];
            cell.Style.Numberformat.Format = format;
        }



        /// <summary>
        /// Belirtilen çalışma sayfasındaki belirli bir sütundaki tüm hücreleri, belirtilen formatla biçimlendirmek için kullanılır.
        /// </summary>
        /// <param name="workSheet">Biçimlendirilecek sütunu içeren çalışma sayfası.</param>
        /// <param name="column">Biçimlendirilecek sütun numarası.</param>
        /// <param name="format">Sütuna uygulanacak format.</param>
        public void FormatCell(ExcelWorksheet workSheet, int column, string format)
        {
            var columnCells = workSheet.Cells[1, column, workSheet.Dimension.End.Row, column];
            columnCells.Style.Numberformat.Format = format;
        }



        /// <summary>
        /// Birbirine bağımlı dropdownlar oluşturmak için kullanılır
        /// <param name="WorkSheet">Açılır listelerin oluşturulacağı çalışma sayfası.</param>
        /// <param name="sourceSheet">Açılır listenin kaynak değerlerinin bulunduğu sayfanın adı.</param>
        /// <param name="length">Kaynak değerlerin sayısı.</param>
        /// <param name="firstCell">İlk açılır listenin hücre adresi.</param>
        /// <param name="secondCell">İkinci açılır listenin hücre adresi.</param>
        /// <param name="minRow">Açılır listelerin uygulanacağı minimum satır numarası.</param>
        /// <param name="maxRow">Açılır listelerin uygulanacağı maksimum satır numarası.</param>
        /// </summary>
        void AddDropdownList(ExcelWorksheet WorkSheet, string sourceSheet, int length, string firstCell, string secondCell, int minRow, int maxRow)
        {
            var range = WorkSheet.Cells[$"{firstCell}{minRow}:{firstCell}{maxRow}"];

            var validation = range.DataValidation.AddListDataValidation();

            var excelColumnNames = ExcelHelper.GenerateExcelColumnNames(length);

            validation.Formula.ExcelFormula = $"={sourceSheet}!$A$1:${excelColumnNames.Last()}$1";

            var secondrange = WorkSheet.Cells[$"{secondCell}{minRow}:{secondCell}{maxRow}"];
            var secondValidation = secondrange.DataValidation.AddListDataValidation();

            secondValidation.Formula.ExcelFormula = $"=INDIRECT({firstCell}2)";
        }



        /// <summary>
        /// Dispose metodunun birden fazla kez çağrılmasını önlemek için kullanılan koruyucu metod.
        /// </summary>
        /// <param name="disposing">Yönetilen kaynakların serbest bırakılıp bırakılmayacağını belirten boolean değer.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Yönetilen kaynakları serbest bırak
                    WorkSheet.Dispose();
                    Excel.Dispose();
                }

                // Yönetilmeyen kaynakları serbest bırak (varsa)

                disposed = true;
            }
        }


    }
}