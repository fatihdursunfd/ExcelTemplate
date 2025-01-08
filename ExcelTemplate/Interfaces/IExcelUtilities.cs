using ExcelTemplate.Enums;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelTemplate.Interfaces
{
    public interface IExcelUtilities : IDisposable
    {
        /// <summary>
        /// Oluþturulan excel dosyasýný belirtilen dosya adý ve yol ile kaydetmek için kullanýlýr.
        /// </summary>
        /// <param name="filename">Kaydedilecek dosyanýn adý.</param>
        /// <param name="path">Dosyanýn kaydedileceði yol.</param>
        void Save(string filename, string path);


        /// <summary>
        /// Excel dosyasýný byte dizisi olarak döndürmek için kullanýlýr.
        /// Bu metod, oluþturulan Excel dosyasýnýn içeriðini byte dizisi formatýnda almanýzý saðlar.
        /// </summary>
        /// <returns>Excel dosyasýnýn byte dizisi.</returns>
        byte[] GetExcelAsByteArray();


        /// <summary>
        /// Excel sayfasýnýn temel stil ayarlarýný yapmak için kullanýlýr.
        /// Tab rengi, varsayýlan satýr yüksekliði, ilk satýr yüksekliði,
        /// yatay ve dikey hizalama ile yazý tipi kalýnlýðý gibi özellikleri ayarlamak için kullanýlýr.
        /// </summary>
        /// <param name="WorkSheet">Stil ayarlarýnýn uygulanacaðý Excel çalýþma sayfasý.</param>
        /// <param name="color">Tabýn rengi (varsayýlan: siyah).</param>
        /// <param name="defaultRowHeight">Varsayýlan satýr yüksekliði (varsayýlan: 12).</param>
        /// <param name="height">Ýlk satýrýn yüksekliði (varsayýlan: 20).</param>
        /// <param name="horizontalAlignment">Ýlk satýrýn yatay hizalamasý (varsayýlan: Center).</param>
        /// <param name="verticalAlignment">Ýlk satýrýn dikey hizalamasý (varsayýlan: Center).</param>
        /// <param name="isBold">Yazý tipinin kalýn olup olmayacaðý (varsayýlan: false).</param>
        void ApplyDefaultStyling(ExcelWorksheet WorkSheet, Color? color = null, int defaultRowHeight = 12, int height = 20, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center, ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Center, bool isBold = false);


        /// <summary>
        /// Belirtilen hücreye stil uygulamak için kullanýlýr.
        /// Hücreye yatay ve dikey hizalama, yazý tipi kalýnlýðý ve arka plan rengi gibi stil özellikleri ayarlanýr.
        /// </summary>
        /// <param name="workSheet">Stilin uygulanacaðý Excel çalýþma sayfasý.</param>
        /// <param name="row">Stilin uygulanacaðý hücrenin satýr numarasý.</param>
        /// <param name="column">Stilin uygulanacaðý hücrenin sütun numarasý.</param>
        /// <param name="horizontalAlignment">Hücrenin yatay hizalamasý.</param>
        /// <param name="verticalAlignment">Hücrenin dikey hizalamasý.</param>
        /// <param name="isBold">Yazý tipinin kalýn olup olmayacaðý.</param>
        /// <param name="backgroundColor">Hücrenin arka plan rengi (isteðe baðlý).</param>
        void ApplyStyleToCell(ExcelWorksheet workSheet, int row, int column, ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment, bool isBold, Color? backgroundColor = null);


        /// <summary>
        /// Yeni bir sayfa eklemek için kullanýlýr.
        /// Sayfanýn adý belirtilir ve isteðe baðlý olarak sayfanýn görünürlüðü ayarlanabilir.
        /// </summary>
        /// <param name="sheetName">Eklenecek sayfanýn adý.</param>
        /// <param name="isHidden">Sayfanýn gizli olup olmayacaðý (varsayýlan: false).</param>
        /// <returns>Eklenen ExcelWorksheet nesnesi.</returns>
        ExcelWorksheet AddSheet(string sheetName, bool isHidden = false);


        /// <summary>
        /// Belirtilen excel çalýþma sayfasýna, verilen sütun adlarýný eklemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Sütun adlarýnýn ekleneceði Excel çalýþma sayfasý.</param>
        /// <param name="columnNames">Eklenecek sütun adlarýný içeren dizi.</param>
        void AddColumns(ExcelWorksheet workSheet, string[] columnNames);


        /// <summary>
        /// Belirtilen excel çalýþma sayfasýna, verilen sütun adlarýný eklemek için kullanýlýr.
        /// Sütun adlarý, bir Dictionary kullanýlarak belirtilir; anahtar sütun numarasýný, deðer ise sütun adýný temsil eder.
        /// </summary>
        /// <param name="workSheet">Sütun adlarýnýn ekleneceði Excel çalýþma sayfasý.</param>
        /// <param name="columns">Eklenecek sütun adlarýný içeren sözlük.</param>
        void AddColumns(ExcelWorksheet workSheet, Dictionary<int, string> columns);


        /// <summary>
        /// Belirtilen Excel çalýþma sayfasýna, verilen hücre aralýðýnda belirli bir doðrulama türünü uygulamak için kullanýlýr.
        /// Kullanýcýdan alýnan verilerin geçerliliðini kontrol etmek amacýyla e-posta, tam sayý, tarih, ondalýk, zaman ve metin uzunluðu gibi doðrulama türleri desteklenmektedir.
        /// </summary>
        /// <param name="worksheet">Doðrulamanýn uygulanacaðý Excel çalýþma sayfasý.</param>
        /// <param name="range">Doðrulamanýn uygulanacaðý hücre aralýðý.</param>
        /// <param name="validationType">Uygulanacak doðrulama türü.</param>
        /// <param name="values">(isteðe baðlý) Doðrulama türü için kullanýlacak deðerler dizisi.</param>
        /// <param name="customFormula">(isteðe baðlý) Özel doðrulama için kullanýlacak formül.</param>
        void ApplyValidation(ExcelWorksheet worksheet, string range, ValidationTypes validationType, string[]? values = null, string? customFormula = null);


        /// <summary>
        /// Bir hücreye yorum eklemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Yorumun ekleneceði çalýþma sayfasý.</param>
        /// <param name="row">Yorumun ekleneceði hücrenin satýr numarasý.</param>
        /// <param name="column">Yorumun ekleneceði hücrenin sütun numarasý.</param>
        /// <param name="comment">Eklenecek yorum metni.</param>
        /// <param name="author">Yorumun yazarý.</param>
        void AddComment(ExcelWorksheet workSheet, int row, int column, string comment, string author);


        /// <summary>
        /// Bir sütuna açýlýr liste(dropdown) eklemek için kullanýlýr.
        /// Açýlýr liste, verilen deðerler dizisinden oluþturulur ve belirtilen hücre aralýðýna uygulanýr.
        /// Excel'deki formül karakter uzunluðu kýsýtý sebebiyle, eðer formül 255 karakterden uzunsa, deðerler yeni bir sayfada tanýmlanýr.
        /// </summary>
        /// <param name="WorkSheet">Açýlýr listenin ekleneceði çalýþma sayfasý.</param>
        /// <param name="values">Açýlýr listede gösterilecek deðerler.</param>
        /// <param name="cell">Açýlýr listenin uygulanacaðý hücre adresi.</param>
        /// <param name="minRow">Açýlýr listenin uygulanacaðý minimum satýr numarasý.</param>
        /// <param name="maxRow">Açýlýr listenin uygulanacaðý maksimum satýr numarasý.</param>
        void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow);


        /// <summary>
        /// Bir sütuna açýlýr liste(dropdown) eklemek için kullanýlýr.
        /// Açýlýr liste, verilen deðerler dizisinden oluþturulur ve belirtilen hücre aralýðýna uygulanýr.
        /// Ayrýca, açýlýr listenin deðerlerinin yer alacaðý yeni bir sayfa oluþturulur.
        /// </summary>
        /// <param name="WorkSheet">Açýlýr listenin ekleneceði çalýþma sayfasý.</param>
        /// <param name="values">Açýlýr listede gösterilecek deðerler.</param>
        /// <param name="cell">Açýlýr listenin uygulanacaðý hücre adresi.</param>
        /// <param name="minRow">Açýlýr listenin uygulanacaðý minimum satýr numarasý.</param>
        /// <param name="maxRow">Açýlýr listenin uygulanacaðý maksimum satýr numarasý.</param>
        /// <param name="sheetName">Deðerlerin yer alacaðý yeni sayfanýn ismi.</param>
        void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow, string sheetName);


        /// <summary>
        /// Birbirine baðýmlý dropdownlar oluþturmak için kullanýlýr.
        /// Bu metod, belirtilen çalýþma sayfasýnda iki hücre arasýnda baðýmlý açýlýr listeler oluþturur.
        /// Ýlk hücrede seçilen deðere göre ikinci hücredeki seçenekler dinamik olarak güncellenir.
        /// </summary>
        /// <param name="worksheet">Açýlýr listelerin oluþturulacaðý çalýþma sayfasý.</param>
        /// <param name="data">Ýlk hücredeki ve ona baðýmlý ikinci hücredeki seçime baðlý olarak gösterilecek deðerlerin listesi.</param>
        /// <param name="firstCell">Ýlk açýlýr listenin hücre adresi.</param>
        /// <param name="secondCell">Ýkinci açýlýr listenin hücre adresi.</param>
        /// <param name="minRow">Açýlýr listelerin uygulanacaðý minimum satýr numarasý.</param>
        /// <param name="maxRow">Açýlýr listelerin uygulanacaðý maksimum satýr numarasý.</param>
        void CreateDependentDropdowns(ExcelWorksheet worksheet, Dictionary<string, List<string>> data, string firstCell, string secondCell, int minRow, int maxRow);


        /// <summary>
        /// Belirtilen çalýþma sayfasýnda bir hücre aralýðýna ad tanýmlamak için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Adýn tanýmlanacaðý çalýþma sayfasý.</param>
        /// <param name="rangeName">Tanýmlanacak adýn ismi.</param>
        /// <param name="cellRange">Adýn uygulanacaðý hücre aralýðý.</param>
        void DefineNamedRange(ExcelWorksheet workSheet, string rangeName, string cellRange);


        /// <summary>
        /// Otomatik olarak sütunlarýn geniþliðini ayarlamak için kullanýlýr.
        /// Minimum geniþlik 10, maksimum geniþlik 100 olarak ayarlanmýþtýr.
        /// </summary>
        /// <param name="workSheet">AutoFit'in uygulanacaðý excel çalýþma sayfasý.</param>
        void SetAutoFit(ExcelWorksheet WorkSheet);


        /// <summary>
        /// Bütün excelin background rengini setlemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Çalýþýlacak olan excel çalýþma sayfasý.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        void SetBackGroundColor(ExcelWorksheet workSheet, Color color);


        /// <summary>
        /// Excelde istenilen satýr ve/veya sütunun background rengini setlemek için kullanýlýr.
        /// Eðer bir satýr numarasý saðlanýrsa, o satýrýn arka plan rengi ayarlanýr.
        /// Eðer bir sütun numarasý saðlanýrsa, o sütunun arka plan rengi ayarlanýr.
        /// Her ikisi de saðlanýrsa, her ikisi de ayarlanýr. Hiçbiri saðlanmazsa, hiçbir iþlem yapýlmaz.
        /// </summary>
        /// <param name="workSheet">Çalýþýlacak olan excel çalýþma sayfasý.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        /// <param name="rowNumber">Arka plan renginin ayarlanacaðý satýr numarasý (isteðe baðlý).</param>
        /// <param name="columnNumber">Arka plan renginin ayarlanacaðý sütun numarasý (isteðe baðlý).</param>
        void SetBackGroundColor(ExcelWorksheet workSheet, Color color, int? rowNumber = null, int? columnNumber = null);


        /// <summary>
        /// Bir hücreye bir deðer yazmak için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Deðerin yazýlacaðý çalýþma sayfasý.</param>
        /// <param name="row">Hücrenin satýr numarasý.</param>
        /// <param name="column">Hücrenin sütun numarasý.</param>
        /// <param name="value">Yazýlacak deðer.</param>
        void WriteCell(ExcelWorksheet workSheet, int row, int column, object value);


        /// <summary>
        /// Bir hücredeki deðeri temizlemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Deðerin temizleneceði çalýþma sayfasý.</param>
        /// <param name="row">Hücrenin satýr numarasý.</param>
        /// <param name="column">Hücrenin sütun numarasý.</param>
        void ClearCell(ExcelWorksheet workSheet, int row, int column);


        /// <summary>
        /// Belirtilen hücredeki deðeri okumak için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Deðerin okunacaðý çalýþma sayfasý.</param>
        /// <param name="row">Hücrenin satýr numarasý.</param>
        /// <param name="column">Hücrenin sütun numarasý.</param>
        /// <returns>Okunan hücre deðeri.</returns>
        object ReadCell(ExcelWorksheet workSheet, int row, int column);


        /// <summary>
        /// Belirtilen hücre aralýðýný birleþtirmek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Birleþtirilecek hücrelerin bulunduðu çalýþma sayfasý.</param>
        /// <param name="fromRow">Birleþtirmenin baþlayacaðý satýr numarasý.</param>
        /// <param name="fromColumn">Birleþtirmenin baþlayacaðý sütun numarasý.</param>
        /// <param name="toRow">Birleþtirmenin biteceði satýr numarasý.</param>
        /// <param name="toColumn">Birleþtirmenin biteceði sütun numarasý.</param>
        void MergeCells(ExcelWorksheet workSheet, int fromRow, int fromColumn, int toRow, int toColumn);


        /// <summary>
        /// Belirtilen çalýþma sayfasýný bir þifre ile korur, yetkisiz deðiþiklikleri engeller.
        /// </summary>
        /// <param name="workSheet">Korunacak çalýþma sayfasý.</param>
        /// <param name="password">Koruma için ayarlanacak þifre.</param>
        void ProtectSheet(ExcelWorksheet workSheet, string password);


        /// <summary>
        /// Çalýþma sayfasýndaki belirli bir hücreye bir formül eklemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Formülün ekleneceði çalýþma sayfasý.</param>
        /// <param name="row">Hücrenin satýr numarasý.</param>
        /// <param name="column">Hücrenin sütun numarasý.</param>
        /// <param name="formula">Hücreye eklenecek formül.</param>
        void AddFormula(ExcelWorksheet workSheet, int row, int column, string formula);


        /// <summary>
        /// Belirli bir hücre aralýðýna formüle dayalý koþullu biçimlendirme eklemek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Koþullu biçimlendirmenin uygulanacaðý çalýþma sayfasý.</param>
        /// <param name="address">Biçimlendirilecek hücre aralýðýnýn adresi.</param>
        /// <param name="formula">Biçimlendirme koþulunu belirleyen formül.</param>
        /// <param name="color">Koþul karþýlandýðýnda uygulanacak arka plan rengi.</param>
        void AddConditionalFormatting(ExcelWorksheet workSheet, string address, string formula, Color color);


        /// <summary>
        /// Belirtilen çalýþma sayfasýndaki belirli bir hücreyi, belirtilen formatla biçimlendirmek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Biçimlendirilecek hücreyi içeren çalýþma sayfasý.</param>
        /// <param name="row">Biçimlendirilecek hücrenin satýr numarasý.</param>
        /// <param name="column">Biçimlendirilecek hücrenin sütun numarasý.</param>
        /// <param name="format">Hücreye uygulanacak format.</param>
        void FormatCell(ExcelWorksheet workSheet, int row, int column, string format);


        /// <summary>
        /// Belirtilen çalýþma sayfasýndaki belirli bir sütundaki tüm hücreleri, belirtilen formatla biçimlendirmek için kullanýlýr.
        /// </summary>
        /// <param name="workSheet">Biçimlendirilecek sütunu içeren çalýþma sayfasý.</param>
        /// <param name="column">Biçimlendirilecek sütun numarasý.</param>
        /// <param name="format">Sütuna uygulanacak format.</param>
        void FormatCell(ExcelWorksheet workSheet, int column, string format);
    }
}