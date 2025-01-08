using ExcelTemplate.Enums;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelTemplate.Interfaces
{
    public interface IExcelUtilities : IDisposable
    {
        /// <summary>
        /// Olu�turulan excel dosyas�n� belirtilen dosya ad� ve yol ile kaydetmek i�in kullan�l�r.
        /// </summary>
        /// <param name="filename">Kaydedilecek dosyan�n ad�.</param>
        /// <param name="path">Dosyan�n kaydedilece�i yol.</param>
        void Save(string filename, string path);


        /// <summary>
        /// Excel dosyas�n� byte dizisi olarak d�nd�rmek i�in kullan�l�r.
        /// Bu metod, olu�turulan Excel dosyas�n�n i�eri�ini byte dizisi format�nda alman�z� sa�lar.
        /// </summary>
        /// <returns>Excel dosyas�n�n byte dizisi.</returns>
        byte[] GetExcelAsByteArray();


        /// <summary>
        /// Excel sayfas�n�n temel stil ayarlar�n� yapmak i�in kullan�l�r.
        /// Tab rengi, varsay�lan sat�r y�ksekli�i, ilk sat�r y�ksekli�i,
        /// yatay ve dikey hizalama ile yaz� tipi kal�nl��� gibi �zellikleri ayarlamak i�in kullan�l�r.
        /// </summary>
        /// <param name="WorkSheet">Stil ayarlar�n�n uygulanaca�� Excel �al��ma sayfas�.</param>
        /// <param name="color">Tab�n rengi (varsay�lan: siyah).</param>
        /// <param name="defaultRowHeight">Varsay�lan sat�r y�ksekli�i (varsay�lan: 12).</param>
        /// <param name="height">�lk sat�r�n y�ksekli�i (varsay�lan: 20).</param>
        /// <param name="horizontalAlignment">�lk sat�r�n yatay hizalamas� (varsay�lan: Center).</param>
        /// <param name="verticalAlignment">�lk sat�r�n dikey hizalamas� (varsay�lan: Center).</param>
        /// <param name="isBold">Yaz� tipinin kal�n olup olmayaca�� (varsay�lan: false).</param>
        void ApplyDefaultStyling(ExcelWorksheet WorkSheet, Color? color = null, int defaultRowHeight = 12, int height = 20, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center, ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Center, bool isBold = false);


        /// <summary>
        /// Belirtilen h�creye stil uygulamak i�in kullan�l�r.
        /// H�creye yatay ve dikey hizalama, yaz� tipi kal�nl��� ve arka plan rengi gibi stil �zellikleri ayarlan�r.
        /// </summary>
        /// <param name="workSheet">Stilin uygulanaca�� Excel �al��ma sayfas�.</param>
        /// <param name="row">Stilin uygulanaca�� h�crenin sat�r numaras�.</param>
        /// <param name="column">Stilin uygulanaca�� h�crenin s�tun numaras�.</param>
        /// <param name="horizontalAlignment">H�crenin yatay hizalamas�.</param>
        /// <param name="verticalAlignment">H�crenin dikey hizalamas�.</param>
        /// <param name="isBold">Yaz� tipinin kal�n olup olmayaca��.</param>
        /// <param name="backgroundColor">H�crenin arka plan rengi (iste�e ba�l�).</param>
        void ApplyStyleToCell(ExcelWorksheet workSheet, int row, int column, ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment, bool isBold, Color? backgroundColor = null);


        /// <summary>
        /// Yeni bir sayfa eklemek i�in kullan�l�r.
        /// Sayfan�n ad� belirtilir ve iste�e ba�l� olarak sayfan�n g�r�n�rl��� ayarlanabilir.
        /// </summary>
        /// <param name="sheetName">Eklenecek sayfan�n ad�.</param>
        /// <param name="isHidden">Sayfan�n gizli olup olmayaca�� (varsay�lan: false).</param>
        /// <returns>Eklenen ExcelWorksheet nesnesi.</returns>
        ExcelWorksheet AddSheet(string sheetName, bool isHidden = false);


        /// <summary>
        /// Belirtilen excel �al��ma sayfas�na, verilen s�tun adlar�n� eklemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">S�tun adlar�n�n eklenece�i Excel �al��ma sayfas�.</param>
        /// <param name="columnNames">Eklenecek s�tun adlar�n� i�eren dizi.</param>
        void AddColumns(ExcelWorksheet workSheet, string[] columnNames);


        /// <summary>
        /// Belirtilen excel �al��ma sayfas�na, verilen s�tun adlar�n� eklemek i�in kullan�l�r.
        /// S�tun adlar�, bir Dictionary kullan�larak belirtilir; anahtar s�tun numaras�n�, de�er ise s�tun ad�n� temsil eder.
        /// </summary>
        /// <param name="workSheet">S�tun adlar�n�n eklenece�i Excel �al��ma sayfas�.</param>
        /// <param name="columns">Eklenecek s�tun adlar�n� i�eren s�zl�k.</param>
        void AddColumns(ExcelWorksheet workSheet, Dictionary<int, string> columns);


        /// <summary>
        /// Belirtilen Excel �al��ma sayfas�na, verilen h�cre aral���nda belirli bir do�rulama t�r�n� uygulamak i�in kullan�l�r.
        /// Kullan�c�dan al�nan verilerin ge�erlili�ini kontrol etmek amac�yla e-posta, tam say�, tarih, ondal�k, zaman ve metin uzunlu�u gibi do�rulama t�rleri desteklenmektedir.
        /// </summary>
        /// <param name="worksheet">Do�rulaman�n uygulanaca�� Excel �al��ma sayfas�.</param>
        /// <param name="range">Do�rulaman�n uygulanaca�� h�cre aral���.</param>
        /// <param name="validationType">Uygulanacak do�rulama t�r�.</param>
        /// <param name="values">(iste�e ba�l�) Do�rulama t�r� i�in kullan�lacak de�erler dizisi.</param>
        /// <param name="customFormula">(iste�e ba�l�) �zel do�rulama i�in kullan�lacak form�l.</param>
        void ApplyValidation(ExcelWorksheet worksheet, string range, ValidationTypes validationType, string[]? values = null, string? customFormula = null);


        /// <summary>
        /// Bir h�creye yorum eklemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Yorumun eklenece�i �al��ma sayfas�.</param>
        /// <param name="row">Yorumun eklenece�i h�crenin sat�r numaras�.</param>
        /// <param name="column">Yorumun eklenece�i h�crenin s�tun numaras�.</param>
        /// <param name="comment">Eklenecek yorum metni.</param>
        /// <param name="author">Yorumun yazar�.</param>
        void AddComment(ExcelWorksheet workSheet, int row, int column, string comment, string author);


        /// <summary>
        /// Bir s�tuna a��l�r liste(dropdown) eklemek i�in kullan�l�r.
        /// A��l�r liste, verilen de�erler dizisinden olu�turulur ve belirtilen h�cre aral���na uygulan�r.
        /// Excel'deki form�l karakter uzunlu�u k�s�t� sebebiyle, e�er form�l 255 karakterden uzunsa, de�erler yeni bir sayfada tan�mlan�r.
        /// </summary>
        /// <param name="WorkSheet">A��l�r listenin eklenece�i �al��ma sayfas�.</param>
        /// <param name="values">A��l�r listede g�sterilecek de�erler.</param>
        /// <param name="cell">A��l�r listenin uygulanaca�� h�cre adresi.</param>
        /// <param name="minRow">A��l�r listenin uygulanaca�� minimum sat�r numaras�.</param>
        /// <param name="maxRow">A��l�r listenin uygulanaca�� maksimum sat�r numaras�.</param>
        void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow);


        /// <summary>
        /// Bir s�tuna a��l�r liste(dropdown) eklemek i�in kullan�l�r.
        /// A��l�r liste, verilen de�erler dizisinden olu�turulur ve belirtilen h�cre aral���na uygulan�r.
        /// Ayr�ca, a��l�r listenin de�erlerinin yer alaca�� yeni bir sayfa olu�turulur.
        /// </summary>
        /// <param name="WorkSheet">A��l�r listenin eklenece�i �al��ma sayfas�.</param>
        /// <param name="values">A��l�r listede g�sterilecek de�erler.</param>
        /// <param name="cell">A��l�r listenin uygulanaca�� h�cre adresi.</param>
        /// <param name="minRow">A��l�r listenin uygulanaca�� minimum sat�r numaras�.</param>
        /// <param name="maxRow">A��l�r listenin uygulanaca�� maksimum sat�r numaras�.</param>
        /// <param name="sheetName">De�erlerin yer alaca�� yeni sayfan�n ismi.</param>
        void AddDropdownList(ExcelWorksheet WorkSheet, string[] values, string cell, int minRow, int maxRow, string sheetName);


        /// <summary>
        /// Birbirine ba��ml� dropdownlar olu�turmak i�in kullan�l�r.
        /// Bu metod, belirtilen �al��ma sayfas�nda iki h�cre aras�nda ba��ml� a��l�r listeler olu�turur.
        /// �lk h�crede se�ilen de�ere g�re ikinci h�credeki se�enekler dinamik olarak g�ncellenir.
        /// </summary>
        /// <param name="worksheet">A��l�r listelerin olu�turulaca�� �al��ma sayfas�.</param>
        /// <param name="data">�lk h�credeki ve ona ba��ml� ikinci h�credeki se�ime ba�l� olarak g�sterilecek de�erlerin listesi.</param>
        /// <param name="firstCell">�lk a��l�r listenin h�cre adresi.</param>
        /// <param name="secondCell">�kinci a��l�r listenin h�cre adresi.</param>
        /// <param name="minRow">A��l�r listelerin uygulanaca�� minimum sat�r numaras�.</param>
        /// <param name="maxRow">A��l�r listelerin uygulanaca�� maksimum sat�r numaras�.</param>
        void CreateDependentDropdowns(ExcelWorksheet worksheet, Dictionary<string, List<string>> data, string firstCell, string secondCell, int minRow, int maxRow);


        /// <summary>
        /// Belirtilen �al��ma sayfas�nda bir h�cre aral���na ad tan�mlamak i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Ad�n tan�mlanaca�� �al��ma sayfas�.</param>
        /// <param name="rangeName">Tan�mlanacak ad�n ismi.</param>
        /// <param name="cellRange">Ad�n uygulanaca�� h�cre aral���.</param>
        void DefineNamedRange(ExcelWorksheet workSheet, string rangeName, string cellRange);


        /// <summary>
        /// Otomatik olarak s�tunlar�n geni�li�ini ayarlamak i�in kullan�l�r.
        /// Minimum geni�lik 10, maksimum geni�lik 100 olarak ayarlanm��t�r.
        /// </summary>
        /// <param name="workSheet">AutoFit'in uygulanaca�� excel �al��ma sayfas�.</param>
        void SetAutoFit(ExcelWorksheet WorkSheet);


        /// <summary>
        /// B�t�n excelin background rengini setlemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">�al���lacak olan excel �al��ma sayfas�.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        void SetBackGroundColor(ExcelWorksheet workSheet, Color color);


        /// <summary>
        /// Excelde istenilen sat�r ve/veya s�tunun background rengini setlemek i�in kullan�l�r.
        /// E�er bir sat�r numaras� sa�lan�rsa, o sat�r�n arka plan rengi ayarlan�r.
        /// E�er bir s�tun numaras� sa�lan�rsa, o s�tunun arka plan rengi ayarlan�r.
        /// Her ikisi de sa�lan�rsa, her ikisi de ayarlan�r. Hi�biri sa�lanmazsa, hi�bir i�lem yap�lmaz.
        /// </summary>
        /// <param name="workSheet">�al���lacak olan excel �al��ma sayfas�.</param>
        /// <param name="color">Arka plan rengi olarak ayarlanacak renk.</param>
        /// <param name="rowNumber">Arka plan renginin ayarlanaca�� sat�r numaras� (iste�e ba�l�).</param>
        /// <param name="columnNumber">Arka plan renginin ayarlanaca�� s�tun numaras� (iste�e ba�l�).</param>
        void SetBackGroundColor(ExcelWorksheet workSheet, Color color, int? rowNumber = null, int? columnNumber = null);


        /// <summary>
        /// Bir h�creye bir de�er yazmak i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">De�erin yaz�laca�� �al��ma sayfas�.</param>
        /// <param name="row">H�crenin sat�r numaras�.</param>
        /// <param name="column">H�crenin s�tun numaras�.</param>
        /// <param name="value">Yaz�lacak de�er.</param>
        void WriteCell(ExcelWorksheet workSheet, int row, int column, object value);


        /// <summary>
        /// Bir h�credeki de�eri temizlemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">De�erin temizlenece�i �al��ma sayfas�.</param>
        /// <param name="row">H�crenin sat�r numaras�.</param>
        /// <param name="column">H�crenin s�tun numaras�.</param>
        void ClearCell(ExcelWorksheet workSheet, int row, int column);


        /// <summary>
        /// Belirtilen h�credeki de�eri okumak i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">De�erin okunaca�� �al��ma sayfas�.</param>
        /// <param name="row">H�crenin sat�r numaras�.</param>
        /// <param name="column">H�crenin s�tun numaras�.</param>
        /// <returns>Okunan h�cre de�eri.</returns>
        object ReadCell(ExcelWorksheet workSheet, int row, int column);


        /// <summary>
        /// Belirtilen h�cre aral���n� birle�tirmek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Birle�tirilecek h�crelerin bulundu�u �al��ma sayfas�.</param>
        /// <param name="fromRow">Birle�tirmenin ba�layaca�� sat�r numaras�.</param>
        /// <param name="fromColumn">Birle�tirmenin ba�layaca�� s�tun numaras�.</param>
        /// <param name="toRow">Birle�tirmenin bitece�i sat�r numaras�.</param>
        /// <param name="toColumn">Birle�tirmenin bitece�i s�tun numaras�.</param>
        void MergeCells(ExcelWorksheet workSheet, int fromRow, int fromColumn, int toRow, int toColumn);


        /// <summary>
        /// Belirtilen �al��ma sayfas�n� bir �ifre ile korur, yetkisiz de�i�iklikleri engeller.
        /// </summary>
        /// <param name="workSheet">Korunacak �al��ma sayfas�.</param>
        /// <param name="password">Koruma i�in ayarlanacak �ifre.</param>
        void ProtectSheet(ExcelWorksheet workSheet, string password);


        /// <summary>
        /// �al��ma sayfas�ndaki belirli bir h�creye bir form�l eklemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Form�l�n eklenece�i �al��ma sayfas�.</param>
        /// <param name="row">H�crenin sat�r numaras�.</param>
        /// <param name="column">H�crenin s�tun numaras�.</param>
        /// <param name="formula">H�creye eklenecek form�l.</param>
        void AddFormula(ExcelWorksheet workSheet, int row, int column, string formula);


        /// <summary>
        /// Belirli bir h�cre aral���na form�le dayal� ko�ullu bi�imlendirme eklemek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Ko�ullu bi�imlendirmenin uygulanaca�� �al��ma sayfas�.</param>
        /// <param name="address">Bi�imlendirilecek h�cre aral���n�n adresi.</param>
        /// <param name="formula">Bi�imlendirme ko�ulunu belirleyen form�l.</param>
        /// <param name="color">Ko�ul kar��land���nda uygulanacak arka plan rengi.</param>
        void AddConditionalFormatting(ExcelWorksheet workSheet, string address, string formula, Color color);


        /// <summary>
        /// Belirtilen �al��ma sayfas�ndaki belirli bir h�creyi, belirtilen formatla bi�imlendirmek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Bi�imlendirilecek h�creyi i�eren �al��ma sayfas�.</param>
        /// <param name="row">Bi�imlendirilecek h�crenin sat�r numaras�.</param>
        /// <param name="column">Bi�imlendirilecek h�crenin s�tun numaras�.</param>
        /// <param name="format">H�creye uygulanacak format.</param>
        void FormatCell(ExcelWorksheet workSheet, int row, int column, string format);


        /// <summary>
        /// Belirtilen �al��ma sayfas�ndaki belirli bir s�tundaki t�m h�creleri, belirtilen formatla bi�imlendirmek i�in kullan�l�r.
        /// </summary>
        /// <param name="workSheet">Bi�imlendirilecek s�tunu i�eren �al��ma sayfas�.</param>
        /// <param name="column">Bi�imlendirilecek s�tun numaras�.</param>
        /// <param name="format">S�tuna uygulanacak format.</param>
        void FormatCell(ExcelWorksheet workSheet, int column, string format);
    }
}