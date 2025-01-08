using ExcelTemplate.Consts;
using ExcelTemplate.Enums;
using ExcelTemplate.Services;

var _excelUtilities = new ExcelUtilities();

_excelUtilities.ApplyDefaultStyling(_excelUtilities.WorkSheet);

_excelUtilities.AddColumns(_excelUtilities.WorkSheet, StaticData.ColumnNames);

_excelUtilities.CreateDependentDropdowns(_excelUtilities.WorkSheet, StaticData.CountriesWithCities, "D", "E", 2, 100);

// Yorum ekleme
_excelUtilities.AddComment(_excelUtilities.WorkSheet, 1, 5, "Ülke seçildikten sonra seçilmelidir.", "FD");

// Hücrelere biçimlendirme ekleme:
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 1, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 2, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 3, ExcelConsts.IntegerFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 4, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 5, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 6, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 7, ExcelConsts.TextFormat);
_excelUtilities.FormatCell(_excelUtilities.WorkSheet, 8, ExcelConsts.DecimalFormat);

// Hücrelere validasyon ekleme:
_excelUtilities.ApplyValidation(_excelUtilities.WorkSheet, "C2:C100", ValidationTypes.Integer);
_excelUtilities.ApplyValidation(_excelUtilities.WorkSheet, "H2:H100", ValidationTypes.Decimal);
_excelUtilities.ApplyValidation(_excelUtilities.WorkSheet, "G2:G100", ValidationTypes.Email);

_excelUtilities.SetAutoFit(_excelUtilities.WorkSheet);

_excelUtilities.ProtectSheet(_excelUtilities.WorkSheet, "1234");

_excelUtilities.Save("test", @"C:\\Users\\Neyasis\\source\\repos\\ExcelTemplate\\ExcelTemplate\\");

_excelUtilities.Dispose();