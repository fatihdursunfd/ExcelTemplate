using ExcelTemplate.Consts;
using OfficeOpenXml;

namespace ExcelTemplate.Helpers
{
    public static class ValidationHelper
    {
        public static void CreateEmailValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddCustomValidation(range);

            validation.Formula.ExcelFormula = "ISNUMBER(MATCH(\"*@*.?*\", " + range.Split(':')[0] + ", 0))";
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Email";
            validation.Error = "Please enter a valid email address.";
        }

        public static void CreateIntegerValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddIntegerValidation(range);

            validation.Formula.Value = int.MinValue;
            validation.Formula2.Value = int.MaxValue;

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Number";
            validation.Error = "Please enter a valid integer.";
        }

        public static void CreateDateValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddDateTimeValidation(range);

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Date";
            validation.Error = "Please enter a valid date.";
        }

        public static void CreateDecimalValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddDecimalValidation(range);

            validation.Formula.Value = ExcelConsts.LargestAllowedNegativeNumber;
            validation.Formula2.Value = ExcelConsts.LargestAllowedPositiveNumber;

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Decimal";
            validation.Error = "Please enter a valid decimal number.";

        }

        public static void CreateTimeValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddTimeValidation(range);

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Time";
            validation.Error = "Please enter a valid time.";
        }

        public static void CreateTextLengthValidation(ExcelWorksheet worksheet, string range)
        {
            var validation = worksheet.DataValidations.AddTextLengthValidation(range);

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Text Length";
            validation.Error = "Please enter text with a valid length.";
        }

        public static void CreateListValidation(ExcelWorksheet worksheet, string range, string[]? values)
        {
            if (values == null)
            {
                return;
            }

            var validation = worksheet.DataValidations.AddListValidation(range);

            foreach (var value in values)
            {
                validation.Formula.Values.Add(value);
            }

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Selection";
            validation.Error = "Please select a value from the list.";
        }

        public static void CreateCustomValidation(ExcelWorksheet worksheet, string range, string? formula)
        {
            var validation = worksheet.DataValidations.AddCustomValidation(range);

            validation.Formula.ExcelFormula = formula;
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Input";
            validation.Error = "Please enter a valid value.";
        }

    }
}