using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RedStar.TimesheetGenerator.Excel
{
    public static class CellExtensions
    {
        public static void AddBlackBorder(this ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Top.Color.SetColor(Color.Black);
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Color.SetColor(Color.Black);
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Color.SetColor(Color.Black);
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Color.SetColor(Color.Black);
        }

        public static void AddThickBlackBorder(this ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Top.Color.SetColor(Color.Black);
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Bottom.Color.SetColor(Color.Black);
            cell.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Left.Color.SetColor(Color.Black);
            cell.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Right.Color.SetColor(Color.Black);
        }

        public static void CenterHorizontallAndVertically(this ExcelRange cell)
        {
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }
}