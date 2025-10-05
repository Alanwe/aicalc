namespace AiCalc.Models;

public readonly record struct CellAddress(string SheetName, int Row, int Column)
{
    public override string ToString()
    {
        var colName = ColumnIndexToName(Column);
        return $"{SheetName}!{colName}{Row + 1}";
    }

    public static bool TryParse(string input, string defaultSheet, out CellAddress address)
    {
        address = default;
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var sheetSplit = input.Split('!');
        string sheet = defaultSheet;
        string cellPart;

        if (sheetSplit.Length == 2)
        {
            sheet = sheetSplit[0];
            cellPart = sheetSplit[1];
        }
        else
        {
            cellPart = sheetSplit[0];
        }

        int column = 0;
        int rowIndex = -1;
        foreach (var ch in cellPart)
        {
            if (char.IsLetter(ch))
            {
                column = column * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            }
            else if (char.IsDigit(ch))
            {
                rowIndex = rowIndex < 0 ? 0 : rowIndex;
                rowIndex = rowIndex * 10 + (ch - '0');
            }
            else
            {
                return false;
            }
        }

        if (rowIndex < 1 || column < 1)
        {
            return false;
        }

        address = new CellAddress(sheet, rowIndex - 1, column - 1);
        return true;
    }

    public static string ColumnIndexToName(int columnIndex)
    {
        int dividend = columnIndex + 1;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}
