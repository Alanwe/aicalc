using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using Windows.UI;

namespace AiCalc.Services;

/// <summary>
/// Provides syntax highlighting for formula text (Phase 5)
/// </summary>
public static class FormulaSyntaxHighlighter
{
    private static readonly SolidColorBrush FunctionBrush = new(Color.FromArgb(0xFF, 0x00, 0x7A, 0xCC)); // Blue
    private static readonly SolidColorBrush CellRefBrush = new(Color.FromArgb(0xFF, 0x09, 0x8D, 0x58)); // Green
    private static readonly SolidColorBrush StringBrush = new(Color.FromArgb(0xFF, 0xA3, 0x15, 0x15)); // Red
    private static readonly SolidColorBrush NumberBrush = new(Color.FromArgb(0xFF, 0x09, 0x88, 0x58)); // Teal
    private static readonly SolidColorBrush OperatorBrush = new(Color.FromArgb(0xFF, 0x66, 0x66, 0x66)); // Gray
    private static readonly SolidColorBrush DefaultBrush = new(Color.FromArgb(0xFF, 0x00, 0x00, 0x00)); // Black

    private static readonly HashSet<string> Operators = new() { "+", "-", "*", "/", "=", "<", ">", "<=", ">=", "<>", "&" };

    /// <summary>
    /// Parse formula and return colored spans for RichEditBox
    /// </summary>
    public static List<FormulaToken> Tokenize(string formula)
    {
        var tokens = new List<FormulaToken>();
        if (string.IsNullOrWhiteSpace(formula))
        {
            return tokens;
        }

        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];

            // String literals
            if (c == '"')
            {
                int start = i;
                i++;
                while (i < formula.Length && formula[i] != '"')
                {
                    i++;
                }
                if (i < formula.Length) i++; // Include closing quote
                tokens.Add(new FormulaToken(start, i - start, FormulaTokenType.String));
                continue;
            }

            // Numbers
            if (char.IsDigit(c) || (c == '.' && i + 1 < formula.Length && char.IsDigit(formula[i + 1])))
            {
                int start = i;
                while (i < formula.Length && (char.IsDigit(formula[i]) || formula[i] == '.'))
                {
                    i++;
                }
                tokens.Add(new FormulaToken(start, i - start, FormulaTokenType.Number));
                continue;
            }

            // Cell references (e.g., A1, B12, Sheet1!A1)
            if (char.IsLetter(c))
            {
                int start = i;

                // Check for sheet reference (Sheet1!)
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_'))
                {
                    i++;
                }
                
                if (i < formula.Length && formula[i] == '!')
                {
                    i++; // Skip '!'
                }

                // Column letters
                int colStart = i;
                while (i < formula.Length && char.IsLetter(formula[i]))
                {
                    i++;
                }

                // Row numbers
                int rowStart = i;
                while (i < formula.Length && char.IsDigit(formula[i]))
                {
                    i++;
                }

                // Determine if it's a cell reference or function
                if (rowStart > colStart && i > rowStart)
                {
                    // Has letters followed by numbers - likely a cell reference
                    tokens.Add(new FormulaToken(start, i - start, FormulaTokenType.CellReference));
                }
                else if (i < formula.Length && formula[i] == '(')
                {
                    // Followed by '(' - it's a function
                    tokens.Add(new FormulaToken(start, i - start, FormulaTokenType.Function));
                }
                else
                {
                    // Just text/identifier
                    tokens.Add(new FormulaToken(start, i - start, FormulaTokenType.Text));
                }
                continue;
            }

            // Operators
            if (i + 1 < formula.Length && Operators.Contains(formula.Substring(i, 2)))
            {
                tokens.Add(new FormulaToken(i, 2, FormulaTokenType.Operator));
                i += 2;
                continue;
            }
            if (Operators.Contains(c.ToString()))
            {
                tokens.Add(new FormulaToken(i, 1, FormulaTokenType.Operator));
                i++;
                continue;
            }

            // Default: skip whitespace and punctuation
            i++;
        }

        return tokens;
    }

    /// <summary>
    /// Get brush color for token type
    /// </summary>
    public static SolidColorBrush GetBrushForTokenType(FormulaTokenType type)
    {
        return type switch
        {
            FormulaTokenType.Function => FunctionBrush,
            FormulaTokenType.CellReference => CellRefBrush,
            FormulaTokenType.String => StringBrush,
            FormulaTokenType.Number => NumberBrush,
            FormulaTokenType.Operator => OperatorBrush,
            _ => DefaultBrush
        };
    }
}

public enum FormulaTokenType
{
    Text,
    Function,
    CellReference,
    String,
    Number,
    Operator
}

public record FormulaToken(int Start, int Length, FormulaTokenType Type);
