using System.Collections.Generic;
using System.Linq;

namespace AiCalc.Services;

public static class FormulaParser
{
    /// <summary>
    /// Splits function argument text into top-level arguments, handling quoted strings,
    /// nested parenthesis, and escaped quotes.
    /// </summary>
    public static IEnumerable<string> SplitArguments(string args)
    {
        if (string.IsNullOrWhiteSpace(args))
        {
            yield break;
        }

        int depth = 0;
        bool inString = false;
        char stringChar = '\0';
        bool lastWasEscape = false;
        var current = new List<char>();

        for (int i = 0; i < args.Length; i++)
        {
            var ch = args[i];

            // Handle string quoting and escapes
            if (inString)
            {
                current.Add(ch);
                if (!lastWasEscape && ch == stringChar)
                {
                    inString = false;
                    stringChar = '\0';
                }
                lastWasEscape = ch == '\\' && !lastWasEscape;
                continue;
            }

            if (ch == '"' || ch == '\'')
            {
                inString = true;
                stringChar = ch;
                current.Add(ch);
                lastWasEscape = false;
                continue;
            }

            // Track parentheses depth for nested function arguments
            if (ch == '(')
            {
                depth++;
                current.Add(ch);
                continue;
            }

            if (ch == ')')
            {
                if (depth > 0)
                    depth--;
                current.Add(ch);
                continue;
            }

            // Top-level commas separate args
            if (ch == ',' && depth == 0)
            {
                yield return new string(current.ToArray());
                current.Clear();
                continue;
            }

            current.Add(ch);
        }

        if (current.Count > 0)
        {
            yield return new string(current.ToArray());
        }
    }
}
