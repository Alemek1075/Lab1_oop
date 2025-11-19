using System.Globalization;
using System.Text.RegularExpressions;

namespace MauiApp1;

public static class Calculator
{
    public static object Evaluate(string expression, Func<string, object> cellResolver)
    {
        if (string.IsNullOrWhiteSpace(expression)) return 0m;

        if (expression.Contains(','))
            throw new ArgumentException("Comma is not allowed");

        if (decimal.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal val))
            return val;

        if (expression.StartsWith("="))
            expression = expression.Substring(1);

        string pattern = @"(mod|div|inc|dec|<>|<=|>=|[+\-*/=<>()])|([a-zA-Z]+[0-9]+)|([0-9]+(?:\.[0-9]+)?)";

        string cleanExpr = expression.Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
        var matches = Regex.Matches(cleanExpr, pattern, RegexOptions.IgnoreCase);

        int totalLength = 0;
        foreach (Match m in matches) totalLength += m.Length;

        if (totalLength != cleanExpr.Length)
            throw new ArgumentException("Invalid characters");

        for (int i = 0; i < matches.Count; i++)
        {
            string token = matches[i].Value.ToLower();
            if (token == "inc" || token == "dec")
            {
                if (i + 1 >= matches.Count || matches[i + 1].Value != "(")
                {
                    throw new ArgumentException("Function requires parentheses");
                }
            }
        }

        var rpn = ToRPN(matches);
        return CalcRPN(rpn, cellResolver);
    }

    private static Queue<string> ToRPN(MatchCollection tokens)
    {
        Queue<string> output = new Queue<string>();
        Stack<string> stack = new Stack<string>();

        foreach (Match match in tokens)
        {
            string token = match.Value;

            if (decimal.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            {
                output.Enqueue(token);
            }
            else if (Regex.IsMatch(token, @"^[a-zA-Z]+[0-9]+$"))
            {
                output.Enqueue(token);
            }
            else
            {
                string op = token.ToLower();

                if (op == "inc" || op == "dec")
                {
                    stack.Push(op);
                }
                else if (op == "(")
                {
                    stack.Push(op);
                }
                else if (op == ")")
                {
                    bool parenFound = false;
                    while (stack.Count > 0)
                    {
                        if (stack.Peek() == "(")
                        {
                            parenFound = true;
                            break;
                        }
                        output.Enqueue(stack.Pop());
                    }

                    if (!parenFound) throw new ArgumentException("Mismatched parentheses");

                    stack.Pop();

                    if (stack.Count > 0)
                    {
                        string top = stack.Peek();
                        if (top == "inc" || top == "dec")
                        {
                            output.Enqueue(stack.Pop());
                        }
                    }
                }
                else
                {
                    while (stack.Count > 0 && GetPriority(stack.Peek()) >= GetPriority(op))
                    {
                        output.Enqueue(stack.Pop());
                    }
                    stack.Push(op);
                }
            }
        }

        while (stack.Count > 0)
        {
            string top = stack.Pop();
            if (top == "(") throw new ArgumentException("Mismatched parentheses");
            output.Enqueue(top);
        }

        return output;
    }

    private static object CalcRPN(Queue<string> rpn, Func<string, object> cellResolver)
    {
        Stack<object> stack = new Stack<object>();

        while (rpn.Count > 0)
        {
            string token = rpn.Dequeue();

            if (decimal.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal d))
            {
                stack.Push(d);
            }
            else if (Regex.IsMatch(token, @"^[a-zA-Z]+[0-9]+$"))
            {
                object resolvedValue = cellResolver(token);
                if (resolvedValue is string s && s.Contains(','))
                    throw new ArgumentException("Cell contains comma");

                stack.Push(resolvedValue);
            }
            else if (token == "inc")
            {
                if (stack.Count < 1) throw new ArgumentException("Missing operand");
                object val = stack.Pop();
                stack.Push(ToDecimal(val) + 1);
            }
            else if (token == "dec")
            {
                if (stack.Count < 1) throw new ArgumentException("Missing operand");
                object val = stack.Pop();
                stack.Push(ToDecimal(val) - 1);
            }
            else
            {
                if (stack.Count < 2) throw new ArgumentException("Missing operands");

                object b = stack.Pop();
                object a = stack.Pop();

                switch (token)
                {
                    case "+":
                        stack.Push(ToDecimal(a) + ToDecimal(b));
                        break;
                    case "-": stack.Push(ToDecimal(a) - ToDecimal(b)); break;
                    case "*": stack.Push(ToDecimal(a) * ToDecimal(b)); break;
                    case "/":
                        decimal divB = ToDecimal(b);
                        if (divB == 0) throw new DivideByZeroException();
                        stack.Push(ToDecimal(a) / divB);
                        break;
                    case "mod":
                        decimal modB = ToDecimal(b);
                        if (modB == 0) throw new DivideByZeroException();
                        stack.Push(ToDecimal(a) % modB);
                        break;
                    case "div":
                        decimal intDivB = ToDecimal(b);
                        if (intDivB == 0) throw new DivideByZeroException();
                        stack.Push((int)(ToDecimal(a) / intDivB));
                        break;
                    case "=": stack.Push(ToDecimal(a) == ToDecimal(b) ? 1m : 0m); break;
                    case "<": stack.Push(ToDecimal(a) < ToDecimal(b) ? 1m : 0m); break;
                    case ">": stack.Push(ToDecimal(a) > ToDecimal(b) ? 1m : 0m); break;
                    case "<=": stack.Push(ToDecimal(a) <= ToDecimal(b) ? 1m : 0m); break;
                    case ">=": stack.Push(ToDecimal(a) >= ToDecimal(b) ? 1m : 0m); break;
                    case "<>": stack.Push(ToDecimal(a) != ToDecimal(b) ? 1m : 0m); break;
                }
            }
        }

        if (stack.Count != 1) throw new ArgumentException("Invalid expression");
        return stack.Pop();
    }

    private static int GetPriority(string op)
    {
        if (op == "inc" || op == "dec") return 4;
        if (op == "*" || op == "/" || op == "mod" || op == "div") return 3;
        if (op == "+" || op == "-") return 2;
        if (op == "=" || op == "<" || op == ">" || op == "<=" || op == ">=" || op == "<>") return 1;
        return 0;
    }

    private static bool IsNumber(object value)
    {
        return value is decimal || value is int || value is double || value is float;
    }

    private static decimal ToDecimal(object value)
    {
        if (value is decimal d) return d;

        if (value is string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0m;

            if (s.Contains(',')) throw new ArgumentException("Comma not allowed");

            if (decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal res))
                return res;
        }

        throw new ArgumentException("Value is not a number");
    }
}