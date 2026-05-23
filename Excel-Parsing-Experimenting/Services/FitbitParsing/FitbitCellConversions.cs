using System.Globalization;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

internal static class FitbitCellConversions
{
    private static readonly NumberStyles IntegerStyles = NumberStyles.Integer | NumberStyles.AllowThousands;
    private static readonly NumberStyles DoubleStyles = NumberStyles.Float | NumberStyles.AllowThousands;

    public static bool IsBlank(object? value)
    {
        return value switch
        {
            null => true,
            DBNull => true,
            string stringValue => string.IsNullOrWhiteSpace(stringValue),
            _ => false
        };
    }

    public static bool TryToString(object? value, out string text)
    {
        text = string.Empty;
        if (IsBlank(value))
        {
            return false;
        }

        text = value switch
        {
            DateTime dateTime => dateTime.ToString(CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty
        };

        return !string.IsNullOrWhiteSpace(text);
    }

    public static bool TryToDateTime(object? value, out DateTime dateTime)
    {
        dateTime = default;

        if (value is DateTime directDateTime)
        {
            dateTime = directDateTime;
            return true;
        }

        if (TryToDouble(value, out var oaDate))
        {
            try
            {
                dateTime = DateTime.FromOADate(oaDate);
                return true;
            }
            catch (ArgumentException)
            {
                // Fall back to string parsing below.
            }
        }

        if (!TryToString(value, out var text))
        {
            return false;
        }

        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dateTime) ||
               DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out dateTime);
    }

    public static bool TryToInt32(object? value, out int number)
    {
        number = default;

        if (value is int directInt)
        {
            number = directInt;
            return true;
        }

        if (value is long directLong && directLong is <= int.MaxValue and >= int.MinValue)
        {
            number = (int)directLong;
            return true;
        }

        if (TryToString(value, out var text) &&
            int.TryParse(text, IntegerStyles, CultureInfo.InvariantCulture, out number))
        {
            return true;
        }

        if (TryToString(value, out text) &&
            int.TryParse(text, IntegerStyles, CultureInfo.CurrentCulture, out number))
        {
            return true;
        }

        if (!TryToDouble(value, out var numericValue))
        {
            return false;
        }

        number = Convert.ToInt32(Math.Round(numericValue, MidpointRounding.AwayFromZero));
        return true;
    }

    public static bool TryToNullableInt32(object? value, out int? number)
    {
        number = null;

        if (IsBlank(value))
        {
            return true;
        }

        if (TryToString(value, out var text) &&
            string.Equals(text, "N/A", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!TryToInt32(value, out var parsedNumber))
        {
            return false;
        }

        number = parsedNumber;
        return true;
    }

    public static bool TryToDouble(object? value, out double number)
    {
        number = default;

        switch (value)
        {
            case null:
            case DBNull:
                return false;
            case double directDouble:
                number = directDouble;
                return true;
            case float directFloat:
                number = directFloat;
                return true;
            case decimal directDecimal:
                number = Convert.ToDouble(directDecimal, CultureInfo.InvariantCulture);
                return true;
            case int directInt:
                number = directInt;
                return true;
            case long directLong:
                number = directLong;
                return true;
        }

        if (!TryToString(value, out var text))
        {
            return false;
        }

        return double.TryParse(text, DoubleStyles, CultureInfo.InvariantCulture, out number) ||
               double.TryParse(text, DoubleStyles, CultureInfo.CurrentCulture, out number);
    }
}
