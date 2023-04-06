public class Program
{
    private static void Main(string[] args)
    {
            if (AskForAcceptance("\nWould like to convert Letter to Number?"))
            {
            Console.WriteLine("\nPlease insert the column you wish to get the number of:");
            var columnLetters = Console.ReadLine();
            Console.WriteLine("Would you like to use lowercase letters aswell?(Y=Yes/N=No default N)");
            var allowSmall = (Console.ReadKey().KeyChar == 'y') ? true : false;

            try
            {
                Console.WriteLine($"\nThe number of the desired column is {GetExcelColumnNumber(columnLetters, allowSmall)}");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }

        if (AskForAcceptance("\nWould like to convert Number to Letter?"))
        {
            Console.WriteLine("\nPlease insert the column you wish to get the number of:");
            string input = Console.ReadLine();
            if (int.TryParse(input, out int number))
            {
                Console.WriteLine($"\nThe letter of the desired column is {GetExcelLetter(number)}");
            }
            else
            {
                Console.WriteLine("Invalid input.");
            }        

        }
    }

    /// <summary>
    /// ASks for key from user in order to determine further actions.
    /// </summary>
    /// <param name="message">The message of the request.</param>
    /// <param name="acceptance">The key for accepting the request (default is y).</param>
    /// <returns>True if user hits the required key</returns>
    private static bool AskForAcceptance(string message,char acceptance='y')
    {
        Console.WriteLine($"{message}\n({acceptance}=Yes anything else = no)\n please enter a key)");
        return Console.ReadKey().KeyChar == acceptance;
    }

    /// <summary>
    /// Return the number of a column determined by a combination of letters.
    /// </summary>
    /// <param name="columnLetter">The letters on top of the excel column.</param>
    /// <param name="allowSmallLetters">Allow small case letters and convert them into uppercase letters or consider them errors.</param>
    /// <returns>The nubmer of the column.</returns>
    /// <exception cref="ArgumentException">Throws exception if a character from the columnLetters is invalid.</exception>
    public static int GetExcelColumnNumber(string columnLetter, bool allowSmallLetters = false)
    {
        if(columnLetter == null)
            throw new ArgumentNullException(nameof(columnLetter));

        if(allowSmallLetters)
        {
            columnLetter = columnLetter.ToUpper();
        }
        var letter = columnLetter.AsEnumerable<char>().ToArray();
        int column = 0;
        for (int i=0;i< letter.Length; i++)
        {
            if (letter[i] < 'A' || letter[i] > 'Z') {
                throw new ArgumentException($"The character format {letter[i]} is not inside the upper case alphabet." +
                    $"\n Please use only english uppercase letters.");
            }
            column = column*26 + (letter[i] - 'A' + 1);
        }
        return column;
    }

    /// <summary>
    /// Converts a column number into a set of letters that would be seen at the top of an excel table.
    /// </summary>
    /// <param name="columnNumber">The number we want to convert.</param>
    /// <returns>The letters associated with a column of that index.</returns>
    public static string GetExcelLetter(int columnNumber)
    {
        if (columnNumber == null)
            throw new ArgumentNullException(nameof(columnNumber));

        var column = "";
        while(columnNumber > 0)
        {
            var currentVal = (columnNumber -1) % 26;
            column = (char)(currentVal + 'A') + column;
            columnNumber = (columnNumber- currentVal)/26;
        }

        return column;
    }
}