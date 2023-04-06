public class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Please insert the column you wish to get the number of:");
        var columnLetters = Console.ReadLine();
        Console.WriteLine("Would you like to use lowercase letters aswell?(Y=Yes/N=No default N)");
        var allowSmall = (Console.ReadKey().KeyChar=='y')?true:false;

        try
        {
            Console.WriteLine($"\nThe number of the desired column is {GetExcelColumnNumber(columnLetters, allowSmall)}");
        } catch (ArgumentException ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
        }
    }

    public static int GetExcelColumnNumber(string columnLetter, bool allowSmallLetters = false)
    {
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

}