namespace ExcelSheetTest
{
    using Xunit;

    public class UnitTest1
    {
        [Theory]
        [InlineData("A",1)]
        [InlineData("B", 2)]
        [InlineData("C", 3)]
        [InlineData("D", 4)]
        [InlineData("E", 5)]
        [InlineData("X", 24)]
        [InlineData("Y", 25)]
        [InlineData("Z", 26)]
        public void TestForSingleLetter_ValueIsPredefined_ShouleBeEqual(string letter,int number)
        {
            // Arrange

            // Act
            var column = Program.GetExcelColumnNumber(letter);

            // Assert
            Assert.Equal(number, column);
        }

        [Theory]
        [InlineData("AA", 27)]
        [InlineData("BBC", 1407)]
        [InlineData("YZ", 676)]
        [InlineData("ZZ", 702)]
        [InlineData("AAC", 705)]
        public void TestForMultipleLetters_ValueIsPredefined_ShouleBeEqual(string letter, int number)
        {
            // Arrange

            // Act
            var column = Program.GetExcelColumnNumber(letter);

            // Assert
            Assert.Equal(number, column);
        }

        [Theory]
        [InlineData("AA",false, false)]
        [InlineData("Bb", true, false)]
        [InlineData("Yz", false, true)]
        [InlineData("BB ", true, true)]
        [InlineData("BA3", true, true)]
        [InlineData("BQ@", false, true)]
        public void MultipleLetters_TestingForLowerCaseAndRandomChars_ShouldThrowErrorWhenCharacterIsNotAllowed(string letter,bool allowsmall, bool shouldThrowException)
        {
            // Arrange

            // Act
            var exception =Record.Exception(() => Program.GetExcelColumnNumber(letter,allowsmall));

            // Assert
            Assert.Equal(exception!=null, shouldThrowException);
        }



        [Theory]
        [InlineData("A", 1)]
        [InlineData("B", 2)]
        [InlineData("C", 3)]
        [InlineData("D", 4)]
        [InlineData("E", 5)]
        [InlineData("X", 24)]
        [InlineData("Y", 25)]
        [InlineData("Z", 26)]
        public void TestNumberToLetter_ValueIsPredefined_ShouleBeEqual(string letter, int number)
        {
            // Arrange

            // Act
            var column = Program.GetExcelLetter(number);

            // Assert
            Assert.Equal(letter, column);
        }

        [Theory]
        [InlineData("AA", 27)]
        [InlineData("BBC", 1407)]
        [InlineData("YZ", 676)]
        [InlineData("ZZ", 702)]
        [InlineData("AAC", 705)]
        public void TestNumberToLetters_ValueIsPredefined_ShouleBeEqual(string letter, int number)
        {
            // Arrange

            // Act
            var column = Program.GetExcelLetter(number);

            // Assert
            Assert.Equal(letter, column);
        }
    }
}