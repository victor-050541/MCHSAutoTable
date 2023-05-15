using MCHSAutoTable;
using NUnit.Framework;

namespace MCHSAutoTableTests;

[TestFixture]
public class DataBaseConnectionTests
{
    [Test]
    public void TestDataBaseConnection()
    {
        // Arrange
        var context = new ApplicationContext();
        // Act & Assert
        Assert.DoesNotThrow(() =>
            context.Database.EnsureCreated()); // Проверяем, что подключение к базе данных не вызывает исключений
    }
}