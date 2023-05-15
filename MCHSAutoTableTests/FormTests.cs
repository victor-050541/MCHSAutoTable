using System.Windows.Forms;
using MCHSAutoTable;

namespace MCHSAutoTableTests;

[TestFixture]
public class FormTests
{
    [Test]
    public void TestFormOpening()
    {
        // Arrange
        var form = new Form1();

        // Act & Assert
        Assert.DoesNotThrow(() => form.ShowDialog()); // Проверяем, что открытие формы не вызывает исключений
    }
    
    [Test]
    public void TestFormInitialization()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        Assert.Multiple(() =>
        {
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
        });
    }
}