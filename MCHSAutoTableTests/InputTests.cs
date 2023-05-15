using System.Windows.Forms;
using MCHSAutoTable;

namespace MCHSAutoTableTests;

[TestFixture]
public class InputTests
{
    [Test]
    public void TestButtonClick()
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
        
        // Arrange
        var button = form.Controls.Find("button1", true).FirstOrDefault() as Button;
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick2()
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
        
        // Arrange
        var button = form.Controls.Find("button2", true).FirstOrDefault() as Button;
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick3()
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
        
        // Arrange
        var button = form.Controls.Find("button3", true).FirstOrDefault() as Button;
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick4()
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
        
        // Arrange
        var button = form.Controls.Find("button4", true).FirstOrDefault() as Button;
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick5()
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
        
        // Arrange
        var button = form.Controls.Find("button5", true).FirstOrDefault() as Button;
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick6()
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
        
        // Arrange
        var button = form.Controls.Find("button6", true).FirstOrDefault() as Button;
        Console.WriteLine(button);
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    // generate tests for 17 buttons 
    
    [Test]
    public void TestButtonClick7()
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
        
        // Arrange
        var button = form.Controls.Find("button7", true).FirstOrDefault() as Button;
        Console.WriteLine(button);
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick8()
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
        
        // Arrange
        var button = form.Controls.Find("button8", true).FirstOrDefault() as Button;
        Console.WriteLine(button);
        // Act
        button.PerformClick();
        // Assert
        Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
    }
    
    [Test]
    public void TestButtonClick9()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        // Assert
        Assert.Multiple(() =>
        {
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
            // Arrange
            var button = form.Controls.Find("button9", true).FirstOrDefault() as Button;
            Console.WriteLine(button);
            // Act
            button.PerformClick();
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
        });
    }
    
    [Test]
    public void TestButtonClick10()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        // Assert
        Assert.Multiple(() =>
        {
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
            // Arrange
            var button = form.Controls.Find("button10", true).FirstOrDefault() as Button;
            Console.WriteLine(button);
            // Act
            button.PerformClick();
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
        });
    }
    
    [Test]
    public void TestButtonClick11()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        // Assert
        Assert.Multiple(() =>
        {
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
            // Arrange
            var button = form.Controls.Find("button11", true).FirstOrDefault() as Button;
            Console.WriteLine(button);
            // Act
            button.PerformClick();
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
        });
    }
    
    [Test]
    public void TestButtonClick12()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        // Assert
        Assert.Multiple(() =>
        {
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
            // Arrange
            var button = form.Controls.Find("button12", true).FirstOrDefault() as Button;
            Console.WriteLine(button);
            // Act
            button.PerformClick();
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
        });
    }
    
    [Test]
    public void TestButtonClick13()
    {
        // Arrange
        var form = new Form1();
        // Act
        form.Show(); // Открываем форму
        // Assert
        Assert.Multiple(() =>
        {
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
            Assert.That(form.Controls, Is.Not.Null); // Проверяем, что элементы управления инициализированы
            // Arrange
            var button = form.Controls.Find("button13", true).FirstOrDefault() as Button;
            Console.WriteLine(button);
            // Act
            button.PerformClick();
            // Assert
            Assert.That(form.Visible, Is.True); // Проверяем, что форма видима
        });
    }


}