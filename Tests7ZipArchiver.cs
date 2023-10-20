using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Windows;

namespace AppiumArchive7ZipTest
{
    public class Tests7ZipArchiver
    {

        private const string AppiumUriString = "http://127.0.0.1:4723/wd/hub";
        private const string ZipLocation = @"C:\Program Files\7-Zip\7zFM.exe";
        private const string tempDirectory = @"C:\temp";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> driverArchiveWindow;
        private AppiumOptions options;
        private AppiumOptions optionsArchiveWindow;

        [SetUp]
        public void OpenApp()
        {
            this.options = new AppiumOptions() { PlatformName = "Windows"};
            options.AddAdditionalCapability("app", ZipLocation);
            this.driver = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);

            this.optionsArchiveWindow = new AppiumOptions() { PlatformName = "Windows" };
            optionsArchiveWindow.AddAdditionalCapability("app", "Root");
            this.driverArchiveWindow = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), optionsArchiveWindow);
            driverArchiveWindow.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            
            if (Directory.Exists(tempDirectory)) 
            {
                Directory.Delete(tempDirectory, true);

            }

            Directory.CreateDirectory(tempDirectory);  
            Thread.Sleep(1000); 
        }

        [TearDown]

        public void CloseApp() 
        {
           driver.Quit();
        }

        [Test]
        public void Test_ArchiveFunctionality()
        {
            var inputFilePath = driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            inputFilePath.SendKeys(@"C:\Program Files\7-Zip\" + Keys.Enter);

            var listFiles = driver.FindElementByClassName("SysListView32");
            listFiles.SendKeys(Keys.Control + "a");

            var buttonAdd = driver.FindElementByName("Add");
            buttonAdd.Click();

            var windowArchive = driverArchiveWindow.FindElementByName("Add to Archive");

            Thread.Sleep(500); 
            var inputArchivePath = windowArchive.FindElementByXPath("//Edit[@Name='Archive:']");
            inputArchivePath.SendKeys(@"C:\temp\archive.7z");

            var dropDownFieldArchiveFormat = windowArchive.FindElementByName("Archive format:");
            dropDownFieldArchiveFormat.SendKeys("7z");    

            var dropDownFieldCompressionLevel = windowArchive.FindElementByName("Compression level:");
            dropDownFieldCompressionLevel.SendKeys("9 - Ultra"); 

            var dropDownFieldCompressionMethod = windowArchive.FindElementByName("Compression method:");
            dropDownFieldCompressionMethod.SendKeys("* LZMA2");

            var buttonOk = windowArchive.FindElementByName("OK");
            buttonOk.Click();

            Thread.Sleep(500);
            inputFilePath.SendKeys(tempDirectory + @"\archive.7z" + Keys.Enter);

            var buttonExtract = driver.FindElementByName("Extract");
            buttonExtract.Click();

            var inputFieldCopyTo = driver.FindElementByName("Copy to:");
            inputFieldCopyTo.SendKeys(tempDirectory + Keys.Enter);
            Thread.Sleep(1000);

            FileAssert.AreEqual(ZipLocation, @"C:\temp\7zFM.exe");
        }
    } 
}