using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace NunitTeste2
{
    [TestClass]
    public class UnitTest1
    {

        readonly IWebDriver driver = new ChromeDriver();
        readonly string caminho = @"C:\Users\dsegatto\Desktop\Relevante\teste.txt";

        [TestInitialize]
        public void Initialize()
        {
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://g1.globo.com/");
        }

        [TestMethod]
        public void ExecuteTest()
        {
            IWebElement element = driver.FindElement(By.XPath(".//*[@id='glb-corpo']/div[4]/div/div/div/div/div[1]/ul/li[3]/a"));
            element.Click();
            getConteudo();
        }


        public void getConteudo()
        {
            var titulo = driver.FindElement(By.ClassName("content-head__title")).Text;
            var corpo = driver.FindElement(By.ClassName("content-text__container")).Text;
            var page = titulo + "\n\n" + corpo;

            salvaArquivo(caminho, page);
        }

        public void salvaArquivo( string pasta, string page)
        {
            File.WriteAllText(pasta, page);
        }

        public static ExchangeService conEmail()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            service.Credentials = new WebCredentials("denilson.segatto.silva@everis.com", "everis@2017");
            return service;
        }

        [TestMethod]
        public void enviaEmail()
        {
            EmailMessage email = new EmailMessage(conEmail());
            email.ToRecipients.Add("denilson.segatto.silva@everis.com");
            email.Subject = "teste";
            email.Body = "Teste";
            email.Attachments.AddFileAttachment(caminho);
            email.SendAndSaveCopy();
        }

        [TestCleanup]
        public void CleanUp()
        {
            driver.Quit();
        }


    }

}
