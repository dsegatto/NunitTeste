using Microsoft.Exchange.WebServices.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;



namespace NunitTeste
{

    class Program
    {

        static void Main(string[] args)
        {
            // Method intentionally left empty.
        }

        readonly IWebDriver driver = new ChromeDriver();

        public void Initialize()
        {
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://g1.globo.com/");            
        }
        public void ExecuteTest()
        {
            IWebElement element = driver.FindElement(By.XPath(".//*[@id='glb-corpo']/div[4]/div/div/div/div/div[1]/ul/li[3]/a"));
            element.Click();

            var titulo = driver.FindElement(By.ClassName("content-head__title")).Text;
            var corpo = driver.FindElement(By.ClassName("content-text__container")).Text;

            var page = "<h1>"+titulo+"</h1>" + "<br/><br/>" + "<p>"+corpo+"</p>";

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.UseDefaultCredentials = true;
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            service.Credentials = new WebCredentials("denilson.segatto.silva@everis.com", "everis@2017", "USERSAD");


            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("denilson.segatto.silva@everis.com");
            email.Subject = "teste";
            email.Body = new MessageBody(page);
            email.SendAndSaveCopy();
        }
 

        public void CleanUp()
        {
            driver.Quit();
        }


    }
}
