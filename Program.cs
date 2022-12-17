//Import all the necessary libraries
//Selenium
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
//Threading voor sleeps
using System.Threading;
using System.Threading.Tasks;
using System.IO;
//Excel en range voor wegscrijven naar excel
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;

namespace WebScraper
{
    class Program
    {
        //3 Functies voor het wegscrijven naar excel van de verschillende scrapes
        public static void writeExcel1(List<string> titles, List<string> uploaders, List<string> views, List<string> posts, List<string> links)
        {
            //maak een variabele waar je de filepath van de file naar waar we willen schrijven in zet
            string filepath = "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.xlsx";

            //excel applicatie declareren
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //workbook (excel document) variabele een naam geven
            Workbook wb;

            //worksheet (sheet in document) variabele een naam geven
            Worksheet ws;

            //open een excel document
            wb = excel.Workbooks.Open(filepath);

            //kies de sheet waarop je wilt werken in de workbook (de sheets onderaan een excel document)
            ws = wb.Worksheets[1];


            for (int i = 0; i < 5; i++)
            {
                //met range en de indexen van de items die we gaan wegschrijven kunnen we aanduiden in welke cellen op de worksheet onze data moet terecht komen. Titels komen hier dus in kolom A (rij 1 - 5). De getallen die ik aangeef met de index van het item in de lijst zijn hetzelfde. Dit doe ik omdat ik 1 waarde per keer aan het blad toevoeg. Ik moet dus een range van 1 cel bepalen hiervoor.
                Range cellRangeTitles = ws.Range[$"A{titles.IndexOf(titles[i]) + 1}:A{titles.IndexOf(titles[i]) + 1}"];

                //plaats voor elke waarde die we tegenkomen in de list deze waarde in het vakje dat hierboven bepaald is
                cellRangeTitles.Value = titles[i];

                Range cellRangeUploaders = ws.Range[$"B{uploaders.IndexOf(uploaders[i]) + 1}:B{uploaders.IndexOf(uploaders[i]) + 1}"];
                cellRangeUploaders.Value = uploaders[i];

                Range cellRangeViews = ws.Range[$"C{views.IndexOf(views[i]) + 1}:C{views.IndexOf(views[i]) + 1}"];
                cellRangeViews.Value = views[i];

                Range cellRangePosts = ws.Range[$"D{posts.IndexOf(posts[i]) + 1}:D{posts.IndexOf(posts[i]) + 1}"];
                cellRangePosts.Value = posts[i];

                Range cellRangeLinks = ws.Range[$"E{links.IndexOf(links[i]) + 1}:E{links.IndexOf(links[i]) + 1}"];
                cellRangeLinks.Value = links[i];
            }

            //sla file met nieuwe data op
            wb.SaveAs(filepath);

            //sluit excel bestand
            wb.Close();
        }

        public static void writeExcel2(List<string> jobs, List<string> companies, List<string> locations, List<string> keywords, List<string> links)
        {
            //maak een variabele waar je de filepath van de file naar waar we willen schrijven in zet
            string filepath = "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.xlsx";

            //excel applicatie declareren
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //workbook (excel document) variabele een naam geven
            Workbook wb;

            //worksheet (sheet in document) variabele een naam geven
            Worksheet ws;

            //open een excel document
            wb = excel.Workbooks.Open(filepath);

            //kies de sheet waarop je wilt werken in de workbook (de sheets onderaan een excel document)
            ws = wb.Worksheets[1];


            for (int i = 0; i < 5; i++)
            {
                //met range en de indexen van de items die we gaan wegschrijven kunnen we aanduiden in welke cellen op de worksheet onze data moet terecht komen. Titels komen hier dus in kolom A (rij 1 - 5). De getallen die ik aangeef met de index van het item in de lijst zijn hetzelfde. Dit doe ik omdat ik 1 waarde per keer aan het blad toevoeg. Ik moet dus een range van 1 cel bepalen hiervoor.
                Range cellRangeJobs = ws.Range[$"A{jobs.IndexOf(jobs[i]) + 1}:A{jobs.IndexOf(jobs[i]) + 1}"];

                //plaats voor elke waarde die we tegenkomen in de list deze waarde in het vakje dat hierboven bepaald is
                cellRangeJobs.Value = jobs[i];

                Range cellRangeCompanies = ws.Range[$"B{companies.IndexOf(companies[i]) + 1}:B{companies.IndexOf(companies[i]) + 1}"];
                cellRangeCompanies.Value = companies[i];

                Range cellRangeLocations = ws.Range[$"C{locations.IndexOf(locations[i]) + 1}:C{locations.IndexOf(locations[i]) + 1}"];
                cellRangeLocations.Value = locations[i];

                Range cellRangeKeywords = ws.Range[$"D{keywords.IndexOf(keywords[i]) + 1}:D{keywords.IndexOf(keywords[i]) + 1}"];
                cellRangeKeywords.Value = keywords[i];

                Range cellRangeLinks = ws.Range[$"E{links.IndexOf(links[i]) + 1}:E{links.IndexOf(links[i]) + 1}"];
                cellRangeLinks.Value = links[i];
            }

            //sla file met nieuwe data op
            wb.SaveAs(filepath);

            //sluit excel bestand
            wb.Close();
        }

        public static void writeExcel3(List<string> products, List<string> prices)
        {
            //maak een variabele waar je de filepath van de file naar waar we willen schrijven in zet
            string filepath = "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.xlsx";

            //excel applicatie declareren
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //workbook (excel document) variabele een naam geven
            Workbook wb;

            //worksheet (sheet in document) variabele een naam geven
            Worksheet ws;

            //open een excel document
            wb = excel.Workbooks.Open(filepath);

            //kies de sheet waarop je wilt werken in de workbook (de sheets onderaan een excel document)
            ws = wb.Worksheets[1];


            for (int i = 0; i < 5; i++)
            {
                //met range en de indexen van de items die we gaan wegschrijven kunnen we aanduiden in welke cellen op de worksheet onze data moet terecht komen. Titels komen hier dus in kolom A (rij 1 - 5). De getallen die ik aangeef met de index van het item in de lijst zijn hetzelfde. Dit doe ik omdat ik 1 waarde per keer aan het blad toevoeg. Ik moet dus een range van 1 cel bepalen hiervoor.
                Range cellRangeTitles = ws.Range[$"A{products.IndexOf(products[i]) + 1}:A{products.IndexOf(products[i]) + 1}"];

                //plaats voor elke waarde die we tegenkomen in de list deze waarde in het vakje dat hierboven bepaald is
                cellRangeTitles.Value = products[i];

                Range cellRangePosts = ws.Range[$"B{prices.IndexOf(prices[i]) + 1}:B{prices.IndexOf(prices[i]) + 1}"];
                cellRangePosts.Value = prices[i];
            }

            //sla file met nieuwe data op
            wb.SaveAs(filepath);

            //sluit excel bestand
            wb.Close();
        }

        //3 Functies voor het wegscrijven naar CSV van de verschillende scrapes
        public static void writeCSV1(string title, string uploader, string view, string post, string link, string filepath)
        {
            try
            {
                //Maak een streamwriter object aan om data weg te scrijven naar een file gedefinieerd door de filepath parameter van de methode. Het false argument betekent dat elke keer dat deze functie wordt uitgevoerd de file herschreven zal worden ipv dat de data geappend wordt.
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filepath, true))
                {
                    //Zet de meegekregen strings op een lijn tekst met komma's ertussen
                    file.WriteLine(title + "," + uploader + "," + view + "," + post + "," + link);
                }
            }
            //Als er iets misgaat, print dan dat er iets is misgegaan en wat er is misgegaan
            catch (Exception ex)
            {
                throw new ApplicationException("Something went wrong: ", ex);
            }
        }

        //Exact hetzelfde als de methode hierboven maar met andere parameternamen
        public static void writeCSV2(string job, string company, string location, string keyword, string link, string filepath)
        {
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filepath, true))
                {
                    file.WriteLine(job + "," + company + "," + location + "," + keyword + "," + link);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Something went wrong: ", ex);
            }
        }

        //Exact hetzelfde als de methode hierboven maar met minder parameters
        public static void writeCSV3(string product, string price, string filepath)
        {
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filepath, true))
                {
                    file.WriteLine(product + "," + price);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Something went wrong: ", ex);
            }
        }

        //Functie voor het scrapen van youtube
        public static void scrape1()
        {
            //Vraag de gebruiker naar wat hij wilt zoeken op youtube en sla de string op in een variabele.
            Console.Write("Enter what to search for on youtube: ");
            string searchtask = Console.ReadLine();

            var chromeOptions = new ChromeOptions();

            //PageLoadStrategy.Normal zorgt ervoor dat het programma wacht tot alles van de webpagina geladen is voor het verdergaat. Dit is nodig omdat het programma anders naar tags op een webpagina begint te zoeken voor die zijn ingeladen, wat tot errors leidt.
            chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

            //Maak een webdriver object aan met selenium en gebruik dat om te navigeren naar de url van de website. De chromeOptions is de PageLoadStrategy van hierboven.
            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Navigate().GoToUrl("https://www.youtube.com");


            //Wanneer je google voor de eerste keer opstart (wat we met selenium technisch gezien elke keer doen) moet je cookies accepteren. Met de volgende 2 lijnen code zoeken we naar de knop om die cookies te accepteren en klikken we erop.
            var cookies = driver.FindElement(By.XPath("//*[@id=\"content\"]/div[2]/div[6]/div[1]/ytd-button-renderer[2]/yt-button-shape/button/yt-touch-feedback-shape/div/div[2]"));
            cookies.Click();

            //Wacht nog een extra seconde tot de pagina volledig is geladen. Zelfs de PageLoadStrategy van hierboven werkt niet altijd volledig.
            System.Threading.Thread.Sleep(1000);

            //Zoek naar de zoekbalk van youtube.
            var searchbar = driver.FindElement(By.CssSelector("input#search"));

            //Vul de zoekterm die de gebruiker bovenaan deze methode heeft meegegeven in in de zoekbalk en enter.
            searchbar.SendKeys(searchtask);
            searchbar.Submit();

            //Wacht nog 2 extra seconden tot de pagina geladen is.
            System.Threading.Thread.Sleep(2000);

            //Maak 6 collecties van selenium IWebElements aan waarin de objecten waaruit we onze data gaan halen in opgeslagen worden. Ik heb online 1 manier gevonden om enkel de eerste 5 items van een type op te halen, maar dit werkte niet voor mij. Ik heb dus alle items van een type moeten ophalen en het op een andere manier moeten oplossen.
            ReadOnlyCollection<IWebElement> titles = driver.FindElements(By.XPath("//*[@id=\"video-title\"]/yt-formatted-string"));
            ReadOnlyCollection<IWebElement> posts = driver.FindElements(By.XPath("//*[@id=\"metadata-line\"]/span[2]"));
            ReadOnlyCollection<IWebElement> uploaders = driver.FindElements(By.XPath(".//*[@id=\"text\"]/a"));
            ReadOnlyCollection<IWebElement> views = driver.FindElements(By.XPath("//*[@id=\"metadata-line\"]/span[1]"));
            ReadOnlyCollection<IWebElement> links = driver.FindElements(By.Id("video-title"));
            ReadOnlyCollection<IWebElement> ads = driver.FindElements(By.XPath("//*[@id=\"ad-badge-container\"]/ytd-badge-supported-renderer/div/span"));

            //Maak verschillende lijsten aan om de data die we ophalen in op te slagen. Dit is nodig om enkel de eerste 5 items te kunnen verkrijgen en voor onze data later te kunnen wegscrijven naar een ander formaat.
            List<string> titleList = new List<string>();
            List<string> uploaderList = new List<string>();
            List<string> viewList = new List<string>();
            List<string> postList = new List<string>();
            List<string> linkList = new List<string>();

            //--------------------------- for loops om data in lijsten te zetten -----------------------------------------------

            //Om de posttijden op te halen die ik wou hebben om te zien hoe up to date een video is moet ik eerst checken of de content in de opgehaalde objecten hierboven eindigt met "ago". Dit is nodig aangezien er van dat tagtype meer tags worden opgehaald dan enkel de tags met posttijden. Op deze manier kan ik de posttijden uit die lijst van objecten filteren.
            for (int i = 0; i < 5; i++)
            {
                IWebElement post = posts[i];
                if (post.Text.EndsWith("ago"))
                {
                    //Voeg de opgehaalde string toe aan de lijst als hij eindigt op "ago".
                    postList.Add(post.Text);
                }
            }

            //For loop om de eerste 5 titels op te halen en aan de lijst toe te voegen.
            for (int i = 0; i < 5; i++)
            {
                IWebElement title = titles[i];
                titleList.Add(title.Text);
            }

            //For loop om de eerste 5 uploaders op te halen en aan de lijst toe te voegen. We beginnen hier vanaf index 2 tot en met index 12 omdat ook bij deze objecten om de beurt een string en een leeg object worden opgehaald. We kunnen de lege objecten dus overslaan door te werken met sprongen van 2. Het eerste zoekresultaat slaan we over omdat bij mijn voorbeeld dit altijd een advertentie was.
            for (int i = 2; i < 12; i += 2)
            {
                IWebElement uploader = uploaders[i];
                uploaderList.Add(uploader.Text);
            }

            //Deze for is om het aantal views op te halen. Zoals hierboven aangegeven haal ik ook tags op die ads zijn. Ik check eerst of er een ad in die lijst zit. Als dit het geval is moet het programma de views van die eerste tag overslaan, anders niet.
            //if (ads.Count > 0)
            //{
            //    for (int i = 1; i < 6; i++)
            //    {
            //        IWebElement view = views[i];
            //        viewList.Add(view.Text);
            //    }
            //}
            //else
            //{
            for (int i = 0; i < 5; i++)
            {
                IWebElement view = views[i];
                viewList.Add(view.Text);
            }
            //}

            //Zelfde geval zoals de vorige for loop maar dan voor de links van de video
            if (ads.Count > 0)
            {
                for (int i = 1; i < 6; i++)
                {
                    IWebElement link = links[i];
                    string linkvalue = link.GetAttribute("href");
                    linkList.Add(linkvalue);
                }
            }
            else
            {
                for (int i = 0; i < 5; i++)
                {
                    //De links van een job staan in een bepaald attribuut van de tag, niet als tekst in de tag. Daarom halen we enkel hier een attribuut op ipv tekst.
                    IWebElement link = links[i];
                    string linkvalue = link.GetAttribute("href");
                    linkList.Add(linkvalue);
                }
            }

            //In deze for loop print ik de opgehaalde data per video af in de console. Ook schrijf ik de data per video weg naar een lijn in een CSV formaat.
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine(titleList[i]);
                Console.WriteLine(uploaderList[i]);
                Console.WriteLine(viewList[i]);
                Console.WriteLine(postList[i]);
                Console.WriteLine(linkList[i]);
                //Print een lege lijn tussen elke video voor de duidelijkheid.
                Console.WriteLine();
                writeCSV1(titleList[i], uploaderList[i], viewList[i], postList[i], linkList[i], "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.txt");
            }

            //Schrijf de data ook weg naar excel.
            writeExcel1(titleList, uploaderList, viewList, postList, linkList);
        }

        //Functie voor het scrapen van ictjobs
        public static void scrape2()
        {
            //Vraag de gebruiker naar wat hij wilt zoeken op ictjob en sla de string op in een variabele.
            Console.Write("Enter what to search for on ictjob: ");
            string searchtask = Console.ReadLine();

            var chromeOptions = new ChromeOptions();

            //PageLoadStrategy.Normal zorgt ervoor dat het programma wacht tot alles van de webpagina geladen is voor het verdergaat. Dit is nodig omdat het programma anders naar tags op een webpagina begint te zoeken voor die zijn ingeladen, wat tot errors leidt.
            chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

            //Maak een webdriver object aan met selenium en gebruik dat om te navigeren naar de url van de website. De chromeOptions is de PageLoadStrategy van hierboven.
            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Navigate().GoToUrl("https://www.ictjob.be/nl/");

            //Wacht nog een extra seconde tot de pagina volledig is geladen. Zelfs de PageLoadStrategy van hierboven werkt niet altijd volledig.
            System.Threading.Thread.Sleep(1000);


            //Zoek naar de zoekbalk van youtube.
            var searchbar = driver.FindElement(By.XPath("//*[@id=\"keywords-input\"]"));

            //Vul de zoekterm die de gebruiker bovenaan deze methode heeft meegegeven in in de zoekbalk en enter.
            searchbar.SendKeys(searchtask);
            searchbar.Submit();

            //Wacht nog 2 seconden nadat de pagina geladen is. Het duurt erg lang om de pagina te laden waarvan we onze data ophalen. Dit zorgt er ook voor dat ons programma af en toe toch vastloopt hier omdat het soms toch begint te zoeken naar tags voor de pagina geladen is.
            System.Threading.Thread.Sleep(2000);

            //Maak 5 collecties van selenium IWebElements aan waarin de objecten waaruit we onze data gaan halen in opgeslagen worden. Ik heb online 1 manier gevonden om enkel de eerste 5 items van een type op te halen, maar dit werkte niet voor mij. Ik heb dus alle items van een type moeten ophalen en het op een andere manier moeten oplossen.
            ReadOnlyCollection<IWebElement> jobs = driver.FindElements(By.ClassName("job-title"));
            ReadOnlyCollection<IWebElement> companies = driver.FindElements(By.ClassName("job-company"));
            ReadOnlyCollection<IWebElement> locations = driver.FindElements(By.CssSelector("span[itemprop=\"addressLocality\"]"));
            ReadOnlyCollection<IWebElement> keywords = driver.FindElements(By.ClassName("job-keywords"));
            ReadOnlyCollection<IWebElement> links = driver.FindElements(By.ClassName("search-item-link"));

            //Maak verschillende lijsten aan om de data die we ophalen in op te slagen. Dit is nodig om enkel de eerste 5 items te kunnen verkrijgen en voor onze data later te kunnen wegscrijven naar een ander formaat.
            List<string> jobList = new List<string>();
            List<string> companyList = new List<string>();
            List<string> locationList = new List<string>();
            List<string> keywordList = new List<string>();
            List<string> linkList = new List<string>();

            //Bij deze for loop voor de namen van de jobs halen we weer elke 2e tag op uit de collectie omdat ook hier de helft van de opgehaalde tags leeg is.
            for (int i = 0; i < 10; i += 2)
            {
                IWebElement job = jobs[i];
                jobList.Add(job.Text);
            }

            //For loop om de eerste 5 van alle andere gegevens op te halen.
            for (int i = 0; i < 5; i++)
            {
                //bedrijfsnamen
                IWebElement company = companies[i];
                companyList.Add(company.Text);
                //locaties
                IWebElement location = locations[i];
                locationList.Add(location.Text);
                //sleutelwoorden
                IWebElement keyword = keywords[i];
                keywordList.Add(keyword.Text);
                //De links van een job staan in een bepaald attribuut van de tag, niet als tekst in de tag. Daarom halen we enkel hier een attribuut op ipv tekst.
                IWebElement link = links[i];
                string linkvalue = link.GetAttribute("href");
                linkList.Add(linkvalue);
            }


            //--------------------------------------------------------------------------------------------------------------

            //In deze for loop print ik de opgehaalde data per job af in de console. Ook schrijf ik de data per video weg naar een lijn in een CSV formaat.
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine(jobList[i]);
                Console.WriteLine(companyList[i]);
                Console.WriteLine(locationList[i]);
                Console.WriteLine(keywordList[i]);
                Console.WriteLine(linkList[i]);
                writeCSV2(jobList[i], companyList[i], locationList[i], keywordList[i], linkList[i], "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.txt");
                //Print een lege lijn tussen elke job voor de duidelijkheid
                Console.WriteLine();
            }

            //Schrijf de data ook weg naar excel.
            writeExcel2(jobList, companyList, locationList, keywordList, linkList);
        }

        //Functie voor het scrapen van coolblue
        public static void scrape3()
        {
            //Vraag de gebruiker naar wat hij wilt zoeken op youtube en sla de string op in een variabele.
            Console.Write("Enter what to search for on Coolblue: ");
            string searchtask = Console.ReadLine();


            var chromeOptions = new ChromeOptions();

            //PageLoadStrategy.Normal zorgt ervoor dat het programma wacht tot alles van de webpagina geladen is voor het verdergaat. Dit is nodig omdat het programma anders naar tags op een webpagina begint te zoeken voor die zijn ingeladen, wat tot errors leidt.
            chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

            //Maak een webdriver object aan met selenium en gebruik dat om te navigeren naar de url van de website. De chromeOptions is de PageLoadStrategy van hierboven.
            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Navigate().GoToUrl("https://www.coolblue.be/nl");


            //Wanneer je google voor de eerste keer opstart (wat we met selenium technisch gezien elke keer doen) moet je cookies accepteren. Met de volgende 2 lijnen code zoeken we naar de knop om die cookies te accepteren en klikken we erop.
            var cookies = driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div/div[1]/div/div[1]/form/div[2]/button"));
            cookies.Click();

            //Wacht nog 1 extra seconde tot de pagina geladen is.
            System.Threading.Thread.Sleep(1000);


            //Zoek naar de zoekbalk van coolblue.
            var searchbar = driver.FindElement(By.XPath("//*[@id=\"search_query\"]"));

            //Vul de zonet gevraagde zoekterm in de zoekbalk.
            searchbar.SendKeys(searchtask);
            //Bevestig de zoekterm.
            searchbar.Submit();

            System.Threading.Thread.Sleep(2000);

            //Klik op de dropdown met rangschikopties.
            driver.FindElement(By.XPath("//*[@id=\"sorting\"]")).Click();

            //Klik op de knop met "prijs laag - hoog" om te rangschikken op prijs. (ik ben een arme student).
            driver.FindElement(By.XPath("//*[@id=\"sorting\"]/option[2]")).Click();

            System.Threading.Thread.Sleep(2000);

            //Maak 3 collecties van selenium IWebElements aan waarin de objecten waaruit we onze data gaan halen in opgeslagen worden. Ik heb online 1 manier gevonden om enkel de eerste 5 items van een type op te halen, maar dit werkte niet voor mij. Ik heb dus alle items van een type moeten ophalen en het op een andere manier moeten oplossen.
            ReadOnlyCollection<IWebElement> products = driver.FindElements(By.XPath("/html/body/main/div[3]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div[3]/div/div/a"));
            ReadOnlyCollection<IWebElement> prices = driver.FindElements(By.CssSelector("strong[class='sales-price__current js-sales-price-current']"));

            //Maak verschillende lijsten aan om de data die we ophalen in op te slagen. Dit is nodig om enkel de eerste 5 items te kunnen verkrijgen en voor onze data later te kunnen wegscrijven naar een ander formaat.
            List<string> productList = new List<string>();
            List<string> priceList = new List<string>();


            //----------------------------------------------------------------------------------------------------------------------

            //For loop om de eerste 5 items van alle collecties op te halen en aan de lijst toe te voegen.
            for (int i = 0; i < 5; i++)
            {
                IWebElement product = products[i];
                IWebElement price = prices[i];
                productList.Add(product.Text);
                priceList.Add(price.Text);
            }

            //In deze for loop print ik de opgehaalde data per video af in de console. Ook schrijf ik de data per video weg naar een lijn in een CSV formaat.
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine(productList[i]);
                Console.WriteLine(priceList[i]);
                Console.WriteLine();
                writeCSV3(productList[i], priceList[i], "C:\\Users\\jenve\\Desktop\\School\\2e bachelor\\DevOps & Security\\webscrape_data.txt");
            }

            //Schrijf de data ook weg naar excel.
            writeExcel3(productList, priceList);
        }

        //Main programma waarin je soort scrape kan kiezen
        static void Main(string[] args)
        {
            //Vraag de gebruiker welke scrape hij/zij wilt doen.
            Console.WriteLine("1 = Youtube Scrape");
            Console.WriteLine("2 = ICTJob Scrape");
            Console.WriteLine("3 = Coolblue Scrape");
            Console.Write("What scrape do you want to do?: ");

            //Sla antwoord op in een variabele.
            string scrapetype = Console.ReadLine();

            //Herhaal vorige stap zolang de gebruiker geen van de opgelijstte opties ingeeft.
            while ((scrapetype != "1") && (scrapetype != "2") && (scrapetype != "3"))
            {
                Console.WriteLine("Please choose a listed option:");
                Console.WriteLine("1 = Youtube Scrape");
                Console.WriteLine("2 = ICTJob Scrape");
                Console.WriteLine("3 = Coolblue Scrape");
                Console.Write("What scrape do you want to do?: ");
                scrapetype = Console.ReadLine();
            }

            //Voeg volgende functie uit als de gebruiker optie 1 kiest.
            if (scrapetype == "1")
            {
                scrape1();
            }

            //Voer deze functie uit als de gebruiker optie 2 kiest.
            if (scrapetype == "2")
            {
                scrape2();
            }

            //Voer scrape3 uit als de gerbuiker optie 3 kiest.
            if (scrapetype == "3")
            {
                scrape3();
            }
        }
    }
}
