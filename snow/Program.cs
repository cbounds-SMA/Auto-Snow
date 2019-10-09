using System;
using System.IO;
using System.Threading;
using System.Xml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace snow
{
    class Program
    {

        static void Main(string[] args)
        {
            string usrpath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString();
            string tktpath = @usrpath + "\\inc-ticket";
            string oldtktpath = usrpath + "\\inc-ticket\\old";
            System.IO.Directory.CreateDirectory(oldtktpath);

            searchDirectory(tktpath);

        }

        private static void searchDirectory(string tktpath)
        {
            string[] fileEntries = Directory.GetFiles(tktpath);
            foreach (string fileName in fileEntries)
                readXmlFile(fileName);
        }

        private static void readXmlFile(string filepath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filepath);
            XmlNodeList ticketNodes = xmlDoc.SelectNodes("Ticket");

            foreach (XmlNode ticketNode in ticketNodes)
            {
                XmlNode userNode = ticketNode.SelectSingleNode("usr");
                XmlNode compNode = ticketNode.SelectSingleNode("computerName");
                XmlNode titleNode = ticketNode.SelectSingleNode("requestTitle");
                XmlNode category = ticketNode.SelectSingleNode("category");
                XmlNode subcategory = ticketNode.SelectSingleNode("sub-category");
                XmlNode subcategory2 = ticketNode.SelectSingleNode("sub-category2");
                XmlNode descriptionNode = ticketNode.SelectSingleNode("description");
                XmlNode isResolvedNode = ticketNode.SelectSingleNode("isResolved");
                XmlNode resolutionTextNode = ticketNode.SelectSingleNode("resolution");
                XmlNode resolutionCausePrimaryCode = ticketNode.SelectSingleNode("causecodePrimary");
                XmlNode resolutionCauseCodeSub1 = ticketNode.SelectSingleNode("causeCodeSub1");
                XmlNode resolutionCode = ticketNode.SelectSingleNode("resolutionCode");
                XmlNode closeByNode = ticketNode.SelectSingleNode("resolvedBy");

                var ticketData = new System.Collections.Generic.Dictionary<string, string>();
                ticketData.Add("usr", userNode.InnerText);
                ticketData.Add("computerName", compNode.InnerText);
                ticketData.Add("title", titleNode.InnerText);
                ticketData.Add("category", category.InnerText);
                ticketData.Add("subcategory", subcategory.InnerText);
                ticketData.Add("subcategory2", subcategory2.InnerText);
                ticketData.Add("description", descriptionNode.InnerText);
                ticketData.Add("isResolved", isResolvedNode.InnerText);

                if (isResolvedNode.InnerText == "true")
                {
                    ticketData.Add("resolutionNotes", resolutionTextNode.InnerText);
                    ticketData.Add("resolutionCausePrimaryCode", resolutionCausePrimaryCode.InnerText);
                    ticketData.Add("resolutionCauseCodeSub1", resolutionCauseCodeSub1.InnerText);
                    ticketData.Add("resolutionCode", resolutionCode.InnerText);
                    ticketData.Add("closedBy", closeByNode.InnerText);
                }

                createTicket(ticketData);
                moveFile(filepath);
                Thread.Sleep(15000);
            }
        }

        private static void moveFile(string filepath)
        {
            Random random = new Random();
            int num = random.Next();
            string usrpath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString();
            string filename = Path.GetFileName(filepath);
            string destinationFile = @usrpath + "\\ticket\\old\\" + filename;
            // To move a file or folder to a new location:
            System.IO.File.Move(filepath, destinationFile);
        }

        private static void createTicket(System.Collections.Generic.Dictionary<string, string> ticketData)
        {

            Random random = new Random();
            int longsleep = random.Next(30000, 60000);
            int mediumsleep = random.Next(7500, 15000);
            int shortsleep = random.Next(1500, 3000);
            //[Start Chrome driver]

            try
            {
                IWebDriver driver = new ChromeDriver();
                //[Start nagivagation to main atos snow sma portal for SSO]
                driver.Url = "https://atosglobal.service-now.com/sp/?id=sso&portal-id=sma";
                Thread.Sleep(mediumsleep);
                driver.Navigate().Refresh();
                Thread.Sleep(mediumsleep);
                driver.Url = "https://atosglobal.service-now.com/ui_page.do?sys_id=da091c880fd8f1001668956f62050e9e";
                //[End naviagation to main atos snow sma portal for SSO wait 15 seconds]\
                Thread.Sleep(mediumsleep);
                driver.Navigate().Refresh();
                Thread.Sleep(mediumsleep);

    

                IWebElement element = driver.FindElement(By.Id("sys_display.user"));
                element.SendKeys(ticketData["usr"]);
                element.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                driver.FindElement(By.XPath("/html/body/div[3]/div/button")).Click();
                Thread.Sleep(mediumsleep);
                IWebElement incContactType = driver.FindElement(By.Id("incident.contact_type"));
                incContactType.SendKeys("Email");
                incContactType.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement incidentCategory = driver.FindElement(By.Id("incident.category"));
                incidentCategory.SendKeys(ticketData["category"]);
                incidentCategory.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement incidentSubCategory1 = driver.FindElement(By.Id("incident.subcategory"));
                incidentSubCategory1.SendKeys(ticketData["subcategory"]);
                incidentSubCategory1.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement incidentSubCategory2 = driver.FindElement(By.Id("incident.u_category3"));
                incidentSubCategory2.SendKeys(ticketData["subcategory2"]);
                incidentSubCategory2.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement cmdb_ci = driver.FindElement(By.Id("sys_display.incident.cmdb_ci"));
                cmdb_ci.SendKeys(ticketData["computerName"]);
                cmdb_ci.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement shortDescription = driver.FindElement(By.Id("incident.short_description"));
                shortDescription.SendKeys(ticketData["title"]);
                shortDescription.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                IWebElement description = driver.FindElement(By.Id("incident.description"));
                description.SendKeys(ticketData["description"]);
                description.SendKeys(Keys.Return);
                Thread.Sleep(shortsleep);
                
                driver.FindElement(By.XPath("//*[@id='6685b1c93744d7849d3b861754990ef8']")).Click();
                Thread.Sleep(longsleep);
                if (ticketData["isResolved"] == "true")
                {
                    IWebElement assigned_to = driver.FindElement(By.Id("sys_display.incident.assigned_to"));
                    assigned_to.SendKeys(ticketData["closedBy"]);
                    assigned_to.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    IWebElement incident_state = driver.FindElement(By.Id("incident.incident_state"));
                    incident_state.SendKeys("Resolved");
                    incident_state.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    driver.FindElement(By.XPath("//*[@id='tabs2_section']/span[4]/span[1]/span[2]")).Click();
                    Thread.Sleep(mediumsleep);
                    IWebElement primary_close_code = driver.FindElement(By.Id("incident.u_sub_close_code"));
                    primary_close_code.SendKeys(ticketData["resolutionCausePrimaryCode"]);
                    primary_close_code.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    IWebElement close_code_2 = driver.FindElement(By.Id("incident.u_sub_close_code_2"));
                    close_code_2.SendKeys(ticketData["resolutionCauseCodeSub1"]);
                    close_code_2.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    IWebElement resolution_code = driver.FindElement(By.Id("incident.u_resolution_code"));
                    resolution_code.SendKeys(ticketData["resolutionCode"]);
                    resolution_code.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    IWebElement close_code = driver.FindElement(By.Id("incident.close_code"));
                    close_code.SendKeys("Resolved");
                    close_code.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    IWebElement close_notes = driver.FindElement(By.Id("incident.close_notes"));
                    close_notes.SendKeys(ticketData["resolutionNotes"]);
                    close_notes.SendKeys(Keys.Return);
                    Thread.Sleep(shortsleep);
                    driver.FindElement(By.XPath("//*[@id='6685b1c93744d7849d3b861754990ef8']")).Click();
                }
                Thread.Sleep(longsleep);
                driver.Close();
            }
            catch
            {
                Console.WriteLine("Error creating ticket");
            }
        }
    }
}