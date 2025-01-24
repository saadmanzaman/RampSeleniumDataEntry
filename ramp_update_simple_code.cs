using System;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Windows.Forms;
using System.Threading;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;

namespace VardaApp
{
    public partial class RampClassUpload
    {

        public const string rampTkn = "RAMP_TOKEN";
        public const string rampURI = "https://api.ramp.com";
        public const string rampEndPoint = "developer/v1/transactions/";

        public const string startDateTime = "2023-12-01T00:00:00.000000";
        public const string endDateTime = "2024-01-01T00:00:00.000000";

        public const string chromeUserProfile = "user-data-dir=C:\\Users\\Your_User_Name\\AppData\\Local\\Google\\Chrome\\User Data";

        //ramp transaction website xpath addresses
        public const string selPath_ID_AcctSection = "accounting-section"; //find by ID



        //TEST XPATHS IN DEVELOPER CONSOLE IN CHROME BEFORE YOU USE IT HERE TO SEE IF IT WORKS
        //IT WILL SAVE YOU A LOT OF SANITY

        //category
        public const string selPath_XPath_QBOCategoryInput = "//div[@id='accounting-section']//span[contains(text(), 'Accounting Category')]//ancestor::div[1]";

        //category item container
        public const string selPath_XPath_QBOCategoryItemContainer = "//div[contains(@class,'RyuListViewMain')]//div[contains(text(),'{0}')]";

        //class
        //public const string selPath_XPath_QBOClassInput = "//label/span[contains(text(),'QuickBooks Class')]/ancestor::label/following-sibling::div/input";
        public const string selPath_XPath_QBOClassInput = "//div[@id='accounting-section']//span[contains(text(), 'Accounting Project Code')]//ancestor::div[1]";
        //class item container
        public const string selPath_XPath_QBOClassItemContainer = "//div[contains(@class,'RyuListViewMain')]//div[contains(text(),'{0}')]";


        //mark ready button
        public const string markReadyButtonText = "//button[@data-id='button-accounting-mark-ready']/span[2]";
        public const string markReadyButton = "//button[@data-id='button-accounting-mark-ready']";

        //synced button
        public const string syncedButton = "//span[text()='Synced']";


        public void updateGL(IWebElement qbCategory, string glcode, Excel.Worksheet currentsheet1, dynamic actions, int i, WebDriverWait wait, IWebDriver driver)
        {
            //Will update GL code

            try
            {
                //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                //js.ExecuteScript("arguments[0].scrollIntoView();", qbCategory);

                string elementText = qbCategory.Text;

                if (elementText.Contains(glcode) != true)
                {
                    wait.Until(ExpectedConditions.ElementToBeClickable(qbCategory)).Click();

                    //qbCategory.Click();
                    Thread.Sleep(3000);//wait for drop down to appear
                                       //qbCategory.SendKeys(glcode);
                    actions.SendKeys(glcode).Perform(); //send keys without element since the focus is already on the element

                    Thread.Sleep(3000);
                    //actions.MoveToElement(qbCategory).SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                    //Thread.Sleep(2000);
                    //find the GL code in the drop down
                    string xPathText = string.Format(selPath_XPath_QBOCategoryItemContainer, glcode);
                    var itemContainerItems = driver.FindElements(By.XPath(xPathText));

                    wait.Until(ExpectedConditions.ElementToBeClickable(itemContainerItems[0])).Click();
                    Thread.Sleep(2000);
                }

            }
            catch
            {
                return;
            }


        }
        public void updateClass(IWebElement qbClass, string classcode, Excel.Worksheet currentsheet1, dynamic actions, int i, WebDriverWait wait, IWebDriver driver)
        {

            try
            {
                string elementText = qbClass.Text;

                if (elementText.Contains(classcode) != true)
                {
                    //Will update the Class code           
                    wait.Until(ExpectedConditions.ElementToBeClickable(qbClass)).Click();

                    Thread.Sleep(3000);//wait for drop down to appear
                                       //qbCategory.SendKeys(glcode);
                    actions.SendKeys(classcode).Perform(); //send keys without element since the focus is already on the element
                    Thread.Sleep(3000);

                    //substitute the xpath text using the project code so it can find it int he list
                    string xPathText = string.Format(selPath_XPath_QBOClassItemContainer, classcode);
                    var itemContainerItems = driver.FindElements(By.XPath(xPathText));

                    //click the project in the list
                    wait.Until(ExpectedConditions.ElementToBeClickable(itemContainerItems[0])).Click();
                    //itemContainerItems[0].Click();
                    Thread.Sleep(2000);

                }

            }
            catch
            {
                return;
            }


        }
        public (string status, IWebElement element) checkStatus(dynamic elements)
        {
            string xx = "Nothing";
            IWebElement n = elements;

            string textReturned = n.Text.ToLower().ToString();

            if (textReturned == "ready")// && ExpectedConditions.ElementToBeClickable(m).Equals(true)
            {
                xx = "Ready";
                //n = m;
            }
            else if (textReturned == "mark ready") // && ExpectedConditions.ElementToBeClickable(m).Equals(true)
            {
                xx = "Mark Ready";
                //n = m;
            }
            else if (textReturned == "synced")
            {
                xx = "Synced";
                //n = m;
            }

            return (xx, n);
        }



        public void updateProgress(string text)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook != null)
            {
                Excel.Worksheet logSheet = null;

                // Try to find the "Log" worksheet
                foreach (Excel.Worksheet sheet in activeWorkbook.Sheets)
                {
                    if (sheet.Name == "Log")
                    {
                        logSheet = sheet;
                        break;
                    }
                }

                // If the "Log" worksheet does not exist, create it
                if (logSheet == null)
                {
                    logSheet = (Excel.Worksheet)activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[activeWorkbook.Sheets.Count]);
                    logSheet.Name = "Log";
                }

                // Find the last used row in the "Log" worksheet
                Excel.Range lastCell = logSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = lastCell.Row;

                // If the last cell is not empty, move to the next row
                if (lastCell.Value2 != null)
                {
                    lastRow++;
                }

                // Write the log message to the next available row
                logSheet.Cells[lastRow, 1] = DateTime.Now.ToString();
                logSheet.Cells[lastRow, 2] = text;

                // Optionally, save the workbook after updating the log
                // activeWorkbook.Save();
            }
            else
            {
                // Handle the case where there is no active workbook
                Debug.WriteLine("There is no active workbook.");
            }
        }

        public void rampUpload()
        {

            Excel.Worksheet currentsheet1 = Globals.ThisAddIn.Application.ActiveSheet;



            if (currentsheet1.Name != "RampTransactionUpdater")
            {
                updateProgress("Sheet not found");
                return;
            }

            //Get sheet variables
            int i = 1;
            int x = currentsheet1.UsedRange.Rows.Count;

            string url;
            string glcode;
            string classcode;

            //updateProgress($"Starting, count of rows is {x}");

            try
            {
                //Set up chrome
                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument(chromeUserProfile);
                //chromeOptions.AddArgument("no-sandbox");
                //chromeOptions.AddArgument("--profile-directory=Profile 1");

                IWebDriver driver = new ChromeDriver(chromeOptions);



                //Loop through sheet
                for (i = 1; i <= x; i++)
                {
                    updateProgress($"Starting on row {i}.");
                    try
                    {
                        var actions = new OpenQA.Selenium.Interactions.Actions(driver);


                        //If there is a value in the 4th column, meaning that it was already reviewed, move to the next line
                        if (currentsheet1.Cells[i, 4].Value != null)
                        {
                            //currentsheet1.Cells[i, 5].Value = "Row isn't empty";
                            updateProgress("This row has been processed, moving onto the next row");
                            continue;
                        }

                        url = currentsheet1.Cells[i, 1].Value;
                        glcode = currentsheet1.Cells[i, 2].Value;
                        classcode = currentsheet1.Cells[i, 3].Value;

                        updateProgress($"GL Code is {glcode}, Classcode is {classcode}.");


                        driver.Url = url;
                        //Thread.Sleep(10000);                        


                        //Start parsing
                        //Wait 10 seconds for the element to be clickable
                        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                        //wait for element to become visible
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(selPath_XPath_QBOCategoryInput)));


                        //drill into category                        
                        var selElement_QBOCategoryInput = driver.FindElement(By.XPath(selPath_XPath_QBOCategoryInput));
                        updateProgress($"Value of category is {selElement_QBOCategoryInput.Text}");

                        //drill into class                        
                        var selElement_QBOClassInput = driver.FindElement(By.XPath(selPath_XPath_QBOClassInput));
                        updateProgress($"Value of class is {selElement_QBOClassInput.Text}");


                        //CHECK PAGE STATUS//
                        /* The page can be either capable of being "Mark Ready", "Ready", or "Synced".
                         * The mark button can be located in multiple locations, so hard coding the location doesn't work. Despite the location change
                         * The "class" name is still the same in all instances. So we will create a variable with all elements that fight the class
                         and loop through them to find some key words */

                        var markButtonsText = driver.FindElements(By.XPath(markReadyButtonText));

                        IWebElement markButton = driver.FindElement(By.XPath(markReadyButton));
                        updateProgress($"Mark button text is {markButtonsText[0].Text}.");
                        var syncButton = driver.FindElements(By.XPath(syncedButton));


                        //if Mark Ready button is missing, check if the transaction is Synced
                        if (markButtonsText.Count == 0)
                        {
                            if (syncButton.Count > 0)
                            {
                                //updateProgress("Synced");
                                currentsheet1.Cells[i, 4].Value = "Synced already";
                                updateProgress($"Transaction is synced already.");
                                continue;
                            }
                        }




                        //use method to find page status and the element with the status
                        string pageStatus = checkStatus(markButtonsText[0]).status;
                        IWebElement markButtonFound = checkStatus(markButtonsText[0]).element;
                        //currentsheet1.Cells[i, 6].Value = "Page Status " + pageStatus;
                        if (pageStatus == "Nothing")
                        {
                            updateProgress($"Page status is {pageStatus}, moving onto next row");
                            continue;
                        }
                        //else
                        //{
                        //    //updateProgress($"Page status is {pageStatus}, continuing.");                            
                        //}



                        IWebElement qbCategory = selElement_QBOCategoryInput; //QuickBooks Category                
                        IWebElement qbClass = selElement_QBOClassInput; //QuickBooks Class

                        //Debug
                        //currentsheet1.Cells[i, 5].Value = "Variable 10 is " + qbCategory.GetAttribute("value");
                        //currentsheet1.Cells[i, 5].Value = "Variable 14 is " + qbClass.GetAttribute("value");
                        //updateProgress($"The QBO Category is {qbCategory.GetAttribute("Value")}. It should be a {glcode}.");
                        //updateProgress($"The QBO Class is {qbClass.GetAttribute("Value")}. It should be {classcode}");

                        //UPDATE IF UPDATE IS NECESSARY, OR CLICK READY//



                        if (pageStatus == "Mark Ready")// && (qbCategory.GetAttribute("Value") != glcode || qbClass.GetAttribute("Value") != classcode))
                        {
                            updateGL(qbCategory, glcode, currentsheet1, actions, i, wait, driver);
                            Thread.Sleep(1000); //wait 1 sec
                            updateClass(qbClass, classcode, currentsheet1, actions, i, wait, driver);
                            Thread.Sleep(1000); //wait 1 sec
                                                //wait.Until(ExpectedConditions.ElementToBeClickable(markButton)).Click();

                            //wait.Until(ExpectedConditions.ElementToBeClickable(markButton));                            


                            currentsheet1.Cells[i, 4].Value = "Updated";
                            updateProgress("Updated and marked ready");

                        }
                        else if (pageStatus == "Ready")// && (qbCategory.GetAttribute("Value") != glcode || qbClass.GetAttribute("Value") != classcode))
                        {
                            wait.Until(ExpectedConditions.ElementToBeClickable(markButton)).Click();
                            //driver.Navigate().Refresh();
                            Thread.Sleep(3000);
                            //wait.Until(ExpectedConditions.ElementToBeClickable(markButton));
                            updateGL(qbCategory, glcode, currentsheet1, actions, i, wait, driver);
                            Thread.Sleep(2000); //wait 1 sec
                            updateClass(qbClass, classcode, currentsheet1, actions, i, wait, driver);
                            Thread.Sleep(1000); //wait 1 sec
                            //wait.Until(ExpectedConditions.ElementToBeClickable(markButton)).Click();

                            //wait.Until(ExpectedConditions.ElementToBeClickable(markButton));
                            currentsheet1.Cells[i, 4].Value = "Incorrectly Marked ready, re-updated and marked ready";
                            //updateProgress("Incorrectly Marked ready, re-updated and marked ready");
                        }
                        else if (pageStatus == "Synced")
                        {
                            currentsheet1.Cells[i, 4].Value = "Already Synced";
                            //updateProgress("Already synced. Moving onto next row.");                            
                            Thread.Sleep(1000); //wait 1 sec
                            continue;
                        }
                        //updateProgress("Moving onto next row");
                    }
                    catch (Exception ex)
                    {
                        currentsheet1.Cells[i, 4].Value = $"Error:{ex.InnerException.Message}";
                        //updateProgress($"Error:{ex.InnerException.Message}");
                        continue;
                    }



                };

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }

    }


}
