using ClosedXML.Excel;
using FNSoftwareBot.Services;
using JobDataScraper.Models;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace JobDataScraper.Services
{
    public class InappJobAtlasService
    {
        #region Variables and Properties

        private Task MainTask;

        private bool _continueRun = true;

        private string _baseUrl = "https://atlantelavoro.inapp.org/atlante_professioni.php";

        private string _fileName = "InappJobAtlas.xlsx";

        #endregion

        public void Main()
        {
            MainTask = Task.Run(() => Run());

            ConsoleKeyInfo consoleKeyInfo;
            do
            {
                consoleKeyInfo = Console.ReadKey(true);

                if (consoleKeyInfo.Key == ConsoleKey.Q)
                {
                    _continueRun = false;

                    Log.WriteLogMessage($"The program will close automatically when all the operation are finished.", LogMessageType.Info);

                    while (!MainTask.IsCompleted)
                    {
                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    }
                }
            } while (consoleKeyInfo.Key != ConsoleKey.Q);

            // Wait for use to close
            Log.WriteLogMessage($"Press 'Enter' to close the window.", LogMessageType.Info);
            Console.ReadLine();
        }

        #region Private Methods

        public void Run()
        {
            List<AdaJobModel> adaJobModelList = new List<AdaJobModel>();

            try
            {
                Log.WriteLogMessage($"InappJobAtlas Scraper - {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}", LogMessageType.Info);
                Log.WriteLogMessage($"To close the program, you must press the letter Q and wait for the program to close itself", LogMessageType.Info);
                Log.WriteLogMessage($"Loading...", LogMessageType.Info);

                //ReadExcel(out List<string> alreadyDoneList);

                using IWebDriver driver =
                    FNSoftwareBot.Services.DriverService.InitializeDriver(
                        Settings.Settings.BOT_BROWSER,
                        Settings.Settings.BOT_HEADLESS,
                        true,
                        TimeSpan.FromSeconds(30),
                        TimeSpan.FromMilliseconds(200),
                        PageLoadStrategy.Default
                    );

                driver.Navigate().GoToUrl(_baseUrl);

                driver.FindElement(By.XPath("//a[@href=\"#first\"]")).Click();

                IWebElement allCategoryButton = driver.FindElement(By.XPath("//a[@href=\"javascript:carica('templates/sql_repertori_settori_QNQR_ajax.php?codice_repertorio=AP&id_cat_cnel=','repertori_all_cnel_AP_tutti')\"]"));
                driver.ScrollElementIntoView(allCategoryButton);

                Thread.Sleep(TimeSpan.FromSeconds(2));

                allCategoryButton.Click();

                Thread.Sleep(TimeSpan.FromSeconds(2));

                List<Tuple<string, string, string>> categoryCompleteList = new List<Tuple<string, string, string>>();

                List<IWebElement> categoryList =
                    driver
                    .FindElements(By.XPath("//div[@id='repertori_all_cnel_AP_tutti']//div[@class='panel panel-primary']"))
                    .Skip(Settings.Settings.BOT_SKIP)
                    .Take(Settings.Settings.BOT_TAKE)
                    .ToList();

                foreach (IWebElement category in categoryList)
                {
                    string categoryName = category.FindElement(By.XPath(".//div[@class='panel-heading']//a")).Text;

                    category.FindElement(By.XPath(".//div[@class='panel-heading']//a")).Click();

                    List<IWebElement> categoryElementList =
                        category
                        .FindElements(By.XPath(".//div[@class='panel-collapse collapse show']//a"))
                        .ToList();

                    foreach (IWebElement categoryElement in categoryElementList)
                    {
                        string categoryElementName = categoryElement.GetAttribute("innerText");
                        string categoryElementLink = categoryElement.GetAttribute("href");

                        categoryElementName = categoryElementName.Trim();
                        categoryElementLink = categoryElementLink.Trim();

                        categoryCompleteList.Add(new Tuple<string, string, string>(categoryName, categoryElementName, categoryElementLink));
                    }
                }

                foreach (Tuple<string, string, string> categoryComplete in categoryCompleteList)
                {
                    Log.WriteLogMessage($"Doing: {categoryComplete.Item3}", LogMessageType.Info);

                    driver.Navigate().GoToUrl(categoryComplete.Item3);

                    List<string> categoryAdaUrlList =
                        driver
                        .FindElements(By.XPath("//a[contains(@href, \"dettaglio_ada_pre\")]"))
                        .Select(we => we.GetAttribute("href"))
                        .ToList();

                    foreach (string categoryAdaUrl in categoryAdaUrlList)
                    {
                        Log.WriteLogMessage($"Doing: {categoryAdaUrl}", LogMessageType.Info);

                        try
                        {
                            driver.Navigate().GoToUrl(categoryAdaUrl);

                            var adaJobModel = new AdaJobModel()
                            {
                                CategoryName = categoryComplete.Item1,
                                CategoryDescription = categoryComplete.Item2
                            };

                            string titleElementText = driver.FindElement(By.XPath("//li[@class='breadcrumb-item active']")).Text;

                            if (titleElementText.StartsWith("/DETTAGLIO"))
                                titleElementText = titleElementText.Replace("/DETTAGLIO", string.Empty);
                            if (titleElementText.StartsWith("DETTAGLIO"))
                                titleElementText = titleElementText.Replace("DETTAGLIO", string.Empty);

                            if (titleElementText.Contains("("))
                                titleElementText = titleElementText.Replace(titleElementText.Substring(titleElementText.IndexOf("("), titleElementText.IndexOf(")") - titleElementText.IndexOf("(") + 1), string.Empty);

                            adaJobModel.Code = titleElementText.Split("-")[0].Trim();
                            adaJobModel.Description = titleElementText.Split("-")[1].Trim();

                            List<IWebElement> containerElementList = driver.FindElements(By.XPath("//section[contains(@class, \"overview-block\")]//div[contains(@class,\"container-fluid\")]")).ToList();

                            foreach (IWebElement containerElement in containerElementList)
                            {
                                try
                                {
                                    List<IWebElement> columnElementList = containerElement.FindElements(By.XPath(".//div[@class=\"col-sm-4\"]")).ToList();

                                    if (columnElementList.Count == 0)
                                        columnElementList = containerElement.FindElements(By.XPath(".//div[@class=\"col-sm-6\"]")).ToList();

                                    List<IWebElement> taskElementList = columnElementList[0].FindElements(By.XPath(".//p")).ToList();

                                    List<AdaTask> adaTaskList = new List<AdaTask>();

                                    foreach (IWebElement taskElement in taskElementList)
                                    {
                                        string taskDescription = taskElement.GetAttribute("innerText");

                                        if (taskDescription.Contains("\n"))
                                            taskDescription = taskDescription.Replace("\n", string.Empty);

                                        if (taskDescription.Contains("\r"))
                                            taskDescription = taskDescription.Replace("\r", string.Empty);

                                        taskDescription = taskDescription.Trim();

                                        adaTaskList.Add(new AdaTask() { Description = taskDescription });
                                    }

                                    if (taskElementList.Count > 1)
                                    {
                                        string expectedResultDescription = columnElementList[1].FindElement(By.XPath(".//p")).GetAttribute("innerText");

                                        if (expectedResultDescription.Contains("\n"))
                                            expectedResultDescription = expectedResultDescription.Replace("\n", string.Empty);

                                        if (expectedResultDescription.Contains("\r"))
                                            expectedResultDescription = expectedResultDescription.Replace("\r", string.Empty);

                                        expectedResultDescription = expectedResultDescription.Trim();

                                        if (adaTaskList.Count > 0)
                                            adaTaskList[0].ExpectedResultDescription = expectedResultDescription;
                                    }

                                    adaJobModel.Tasks.AddRange(adaTaskList);
                                }
                                catch (Exception) { }

                            }

                            List<IWebElement> associationList = driver.FindElements(By.XPath("//div[@id='accordion-cp2011']//tbody//tr")).ToList();

                            foreach (IWebElement association in associationList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = association.FindElements(By.XPath(".//td")).ToList();

                                    string associationCode = columnList[0].GetAttribute("innerText");
                                    string associationName = columnList[1].GetAttribute("innerText");

                                    if (associationCode.Contains("\n"))
                                        associationCode = associationCode.Replace("\n", string.Empty);

                                    if (associationCode.Contains("\r"))
                                        associationCode = associationCode.Replace("\r", string.Empty);

                                    associationCode = associationCode.Trim();

                                    if (associationName.Contains("\n"))
                                        associationName = associationName.Replace("\n", string.Empty);

                                    if (associationName.Contains("\r"))
                                        associationName = associationName.Replace("\r", string.Empty);

                                    associationName = associationName.Trim();

                                    adaJobModel.Associations.Add(new AdaJobAssociation()
                                    {
                                        Code = associationCode,
                                        Name = associationName
                                    });
                                }
                                catch (Exception) { }
                            }

                            adaJobModelList.Add(adaJobModel);

                            Log.WriteLogMessage($"Done: {categoryAdaUrl}", LogMessageType.Success);

                        }
                        catch (Exception ex)
                        {
                            Log.WriteLogMessage($"Error: {categoryAdaUrl} - {ex.Message}", LogMessageType.Error);
                        }

                        Thread.Sleep(TimeSpan.FromSeconds(2));
                    }

                    Log.WriteLogMessage($"Done: {categoryComplete.Item3}", LogMessageType.Success);

                    Log.WriteLogMessage($"Writing File", LogMessageType.Info);

                    WriteExcel(adaJobModelList);

                    adaJobModelList.Clear();

                    Log.WriteLogMessage($"File Wrote", LogMessageType.Success);

                    Thread.Sleep(TimeSpan.FromSeconds(1));
                }

                //List<string> adaUrlList =
                //    driver
                //    .FindElements(By.XPath("//a[contains(@href, 'dettaglio_ada.php')]"))
                //    .Skip(Settings.Settings.BOT_SKIP)
                //    .Take(Settings.Settings.BOT_TAKE)
                //    .Select(we => we.GetAttribute("href"))
                //    .ToList();

                //foreach (string adaUrl in adaUrlList)
                //{
                //    Log.WriteLogMessage($"Doing: {adaUrl}", LogMessageType.Info);

                //    driver.Navigate().GoToUrl(adaUrl);

                //    var adaJobModel = new AdaJobModel();

                //    string titleElementText = driver.FindElement(By.XPath("//li[@class='breadcrumb-item active']")).Text;

                //    if (titleElementText.StartsWith("/DETTAGLIO"))
                //        titleElementText = titleElementText.Replace("/DETTAGLIO", string.Empty);
                //    if (titleElementText.StartsWith("DETTAGLIO"))
                //        titleElementText = titleElementText.Replace("DETTAGLIO", string.Empty);

                //    if (titleElementText.Contains("("))
                //        titleElementText = titleElementText.Replace(titleElementText.Substring(titleElementText.IndexOf("("), titleElementText.IndexOf(")") - titleElementText.IndexOf("(") + 1), string.Empty);

                //    adaJobModel.Code = titleElementText.Split("-")[0].Trim();
                //    adaJobModel.Description = titleElementText.Split("-")[1].Trim();

                //    List<IWebElement> taskList = driver.FindElements(By.XPath("//div[@id='accordion-attivitaList']//tbody//tr")).ToList();

                //    foreach (IWebElement task in taskList)
                //    {
                //        try
                //        {
                //            string taskText = task.GetAttribute("innerText");

                //            if (taskText.Contains("\n"))
                //                taskText = taskText.Replace("\n", string.Empty);

                //            if (taskText.Contains("\r"))
                //                taskText = taskText.Replace("\r", string.Empty);

                //            taskText = taskText.Trim();

                //            adaJobModel.Tasks.Add(taskText);
                //        }
                //        catch (Exception) { }
                //    }

                //    List<IWebElement> associationList = driver.FindElements(By.XPath("//div[@id='accordion-cp2011']//tbody//tr")).ToList();

                //    foreach (IWebElement association in associationList)
                //    {
                //        try
                //        {
                //            List<IWebElement> columnList = association.FindElements(By.XPath(".//td")).ToList();

                //            string associationCode = columnList[0].GetAttribute("innerText");
                //            string associationName = columnList[1].GetAttribute("innerText");

                //            if (associationCode.Contains("\n"))
                //                associationCode = associationCode.Replace("\n", string.Empty);

                //            if (associationCode.Contains("\r"))
                //                associationCode = associationCode.Replace("\r", string.Empty);

                //            associationCode = associationCode.Trim();

                //            if (associationName.Contains("\n"))
                //                associationName = associationName.Replace("\n", string.Empty);

                //            if (associationName.Contains("\r"))
                //                associationName = associationName.Replace("\r", string.Empty);

                //            associationName = associationName.Trim();

                //            adaJobModel.Associations.Add(new AdaJobAssociation()
                //            {
                //                Code = associationCode,
                //                Name = associationName
                //            });
                //        }
                //        catch (Exception) { }
                //    }

                //    adaJobModelList.Add(adaJobModel);

                //    Log.WriteLogMessage($"Done: {adaUrl}", LogMessageType.Success);

                //    if (!_continueRun) break;
                //}
            }
            catch (Exception ex)
            {
                Log.WriteLogMessage($"Error: {ex.Message}", LogMessageType.Error);
                Console.ReadLine();
            }
            finally
            {
                Log.WriteLogMessage($"Program End", LogMessageType.Info);
            }
        }

        //private void ReadExcel(out List<string> alreadyDoneJobModelList)
        //{
        //    alreadyDoneJobModelList = new List<string>();

        //    if (File.Exists(_fileName))
        //    {
        //        XLWorkbook workBook = new XLWorkbook(_fileName);

        //        IXLWorksheet workSheet = workBook.Worksheets.FirstOrDefault(ws => ws.Name == "Jobs");

        //        if (workSheet != null)
        //        {
        //            for (int i = 2; i < workSheet.LastRowUsed().RowNumber(); i++)
        //            {
        //                alreadyDoneJobModelList.Add($"{Convert.ToString(workSheet.Cell(i, 1).Value)} - {Convert.ToString(workSheet.Cell(i, 2).Value)}");
        //            }
        //        }
        //    }

        //    alreadyDoneJobModelList = alreadyDoneJobModelList.Distinct().ToList();
        //}

        private void WriteExcel(List<AdaJobModel> adaJobModelList)
        {
            if (adaJobModelList == null || adaJobModelList.Count == 0) return;

            try
            {
                XLWorkbook workBook = null;

                try
                {
                    workBook = new XLWorkbook(_fileName);
                }
                catch (Exception)
                {
                    workBook = new XLWorkbook();
                }

                foreach (AdaJobModel adaJobModel in adaJobModelList)
                {
                    var workSheet = workBook.Worksheets.FirstOrDefault(ws => ws.Name == "AdaJobs");

                    if (workSheet == null)
                    {
                        workSheet = workBook.Worksheets.Add("AdaJobs");

                        // Header
                        workSheet.Range(1, 1, 1, 8).AddToNamed("Header");
                        workSheet.Cell(1, 1).Value = "Category";
                        workSheet.Cell(1, 2).Value = "Description";
                        workSheet.Cell(1, 3).Value = "ADA Code";
                        workSheet.Cell(1, 4).Value = "ADA Description";
                        workSheet.Cell(1, 5).Value = "Task";
                        workSheet.Cell(1, 6).Value = "Expected Result";
                        workSheet.Cell(1, 7).Value = "ISTAT Code";
                        workSheet.Cell(1, 8).Value = "ISTAT Description";
                    }

                    int rowNumber = workSheet.LastRowUsed().RowNumber() + 1;
                    int maxCount = new[] { adaJobModel.Tasks.Count, adaJobModel.Associations.Count }.Max();

                    for (int i = 0; i < maxCount; i++, rowNumber++)
                    {
                        workSheet.Cell(rowNumber, 1).Value = adaJobModel.CategoryName;
                        workSheet.Cell(rowNumber, 2).Value = adaJobModel.CategoryDescription;
                        workSheet.Cell(rowNumber, 3).Value = adaJobModel.Code;
                        workSheet.Cell(rowNumber, 4).Value = adaJobModel.Description;

                        if (adaJobModel.Tasks.Count > i)
                        {
                            workSheet.Cell(rowNumber, 5).Value = adaJobModel.Tasks[i].Description;
                            if (!string.IsNullOrEmpty(adaJobModel.Tasks[i].ExpectedResultDescription))
                                workSheet.Cell(rowNumber, 6).Value = adaJobModel.Tasks[i].ExpectedResultDescription;
                        }

                        if (adaJobModel.Associations.Count > i)
                        {
                            workSheet.Cell(rowNumber, 7).Value = adaJobModel.Associations[i].Code;
                            workSheet.Cell(rowNumber, 8).Value = adaJobModel.Associations[i].Name;
                        }
                    }
                }

                // Header Style
                var headerStyle = workBook.Style;
                headerStyle.Font.Bold = true;
                headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workBook.NamedRanges.NamedRange("Header").Ranges.Style = headerStyle;

                workBook.SaveAs(_fileName);
            }
            catch (Exception ex)
            {
                Log.WriteLogMessage($"{ex.Message}", LogMessageType.Error);
            }
        }

        #endregion
    }
}
