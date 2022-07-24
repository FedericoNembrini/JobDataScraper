using ClosedXML.Excel;
using FNSoftwareBot.Services;
using JobDataScraper.Models;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace JobDataScraper.Services
{
    public class InappService
    {
        #region Variables and Properties

        private Task MainTask;

        private bool _continueRun = true;

        private string _baseUrl = "https://professionioccupazione.isfol.it/professioni_navigazione.php?tipo_ricerca=1&testo_percorso=NAVIGAZIONE%20PER%20PROFESSIONE&link_percorso=professioni_navigazione.php";

        private string _fileName = "InappJobs.xlsx";

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
            List<JobModel> jobModelList = new List<JobModel>();

            try
            {
                Log.WriteLogMessage($"InappJobs Scraper - {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}", LogMessageType.Info);
                Log.WriteLogMessage($"To close the program, you must press the letter Q and wait for the program to close itself", LogMessageType.Info);
                Log.WriteLogMessage($"Loading...", LogMessageType.Info);

                ReadExcel(out List<string> alreadyDoneList);

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

                List<string> alphabetUrlList =
                    driver.FindElements(By.XPath("//a[contains(@href, 'lettera=')]"))
                    .Skip(Settings.Settings.BOT_SKIP)
                    .Take(Settings.Settings.BOT_TAKE)
                    .Select(we => we.GetAttribute("href"))
                    .ToList();

                foreach (string alphabetUrl in alphabetUrlList)
                {
                    Log.WriteLogMessage($"Start Search: {alphabetUrl}", LogMessageType.Info);

                    driver.Navigate().GoToUrl(alphabetUrl);

                    List<string> pageList =
                        driver.FindElements(By.XPath("//a[contains(@href, '?page=')]"))
                        .Skip(1)
                        .Select(we => we.GetAttribute("href"))
                        .ToList();

                    int counter = -1;
                    do
                    {
                        Log.WriteLogMessage($"Start Page: {counter + 2}", LogMessageType.Info);

                        if (counter > -1)
                            driver.Navigate().GoToUrl(pageList[counter]);

                        List<Tuple<string, string>> resultList =
                            driver.FindElements(By.XPath("//a[contains(@href, 'scheda.php')]"))
                            .Select(we => new Tuple<string, string>(we.Text, we.GetAttribute("href")))
                            .ToList();

                        foreach (var result in resultList)
                        {
                            if (alreadyDoneList.Contains(result.Item1)) continue;

                            Log.WriteLogMessage($"Doing: {result.Item1}", LogMessageType.Info);

                            var jobModel = new JobModel();

                            driver.Navigate().GoToUrl(result.Item2);

                            IWebElement jobTitleElement = driver.FindElement(By.XPath("//div[@class='headcentroscheda']"));

                            jobModel.Id = result.Item1.Substring(0, result.Item1.IndexOf("-")).Trim();
                            jobModel.Name = result.Item1.Substring(result.Item1.IndexOf("-") + 1).Trim();
                            jobModel.Description = jobTitleElement.Text.Substring(jobTitleElement.Text.IndexOf("-") + 1).Trim();

                            List<IWebElement> informationList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 2]")).ToList();

                            foreach (IWebElement information in informationList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = information.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobTasks.Add(new JobTask()
                                    {
                                        Name = columnList[1].Text,
                                        Importance = Convert.ToDecimal(columnList[2].FindElements(By.XPath(".//td"))[0].Text.Trim()),
                                        Frequency = Convert.ToDecimal(columnList[2].FindElements(By.XPath(".//td"))[3].Text.Trim())
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url + "&id_menu=2");

                            List<IWebElement> knowledgeList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 2]")).ToList();

                            foreach (IWebElement knowledge in knowledgeList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = knowledge.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobKnowledges.Add(new JobKnowledge()
                                    {
                                        Name = columnList[3].FindElements(By.XPath(".//span"))[0].Text,
                                        Description = columnList[3].FindElements(By.XPath(".//span"))[1].Text,
                                        Importance = Convert.ToDecimal(columnList[0].Text.Trim()),
                                        Complexity = Convert.ToDecimal(columnList[4].Text.Trim())
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=2", "id_menu=3"));

                            List<IWebElement> skillList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 2]")).ToList();

                            foreach (IWebElement skill in skillList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = skill.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobSkills.Add(new JobSkill()
                                    {
                                        Name = columnList[3].FindElements(By.XPath(".//span"))[0].Text,
                                        Description = columnList[3].FindElements(By.XPath(".//span"))[1].Text,
                                        Importance = Convert.ToDecimal(columnList[0].Text.Trim()),
                                        Complexity = Convert.ToDecimal(columnList[4].Text.Trim())
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=3", "id_menu=4"));

                            List<IWebElement> attitudeList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 2]")).ToList();

                            foreach (IWebElement skill in attitudeList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = skill.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobAttitudes.Add(new JobAttitude()
                                    {
                                        Name = columnList[3].FindElements(By.XPath(".//span"))[0].Text,
                                        Description = columnList[3].FindElements(By.XPath(".//span"))[1].Text,
                                        Importance = Convert.ToDecimal(columnList[0].Text.Trim()),
                                        Complexity = Convert.ToDecimal(columnList[4].Text.Trim())
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=4", "id_menu=5"));

                            List<IWebElement> generalisedActivityList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 2]")).ToList();

                            foreach (IWebElement skill in generalisedActivityList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = skill.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobGeneralisedActivities.Add(new JobGeneralisedActivity()
                                    {
                                        Name = columnList[3].FindElements(By.XPath(".//span"))[0].Text,
                                        Description = columnList[3].FindElements(By.XPath(".//span"))[1].Text,
                                        Importance = Convert.ToDecimal(columnList[0].Text.Trim()),
                                        Complexity = Convert.ToDecimal(columnList[4].Text.Trim())
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=5", "id_menu=7"));

                            List<IWebElement> styleList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 1]")).ToList();

                            foreach (IWebElement skill in styleList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = skill.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobStyles.Add(new JobStyle()
                                    {
                                        Name = columnList[0].FindElements(By.XPath(".//span"))[0].Text,
                                        Description = columnList[0].FindElements(By.XPath(".//span"))[1].Text,
                                        Importance = Convert.ToDecimal(columnList[1].Text.Trim()),
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=7", "id_menu=10"));

                            List<IWebElement> jobExampleList = driver.FindElements(By.XPath("//table[@class='grafico']//tr[count(td) > 1]")).ToList();

                            foreach (IWebElement jobExample in jobExampleList)
                            {
                                try
                                {
                                    List<IWebElement> columnList = jobExample.FindElements(By.XPath(".//td")).ToList();

                                    jobModel.JobExamples.Add(new JobExample()
                                    {
                                        Name = columnList[1].Text
                                    });
                                }
                                catch (Exception) { }
                            }

                            driver.Navigate().GoToUrl(driver.Url.Replace("id_menu=10", "id_menu=11"));

                            try
                            {
                                List<IWebElement> columnList = driver.FindElements(By.XPath("//td[@class='titolo']")).ToList();

                                jobModel.JobEQFs.Add(new JobEQF()
                                {
                                    Name = columnList[0].Text.Replace(columnList[0].FindElement(By.XPath(".//span")).Text, string.Empty).Trim(),
                                    Value = Convert.ToDecimal(columnList[0].FindElement(By.XPath(".//span")).Text)
                                });
                                jobModel.JobEQFs.Add(new JobEQF()
                                {
                                    Name = columnList[2].Text.Replace(columnList[2].FindElement(By.XPath(".//span")).Text, string.Empty).Trim(),
                                    Value = Convert.ToDecimal(columnList[2].FindElement(By.XPath(".//span")).Text)
                                });
                                jobModel.JobEQFs.Add(new JobEQF()
                                {
                                    Name = columnList[3].Text.Replace(columnList[3].FindElement(By.XPath(".//span")).Text, string.Empty).Trim(),
                                    Value = Convert.ToDecimal(columnList[3].FindElement(By.XPath(".//span")).Text)
                                });
                                jobModel.JobEQFs.Add(new JobEQF()
                                {
                                    Name = columnList[1].Text.Replace(columnList[1].FindElement(By.XPath(".//span//span")).Text, string.Empty).Trim(),
                                    Value = Convert.ToDecimal(columnList[1].FindElement(By.XPath(".//span//span")).Text)
                                });
                            }
                            catch (Exception) { }

                            jobModelList.Add(jobModel);

                            Log.WriteLogMessage($"Done: {jobModel.Name}", LogMessageType.Success);

                            if (!_continueRun) break;
                        }

                        Log.WriteLogMessage($"Writing File", LogMessageType.Info);

                        WriteExcel(jobModelList);

                        Log.WriteLogMessage($"File Wrote", LogMessageType.Success);

                        jobModelList.Clear();

                        if (!_continueRun) return;

                        Log.WriteLogMessage($"End Page: {counter + 2}", LogMessageType.Success);

                        counter++;
                    } while (counter < pageList.Count);

                    Log.WriteLogMessage($"End Search: {alphabetUrl}", LogMessageType.Success);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLogMessage($"Error: {ex.Message}", LogMessageType.Error);
            }
            finally
            {
                Log.WriteLogMessage($"Program End", LogMessageType.Info);
            }
        }

        private void ReadExcel(out List<string> alreadyDoneJobModelList)
        {
            alreadyDoneJobModelList = new List<string>();

            if (File.Exists(_fileName))
            {
                XLWorkbook workBook = new XLWorkbook(_fileName);

                IXLWorksheet workSheet = workBook.Worksheets.FirstOrDefault(ws => ws.Name == "Jobs");

                if (workSheet != null)
                {
                    for (int i = 2; i < workSheet.LastRowUsed().RowNumber(); i++)
                    {
                        alreadyDoneJobModelList.Add($"{Convert.ToString(workSheet.Cell(i, 1).Value)} - {Convert.ToString(workSheet.Cell(i, 2).Value)}");
                    }
                }
            }

            alreadyDoneJobModelList = alreadyDoneJobModelList.Distinct().ToList();
        }

        private void WriteExcel(List<JobModel> jobModelList)
        {
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

                foreach (JobModel jobModel in jobModelList)
                {
                    var workSheet = workBook.Worksheets.FirstOrDefault(ws => ws.Name == "Jobs");

                    if (workSheet == null)
                    {
                        workSheet = workBook.Worksheets.Add("Jobs");

                        // Header
                        workSheet.Range(1, 1, 1, 25).AddToNamed("Header");
                        workSheet.Cell(1, 1).Value = "Id";
                        workSheet.Cell(1, 2).Value = "Name";

                        workSheet.Cell(1, 3).Value = "Task Name";
                        workSheet.Cell(1, 4).Value = "Task Importance";
                        workSheet.Cell(1, 5).Value = "Task Frequency";

                        workSheet.Cell(1, 7).Value = "Knowledge";
                        workSheet.Cell(1, 8).Value = "Importancy";

                        workSheet.Cell(1, 10).Value = "Skill";
                        workSheet.Cell(1, 11).Value = "Importancy";

                        workSheet.Cell(1, 13).Value = "Attitude";
                        workSheet.Cell(1, 14).Value = "Importancy";

                        workSheet.Cell(1, 16).Value = "Generalised Task";
                        workSheet.Cell(1, 17).Value = "Importancy";

                        workSheet.Cell(1, 19).Value = "Task Style";
                        workSheet.Cell(1, 20).Value = "Importancy";

                        workSheet.Cell(1, 22).Value = "Job Example";

                        workSheet.Cell(1, 24).Value = "EQF Name";
                        workSheet.Cell(1, 25).Value = "EQF Value";
                    }

                    int rowNumber = workSheet.LastRowUsed().RowNumber() + 1;
                    int maxCount = new[] { jobModel.JobTasks.Count, jobModel.JobKnowledges.Count, jobModel.JobSkills.Count, jobModel.JobAttitudes.Count, jobModel.JobGeneralisedActivities.Count, jobModel.JobStyles.Count, jobModel.JobExamples.Count, jobModel.JobEQFs.Count }.Max();

                    for (int i = 0; i < maxCount; i++, rowNumber++)
                    {
                        workSheet.Cell(rowNumber, 1).Value = jobModel.Id;
                        workSheet.Cell(rowNumber, 2).Value = jobModel.Name;

                        if (jobModel.JobTasks.Count > i)
                        {
                            workSheet.Cell(rowNumber, 3).Value = jobModel.JobTasks[i].Name;
                            workSheet.Cell(rowNumber, 4).Value = jobModel.JobTasks[i].Importance;
                            workSheet.Cell(rowNumber, 5).Value = jobModel.JobTasks[i].Frequency;
                        }

                        if (jobModel.JobKnowledges.Count > i)
                        {
                            workSheet.Cell(rowNumber, 7).Value = jobModel.JobKnowledges[i].Name;
                            workSheet.Cell(rowNumber, 8).Value = jobModel.JobKnowledges[i].Importance;
                        }

                        if (jobModel.JobSkills.Count > i)
                        {
                            workSheet.Cell(rowNumber, 10).Value = jobModel.JobSkills[i].Name;
                            workSheet.Cell(rowNumber, 11).Value = jobModel.JobSkills[i].Importance;
                        }

                        if (jobModel.JobAttitudes.Count > i)
                        {
                            workSheet.Cell(rowNumber, 13).Value = jobModel.JobAttitudes[i].Name;
                            workSheet.Cell(rowNumber, 14).Value = jobModel.JobAttitudes[i].Importance;
                        }

                        if (jobModel.JobGeneralisedActivities.Count > i)
                        {
                            workSheet.Cell(rowNumber, 16).Value = jobModel.JobGeneralisedActivities[i].Name;
                            workSheet.Cell(rowNumber, 17).Value = jobModel.JobGeneralisedActivities[i].Importance;
                        }

                        if (jobModel.JobStyles.Count > i)
                        {
                            workSheet.Cell(rowNumber, 19).Value = jobModel.JobStyles[i].Name;
                            workSheet.Cell(rowNumber, 20).Value = jobModel.JobStyles[i].Importance;
                        }

                        if (jobModel.JobExamples.Count > i)
                        {
                            workSheet.Cell(rowNumber, 22).Value = jobModel.JobExamples[i].Name;
                        }

                        if (jobModel.JobEQFs.Count > i)
                        {
                            workSheet.Cell(rowNumber, 24).Value = jobModel.JobEQFs[i].Name;
                            workSheet.Cell(rowNumber, 25).Value = jobModel.JobEQFs[i].Value;
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
