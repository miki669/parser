using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace parser
{
    public partial class Science_RU_day_Xak : Form
    {
        public Science_RU_day_Xak()
        {
            InitializeComponent();
        }

        private static IWebDriver driver = null;


        void Log(object x)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("[" + DateTime.Now + "]  " + x);
            Console.ForegroundColor = ConsoleColor.Red;
        }
        void SUCCESSLog(object x)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("[" + DateTime.Now + "]  " + x);
            Console.ForegroundColor = ConsoleColor.Blue;
        }
        void ERRORLog(object x)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("[" + DateTime.Now + "]  " + x);
            Console.ForegroundColor = ConsoleColor.Red;
        }
        class pars
        {
            public string _title { get; set; }
            public object _avtor { get; set; }
            public string _ISBN { get; set; }
            public string _type { get; set; }
            public string _papent_number { get; set; }

        }
        private void start_Click(object sender, EventArgs e)
        {

            main();
        }



        void main()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("'excludeSwitches', ['enable-logging']");
            options.AddArgument("--log-level=3");
            options.AddArgument("--window-size=1920,920");
            options.AddArgument("--disable-blink-features=AutomationControlled");
            options.AddExcludedArgument("enable-automation");
            IWebDriver driver = new ChromeDriver(options);
            Excel.Application excel_app = new Excel.Application();
            excel_app.Visible = true;
            excel_app.Workbooks.Add();
            Excel._Worksheet sheet = excel_app.ActiveSheet;
            sheet.Cells[1, 1] = "Ссылка на статью";
            sheet.Cells[1, 2] = "Название";
            sheet.Cells[1, 3] = "ТИП ИЗДАНИЯ";
            sheet.Cells[1, 4] = "ЯЗЫК";
            sheet.Cells[1, 5] = "Тип публикации";
            driver.Navigate().GoToUrl("https://www.elibrary.ru/org_items.asp?orgsid=1020%22");
            Thread.Sleep(1000);
            int row = 2;
            int stranica = 4; // 1 страница на сайте в хпатхе == 3
            Go();
            void Go()
            {
                for (int id = 4; id <= 103; id++)
                {
                    if (id == 103)
                    {


                        next_page();
                        Go();
                    }
                    try
                    {
                        IWebElement element = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[{id}]"));
                        Log($"Post {id}:" + element.Text);
                        IWebElement href = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[{id}]/td[2]/a"));
                        string link = href.GetAttribute("href");
                        SUCCESSLog(link);
                        sheet.Cells[row, 1] = link;
                        try
                        {

                            href.Click();
                            Thread.Sleep(500);
                            try
                            {
                                string Название = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/table[2]/tbody/tr/td[2]/span/b/p")).Text;
                                sheet.Cells[row, 2] = Название;
                            }
                            catch { sheet.Cells[row, 2] = "Ошибка получения"; }
                            try
                            {
                                string Тип = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[1]/td/font[1]")).Text;
                                sheet.Cells[row, 3] = Тип;

                            }
                            catch { sheet.Cells[row, 3] = "Ошибка получения"; }
                            try
                            {
                                string Язык = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[3]/tbody/tr[1]/td/font[2]")).Text;
                                sheet.Cells[row, 4] = Язык;
                            }
                            catch { sheet.Cells[row, 4] = "Ошибка получения"; }

                            try
                            {
                                string ISBN = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[3]/tbody/tr[1]/td/font[3]")).Text;
                                sheet.Cells[row, 5] = "Статья, ISBN: " + ISBN;
                            }
                            catch
                            {
                                try
                                {
                                    string ISBN_2 = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[1]/td/font[3]")).Text;
                                    sheet.Cells[row, 5] = "Статья, ISBN: " + ISBN_2;
                                }
                                catch 
                                {
                                    try
                                    {
                                        
                                        //string contains_patent = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[1]/td/text()[2]")).Text;
                                        string patent_number = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[2]/td/font[1]")).Text;
                                        sheet.Cells[row, 5] = "Патент: " + patent_number;


                                    }
                                    catch
                                    {
                                        string contains_patent = driver
                                            .FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[1]/td/text()[2]")).Text;
                                        if (contains_patent.Contains("патент"))
                                        {
                                            string patent_number_2 = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/table[2]/tbody/tr[1]/td/font[2]")).Text;
                                            sheet.Cells[row, 5] = "Патент: " + patent_number_2;
                                        }
                                    }

                                }

                            }

                            driver.Navigate().Back();

                            Thread.Sleep(500);

                        }
                        catch
                        {/* MessageBox.Show("1")*/
                            ; driver.Navigate().Back();
                        }



                        row++;



                    }
                    catch
                    {
                        Console.WriteLine("С этой страницей все");
                        try
                        {

                            IWebElement next = driver.FindElement(By.XPath(
                                $"/html/body/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[4]/table/tbody/tr/td[13]"));
                            next.Click();
                            id = 4;
                            Go();

                        }
                        catch (Exception ex)
                        {

                            ERRORLog(ex.Message);
                            IWebElement next = driver.FindElement(By.XPath($"/html/body/div[2]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[4]/table/tbody/tr/td[13]"));
                            next.Click();
                            id = 4;
                            Go();



                        }
                    }


                }
            }

            void next_page()
            {
                try
                {
                    stranica++;
                    IWebElement next = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[4]/table/tbody/tr/td[13]"));
                    next.Click();
                }
                catch (Exception ex)
                {
                    ERRORLog(ex.Message);
                    IWebElement next = driver.FindElement(By.XPath($"/html/body/div[2]/table/tbody/tr/td/table[1]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[4]/table/tbody/tr/td[13]"));
                    next.Click();
                    Go();
                }

            }



        }
    }


}
