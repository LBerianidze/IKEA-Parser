using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IKEA_Parser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            this.InitializeComponent();
        }

        private List<string> failed = new List<string>();
        private List<Exception> exceptions = new List<Exception>();
        private OpenQA.Selenium.Chrome.ChromeDriver driver;
        private Dictionary<string, int> articuls = new Dictionary<string, int>();
        private string FillToEight(string item)
        {
            return item.Insert(0, new string('0', 8 - item.Length));
        }
        private void openFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Excel File|*.xls"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.articuls.Clear();
                this.label1.Text = ofd.FileName;
                NPOI.HSSF.UserModel.HSSFWorkbook workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(new FileStream(ofd.FileName, FileMode.Open));
                NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                int rowCount = sheet.LastRowNum;
                //articuls.Add("90368598", 1);
                //articuls.Add("30417790", 1);
                //articuls.Add("10370699", 1);
                //articuls.Add("40369798", 1);
                //articuls.Add("90368975", 1);
                //articuls.Add("30388735", 6);
                //articuls.Add("90199204", 4);
                //return;
                for (int i = 1; i < rowCount; i++)
                {
                    try
                    {
                        NPOI.SS.UserModel.ICell cell = sheet.GetRow(i).GetCell(0);
                        if (cell.CellType == NPOI.SS.UserModel.CellType.String)
                        {
                            if (!Regex.IsMatch(cell.StringCellValue, @"\p{IsCyrillic}"))
                            {
                                this.articuls.Add(sheet.GetRow(i).GetCell(0).StringCellValue.ToLower(), (int)sheet.GetRow(i).GetCell(1).NumericCellValue);
                            }
                        }
                        else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                        {
                            this.articuls.Add(this.FillToEight(sheet.GetRow(i).GetCell(0).NumericCellValue.ToString()), (int)sheet.GetRow(i).GetCell(1).NumericCellValue);
                        }
                    }
                    catch (Exception)
                    {
                        string val = "";
                        int num = (int)sheet.GetRow(i).GetCell(1).NumericCellValue;
                        NPOI.SS.UserModel.ICell cell = sheet.GetRow(i).GetCell(0);
                        if (cell.CellType == NPOI.SS.UserModel.CellType.String)
                        {
                            if (!Regex.IsMatch(cell.StringCellValue, @"\p{IsCyrillic}"))
                            {
                                val = sheet.GetRow(i).GetCell(0).StringCellValue.ToLower();
                            }
                        }
                        else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                        {
                            val = this.FillToEight(sheet.GetRow(i).GetCell(0).NumericCellValue.ToString());
                        }
                        this.articuls[val] += num;
                    }
                }
                this.label3.Text = this.articuls.Count.ToString();
            }
        }

        private void MovePage(string url)
        {
            try
            {
                this.driver.Url = url;
            }
            catch
            {
                bool pageisloaded = false;
                while (!pageisloaded)
                {
                    pageisloaded = this.TryCatch();
                }
            }
        }
        private bool TryCatch()
        {
            try
            {
                IWait<IWebDriver> wait = new WebDriverWait(this.driver, TimeSpan.FromSeconds(3));
                bool Func(IWebDriver driver1)
                {
                    return ((IJavaScriptExecutor)this.driver).ExecuteScript("return document.readyState").Equals("complete");
                }

                wait.Until(Func);
                return true;
            }
            catch
            {
                if (this.driver == null)
                {
                    return true;
                }
                return false;
            }
        }
        public List<HtmlAgilityPack.HtmlNode> GetAllNodes(HtmlAgilityPack.HtmlNode htmlNode)
        {
            List<HtmlAgilityPack.HtmlNode> htmlNodes = new List<HtmlAgilityPack.HtmlNode>
            {
                htmlNode
            };
            foreach (var item in htmlNode.ChildNodes)
            {
                htmlNodes.AddRange(GetAllNodes(item));
            }
            return htmlNodes;
        }
        public HtmlAgilityPack.HtmlNode FindElementByTagNameAndInnerText(HtmlAgilityPack.HtmlNode htmlDocument)
        {
            foreach (var item in GetAllNodes(htmlDocument))
            {
                if (item.Attributes.FirstOrDefault(t => t.Name == "class" && t.Value == "radio-button-field") != null)
                {
                    return item;
                }
            }
            return null;
        }
        //product-button-container
        public HtmlAgilityPack.HtmlNode FindElementByClass(HtmlAgilityPack.HtmlNode htmlDocument)
        {
            foreach (var item in GetAllNodes(htmlDocument))
            {
                if (item.Attributes.FirstOrDefault(t => t.Name == "class")?.Value == "product-button-container")
                {
                    return item;
                }
            }
            return null;
        }
        public HtmlAgilityPack.HtmlNode FindElementByContentAndInnerText(HtmlAgilityPack.HtmlNode htmlDocument, string v)
        {
            foreach (var item in GetAllNodes(htmlDocument))
            {
                if (item.InnerText == v)
                {
                    return item;
                }
            }
            return null;
        }
        private void Parse()
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddExtension(Application.StartupPath + "\\blocker.crx");
            options.AddArgument("start-maximized"); // https://stackoverflow.com/a/26283818/1689770
            options.AddArgument("enable-automation"); // https://stackoverflow.com/a/43840128/1689770
            //options.AddArgument("--headless"); // only if you are ACTUALLY running headless
            options.AddArgument("--no-sandbox"); //https://stackoverflow.com/a/50725918/1689770
            options.AddArgument("--disable-infobars"); //https://stackoverflow.com/a/43840128/1689770
            options.AddArgument("--disable-dev-shm-usage"); //https://stackoverflow.com/a/50725918/1689770
            options.AddArgument("--disable-browser-side-navigation"); //https://stackoverflow.com/a/49123152/1689770
            options.AddArgument("--disable-gpu"); //https://stackoverflow.com/questions/51959986/how-to-solve-selenium-chromedriver-timed-out-receiving-message-from-renderer-exc
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.EnableVerboseLogging = false;
            this.driver = new OpenQA.Selenium.Chrome.ChromeDriver(service, options);
            this.driver.Manage().Window.Maximize();
            string url = "https://www.ikea.com/ru/ru/p/svalk-bokal-dlya-vina-prozrachnoe-steklo-{0}/";
            int rowIndex = 1;
            this.driver.Url = "https://www.ikea.com/ru/ru/customer-service/";
            Thread.Sleep(5000);
            while (true)
            {
                try
                {
                    this.driver.FindElementById("iconHeaderIcon").Click();
                    break;
                }
                catch { }
            }
            try
            {
                this.driver.FindElementByXPath("//*[@id=\"range-modal-mount-node\"]/div/div[3]/div/span/div/button[1]").Click();
            }
            catch(Exception ex)
            {

            }
            Thread.Sleep(1000);
            driver.FindElementById("unit-338").Click();
            Thread.Sleep(1000);
            driver.FindElementByXPath("//*[@id=\"range-modal-mount-node\"]/div/div[3]/div/span/div/button").Click();
            Thread.Sleep(1500);
            for (int i = 0; i < this.articuls.Count; i++)
            {
                Stopwatch st = new Stopwatch();
                st.Start();
                this.driver.Url = string.Format(url, this.articuls.ElementAt(i).Key);
                while (true)
                {
                    try
                    {
                        bool res = new WebDriverWait(this.driver, TimeSpan.FromMilliseconds(250)).Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState")).Equals("complete");
                        break;
                    }
                    catch
                    {

                    }
                }
                Thread.Sleep(1000);
                int tryCount = 0;
                string scrollablePath = "";
                while (true)
                {
                    try
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(driver.PageSource);
                        var elem = FindElementByClass(doc.DocumentNode);
                        var basket_span = FindElementByContentAndInnerText(elem, "Добавить в корзину");
                        if (basket_span != null)
                        {
                            scrollablePath = basket_span.ParentNode.ParentNode.XPath;
                            driver.FindElementByXPath(basket_span.ParentNode.ParentNode.XPath).Click();
                            // button--processing
                           while ( driver.FindElementByXPath(basket_span.ParentNode.ParentNode.XPath).GetAttribute("class").Contains("button--processing"))
                            {

                            }
                            Thread.Sleep(200);
                        }
                        else
                        {
                            throw new Exception();
                        }
                        break;
                    }
                    catch (Exception)
                    {
                        string source = this.driver.PageSource;
                        if (source.Contains("Пока недоступно онлайн") || source.Contains("Страница, которую вы запрашивали, не найдена.") || source.Contains("product-missing") || source.Contains("Нет в наличии в"))
                        {
                            this.BeginInvoke(new ThreadStart(() => this.dataGridView1.Rows.Add(rowIndex++, this.articuls.ElementAt(i).Key, this.articuls.ElementAt(i).Value)));
                            this.failed.Add(this.articuls.ElementAt(i).Key);
                            break;
                        }
                        else if (tryCount++ >= 10)
                        {
                            this.BeginInvoke(new ThreadStart(() => this.dataGridView1.Rows.Add(rowIndex++, this.articuls.ElementAt(i).Key, this.articuls.ElementAt(i).Value)));
                            this.failed.Add(this.articuls.ElementAt(i).Key);
                            break;
                        }
                        else
                        {
                            try
                            {
                                IWebElement element = driver.FindElementByXPath(scrollablePath);

                                driver.ExecuteScript(string.Format("window.scrollTo(0,{0})", element.Location.Y));
                            }
                            catch (Exception ex1)
                            {
                            }
                        }
                        Thread.Sleep(1000);
                    }
                }
                this.BeginInvoke(new ThreadStart(() => this.label4.Text = (i + 1).ToString()));
                st.Stop();
                if (i != 0)
                {
                    this.BeginInvoke(new ThreadStart(() => this.label7.Text = st.Elapsed.TotalSeconds.ToString()));

                }
                this.BeginInvoke(new ThreadStart(() => this.label8.Text = (this.failed.Count).ToString()));
                Thread.Sleep(500);
            }
            try
            {
                IWebElement element = driver.FindElementByXPath("/html/body/header/div/div/div/ul/li[5]/a");

                driver.ExecuteScript(string.Format("window.scrollTo(0,{0})", element.Location.Y));
            }
            catch (Exception ex1)
            {
            }
            //driver.FindElementByXPath("/html/body/header/div/div/div/ul/li[5]/a").Click();
            Invoke(new ThreadStart(() => { MessageBox.Show("Перейдите в корзину, выберите список и нажмите OK", "Парсинг", MessageBoxButtons.OK, MessageBoxIcon.Information); }));
            Thread.Sleep(5000);
            for (int i = 0; i < articuls.Count; i++)
            {
                try
                {
                    if (articuls.ElementAt(i).Value > 1)
                    {
                        string articul = articuls.ElementAt(i).Key;
                        driver.FindElementByXPath(string.Format("//*[@id=\"js_qty_{0}\"]/option[{1}]", articul, articuls.ElementAt(i).Value)).Click();
                        Thread.Sleep(1000);
                        while (driver.PageSource.Contains("product__contents product__blur")) { }
                    }
                }
                catch (Exception ex)
                {

                }
            }

        }
        private async void startParsingButton_Click(object sender, EventArgs e)
        {
            this.failed.Clear();
            this.dataGridView1.Rows.Clear();
            await Task.Run(this.Parse);
            MessageBox.Show("Обработка завершена", "Парсинг", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void closeChrome_Button_Click(object sender, EventArgs e)
        {
            try
            {
                this.driver.Quit();
            }
            catch
            {

            }
        }
    }
}
