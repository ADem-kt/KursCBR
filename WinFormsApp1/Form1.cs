using System.Xml.Linq;
using Spire.Xls;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {

        DateTime dateTime;
        static HttpClient httpClient = new HttpClient();
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            textBox1.Text = "start..."+Environment.NewLine;
            dateTime = dateTimePicker1.Value;
            textBox1.Text += $"Получаю информацию с сайта ЦБР за дату {dateTime.Date.ToString("dd'/'MM'/'yyyy")}..." + Environment.NewLine;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, $"http://www.cbr.ru/scripts/XML_daily.asp?date_req={dateTime.Date.ToString("dd'/'MM'/'yyyy")}");
            using HttpResponseMessage response = await httpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            textBox1.Text += $"Получил" + Environment.NewLine;
            textBox1.Text += $"Создаю xls..." + Environment.NewLine;
            
            await Task.Run(() => creatXLS(content, dateTime));
            textBox1.Text += $"Создал xls в " + Directory.GetCurrentDirectory() + Environment.NewLine;
            textBox1.Text += "end." + Environment.NewLine;
            button1.Enabled = true;
        }
        private void creatXLS(string xml, DateTime dt)
        {

            XDocument xdoc = XDocument.Parse(xml);
            XElement ValCurs = xdoc.Element("ValCurs");
            //Создание объекта Workbook 
            Workbook workbook = new Workbook();
            string dec_sep = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            string message = textBox1.Text;
            int k = 0;
            int i = 1;
            Worksheet worksheet = workbook.Worksheets[k];
            worksheet.Name = "Рубль";
            foreach (XElement elValute in ValCurs.Elements("Valute"))
            {                
                //Запись данных в определенные ячейки
                worksheet.Range[i, 1].Value = elValute.Element("Name").Value;
                worksheet.Range[i, 2].Value = elValute.Element("VunitRate").Value;
                worksheet.AllocatedRange.AutoFitColumns();
                i++;
            }
            k++;
            foreach (XElement elWorksheet in ValCurs.Elements("Valute"))
            {
                workbook.CreateEmptySheet();
                worksheet = workbook.Worksheets[k];
                worksheet.Name = elWorksheet.Element("Name").Value;
                i = 1;
                foreach (XElement elValute in ValCurs.Elements("Valute"))
                {                    
                    //Запись данных в определенные ячейки
                    worksheet.Range[i, 1].Value = elValute.Element("Name").Value;
                    worksheet.Range[i, 2].Value = (double.Parse(elWorksheet.Element("VunitRate").Value.Replace(",", dec_sep).Replace(".", dec_sep)) / double.Parse(elValute.Element("VunitRate").Value.Replace(",", dec_sep).Replace(".", dec_sep))).ToString();//
                    worksheet.AllocatedRange.AutoFitColumns();
                    i++;

                    
                    this.Invoke((MethodInvoker)delegate
                    {

                        textBox1.Text = message + $"Выполнено {Math.Round((Convert.ToDouble(k) / Convert.ToDouble(ValCurs.Elements("Valute").Count())) * 100, 2)}%"+Environment.NewLine;

                    });
                }

                k++;
            }
            workbook.SaveToFile($"{dt.Date.ToString("yyyyMMdd")}.xlsx", ExcelVersion.Version2016);
            
        }
    }
}
