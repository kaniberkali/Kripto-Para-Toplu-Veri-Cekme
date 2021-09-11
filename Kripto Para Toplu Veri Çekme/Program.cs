using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Kripto_Para_Toplu_Veri_Çekme
{
    class Program
    {
        static void Main(string[] args)
        {
            Random rnd = new Random();
            Console.Title = "Kripto Para Toplu Veri Çekme";
            Console.ForegroundColor = ConsoleColor.Green;
            OpenQA.Selenium.Chrome.ChromeDriver drv;

            Console.WriteLine("Website : kodzamani.weebly.com");
            Console.WriteLine("İnstagram : @kodzamani.tk");

            ChromeOptions option = new ChromeOptions();
           option.AddArgument("headless");
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            drv = new ChromeDriver(service,option);
            basadon:
            List<string> veriler = new List<string>();
            Console.Clear();
            drv.Navigate().GoToUrl("https://www.coingecko.com/tr/");
            Console.Write("Coin ismi girin :");
            string coin = Console.ReadLine();
            drv.FindElement(By.XPath("//html/body/div[2]/div[3]/div[2]/div[1]/div/div/input")).SendKeys(coin);
            Thread.Sleep(500);
            Console.Clear();
           int count = drv.FindElements(By.XPath("//li[@class='text-sm mt-1']/a/span/span[2]")).Count;
            for(int i=1;i<=count;i++)
            {
                string coinler = drv.FindElements(By.XPath("//li[@class='text-sm mt-1']/a/span/span[2]"))[i-1].Text;
                Console.WriteLine("{" + i + "} " + coinler);
            }
            Console.Write("Seçiminiz : ");
            try
            {
                int seçim = Convert.ToInt32(Console.ReadLine());
                if (seçim <= count && seçim >= 1)
                {
                    drv.Navigate().GoToUrl("https://www.coingecko.com/tr/coins/"+ drv.FindElements(By.XPath("//li[@class='text-sm mt-1']/a"))[seçim - 1].GetAttribute("href").Replace("https://www.coingecko.com/tr/search_redirect?id=","").Split('&')[0] + "/historical_data/usd?end_date=3000-02-24&start_date=2000-02-24");
                    Console.Clear();
                    Thread.Sleep(3000);
                    int altcount = drv.FindElements(By.XPath("//th[@class='font-semibold text-center']")).Count;
                    Console.WriteLine("Toplanacak Toplam Veri :" + altcount);
                    for (int i = 0; i < altcount; i++)
                    {
                        string tarih = drv.FindElements(By.XPath("//tr/th[@class='font-semibold text-center']"))[i].Text;
                        string piyasadegeri = drv.FindElements(By.XPath("//tr/td[1]"))[i].Text;
                        string hacim = drv.FindElements(By.XPath("//tr/td[2]"))[i].Text;
                        string ac = drv.FindElements(By.XPath("//tr/td[3]"))[i].Text;
                        string kapat = drv.FindElements(By.XPath("//tr/td[4]"))[i].Text;
                        Console.WriteLine(altcount+"/"+(i+1)+") Tarih :{0} Piyasa Değeri:{1} Hacim:{2} Aç:{3} Kapat:{4}", tarih, piyasadegeri, hacim, ac, kapat);
                        veriler.Add(tarih + ":" + piyasadegeri + ":" + hacim + ":" + ac + ":" + kapat);
                    }
                    Console.WriteLine("Toplanan Toplam Veri :" + veriler.Count);
                }
                else
                    goto basadon;
            }
            catch
            {
                goto basadon;
            }
            Console.WriteLine("------------------------------------------------------------------");
            Console.WriteLine("Tüm işlemler başarıyla bitirildi.");
            Console.WriteLine("-------------------------------------");
            Console.WriteLine("{1} Verileri Excel Sayfasına Aktar.");
            Console.WriteLine("{2} Yeni bir kripto para verisi çek.");
            Console.WriteLine("{3} Çıkış Yap.");
            Console.Write("Seçim :");
            string altseçim = Console.ReadLine();
            int sayi = rnd.Next(999999999);
            if (altseçim =="1")
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                       Console.WriteLine("Bilgisayarınızda excel kurulu değil.");
                        return;
                    }
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1, 6] = "Bu veriler Kripto Para Toplu Veri Çekme programı ile çekilmiştir website : kodzamani.weebly.com, instagram : @kodzamani.tk";
                    xlWorkSheet.Cells[2, 1] = "Tarih";
                    xlWorkSheet.Cells[2, 2] = "Piyasa Değeri";
                    xlWorkSheet.Cells[2, 3] = "Hacim";
                    xlWorkSheet.Cells[2, 4] = "Aç";
                    xlWorkSheet.Cells[2, 5] = "Kapat";
                    Console.WriteLine("Tarih, Piyasa Değeri, Hacim, Aç, Kapat");
                    Console.Clear();
                    for (int i = 3; i <= veriler.Count+2; i++)
                    {
                        try
                        {
                            xlWorkSheet.Cells[i, 1] = veriler[i- 3].Split(':')[0];
                            xlWorkSheet.Cells[i, 2] = veriler[i - 3].Split(':')[1];
                            xlWorkSheet.Cells[i, 3] = veriler[i - 3].Split(':')[2];
                            xlWorkSheet.Cells[i, 4] = veriler[i - 3].Split(':')[3];
                            xlWorkSheet.Cells[i, 5] = veriler[i - 3].Split(':')[4];
                            Console.WriteLine("Hücre :(" + i + ",1)" + veriler[i - 3].Split(':')[0]);
                            Console.WriteLine("Hücre :(" + i + ",2)" + veriler[i - 3].Split(':')[1]);
                            Console.WriteLine("Hücre :(" + i + ",3)" + veriler[i - 3].Split(':')[2]);
                            Console.WriteLine("Hücre :(" + i + ",4)" + veriler[i - 3].Split(':')[3]);
                            Console.WriteLine("Hücre :(" + i + ",5)" + veriler[i - 3].Split(':')[4]);
                            Console.WriteLine("---------------------------------------------------");
                        }
                        catch { }
                    }
                    xlWorkBook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "kodzamani-"+sayi+".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    Console.WriteLine("Excel dosyası başarıyla oluşturuldu.", "@kodzamani.tk");
                    Process.Start(AppDomain.CurrentDomain.BaseDirectory + "kodzamani-" + sayi + ".xls");
                }
                catch
                {
                    Console.WriteLine("Excel dosyası oluşturulamadı.", "@kodzamani.tk");
                }
                Console.WriteLine("------------------------------------");
                Console.WriteLine("{1} Yeni bir kripto para verisi çek.");
                Console.WriteLine("{2} Çıkış Yap.");
                Console.Write("Seçim :");
                string seçim = Console.ReadLine();
                if (seçim == "1")
                    goto basadon;
                if (seçim == "2")
                    drv.Quit();
            }
            if (altseçim == "2")
                goto basadon;
            if (altseçim == "3")
                drv.Quit();
        }
    }
}
