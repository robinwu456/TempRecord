using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace TempRecord {
    class Program {
        string urlCRM = "http://192.168.1.245/CRM";
        string account;
        string password;
        string inputTemp;

        static void Main(string[] args) {
            Program p = new Program();

            // 載入帳密
            FileInfo exeFile = new FileInfo(System.Reflection.Assembly.GetEntryAssembly().Location);
            string iniFileName = new FileInfo(exeFile.DirectoryName + "/Settings.ini").FullName;
            Console.WriteLine("import " + iniFileName);
            FileUtil.IniFile ini = new FileUtil.IniFile(iniFileName);
            p.account = ini.Read("Account");
            p.password = ini.Read("Password");
            p.inputTemp = ini.Read("Temperature");

            // 溫度登錄流程
            p.AutoRecordTemp();
        }

        void AutoRecordTemp() {
            IWebDriver driver = new ChromeDriver();

            try {
                // 開起官網
                Console.WriteLine("開啟CRM：{0}", urlCRM);
                driver.Navigate().GoToUrl(urlCRM);

                // 輸入帳密
                Console.WriteLine("輸入帳密\n帳號：{0}\n密碼：{1}", account, password);
                var txtAccount = driver.FindElement(By.Id("Account"));
                var txtPwd = driver.FindElement(By.Id("Password"));
                txtAccount.SendKeys(account);
                txtPwd.SendKeys(password);
                Console.WriteLine("CRM登入");
                txtPwd.SendKeys(OpenQA.Selenium.Keys.Enter);


                // 開啟體溫回報視窗
                Console.WriteLine("開起溫度視窗");
                var cmdTempRecord = driver.FindElement(By.Id("showCovid19_CONFIRM"));
                cmdTempRecord.Click();

                // 出勤狀況選擇
                Console.WriteLine("輸入溫度");
                var cboPeriods = driver.FindElement(By.Id("ddlCovid19_PERIODS"));
                var periods = new SelectElement(cboPeriods);
                periods.SelectByText("上班");
                var cboAddend = driver.FindElement(By.Id("ddlCovid19_ATTEND"));
                var attend = new SelectElement(cboAddend);
                attend.SelectByText("正常出勤");

                // 溫度輸入
                var txtTemp = driver.FindElement(By.Id("txtCovid19_BODY_TEMPERATURE"));
                txtTemp.SendKeys(inputTemp);

                // 送出
                Console.WriteLine("送出");
                var cmdSubmit = driver.FindElement(By.XPath("//button[.='送出']"));
                cmdSubmit.Click();

                FileUtil.AppendLog("登錄溫度：" + inputTemp);
                MessageBox.Show("溫度登錄完畢", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            } catch (Exception ex) {
                Console.WriteLine("失敗原因：" + ex.ToString());
                FileUtil.AppendLog("登錄失敗：" + ex.ToString());
                MessageBox.Show("Fail", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                Console.WriteLine("關閉");
                driver.Quit();
            }
        }
    }
}
