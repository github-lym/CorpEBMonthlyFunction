using ClosedXML.Excel;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CorpEBMonthlyFunction
{
    class Program
    {
        static string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        static IniManager iniManager = new IniManager(Path.Combine(assemblyPath, "setting.ini"));
        static Logger log = LogManager.GetCurrentClassLogger();
        static DateTime today = DateTime.Now;

        static void Main(string[] args)
        {
            var dt = GetFunction();
            GenExcel(dt);            
        }

        /// <summary>
        /// 取得所有功能
        /// </summary>
        /// <returns></returns>
        public static DataTable GetFunction()
        {
            DataTable dt = new DataTable();
            string sql = @"SELECT
    u.CUM_BankCode,
	u.CUM_BankChineseName,
    t.TXID,
    t.TXNAME,
	CASE 
        WHEN s.TxId IS NULL THEN '啟用'
        ELSE '停用'
    END AS ONLINE
FROM
    (
        SELECT CUM_BankCode,CUM_BankChineseName
        FROM [CorporateBank].[dbo].[vwBankUnitInfo]
		WHERE RIGHT(CUM_BankCode, 2) = '00'
    ) AS u
CROSS JOIN
    [CorporateBank].[dbo].[AATx] AS t
LEFT JOIN 
	[CorporateBank].[dbo].[AATxStopBank] as s
	on u.CUM_BankCode = s.StopBank and t.TxId = s.TxId
	order by 1,3";

            try
            {
                Console.WriteLine("開始取得資料");
                log.Debug("開始取得資料");
                string connectionString = iniManager.ReadIniFile("DBOption", "ConnString", string.Empty);

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        cmd.CommandType = CommandType.Text;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("GetFunction : " + ex.ToString());
                log.Fatal("GetFunction : " + ex.ToString());
                throw;
            }

            return dt;
        }


        public static void GenExcel(DataTable dt)
        {
            //var ms = new MemoryStream();
            try
            {
                Console.WriteLine($"開始Excel處理({today.ToString("yyyyMM")})");
                log.Debug($"開始Excel處理({today.ToString("yyyyMM")})");

                using (MemoryStream ms = new MemoryStream())
                {
                    using (var wbook = new XLWorkbook())
                    {
                        wbook.Worksheets.Add(dt, "Sheet1"); //直接把DB匯入excel
                        var worksheet = wbook.Worksheet(1);
                        worksheet.Columns().AdjustToContents(); //根據cell大小調整寬度

                        wbook.SaveAs(ms);

                        SendMail(ms);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("GenExcel : " + ex.ToString());
                log.Fatal("GenExcel : " + ex.ToString());
                throw;
            }
        }

        static void SendMail(MemoryStream ms)
        {
            try
            {

                log.Debug($"開始寄mail");
                Console.WriteLine($"開始寄mail");

                string strFom = iniManager.ReadIniFile("MailOption", "UserFrom", string.Empty);
                string strMailAccount = iniManager.ReadIniFile("MailOption", "MailAccount", string.Empty);
                string strMailPW = iniManager.ReadIniFile("MailOption", "MailPassword", string.Empty);
                string strSSL = iniManager.ReadIniFile("MailOption", "MailSSL", string.Empty);
                string strSmtp = iniManager.ReadIniFile("MailOption", "SMTP", string.Empty);
                int strSmtpPort = int.TryParse(iniManager.ReadIniFile("MailOption", "SMTPPort", string.Empty), out strSmtpPort) ? strSmtpPort : 25;
                string strMailTo = iniManager.ReadIniFile("MailOption", "SendTo", string.Empty);

                string strFromName = string.Empty;

                using (System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage())
                {

                    message.SubjectEncoding = System.Text.Encoding.UTF8;
                    message.BodyEncoding = System.Text.Encoding.UTF8;
                    string[] mailTo = strMailTo.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < mailTo.Length; i++)
                        message.To.Add(mailTo[i]);

                    message.Bcc.Add("jamie.lin@afisc.com.tw");

                    message.From = new System.Net.Mail.MailAddress(strFom, strFromName);
                    message.Subject = $"新企網銀單位功能啟用清單 {today.ToString("yyyyMM")} 版";
                    message.IsBodyHtml = true;

                    //string sCustTemplete = HttpContext.Current.Server.MapPath("~/Document/" + strTemplete);
                    //using (System.IO.TextReader txtRead = new System.IO.StreamReader(sCustTemplete))
                    //{
                    //    message.Body = txtRead.ReadToEnd();
                    //}
                    string filename = $"新企網銀單位功能啟用清單 {today.ToString("yyyyMM")} 版" + ".xlsx";
                    ms.Position = 0;
                    var att = new System.Net.Mail.Attachment(ms, filename, "application/vnd.ms-excel");
                    message.Attachments.Add(att);
                    message.Body = $"附件為 新企網銀單位功能啟用清單 {today.ToString("yyyyMM")} 版 ";
                    message.Body += "<br>";
                    string strBody = string.Empty;
                    message.Body += "請卓參。";

                    using (System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(strSmtp, strSmtpPort))
                    {
                        smtp.Timeout = 60 * 1000;
                        if (strSSL == "N")
                        {
                            smtp.EnableSsl = false;
                        }
                        else
                            smtp.EnableSsl = true;
                        if (!string.IsNullOrEmpty(strMailAccount) && !string.IsNullOrEmpty(strMailPW))
                            smtp.Credentials = new NetworkCredential(strMailAccount, strMailPW);
                        smtp.Send(message);
                    }
                }

                Console.WriteLine("寄mail完成");
                log.Debug($"寄mail完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SendMail : {ex}");
                log.Fatal($"SendMail : {ex}");
            }
        }

    }
}
