using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FCWindowService
{
     
    public partial class FC : ServiceBase
    {
        private Timer timer1 = null;
        public FC()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer1 = new Timer();
            this.timer1.Interval = 60000;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            timer1.Enabled = true;
        }
        protected override void OnStop()
        {
            timer1.Enabled = false;
        }
        public void timer1_Tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            library db = new library();
            try
            {
                DataTable dt = db.SelectRecords(@"IF OBJECT_ID('tempdb..#Results') IS NOT NULL DROP TABLE #Results 
;with tbl as (
select wv_FC_Reporting_Total_Amount.Year,OCM_Province.ProvinceEngName,OCM_Province.ProvinceID,isnull(sum(TotalAmount),0) as TotalAmount,SeasonId,SeasonName,isnull(sum(RecievedAmount),0) RecievedAmount,isnull(sum(Balance),0) as Balance
  from  wv_FC_Reporting_Total_Amount inner join FC_ExtensionWorkerInfo on wv_FC_Reporting_Total_Amount.ExtWId=FC_ExtensionWorkerInfo.ExtWID inner join 
 OCM_Province on OCM_Province.ProvinceID=FC_ExtensionWorkerInfo.ProvinceID 
 group by ProvinceEngName,Year,SeasonId,SeasonName,OCM_Province.ProvinceID
 )
 select tbl.ProvinceEngName,tbl.TotalAmount as [Collectable Amount],tbl.RecievedAmount as [Paid Amount],tbl.Balance as [Balance Amount],tbl.Year into #Results from tbl
 select * from #Results 
 drop table #Results ");
                
                DataTable dtEmail = db.SelectRecords(@"select Email from tbl_PC 
where Email is not null
union all
select email from tbl_Region
where Email is not null
union all
select email from aspnet_Membership where ApplicationId like '34764d90-4054-47ea-ab95-17c5d32a6136' and email is not null");

                if (dt.Rows.Count > 0 )
                {
                    byte[] bytes;

                    string emails = "";
                    if (dtEmail.Rows.Count > 0)
                    {
                        foreach (DataRow rw in dtEmail.Rows)
                        {
                            emails += rw["Email"].ToString() + ",";
                        }
                    }
                    emails += "omar.safiullah@gmail.com,grkinyanjui@gmail.com,khalid.ferdaus@yahoo.com,usman.safi@mail.gov.af,arshidir@live.com,h_saleemsafi@yahoo.com,samadsame@gmail.com,khalid.ibrahimsafi@gmail.com,shakir_vet@yahoo.com,shaimaahadi2003@yahoo.com,ramaraorv@yahoo.co.in";
                    MailMessage mm = new MailMessage("fcmis.nhlp@gmail.com","omar.safiullah@gmail.com");
                    mm.Subject = "eWeeklyReport:Farmer Contribution Report By Province for NHLP/MAIL";
                    string body = "<strong>Dear Observers,</strong>  Thank you for seeing online report of  <strong>Farmer Contribution Managment Information System</strong>. Please find the attached reports which is generated from date :<strong>" + DateTime.Now.ToString() + "  date </strong> by Province .Columns are in  Province ,Collectable Amount , Paid Amount and Balance order respectively.";
                    body += "<br/>If you are unable to read the PDF file, please download the latest version of the Adobe Acrobat Reader:<br/>http://get.adobe.com/reader/</br>";
                    body += "<br/>For further information or queries, please contact<strong> NHLP MIS Department</strong>, or go for online on <strong>http:// 103.13.66.210:88</strong> ";
                    body += "<br/>------------------------------------------------------------------------------<br/><br/>This is an automatically generated message; please do not reply to this email.";
                    mm.Body = body;
                    if (dt.Rows.Count > 0)
                    {
                        bytes = SendPDFEmail(dt);
                        mm.Attachments.Add(new Attachment(new MemoryStream(bytes), "FCMIS" + DateTime.Now.ToShortDateString() + "PaidAmountreport.pdf"));
                    }


                    mm.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    NetworkCredential NetworkCred = new NetworkCredential();
                    NetworkCred.UserName = "fcmis.nhlp@gmail.com";
                    NetworkCred.Password = "safi_khan123";
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    smtp.Send(mm);
                    
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                db.Connection.Close();
                timer1.Stop();
                timer1.Enabled = false;
                ((Timer)sender).Dispose();
            }
        }
        private byte[] SendPDFEmail(DataTable dt)
        {
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2'>");
                    sb.Append("<tr><td align='center' ></td align='center' ><td align='center'>Farmer Contribution Managment Information System</td><td align='center' ></td></tr>");
                    sb.Append("<tr><td align='center' style='background-color: #18B5F0' colspan = '3'><b>NHLP Farmer Contribution Report</b></td></tr>");
                    sb.Append("<tr><td colspan = '3'></td></tr>");
                    sb.Append("<tr><td><b>Farmer Contribution Detail By  Province</b>");

                    sb.Append("</td><td colspan = '2'><b>Report Date: </b>");
                    sb.Append("" + DateTime.Now);

                    sb.Append(" </td></tr>");

                    sb.Append("</table>");
                    sb.Append("<br />");
                    sb.Append("<table border = '1'>");
                    sb.Append("<tr>");
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName != "reportdate")
                        {
                            sb.Append("<td bgcolor='orange'><font color='white'>");
                            sb.Append("<strong>" + column.ColumnName + "</strong>");
                            sb.Append("</font></td>");
                        }
                    }
                    sb.Append("</tr>");
                    string pName = null;
                    double collectable = 0, paid = 0, balnce = 0;
                    int rowslenght = dt.Rows.Count;
                    foreach (DataRow row in dt.Rows)
                    {
                        rowslenght--;
                        if (pName == null)
                            pName = row["ProvinceEngName"].ToString();
          

                        collectable += Convert.ToDouble(row["Collectable Amount"].ToString());
                        paid += Convert.ToDouble(row["Paid Amount"].ToString());
                        balnce += Convert.ToDouble(row["Balance Amount"].ToString());
                        sb.Append("<tr>");
                        foreach (DataColumn column in dt.Columns)
                        {
                            if (column.ColumnName != "reportdate")
                            {
                                sb.Append("<td>");
                                sb.Append(row[column]);
                                sb.Append("</td>");
                            }
                        }
                        sb.Append("</tr>");

                    }

                    //sb.Append("<tr  ><td ><b>Summary </b></td><td><b>" + collectable.ToString() + "</b></td><td><b>" + paid.ToString() + "</b></td><td><b>" + balnce.ToString() + "</b></td></tr></table>");

                     sb.Append("</table>");
                    StringReader sr = new StringReader(sb.ToString());

                    Document pdfDoc = new Document(iTextSharp.text.PageSize.A4.Rotate());
                    HTMLWorker htmlparser = new HTMLWorker(pdfDoc);

                    // Net

                    //1st Pdf
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                        pdfDoc.Open();
                        htmlparser.Parse(sr);
                        pdfDoc.Close();
                        byte[] bytes = memoryStream.ToArray();
                        memoryStream.Close();
                        return bytes;

                    }
                }
            }
        }
    }
}
