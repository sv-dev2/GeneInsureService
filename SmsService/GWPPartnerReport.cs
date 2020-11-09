using InsuranceClaim.Models;
using SimplePdfReport.Reporting;
using SimplePdfReport.Reporting.MigraDoc;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SmsService
{
    class GWPPartnerReport
    {

        string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();


        public void InitReports()
        {
            //getAll partners and genarated Data For the report


            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                SqlCommand com = new SqlCommand("select * from Partners where Status= 1");
                com.CommandType = CommandType.Text;
                com.Connection = con;
                List<PartnerModel> partnerModels = new List<PartnerModel>();



                using (SqlDataReader reader = com.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        //PartnerModel model = new PartnerModel
                        //{
                        //    Id = reader.GetInt32(0),
                        //    PartnerName = reader.GetString(1),

                        //};

                        PartnerModel model = new PartnerModel();
                        model.Id = reader["Id"] == null ? 0 : Convert.ToInt32(reader["Id"]);
                        model.PartnerName = reader["PartnerName"] == null ? "" : Convert.ToString(reader["PartnerName"]);

                        SendReport(model);

                    }
                }
            }

        }


        public void SendReport(PartnerModel model)
        {
            var reportService = new ReportPdf();


            string startDate = DateTime.Now.AddDays(-2).ToString("MM/dd/yyyy");
            string endDate = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy");
            List<ALMParnterSearchModelsData> listData = new List<ALMParnterSearchModelsData>();

            var yesterDayDate = DateTime.Now.AddDays(-1).ToString("MM-dd-yyyy");

            //  string partnerData = "select PolicyDetail.PolicyNumber as PolicyNumber, PartnerCommissions.CommissionPercentage ,Branch.BranchName as BranchName,  convert(varchar, PolicyDetail.CreatedOn, 106) as CreatedOn , PaymentInformation.PaymentId, VehicleDetail.Premium,cast(VehicleDetail.Premium * PartnerCommissions.CommissionPercentage as decimal(10, 2)) as Comission_Amount, Customer.ALMId from PolicyDetail join Customer on Customer.Id = PolicyDetail.CustomerId join VehicleDetail on VehicleDetail.PolicyId = PolicyDetail.Id join PaymentInformation on PaymentInformation.PolicyId = PolicyDetail.Id join Branch on VehicleDetail.ALMBranchId = Branch.Id join Partners on Partners.Id = Branch.PartnerId join PartnerCommissions on CommissionEffectiveDate <= PolicyDetail.CreatedOn and PartnerCommissions.PartnerId = branch.PartnerId where PolicyDetail.CreatedOn BETWEEN '" + startDate + "' AND '" + endDate + "' and branch.PartnerId=" + model.Id;

            string partnerData = "select PolicyDetail.PolicyNumber as PolicyNumber, Branch.BranchName as BranchName, ";
            partnerData += " convert(varchar, PolicyDetail.CreatedOn, 106) as CreatedOn, PaymentInformation.PaymentId, VehicleDetail.Premium, ";
            partnerData += " cast(VehicleDetail.Premium * ( select top 1  CommissionPercentage  from PartnerCommissions as b  where PartnerId=" + model.Id + " and b.CommissionEffectiveDate<= PolicyDetail.CreatedOn ";
            partnerData += " or b.CommissionEffectiveDate < (select top 1 CommissionEffectiveDate from PartnerCommissions where b.Id = b.Id+1  )) as decimal(10, 2)) as Comission_Amount,  ";
            partnerData += " Customer.ALMId  from PolicyDetail  join Customer on Customer.Id = PolicyDetail.CustomerId join VehicleDetail on VehicleDetail.PolicyId = PolicyDetail.Id  ";
            partnerData += " join PaymentInformation on PaymentInformation.PolicyId = PolicyDetail.Id  join Branch on VehicleDetail.ALMBranchId = Branch.Id  ";
            partnerData += " join Partners on Partners.Id = Branch.PartnerId where convert(varchar,PolicyDetail.CreatedOn,110)=convert(varchar,'" + yesterDayDate + "',110) and branch.PartnerId=" + model.Id;

            // convert(varchar(5),PolicyDetail.CreatedOn,110)=convert(varchar(5),getdate(),110)
            Library.WriteErrorLog("Query: " + partnerData);

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                SqlCommand com = new SqlCommand(partnerData);
                com.CommandType = CommandType.Text;
                com.Connection = con;

                using (SqlDataReader reader = com.ExecuteReader())
                {

                    if (reader.Read())
                    {

                        ALMParnterSearchModelsData data = new ALMParnterSearchModelsData();
                        data.BranchName = reader["BranchName"] == null ? "" : Convert.ToString(reader["BranchName"]);
                        data.PolicyNumber = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);

                        listData.Add(data);

                    }
                }
            }



            StructureSet structureSet = new StructureSet
            {
                Id = model.PartnerName + " ALM Daily Report ",
                Structures = listData.ToArray()
            };

            var reportData = CreateReportData(model, structureSet);
            var path = GetTempPdfPath();
            reportService.Export(path, reportData);

            //TODO list of Partners to recive emails
            var FromMailAddress = System.Configuration.ConfigurationManager.AppSettings["SendEmailFrom"].ToString();
            var password = System.Configuration.ConfigurationManager.AppSettings["SendEmailFromPassword"].ToString();
            var smtpAddress = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SendEmailSMTP"]);


            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(FromMailAddress);
            mail.To.Add("it@gene.co.zw");
            mail.Subject = "ALM Partner Daily Report ";
            mail.Body = "Please check attached file.";

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(path);
            attachment.Name = "ALM Partner Daily Report";
            mail.Attachments.Add(attachment);
            var smtp = new System.Net.Mail.SmtpClient();
            {
                smtp.Host = smtpAddress;
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtp.Credentials = new System.Net.NetworkCredential(FromMailAddress, password);
            }
            smtp.Send(mail);

        }

        private ReportData CreateReportData(PartnerModel model, StructureSet structureSet)
        {
            return new ReportData
            {
                Patient = model,
                StructureSet = structureSet
            };
        }

        private string GetTempPdfPath()
        {
            return Path.GetTempFileName() + ".pdf";
        }
    }
}
