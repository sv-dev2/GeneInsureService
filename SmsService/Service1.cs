﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SmsService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer1 = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            Library.WriteErrorLog("Reminder Email ");
            timer1 = new Timer();
           // this.timer1.Interval = 500000; //every 5 min min
                                             //  this.timer1.Interval = 150000; //every 30 secs //3600000 :one hour
            this.timer1.Interval = 3600000; //every 1 hour
                                           // this.timer1.Interval = 86400000; // 24 Hours
                                           // this.timer1.Interval = 180000;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            timer1.Enabled = true;
            //Library.WriteErrorLog("window service started");

        }

        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            //Write code here to do some job depends on your requirement

            TimeZoneInfo Zimwabe_Standard_Time = TimeZoneInfo.FindSystemTimeZoneById("South Africa Standard Time");
            DateTime dateTime_Zimwabe = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, Zimwabe_Standard_Time);

            var resHour = dateTime_Zimwabe.ToShortTimeString().Split(':');
            var resAmPm = dateTime_Zimwabe.ToShortTimeString().Split(' ');

            //var resHour = DateTime.Now.ToShortTimeString().Split(':');
            //var resAmPm = DateTime.Now.ToShortTimeString().Split(' ');

            string time = (resHour[0] + resAmPm[1]).ToString().ToLower();

            Library.WriteErrorLog("Zimbabwe time: " + time);

            if (time == "11pm")
            {
                Library.WriteErrorLog("Zimbabwe time2: " + time);
                SendZinaraDailyReport();
                SendFWPSummaryReport();
                GetRecieptReport();
                GWPPartnerReport gWPPartnerReport = new GWPPartnerReport();
                gWPPartnerReport.InitReports();
                GetCustomerDate();
                ReminderEmailSms();
                Library.WriteErrorLog("Timer ticked and some job has been done successfully");
            }


            //GetRecieptReport();
            //  SendZinaraDailyReport();
            // SendFWPSummaryReport();
            //GWPPartnerReport gWPPartnerReport = new GWPPartnerReport();
            //gWPPartnerReport.InitReports();

        }

        protected override void OnStop()
        {
            timer1.Enabled = false;
            Library.WriteErrorLog("window service stopped");
        }


        public void GetCustomerDate()
        {

            try
            {
                Library.WriteErrorLog("date:" + DateTime.Now);
                string body = ReadBirthdayMessage();

                CustomerLogic objsms = new CustomerLogic();
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

                //String day = DateTime.Now.Day.ToString();
                //String Month = DateTime.Now.Month.ToString();

                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                //  string queryString = "select CONVERT(varchar, DateOfBirth, 101) as DateOfBirth,Countrycode,PhoneNumber from Customer where CONVERT(varchar, DateOfBirth, 101)=CONVERT(varchar, GETDATE(), 101)";
                string queryString = "select CONVERT(varchar, DateOfBirth, 101) as DateOfBirth,Countrycode,PhoneNumber from Customer where convert(varchar(5),DateOfBirth,110)=convert(varchar(5),getdate(),110)";
                SqlCommand command = new SqlCommand(queryString, con);
                DataTable dt = new DataTable();
                dt.Load(command.ExecuteReader());


                Library.WriteErrorLog("Read the records and number of records=" + dt.Rows.Count);


                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string CountryCode = Convert.ToString(dt.Rows[0]["Countrycode"]);
                            string contact = Convert.ToString(dt.Rows[0]["PhoneNumber"]);
                            CountryCode = CountryCode.Replace("+", string.Empty);
                            string phonenumber = CountryCode + contact;


                            //call method

                            Library.WriteErrorLog("phone: " + phonenumber + "Body: " + body);

                            objsms.SendSMS(phonenumber, body);



                        }

                    }
                }

                con.Close();



            }
            catch (Exception ex)
            {
                Library.WriteErrorLog("Exception occured :" + ex.Message);
            }

        }

        public string ReadBirthdayMessage()
        {
            string message = "test";
            try
            {
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
                string queryString = "select top 1 Message from BirthdayMessage";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(queryString, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            message = (string)reader["Message"];
                            Library.WriteErrorLog("mesage  :" + message);
                        }
                        reader.Close();
                    }
                    connection.Close();


                }
            }
            catch (Exception ex)
            {
                Library.WriteErrorLog(" ReadBirthdayMessage  :" + ex.Message);
            }
            return message;
        }


        public void ReminderEmailSms()
        {
            try
            {
                Library.WriteErrorLog("EMAIL/SMS window service started");

                //string body = ReadEmailMessage();
                string filepath = System.Configuration.ConfigurationManager.AppSettings["urlPath"];
                //Library.WriteErrorLog("Get File path");
                CustomerLogic objsmsemail = new CustomerLogic();
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

                SqlConnection con = new SqlConnection(connectionString);

                con.Open();
                //DataTable LicenceTickets = GetLicenceTickets();
                //Library.WriteErrorLog("LicenceTickets");
                DataTable queryString = GetEmailDetail();

                DataTable SmsList = GetSMSDetail();
                //DataTable VehicleModels = Getvehiclemodel();
                //DataTable LicenceTickets = GetLicenceTickets();
                //var now = DateTime.Now;

                foreach (DataRow dr in queryString.Rows)
                {
                    //VehicleDetail model = new VehicleDetail();
                    List<VehicleDetail> vehicalList = new List<VehicleDetail>();
                    Library.WriteErrorLog("Read the Email records and number of records=" + dr.RowState);

                    //string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
                    var query = "Select VehicleDetail.Id, VehicleDetail.SumInsured, VehicleDetail.Premium, VehicleDetail.RenewalDate, VehicleMake.MakeDescription, VehicleModel.ModelDescription, Customer.UserID, Customer.FirstName, Customer.LastName, Customer.AddressLine1, Customer.AddressLine2, Customer.IsCustomEmail, PolicyDetail.PolicyNumber, AspNetUsers.Email, AspNetUsers.PhoneNumber";
                    query += " FROM VehicleDetail INNER JOIN ";
                    query += "   VehicleMake ON VehicleDetail.MakeId = VehicleMake.MakeCode INNER JOIN ";
                    query += "    VehicleModel ON VehicleDetail.ModelId = VehicleModel.ModelCode INNER JOIN";
                    query += "     Customer ON VehicleDetail.CustomerId = Customer.Id INNER JOIN";
                    query += "      PolicyDetail ON VehicleDetail.PolicyId = PolicyDetail.Id AND Customer.Id = PolicyDetail.CustomerId INNER JOIN";
                    query += "      AspNetUsers ON Customer.UserID = AspNetUsers.Id ";
                    query += "     Where CONVERT(varchar,VehicleDetail.RenewalDate, 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101) and LicExpiryDate is  null";

                    //SqlCommand cmd = new SqlCommand(query, con);
                    //Library.WriteErrorLog("Read the records and number of records=" + query);
                    using (SqlCommand command = new SqlCommand(query, con))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            VehicleDetail model = new VehicleDetail();
                            model.VehicalId = reader["Id"] == null ? 0 : Convert.ToInt32(reader["Id"]);
                            model.RenewalDate = reader["RenewalDate"] == null ? DateTime.MinValue : Convert.ToDateTime(reader["RenewalDate"]);
                            model.MakeDescription = reader["MakeDescription"] == null ? "" : Convert.ToString(reader["MakeDescription"]);
                            model.ModelDescription = reader["ModelDescription"] == null ? "" : Convert.ToString(reader["ModelDescription"]);
                            model.FirstName = reader["FirstName"] == null ? "" : Convert.ToString(reader["FirstName"]);
                            model.LastName = reader["LastName"] == null ? "" : Convert.ToString(reader["LastName"]);
                            model.AddressLine1 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            model.AddressLine2 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            model.PolicyNumber = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);
                            model.UserId = reader["UserId"] == null ? "" : Convert.ToString(reader["UserId"]);
                            model.Email = reader["Email"] == null ? "" : Convert.ToString(reader["Email"]);
                            model.PhoneNumber = reader["PhoneNumber"] == null ? "" : Convert.ToString(reader["PhoneNumber"]);
                            model.IsCustomEmail = reader["IsCustomEmail"] == null ? false : Convert.ToBoolean(reader["IsCustomEmail"]);
                            //message = (string)reader["Message"];
                            //Library.WriteErrorLog("mesage  :" + message);

                            vehicalList.Add(model);
                        }
                        reader.Close();
                    }
                    foreach (var item in vehicalList)
                    {
                        Library.WriteErrorLog("Read the records and number of records=" + item.UserId);

                        var paths = (AppDomain.CurrentDomain.BaseDirectory + @"EmailTemplate\RenewalReminderEmail.html");

                        //string ReminderEmailPath = "/SmaService/SmsService/SmsService/EmailTemplate/RenewalReminderEmail.html";


                        string EmailBody2 = System.IO.File.ReadAllText(paths);

                        var _body = EmailBody2.Replace("##RenewDate##", item.RenewalDate.ToString("dd/MM/yyyy")).Replace("##path##", filepath).Replace("##FirstName##", item.FirstName).Replace("##LastName##", item.LastName).Replace("##PolicyNumber##", item.PolicyNumber).Replace("##Make##", item.MakeDescription).Replace("##Model##", item.ModelDescription).Replace("##Address1##", item.AddressLine1).Replace("##Address2##", item.AddressLine2).Replace("##numberofDays##", "" + dr["NoOfDays"] + "").Replace("##PolicyNumber##", item.PolicyNumber);
                        var email = item.Email;

                        try
                        {
                            if (!item.IsCustomEmail)
                            {
                                objsmsemail.SendEmail(email, "", "", "Renew/Repay Next Term Premium of Your Policy | " + dr["NoOfDays"] + " ", _body, null);
                            }

                            //Library.WriteErrorLog("EMAIL send successfully");
                        }
                        catch (Exception ex)
                        {
                            //ReminderFailed(body, user.Email, "Renew/Repay Next Term Premium of Your Policy | 21 Days Left", Convert.ToInt32(ePolicyRenewReminderType.Email));
                            Library.WriteErrorLog("EMAIL send unsuccessfully");

                        }


                    }
                    vehicalList.Clear();
                }




                foreach (DataRow dr in SmsList.Rows)
                {

                    Library.WriteErrorLog("Read the SMS records and number of records=" + dr.RowState);
                    List<VehicleDetail> _vehicalList = new List<VehicleDetail>();

                    var query = "Select VehicleDetail.Id, VehicleDetail.SumInsured, VehicleDetail.Premium, VehicleDetail.RenewalDate, VehicleMake.MakeDescription, VehicleModel.ModelDescription, Customer.UserID, Customer.FirstName, Customer.LastName, Customer.AddressLine1, Customer.AddressLine2, PolicyDetail.PolicyNumber, AspNetUsers.Email,Customer.CountryCode, AspNetUsers.PhoneNumber";
                    query += " FROM VehicleDetail INNER JOIN ";
                    query += "   VehicleMake ON VehicleDetail.MakeId = VehicleMake.MakeCode INNER JOIN ";
                    query += "    VehicleModel ON VehicleDetail.ModelId = VehicleModel.ModelCode INNER JOIN";
                    query += "     Customer ON VehicleDetail.CustomerId = Customer.Id INNER JOIN";
                    query += "      PolicyDetail ON VehicleDetail.PolicyId = PolicyDetail.Id AND Customer.Id = PolicyDetail.CustomerId INNER JOIN";
                    query += "      AspNetUsers ON Customer.UserID = AspNetUsers.Id ";
                    query += "     Where CONVERT(varchar,VehicleDetail.RenewalDate, 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101) and LicExpiryDate is  null";

                    using (SqlCommand command = new SqlCommand(query, con))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            VehicleDetail _models = new VehicleDetail();
                            _models.VehicalId = reader["Id"] == null ? 0 : Convert.ToInt32(reader["Id"]);
                            _models.RenewalDate = reader["RenewalDate"] == null ? DateTime.MinValue : Convert.ToDateTime(reader["RenewalDate"]);
                            _models.MakeDescription = reader["MakeDescription"] == null ? "" : Convert.ToString(reader["MakeDescription"]);
                            _models.ModelDescription = reader["ModelDescription"] == null ? "" : Convert.ToString(reader["ModelDescription"]);
                            _models.FirstName = reader["FirstName"] == null ? "" : Convert.ToString(reader["FirstName"]);
                            _models.LastName = reader["LastName"] == null ? "" : Convert.ToString(reader["LastName"]);
                            _models.AddressLine1 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            _models.AddressLine2 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            _models.PolicyNumber = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);
                            _models.UserId = reader["UserId"] == null ? "" : Convert.ToString(reader["UserId"]);
                            _models.Email = reader["Email"] == null ? "" : Convert.ToString(reader["Email"]);
                            _models.PhoneNumber = reader["PhoneNumber"] == null ? "" : Convert.ToString(reader["PhoneNumber"]);
                            _models.CountryCode = reader["CountryCode"] == null ? "" : Convert.ToString(reader["CountryCode"]);

                            _vehicalList.Add(_models);
                        }
                        reader.Close();
                    }
                    foreach (var item in _vehicalList)
                    {

                        //   string _body = "Hello " + item.FirstName + "\nYour Vehicle " + item.MakeDescription + " " + item.ModelDescription + " will Expire in " + dr["NoOfDays"] + " days i.e on " + item.RenewalDate + ". Please Renew/Repay for your next Payment Term before the Renewal date of " + item.RenewalDate + " to continue your services otherwise your vehicle will get Lapsed." + "\nThank you.";

                        string _body = "Dear Customer, Please be advised that your motor vehicle insurance is due for renewals. Our renewals team will call you to assist with the renewal process.";


                        try
                        {
                            //objsmsemail.SendSMS(_models.PhoneNumber, _body);

                            objsmsemail.SendSMS(item.CountryCode.Replace("+", "") + item.PhoneNumber.TrimStart('0'), _body);

                            Library.WriteErrorLog("SMS send successfully");
                        }
                        catch (Exception ex)
                        {
                            Library.WriteErrorLog("SMS send unsuccessfully");

                        }

                    }
                    _vehicalList.Clear();

                }


                //for licensing 
                foreach (DataRow dr in queryString.Rows)
                {

                    //VehicleDetail model = new VehicleDetail();
                    List<VehicleDetail> vehicalList = new List<VehicleDetail>();
                    Library.WriteErrorLog("Read the Email records and number of records=" + dr.RowState);

                    //string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
                    var query = "Select VehicleDetail.Id, VehicleDetail.SumInsured, VehicleDetail.Premium, VehicleDetail.RenewalDate, VehicleMake.MakeDescription, VehicleModel.ModelDescription, Customer.UserID, Customer.FirstName, Customer.LastName, Customer.AddressLine1, Customer.AddressLine2, Customer.IsCustomEmail, PolicyDetail.PolicyNumber, AspNetUsers.Email, AspNetUsers.PhoneNumber";
                    query += " FROM VehicleDetail INNER JOIN ";
                    query += "   VehicleMake ON VehicleDetail.MakeId = VehicleMake.MakeCode INNER JOIN ";
                    query += "    VehicleModel ON VehicleDetail.ModelId = VehicleModel.ModelCode INNER JOIN";
                    query += "     Customer ON VehicleDetail.CustomerId = Customer.Id INNER JOIN";
                    query += "      PolicyDetail ON VehicleDetail.PolicyId = PolicyDetail.Id AND Customer.Id = PolicyDetail.CustomerId INNER JOIN";
                    query += "      AspNetUsers ON Customer.UserID = AspNetUsers.Id ";
                    query += "     Where CONVERT(varchar,CONVERT(datetime, VehicleDetail.LicExpiryDate), 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101) and LicExpiryDate is not null";

                    //SqlCommand cmd = new SqlCommand(query, con);
                    //Library.WriteErrorLog("Read the records and number of records=" + query);
                    using (SqlCommand command = new SqlCommand(query, con))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            VehicleDetail model = new VehicleDetail();
                            model.VehicalId = reader["Id"] == null ? 0 : Convert.ToInt32(reader["Id"]);
                            model.RenewalDate = reader["RenewalDate"] == null ? DateTime.MinValue : Convert.ToDateTime(reader["RenewalDate"]);
                            model.MakeDescription = reader["MakeDescription"] == null ? "" : Convert.ToString(reader["MakeDescription"]);
                            model.ModelDescription = reader["ModelDescription"] == null ? "" : Convert.ToString(reader["ModelDescription"]);
                            model.FirstName = reader["FirstName"] == null ? "" : Convert.ToString(reader["FirstName"]);
                            model.LastName = reader["LastName"] == null ? "" : Convert.ToString(reader["LastName"]);
                            model.AddressLine1 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            model.AddressLine2 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            model.PolicyNumber = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);
                            model.UserId = reader["UserId"] == null ? "" : Convert.ToString(reader["UserId"]);
                            model.Email = reader["Email"] == null ? "" : Convert.ToString(reader["Email"]);
                            model.PhoneNumber = reader["PhoneNumber"] == null ? "" : Convert.ToString(reader["PhoneNumber"]);
                            model.IsCustomEmail = reader["IsCustomEmail"] == null ? false : Convert.ToBoolean(reader["IsCustomEmail"]);
                            //message = (string)reader["Message"];
                            //Library.WriteErrorLog("mesage  :" + message);

                            vehicalList.Add(model);
                        }
                        reader.Close();
                    }
                    foreach (var item in vehicalList)
                    {
                        Library.WriteErrorLog("Read the records and number of records=" + item.UserId);

                        var paths = (AppDomain.CurrentDomain.BaseDirectory + @"EmailTemplate\RenewalReminderEmail.html");

                        //string ReminderEmailPath = "/SmaService/SmsService/SmsService/EmailTemplate/RenewalReminderEmail.html";


                        string EmailBody2 = System.IO.File.ReadAllText(paths);

                        var _body = EmailBody2.Replace("##RenewDate##", item.RenewalDate.ToString("dd/MM/yyyy")).Replace("##path##", filepath).Replace("##FirstName##", item.FirstName).Replace("##LastName##", item.LastName).Replace("##PolicyNumber##", item.PolicyNumber).Replace("##Make##", item.MakeDescription).Replace("##Model##", item.ModelDescription).Replace("##Address1##", item.AddressLine1).Replace("##Address2##", item.AddressLine2).Replace("##numberofDays##", "" + dr["NoOfDays"] + "").Replace("##PolicyNumber##", item.PolicyNumber);
                        var email = item.Email;

                        try
                        {
                            if (!item.IsCustomEmail)
                            {
                                objsmsemail.SendEmail(email, "", "", "Renew/Repay Next Term Premium of Your Policy | " + dr["NoOfDays"] + " ", _body, null);
                            }

                            //Library.WriteErrorLog("EMAIL send successfully");
                        }
                        catch (Exception ex)
                        {
                            //ReminderFailed(body, user.Email, "Renew/Repay Next Term Premium of Your Policy | 21 Days Left", Convert.ToInt32(ePolicyRenewReminderType.Email));
                            Library.WriteErrorLog("EMAIL send unsuccessfully");

                        }


                    }
                    vehicalList.Clear();

                }




                foreach (DataRow dr in SmsList.Rows)
                {

                    Library.WriteErrorLog("Read the SMS records and number of records=" + dr.RowState);
                    List<VehicleDetail> _vehicalList = new List<VehicleDetail>();

                    var query = "Select VehicleDetail.Id, VehicleDetail.SumInsured, VehicleDetail.Premium, VehicleDetail.RenewalDate, VehicleMake.MakeDescription, VehicleModel.ModelDescription, Customer.UserID, Customer.FirstName, Customer.LastName, Customer.AddressLine1, Customer.AddressLine2, PolicyDetail.PolicyNumber, AspNetUsers.Email,Customer.CountryCode, AspNetUsers.PhoneNumber";
                    query += " FROM VehicleDetail INNER JOIN ";
                    query += "   VehicleMake ON VehicleDetail.MakeId = VehicleMake.MakeCode INNER JOIN ";
                    query += "    VehicleModel ON VehicleDetail.ModelId = VehicleModel.ModelCode INNER JOIN";
                    query += "     Customer ON VehicleDetail.CustomerId = Customer.Id INNER JOIN";
                    query += "      PolicyDetail ON VehicleDetail.PolicyId = PolicyDetail.Id AND Customer.Id = PolicyDetail.CustomerId INNER JOIN";
                    query += "      AspNetUsers ON Customer.UserID = AspNetUsers.Id ";
                    query += "     Where CONVERT(varchar,CONVERT(datetime, VehicleDetail.LicExpiryDate), 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101) and LicExpiryDate is not null";

                    using (SqlCommand command = new SqlCommand(query, con))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            VehicleDetail _models = new VehicleDetail();
                            _models.VehicalId = reader["Id"] == null ? 0 : Convert.ToInt32(reader["Id"]);
                            _models.RenewalDate = reader["RenewalDate"] == null ? DateTime.MinValue : Convert.ToDateTime(reader["RenewalDate"]);
                            _models.MakeDescription = reader["MakeDescription"] == null ? "" : Convert.ToString(reader["MakeDescription"]);
                            _models.ModelDescription = reader["ModelDescription"] == null ? "" : Convert.ToString(reader["ModelDescription"]);
                            _models.FirstName = reader["FirstName"] == null ? "" : Convert.ToString(reader["FirstName"]);
                            _models.LastName = reader["LastName"] == null ? "" : Convert.ToString(reader["LastName"]);
                            _models.AddressLine1 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            _models.AddressLine2 = reader["AddressLine1"] == null ? "" : Convert.ToString(reader["AddressLine1"]);
                            _models.PolicyNumber = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);
                            _models.UserId = reader["UserId"] == null ? "" : Convert.ToString(reader["UserId"]);
                            _models.Email = reader["Email"] == null ? "" : Convert.ToString(reader["Email"]);
                            _models.PhoneNumber = reader["PhoneNumber"] == null ? "" : Convert.ToString(reader["PhoneNumber"]);
                            _models.CountryCode = reader["CountryCode"] == null ? "" : Convert.ToString(reader["CountryCode"]);

                            _vehicalList.Add(_models);
                        }
                        reader.Close();
                    }
                    foreach (var item in _vehicalList)
                    {

                        //   string _body = "Hello " + item.FirstName + "\nYour Vehicle " + item.MakeDescription + " " + item.ModelDescription + " will Expire in " + dr["NoOfDays"] + " days i.e on " + item.RenewalDate + ". Please Renew/Repay for your next Payment Term before the Renewal date of " + item.RenewalDate + " to continue your services otherwise your vehicle will get Lapsed." + "\nThank you.";

                        string _body = "Dear Customer, Please be advised that your motor vehicle insurance is due for renewals. Our renewals team will call you to assist with the renewal process.";


                        try
                        {
                            //objsmsemail.SendSMS(_models.PhoneNumber, _body);

                            objsmsemail.SendSMS(item.CountryCode.Replace("+", "") + item.PhoneNumber.TrimStart('0'), _body);

                            Library.WriteErrorLog("SMS send successfully");
                        }
                        catch (Exception ex)
                        {
                            Library.WriteErrorLog("SMS send unsuccessfully");

                        }

                    }
                    _vehicalList.Clear();

                }

                con.Close();
                Library.WriteErrorLog("EMAIL/SMS window service successful");
            }
            catch (Exception ex)
            {

                Library.WriteErrorLog("EMAIL/SMS window service unsuccessful");
            }
        }

        public DataTable GetLicenceTickets()
        {
            DataTable table = new DataTable();
            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
            //var LicenceTickets = InsuranceContext.LicenceTickets.All(where: $"CAST(CreatedDate as date) <= '{DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")}'");
            var data = "Select * from LicenceTicket where CAST (CreatedDate as date)<= getdate()";

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(data, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(table);
            connection.Close();
            return table;
        }
        public DataTable GetEmailDetail()
        {
            DataTable emaildata = new DataTable();
            try
            {

                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
                var getemail = "Select * from PolicyRenewReminderSetting where Email = 1";

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand cmd = new SqlCommand(getemail, connection);
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.Fill(emaildata);
                connection.Close();
                Library.WriteErrorLog("Email Detail successfully");

            }
            catch (Exception ex)
            {

                Library.WriteErrorLog("Email Detail unsuccessfully");
            }
            return emaildata;
        }

        public DataTable GetSMSDetail()
        {
            DataTable smsdata = new DataTable();
            try
            {
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
                var getsms = "Select * from PolicyRenewReminderSetting where SMS = 1";

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand cmd = new SqlCommand(getsms, connection);
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.Fill(smsdata);
                connection.Close();

                Library.WriteErrorLog("sms Detail successfully");
            }
            catch (Exception ex)
            {

                Library.WriteErrorLog("Sms Detail unsuccessfully");
            }
            return smsdata;

        }

        public DataTable Getvehiclemake()
        {
            DataTable vehicledetail = new DataTable();

            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
            var getvehicle = "Select * From VehicleMake";

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(getvehicle, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(vehicledetail);
            connection.Close();

            return vehicledetail;
        }
        public DataTable Getvehiclemodel()
        {


            DataTable vehiclemodeldetail = new DataTable();

            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
            var getvehiclemodel = "Select * From VehicleModel";

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(getvehiclemodel, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(vehiclemodeldetail);
            connection.Close();


            return vehiclemodeldetail;
        }
        public string ReadEmailMessage()
        {
            string message = "test";
            try
            {
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

                string queryString = "select top 1 Message from BirthdayMessage";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(queryString, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        // Call Read before accessing data.
                        while (reader.Read())
                        {
                            message = (string)reader["Message"];
                            Library.WriteErrorLog("mesage  :" + message);
                        }
                        reader.Close();
                    }
                    connection.Close();


                }
            }
            catch (Exception ex)
            {
                Library.WriteErrorLog(" ReadBirthdayMessage  :" + ex.Message);
            }

            return message;
        }


        #region
        private void SendGWPExcelFile()
        {
            DataTable dataTable = GetGWPData();
            string destFilePath = "";
            // destFilePath = @"C:\inetpub\GeneWebsite_latest\CsvFile\GwpReport.csv";


            string uniqueId = Guid.NewGuid().ToString();

            string CsvFileFolder = @"C:\inetpub\GeneWebsite_latest\CsvFile\" + uniqueId;

            if (!Directory.Exists(CsvFileFolder))
            {
                Directory.CreateDirectory(CsvFileFolder);
            }

            var filepath = CsvFileFolder + @"\GwpReport.csv";
            using (StreamWriter writer = new StreamWriter(new FileStream(filepath,
            FileMode.Create, FileAccess.Write)))
            {
            }

            //  destFilePath = Server.MapPath("~/CsvFile/GwpReport.csv");

            destFilePath = filepath;


            // Initilization  
            bool isSuccess = false;
            StreamWriter sw = null;

            try
            {
                // Initialization.  
                StringBuilder stringBuilder = new StringBuilder();

                // Saving Column header.  
                stringBuilder.Append(string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList()) + "\n");

                // Saving rows.  
                dataTable.AsEnumerable().ToList<DataRow>().ForEach(row => stringBuilder.Append(string.Join(",", row.ItemArray) + "\n"));

                // Initialization.  
                string fileContent = stringBuilder.ToString();
                sw = new StreamWriter(new FileStream(destFilePath, FileMode.Create, FileAccess.Write));

                // Saving.  
                sw.Write(fileContent);

                // Settings.  
                isSuccess = true;

                CustomerLogic objsmsemail = new CustomerLogic();
                List<string> _attachements = new List<string>();
                //urlPath

                string urlPath = System.Configuration.ConfigurationManager.AppSettings["urlPath"];

                // string path = urlPath + @"CsvFile/GwpReport.csv";

                string path = urlPath + @"/CsvFile/" + uniqueId + "/GwpReport.csv";

                _attachements.Add(path);

                //  string body = "Please check attached =" + DateTime.Now.ToShortDateString() + " GWP Report";


                StringBuilder mailBody = new StringBuilder();
                mailBody.AppendFormat("<h1>Please click below link to get gwp report.</h1>");
                mailBody.AppendFormat("<p><a href='" + path + "'>GWPReport</a></p>");

                string ccList = "charles@gene.co.zw, emmanuel@gene.co.zw, masimba@gene.co.zw, blessing@gene.co.zw";

                objsmsemail.SendEmail("it@gene.co.zw", ccList, "", "GWPReport_" + DateTime.Now.ToShortDateString(), mailBody.ToString(), _attachements);

                //it@gene.co.zw



            }
            catch (Exception ex)
            {
                // Info.  
                throw ex;
            }
            finally
            {
                // Closing.  
                sw.Flush();
                sw.Dispose();
                sw.Close();
            }
            // Info.  
            //return isSuccess;
        }
        public DataTable GetGWPData()
        {
            int PayLater = 7;
            DataTable table = new DataTable();
            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
            //var LicenceTickets = InsuranceContext.LicenceTickets.All(where: $"CAST(CreatedDate as date) <= '{DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")}'");

            var yesterdayDate = DateTime.Now.AddDays(-1);

            //   var yesterdayDate = DateTime.Now.AddMonths(-2);

            var query = " select PolicyDetail.PolicyNumber as Policy_Number, Customer.ALMId, case when Customer.ALMId is null  then  [dbo].fn_GetUserCallCenterAgent(SummaryDetail.CreatedBy) else [dbo].fn_GetUserALM(Customer.BranchId) end  as PolicyCreatedBy, Customer.FirstName + ' ' + Customer.LastName as Customer_Name,VehicleDetail.TransactionDate as Transaction_date, ";
            query += "  case when Customer.id=SummaryDetail.CreatedBy then [dbo].fn_GetUserBranch(Customer.id) else [dbo].fn_GetUserBranch(SummaryDetail.CreatedBy) end as BranchName, ";
            query += " VehicleDetail.CoverNote as CoverNoteNum, PaymentMethod.Name as Payment_Mode, PaymentTerm.Name as Payment_Term,CoverType.Name as CoverType, Currency.Name as Currency, ";
            query += " VehicleDetail.Premium + VehicleDetail.StampDuty + VehicleDetail.ZTSCLevy as Premium_due, VehicleDetail.StampDuty as Stamp_duty, VehicleDetail.ZTSCLevy as ZTSC_Levy, ";
            query += " cast(VehicleDetail.Premium * 30 / 100 as decimal(10, 2))    as Comission_Amount, VehicleDetail.IncludeRadioLicenseCost, ";
            query += " CASE WHEN IncludeRadioLicenseCost = 1 THEN VehicleDetail.RadioLicenseCost else 0 end as RadioLicenseCost, VehicleDetail.VehicleLicenceFee as Zinara_License_Fee, ";
            query += " VehicleDetail.RenewalDate as PolicyRenewalDate, VehicleDetail.IsActive, VehicleDetail.RenewPolicyNumber as RenewPolicyNumber ";
            query += "  from PolicyDetail ";
            query += " join Customer on PolicyDetail.CustomerId = Customer.Id ";
            query += " join VehicleDetail on PolicyDetail.Id = VehicleDetail.PolicyId ";
            query += "join SummaryVehicleDetail on VehicleDetail.id = SummaryVehicleDetail.VehicleDetailsId ";
            query += " join SummaryDetail on SummaryDetail.id = SummaryVehicleDetail.SummaryDetailId ";
            // query += "  join PaymentInformation on SummaryDetail.Id=PaymentInformation.SummaryDetailId ";
            query += " join PaymentMethod on SummaryDetail.PaymentMethodId = PaymentMethod.Id ";
            query += "join PaymentTerm on VehicleDetail.PaymentTermId = PaymentTerm.Id ";
            query += " left join CoverType on VehicleDetail.CoverTypeId = CoverType.Id ";
            query += " left join Currency on VehicleDetail.CurrencyId = Currency.Id ";
            query += " left join BusinessSource on BusinessSource.Id = VehicleDetail.BusinessSourceDetailId ";
            query += " left   join SourceDetail on VehicleDetail.BusinessSourceDetailId = SourceDetail.Id join AspNetUsers on AspNetUsers.id=customer.UserID join AspNetUserRoles on AspNetUserRoles.UserId=AspNetUsers.Id ";
            query += " where (VehicleDetail.IsActive = 1 or VehicleDetail.IsActive = null) and SummaryDetail.isQuotation=0 and SummaryDetail.PaymentMethodId <>" + PayLater + " and  CONVERT(date, VehicleDetail.TransactionDate) = convert(date, '" + yesterdayDate.ToShortDateString() + "', 101)  order by  VehicleDetail.Id desc ";

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(query, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(table);
            connection.Close();

            if (table.Rows.Count > 0)
            {
                Library.WriteErrorLog("sum of total: ");
                DataRow row = table.NewRow();
                row["Currency"] = "Total";
                row["Premium_due"] = table.Compute("Sum(Premium_due)", "Premium_due > 0");
                row["Stamp_duty"] = table.Compute("Sum(Stamp_duty)", "Stamp_duty > 0");
                row["ZTSC_Levy"] = table.Compute("Sum(ZTSC_Levy)", "ZTSC_Levy > 0");

                row["RadioLicenseCost"] = table.Compute("Sum(RadioLicenseCost)", "RadioLicenseCost > 0");
                row["Zinara_License_Fee"] = table.Compute("Sum(Zinara_License_Fee)", "Zinara_License_Fee > 0");
                table.Rows.Add(row);
            }



            Library.WriteErrorLog("row count: " + table.Rows.Count);

            return table;
        }

        #endregion

        #region daily gwp report

        public void SendZinaraDailyReport()
        {
            Insurance.Service.WeeklyGWPService service = new Insurance.Service.WeeklyGWPService();
            service.SendWeeklyGwpFile();
        }

        public void SendFWPSummaryReport()
        {
            Insurance.Service.WeeklyGWPService service = new Insurance.Service.WeeklyGWPService();
            service.SendWeeklyReport();
        }



        public void GetRecieptReport()
        {

            string startDate = DateTime.Now.AddMonths(-4).ToShortDateString();
            string endDate = DateTime.Now.ToShortDateString();

            Library.WriteErrorLog(" GetRecieptReport :");

            var query = " select [dbo].[fn_GetUserCallCenterAgent] (SummaryDetail.CreatedBy) as AgentName,  PolicyDetail.id as PolicyId,  PolicyDetail.PolicyNumber, VehicleDetail.RegistrationNo,VehicleDetail.TransactionDate, SummaryDetail.CreatedOn, Customer.FirstName + ' ' + Customer.LastName as CustomerName, SummaryDetail.TotalPremium  from PolicyDetail join VehicleDetail on PolicyDetail.Id = VehicleDetail.PolicyId ";
            query += " join SummaryVehicleDetail on VehicleDetail.Id = SummaryVehicleDetail.VehicleDetailsId ";
            query += " join SummaryDetail on SummaryDetail.Id = SummaryVehicleDetail.SummaryDetailId join Customer on VehicleDetail.CustomerId=Customer.Id ";
            query += " where(VehicleDetail.IsActive = 1 or VehicleDetail.IsActive = null) and SummaryDetail.isQuotation = 0 and [dbo].[fn_GetUserBranch] (SummaryDetail.CreatedBy) ='Gene Call Centre' and(CONVERT(date, VehicleDetail.TransactionDate) >= convert(date, '" + startDate + "', 101)  and CONVERT(date, VehicleDetail.TransactionDate) <= convert(date, '" + endDate + "', 101)) order by VehicleDetail.TransactionDate  desc ";

            Library.WriteErrorLog("GetRecieptReport Query :" + query);

            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

            var result = new List<RecieptModel>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    // Call Read before accessing data.
                    while (reader.Read())
                    {
                        RecieptModel model = new RecieptModel();

                        model.AgentName = reader["AgentName"] == null ? "" : Convert.ToString(reader["AgentName"]);
                        model.Policy_Number = reader["PolicyNumber"] == null ? "" : Convert.ToString(reader["PolicyNumber"]);
                        model.VRN = reader["RegistrationNo"] == null ? "" : Convert.ToString(reader["RegistrationNo"]);
                        model.Transaction_date = reader["CreatedOn"] == null ? "" : Convert.ToString(reader["CreatedOn"]);
                        model.Customer_Name = reader["CustomerName"] == null ? "" : Convert.ToString(reader["CustomerName"]);
                        model.Premium_due = reader["TotalPremium"] == null ? 0 : Convert.ToDecimal(reader["TotalPremium"]);
                        model.PolicyId = reader["PolicyId"] == null ? 0 : Convert.ToInt32(reader["PolicyId"]);

                        result.Add(model);

                    }
                    reader.Close();
                }
                connection.Close();
            }



            var query2 = "select * from ReceiptModuleHistory where (CONVERT(date, ReceiptModuleHistory.CreatedOn) >= convert(date, '" + startDate + "', 101) ";
            query2 += " and CONVERT(date, ReceiptModuleHistory.CreatedOn) <= convert(date, '" + endDate + "', 101)) and (ReceiptModuleHistory.IsActive is null or ReceiptModuleHistory.IsActive=1) order by ReceiptModuleHistory.CreatedOn desc";


            Library.WriteErrorLog(" GetRecieptReport start query 2");

            var recieptList = new List<RecieptModel>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query2, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    // Call Read before accessing data.
                    while (reader.Read())
                    {
                        RecieptModel model = new RecieptModel();
                        model.Transaction_date = reader["CreatedOn"] == null ? "" : Convert.ToString(reader["CreatedOn"]);
                        model.PolicyId = reader["PolicyId"] == null ? 0 : Convert.ToInt32(reader["PolicyId"]);
                        recieptList.Add(model);

                    }
                    reader.Close();
                }
                connection.Close();
            }


            List<RecieptDetail> list = new List<RecieptDetail>();
            DateTime dtCurrent = Convert.ToDateTime(DateTime.Now.ToShortDateString());

            foreach (var item in result)
            {
                DateTime policyCreatedDate = Convert.ToDateTime(item.Transaction_date);
                TimeSpan t = dtCurrent.Subtract(policyCreatedDate);
                //if (t.TotalDays == 7 || t.TotalDays == 14 || t.TotalDays == 21 || t.TotalDays == 30)
                //{
                var receiptDetail = recieptList.FirstOrDefault(c => c.PolicyId == item.PolicyId);
                if (receiptDetail == null)
                {
                    item.Days = Convert.ToInt32(t.TotalDays).ToString();

                    //RecieptDetail detail = new RecieptDetail
                    //{
                    //    AgentName = item.AgentName,
                    //    Policy_Number = item.Policy_Number,
                    //    Customer_Name = item.Customer_Name,
                    //    Premium_due = item.Premium_due,
                    //    VRN = item.VRN,
                    //    Transaction_date = item.Transaction_date
                    //};


                    RecieptDetail detail = new RecieptDetail();
                    detail.AgentName = item.AgentName;
                    detail.Policy_Number = item.Policy_Number;
                    detail.Customer_Name = item.Customer_Name;
                    detail.Premium_due = item.Premium_due;
                    detail.VRN = item.VRN;
                    detail.Transaction_date = item.Transaction_date;

                    if (t.TotalDays < 8)
                        detail.Days7 = Convert.ToInt32(item.Days).ToString();
                    else if (t.TotalDays < 15)
                        detail.Days14 = Convert.ToInt32(item.Days).ToString();
                    else if (t.TotalDays < 22)
                        detail.Days21 = Convert.ToInt32(item.Days).ToString();
                    else
                        detail.Days22 = Convert.ToInt32(item.Days).ToString();

                    list.Add(detail);
                }

            }


            Library.WriteErrorLog("GetRecieptReport Count: " + list.Count());

            MemoryStream outputStream = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(outputStream))
            {
                var src = DateTime.Now;
                var hm = new DateTime(src.Year, src.Month, src.Day, src.Hour, src.Minute, 0);


                // export each facility's rollup and detail to tabs in Excel (two tabs per facility)
                ExcelWorksheet facilityWorksheet = package.Workbook.Worksheets.Add("Unreceipt Report");
                facilityWorksheet.Cells["A1"].LoadFromText("Receipt Report").Style.Font.Bold = true;
                facilityWorksheet.Cells["A3"].Value = "Report Generated Date: " + DateTime.Now.ToString("HH:mm") + " " + DateTime.Now.ToShortDateString();
                facilityWorksheet.Cells[4, 1].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.Light1);

                //facilityDetail.Cells.LoadFromDataTable(dataTable, true);

                package.Save();

                outputStream.Position = 0;

                List<Stream> _attachements = new List<Stream>();
                List<InsuranceClaim.Models.AttachmentModel> attachmentModels = new List<InsuranceClaim.Models.AttachmentModel>();
                InsuranceClaim.Models.AttachmentModel attachment = new InsuranceClaim.Models.AttachmentModel();
                attachment.Attachment = outputStream;
                attachment.Name = "Unreceipt Report.xlsx";
                attachmentModels.Add(attachment);

                _attachements.Add(outputStream);
                Debug.WriteLine("************Attached*************");

                StringBuilder mailBody = new StringBuilder();
                mailBody.AppendFormat("<div>Please Find Attached.</div>");

                Debug.WriteLine("***********SendEmail**************");

                string email = System.Configuration.ConfigurationManager.AppSettings["gwpemail"];

                 //email = "kindlebit.net@gmail.com";


                Insurance.Service.EmailService objEmailService = new Insurance.Service.EmailService();
                objEmailService.SendAttachedEmail(email, "", "", "Unreceipt Report - " + DateTime.Now.ToLongDateString(), mailBody.ToString(), attachmentModels);

                Library.WriteErrorLog(" GetRecieptReport mail sent");
            }
        }


        #endregion



    }
}
