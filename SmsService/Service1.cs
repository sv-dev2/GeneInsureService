﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
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
            // this.timer1.Interval = 30000; //every 30 secs
            this.timer1.Interval = 86400000; // 24 Hours
            //this.timer1.Interval = 60000;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            timer1.Enabled = true;
            //Library.WriteErrorLog("window service started");

            //do
            //{

            //    //call method
            //    GetCustomerDate();
            //    //System.Threading.Thread.Sleep(86400000);
            //    System.Threading.Thread.Sleep(30000);
            //    Library.WriteErrorLog("Message sent successfully");
            //}
            //while (true);



        }

        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            //Write code here to do some job depends on your requirement
            GetCustomerDate();
            ReminderEmailSms();
            Library.WriteErrorLog("Timer ticked and some job has been done successfully");
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
                    query += "     Where CONVERT(varchar,VehicleDetail.RenewalDate, 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101)";

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

                        var paths=   (AppDomain.CurrentDomain.BaseDirectory + @"EmailTemplate\RenewalReminderEmail.html");

                        //string ReminderEmailPath = "/SmaService/SmsService/SmsService/EmailTemplate/RenewalReminderEmail.html";
                       

                        string EmailBody2 = System.IO.File.ReadAllText(paths);
                        
                        var _body = EmailBody2.Replace("##RenewDate##", item.RenewalDate.ToString("dd/MM/yyyy")).Replace("##path##", filepath).Replace("##FirstName##", item.FirstName).Replace("##LastName##", item.LastName).Replace("##PolicyNumber##", item.PolicyNumber).Replace("##Make##", item.MakeDescription).Replace("##Model##", item.ModelDescription).Replace("##Address1##", item.AddressLine1).Replace("##Address2##", item.AddressLine2).Replace("##numberofDays##", "" + dr["NoOfDays"] + "").Replace("##PolicyNumber##", item.PolicyNumber);
                        var email = item.Email;

                        try
                        {
                            if(!item.IsCustomEmail)
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
                    query += "     Where CONVERT(varchar,VehicleDetail.RenewalDate, 101) = CONVERT(varchar, DATEADD(day, " + dr["NoOfDays"] + ", GETDATE()), 101)";

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

        //public DataTable GetVehicleList()
        //{
        //    DataTable vehicletlist = new DataTable();
        //    string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();
        //    var list = "Select * from "



        //    return vehicletlist;
        //}

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


    }
}