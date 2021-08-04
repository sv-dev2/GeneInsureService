
using InsuranceClaim.Models;

using OfficeOpenXml;
using SmsService;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Insurance.Service
{
    public class WeeklyGWPService
    {
        public void SendWeeklyGwpFile()
        {
            List<GrossWrittenPremiumReportModels> ListGrossWrittenPremiumReport = new List<GrossWrittenPremiumReportModels>();
            ListGrossWrittenPremiumReportModels _ListGrossWrittenPremiumReport = new ListGrossWrittenPremiumReportModels();
            _ListGrossWrittenPremiumReport.ListGrossWrittenPremiumReportdata = new List<GrossWrittenPremiumReportModels>();

            //var yesterdayDate = DateTime.Now.AddDays(-1);
            var yesterdayDate = DateTime.Now;
            int PayLater = 7;

            var query = " select PolicyDetail.PolicyNumber as Policy_Number, Customer.ALMId, case when Customer.ALMId is null  then  [dbo].fn_GetUserCallCenterAgent(SummaryDetail.CreatedBy) else [dbo].fn_GetUserALM(Customer.BranchId) end  as PolicyCreatedBy, Customer.FirstName + ' ' + Customer.LastName as Customer_Name,VehicleDetail.TransactionDate as Transaction_date, ";
            query += "  case when Customer.id=SummaryDetail.CreatedBy then [dbo].fn_GetUserBranch(Customer.id) else [dbo].fn_GetUserBranch(SummaryDetail.CreatedBy) end as BranchName, ";
            query += " VehicleDetail.CoverNote as CoverNoteNum, PaymentMethod.Name as Payment_Mode, PaymentTerm.Name as Payment_Term,CoverType.Name as CoverType, Currency.Name as Currency, ";
            query += " VehicleDetail.Premium + VehicleDetail.StampDuty + VehicleDetail.ZTSCLevy as Premium_due, VehicleDetail.StampDuty as Stamp_duty, VehicleDetail.ZTSCLevy as ZTSC_Levy, ";
            query += " cast(VehicleDetail.Premium * 30 / 100 as decimal(10, 2))    as Comission_Amount, VehicleDetail.IncludeRadioLicenseCost, ";
            query += " CASE WHEN IncludeRadioLicenseCost = 1 THEN VehicleDetail.RadioLicenseCost else 0 end as RadioLicenseCost, VehicleDetail.VehicleLicenceFee as Zinara_License_Fee, ";
            query += " VehicleDetail.RenewalDate as PolicyRenewalDate, VehicleDetail.IsActive, VehicleDetail.RenewPolicyNumber as RenewPolicyNumber, ";
            query += " VehicleDetail.BusinessSourceDetailId, SummaryDetail.id as SummaryDetailId, BusinessSource.Source as BusinessSourceName, SourceDetail.FirstName + ' ' + SourceDetail.LastName as SourceDetailName from PolicyDetail ";
            query += " join Customer on PolicyDetail.CustomerId = Customer.Id ";
            query += " join VehicleDetail on PolicyDetail.Id = VehicleDetail.PolicyId ";
            query += "join SummaryVehicleDetail on VehicleDetail.id = SummaryVehicleDetail.VehicleDetailsId ";
            query += " join SummaryDetail on SummaryDetail.id = SummaryVehicleDetail.SummaryDetailId ";
            //query += "  join PaymentInformation on SummaryDetail.Id=PaymentInformation.SummaryDetailId ";
            query += " join PaymentMethod on SummaryDetail.PaymentMethodId = PaymentMethod.Id ";
            query += "join PaymentTerm on VehicleDetail.PaymentTermId = PaymentTerm.Id ";
            query += " left join CoverType on VehicleDetail.CoverTypeId = CoverType.Id ";
            query += " left join Currency on VehicleDetail.CurrencyId = Currency.Id ";
            query += " left join BusinessSource on BusinessSource.Id = VehicleDetail.BusinessSourceDetailId ";
            query += " left   join SourceDetail on VehicleDetail.BusinessSourceDetailId = SourceDetail.Id join AspNetUsers on AspNetUsers.id=customer.UserID join AspNetUserRoles on AspNetUserRoles.UserId=AspNetUsers.Id ";
            query += " where (VehicleDetail.IsActive = 1 or VehicleDetail.IsActive = null) and SummaryDetail.isQuotation=0 and SummaryDetail.PaymentMethodId <>" + PayLater + " and CONVERT(date, VehicleDetail.TransactionDate) = convert(date, '" + yesterdayDate.ToString("MM/dd/yyyy") + "', 101) order by  VehicleDetail.Id desc ";
            Debug.WriteLine("***********************");
            Debug.WriteLine(yesterdayDate.ToString("MM/dd/yyyy"));
            Debug.WriteLine("***********************");


            DataTable table = new DataTable();
            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(query, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(table);
            connection.Close();

            ListGrossWrittenPremiumReport = ConvertDataTable<GrossWrittenPremiumReportModels>(table);


            try
            {


                DataTable dt = GetListOfBranch();
                List<BranchModel> obj = ConvertDataTable<BranchModel>(dt);


                List<ZinaraReportModel> reportModelsList = new List<ZinaraReportModel>();
                obj.ForEach(x =>
                {
                    ZinaraReportModel model = new ZinaraReportModel();
                    var count = ListGrossWrittenPremiumReport.FindAll(p => p.BranchName == x.BranchName).Count();
                    var total = ListGrossWrittenPremiumReport.Where(p => p.BranchName == x.BranchName).Sum(item => item.Zinara_License_Fee);
                    // var zinaraAmount = total.Sum(item => item.Zinara_License_Fee);
                    model.BranchName = x.BranchName;
                    model.count = count;
                    model.ZinaraAmount = total;
                    reportModelsList.Add(model);
                });

                GenerateExcel(reportModelsList);

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

        }

        public DataTable GetListOfBranch()
        {

            var branchQuery = "select * from Branch where id != 6";
            DataTable table = new DataTable();
            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(branchQuery, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(table);
            connection.Close();

            return table;
        }


        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }



        private static T GetItem<T>(DataRow dr)
        {

            string columnName = "";
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();
            try
            {


            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    //in case you have a enum/GUID datatype in your model
                    //We will check field's dataType, and convert the value in it.
                    if (pro.Name == column.ColumnName)
                    {
                        try
                        {
                                columnName = column.ColumnName;


                            var convertedValue = GetValueByDataType(pro.PropertyType, dr[column.ColumnName]);
                            // pro.SetValue(obj, convertedValue, null);

                            SetValue(obj, column.ColumnName, convertedValue);

                        }
                        catch (Exception e)
                        {
                            //ex handle code                   
                            throw;
                        }
                        //pro.SetValue(obj, dr[column.ColumnName], null);
                    }
                    else
                        continue;
                }
            }

            }
            catch (Exception ex)
            {
                Library.WriteErrorLog(" ReadBirthdayMessage  :" + ex.Message + "col "+ columnName);
            }

            return obj;
        }
        //private static T GetItem<T>(DataRow dr)
        //{

        //    Type temp = typeof(T);
        //    T obj = Activator.CreateInstance<T>();

        //    try
        //    {
        //        foreach (DataColumn column in dr.Table.Columns)
        //        {
        //            foreach (PropertyInfo pro in temp.GetProperties())
        //            {
        //                if (pro.Name == column.ColumnName)
        //                    pro.SetValue(obj, dr[column.ColumnName], null);
        //                else
        //                    continue;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Library.WriteErrorLog(" read property  :" + ex.Message);
        //    }



        //    return obj;
        //}


        private static object GetValueByDataType(Type propertyType, object o)
        {
            if (o.ToString() == "null")
            {
                return null;
            }
            if (propertyType == (typeof(Guid)) || propertyType == typeof(Guid?))
            {
                return Guid.Parse(o.ToString());
            }
            else if (propertyType == typeof(int) || propertyType.IsEnum)
            {
                return Convert.ToInt32(o);
            }
            else if (propertyType == typeof(decimal))
            {
                return Convert.ToDecimal(o);
            }
            else if (propertyType == typeof(long))
            {
                return Convert.ToInt64(o);
            }
            else if (propertyType == typeof(bool) || propertyType == typeof(bool?))
            {
                return Convert.ToBoolean(o);
            }
            else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
            {
                return Convert.ToDateTime(o);
            }
            return o.ToString();
        }


        public static void SetValue(object inputObject, string propertyName, object propertyVal)
        {
            //find out the type
            Type type = inputObject.GetType();

            //get the property information based on the type
            System.Reflection.PropertyInfo propertyInfo = type.GetProperty(propertyName);

            //find the property type
            Type propertyType = propertyInfo.PropertyType;

            //Convert.ChangeType does not handle conversion to nullable types
            //if the property type is nullable, we need to get the underlying type of the property
            var targetType = IsNullableType(propertyType) ? Nullable.GetUnderlyingType(propertyType) : propertyType;

            //Returns an System.Object with the specified System.Type and whose value is
            //equivalent to the specified object.
            propertyVal = Convert.ChangeType(propertyVal, targetType);

            //Set the value of the property
            propertyInfo.SetValue(inputObject, propertyVal, null);

        }
        private static bool IsNullableType(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }


        public static void GenerateExcel(List<ZinaraReportModel> grossWrittenPremiumReports)
        {

            StreamWriter sw = null;

            try
            {
                int indexList = grossWrittenPremiumReports.Count();
                int totalsIndex = indexList + 6;
                var zinaraAmount = grossWrittenPremiumReports.Sum(x => x.ZinaraAmount);
                var zinarCount = grossWrittenPremiumReports.Sum(x => x.count);

                MemoryStream outputStream = new MemoryStream();
                using (ExcelPackage package = new ExcelPackage(outputStream))
                {

                    // export each facility's rollup and detail to tabs in Excel (two tabs per facility)
                    ExcelWorksheet facilityWorksheet = package.Workbook.Worksheets.Add("Zinara Report");
                    facilityWorksheet.Cells["A1"].LoadFromText("Zinara Report").Style.Font.Bold = true;
                    facilityWorksheet.Cells["A3"].Value = "Report Generated Date: " + DateTime.Now.ToString();

                    facilityWorksheet.Cells["A5"].LoadFromCollection(grossWrittenPremiumReports, true, OfficeOpenXml.Table.TableStyles.Light1);
                    facilityWorksheet.Cells["A" + totalsIndex.ToString()].LoadFromText("TOTALS").Style.Font.Bold = true;
                    facilityWorksheet.Cells["B" + totalsIndex.ToString()].LoadFromText(zinaraAmount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["C" + totalsIndex.ToString()].LoadFromText(zinarCount.ToString()).Style.Font.Bold = true;
                    package.Save();

                    outputStream.Position = 0;

                    List<Stream> _attachements = new List<Stream>();
                    List<AttachmentModel> attachmentModels = new List<AttachmentModel>();
                    AttachmentModel attachment = new AttachmentModel();
                    attachment.Attachment = outputStream;
                    attachment.Name = "Zinara Report.xlsx";
                    attachmentModels.Add(attachment);


                    _attachements.Add(outputStream);
                    Debug.WriteLine("************Attached*************");

                    StringBuilder mailBody = new StringBuilder();
                    mailBody.AppendFormat("<div>Please Find Attached.</div>");

                    Debug.WriteLine("***********SendEmail**************");

                    string email = System.Configuration.ConfigurationManager.AppSettings["gwpemail"];
                    string ccmail = System.Configuration.ConfigurationManager.AppSettings["gwpccmail"];

                    //gwpccmail
                    Insurance.Service.EmailService objEmailService = new Insurance.Service.EmailService();
                    objEmailService.SendAttachedEmail(email, ccmail, "", "Zinara Report - " + DateTime.Now.ToShortDateString(), mailBody.ToString(), attachmentModels);


                    Library.WriteErrorLog("Zinara Report - " + DateTime.Now.ToShortDateString());
                }



            }
            catch (Exception ex)
            {

            }
            finally
            {

            } 
        }

        public static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }


        public void SendWeeklyReport()
        {

            Library.WriteErrorLog("Start Summary GWP Report");

            var dtOneMonthBack = DateTime.Now;
            int year = dtOneMonthBack.Year;
            int month = dtOneMonthBack.Month;

            string firstWeekStart = dtOneMonthBack.ToString("MM") + "/01/" + DateTime.Now.Year.ToString();
            string firstWeekEnd = dtOneMonthBack.ToString("MM") + "/07/" + DateTime.Now.Year.ToString();
            string secondWeekStart = dtOneMonthBack.ToString("MM") + "/08/" + DateTime.Now.Year.ToString();
            string secondWeekEnd = dtOneMonthBack.ToString("MM") + "/14/" + DateTime.Now.Year.ToString();

            string thirdWeekStart = dtOneMonthBack.ToString("MM") + "/15/" + DateTime.Now.Year.ToString();
            string thirdWeekEnd = dtOneMonthBack.ToString("MM") + "/21/" + DateTime.Now.Year.ToString();
            string fourthWeekStart = dtOneMonthBack.ToString("MM") + "/22/" + DateTime.Now.Year.ToString();
            DateTime lastDate = new DateTime(year, month,
                                    DateTime.DaysInMonth(year, month));
            string fourthWeekEnd = lastDate.ToString("MM/dd/yyyy");


            //string firstWeekStart = dtOneMonthBack.ToString("MM/dd/yyyy");
            //string firstWeekEnd = dtOneMonthBack.AddDays(7).ToString("MM/dd/yyyy");

            //string secondWeekStart = dtOneMonthBack.AddDays(8).ToString("MM/dd/yyyy");
            //string secondWeekEnd = dtOneMonthBack.AddDays(14).ToString("MM/dd/yyyy");

            //string thirdWeekStart = dtOneMonthBack.AddDays(15).ToString("MM/dd/yyyy");
            //string thirdWeekEnd = dtOneMonthBack.AddDays(21).ToString("MM/dd/yyyy");
            //string fourthWeekStart = dtOneMonthBack.AddDays(22).ToString("MM/dd/yyyy");
            ////DateTime lastDate = new DateTime(year, month,
            ////                        DateTime.DaysInMonth(year, month));
            //string fourthWeekEnd = DateTime.Now.ToString("MM/dd/yyyy");

            List<GrossWrittenPremiumReportModels> ListGrossWrittenPremiumReport = new List<GrossWrittenPremiumReportModels>();
            ListGrossWrittenPremiumReportModels _ListGrossWrittenPremiumReport = new ListGrossWrittenPremiumReportModels();
            _ListGrossWrittenPremiumReport.ListGrossWrittenPremiumReportdata = new List<GrossWrittenPremiumReportModels>();
            try
            {
                ListGrossWrittenPremiumReport = getGWPData(firstWeekStart, firstWeekEnd);
                Library.WriteErrorLog("Summary GWP count: " + ListGrossWrittenPremiumReport.Count());


                var report2 = getGWPData(secondWeekStart, secondWeekEnd);
                Library.WriteErrorLog("Summary GWP count: " + report2.Count());

                var report3 = getGWPData(thirdWeekStart, thirdWeekEnd);
                Library.WriteErrorLog("Summary GWP count: " + report3.Count());

                var report4 = getGWPData(fourthWeekStart, fourthWeekEnd);
                Library.WriteErrorLog("Summary GWP count: " + report4.Count());

                var report5 = getGWPData(firstWeekStart, fourthWeekEnd);
                Library.WriteErrorLog("Summary GWP count: " + report5.Count());


                DataTable dt = GetAllBranch();
                List<BranchModel> branches = ConvertDataTable<BranchModel>(dt);

                List<WeeklyGWPModel> weeklyGWPModels = new List<WeeklyGWPModel>();
                branches.ForEach(x =>
                {
                    WeeklyGWPModel model = new WeeklyGWPModel();
                    if (x.BranchName == "Online")
                    {
                        var count = ListGrossWrittenPremiumReport.FindAll(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Count();
                        var total = ListGrossWrittenPremiumReport.Where(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Sum(item => item.Premium_due);
                        var count2 = report2.FindAll(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Count();
                        var total2 = report2.Where(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Sum(item => item.Premium_due);
                        var count3 = report3.FindAll(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Count();
                        var total3 = report3.Where(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Sum(item => item.Premium_due);
                        var count4 = report4.FindAll(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Count();
                        var total4 = report4.Where(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Sum(item => item.Premium_due);
                        var count5 = report5.FindAll(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Count();
                        var total5 = report5.Where(p => p.BranchName == x.BranchName || p.BranchName == "" || p.BranchName == null).Sum(item => item.Premium_due);

                        model.BranchName = x.BranchName;
                        model.FirstWeekCount = count;
                        model.FirstWeekValue = total;
                        model.SecondWeekCount = count2;
                        model.SecondWeekValue = total2;
                        model.ThirdWeekCount = count3;
                        model.ThirdWeekValue = total3;
                        model.FourWeekCount = count4;
                        model.FourWeekValue = total4;
                        model.TotalMonthCount = count5;
                        model.TotalMonthValue = total5;
                        weeklyGWPModels.Add(model);
                    }
                    else
                    {
                        var count = ListGrossWrittenPremiumReport.FindAll(p => p.BranchName == x.BranchName).Count();
                        var total = ListGrossWrittenPremiumReport.Where(p => p.BranchName == x.BranchName).Sum(item => item.Premium_due);
                        var count2 = report2.FindAll(p => p.BranchName == x.BranchName).Count();
                        var total2 = report2.Where(p => p.BranchName == x.BranchName).Sum(item => item.Premium_due);
                        var count3 = report3.FindAll(p => p.BranchName == x.BranchName).Count();
                        var total3 = report3.Where(p => p.BranchName == x.BranchName).Sum(item => item.Premium_due);
                        var count4 = report4.FindAll(p => p.BranchName == x.BranchName).Count();
                        var total4 = report4.Where(p => p.BranchName == x.BranchName).Sum(item => item.Premium_due);
                        var count5 = report5.FindAll(p => p.BranchName == x.BranchName).Count();
                        var total5 = report5.Where(p => p.BranchName == x.BranchName).Sum(item => item.Premium_due);

                        model.BranchName = x.BranchName;
                        model.FirstWeekCount = count;
                        model.FirstWeekValue = total;
                        model.SecondWeekCount = count2;
                        model.SecondWeekValue = total2;
                        model.ThirdWeekCount = count3;
                        model.ThirdWeekValue = total3;
                        model.FourWeekCount = count4;
                        model.FourWeekValue = total4;
                        model.TotalMonthCount = count5;
                        model.TotalMonthValue = total5;
                        weeklyGWPModels.Add(model);
                    }


                });

                GenerateExcel2(weeklyGWPModels);
            }
            catch (Exception ex)
            {
                Library.WriteErrorLog("Summary GWP Report Count - " + ex.Message);
            }

        }

        public static void GenerateExcel2(List<WeeklyGWPModel> grossWrittenPremiumReports)
        {
            try
            {
                Library.WriteErrorLog("start GenerateExcel2 ");

                var dtOneMonthBack = DateTime.Now;
                int year = dtOneMonthBack.Year;
                int month = dtOneMonthBack.Month;

                string firstWeekEnd = DateTime.Now.ToString("MM") + "/07/" + DateTime.Now.Year.ToString();
                string secondWeekEnd = DateTime.Now.ToString("MM") + "/14/" + DateTime.Now.Year.ToString();
                string thirdWeekEnd = DateTime.Now.ToString("MM") + "/21/" + DateTime.Now.Year.ToString();
                DateTime lastDate = new DateTime(year, month,
                                    DateTime.DaysInMonth(year, month));

                //string firstWeekEnd = dtOneMonthBack.AddDays(7).ToString("MM/dd/yyyy");
                //string secondWeekEnd = dtOneMonthBack.AddDays(14).ToString("MM/dd/yyyy");
                //string thirdWeekEnd = dtOneMonthBack.AddDays(21).ToString("MM/dd/yyyy");
                //DateTime lastDate = DateTime.Now;

                int firstTotalCount = grossWrittenPremiumReports.Sum(x => x.FirstWeekCount);
                int secondTotalCount = grossWrittenPremiumReports.Sum(x => x.SecondWeekCount);
                int thirdTotalCount = grossWrittenPremiumReports.Sum(x => x.ThirdWeekCount);
                int fourthTotalCount = grossWrittenPremiumReports.Sum(x => x.FourWeekCount);
                int TotalCount = grossWrittenPremiumReports.Sum(x => x.TotalMonthCount);

                decimal? firstTotalValue = grossWrittenPremiumReports.Sum(x => x.FirstWeekValue);
                decimal? secondTotalValue = grossWrittenPremiumReports.Sum(x => x.SecondWeekValue);
                decimal? thirdTotalValue = grossWrittenPremiumReports.Sum(x => x.ThirdWeekValue);
                decimal? fourthTotalValue = grossWrittenPremiumReports.Sum(x => x.FourWeekValue);
                decimal? TotalValue = grossWrittenPremiumReports.Sum(x => x.TotalMonthValue);

                int indexList = grossWrittenPremiumReports.Count();
                int totalsIndex = indexList + 7;

                MemoryStream outputStream = new MemoryStream();
                using (ExcelPackage package = new ExcelPackage(outputStream))
                {
                    var src = DateTime.Now;
                    var hm = new DateTime(src.Year, src.Month, src.Day, src.Hour, src.Minute, 0);


                    // export each facility's rollup and detail to tabs in Excel (two tabs per facility)
                    ExcelWorksheet facilityWorksheet = package.Workbook.Worksheets.Add("Summary GWP Report");
                    facilityWorksheet.Cells["A1"].LoadFromText("GWP Summary Report").Style.Font.Bold = true;
                    facilityWorksheet.Cells["A3"].Value = "Report Generated Date: " + DateTime.Now.ToString("HH:mm") + " " + DateTime.Now.ToShortDateString();
                    facilityWorksheet.Cells["A5"].LoadFromText("Week Ending").Style.Font.Bold = true;
                    facilityWorksheet.Cells["B5"].Value = firstWeekEnd;
                    facilityWorksheet.Cells["D5"].Value = secondWeekEnd;
                    facilityWorksheet.Cells["F5"].Value = thirdWeekEnd;
                    facilityWorksheet.Cells["H5"].Value = lastDate.ToString("MM/dd/yyyy");
                    facilityWorksheet.Cells[6, 1].LoadFromCollection(grossWrittenPremiumReports, true, OfficeOpenXml.Table.TableStyles.Light1);
                    facilityWorksheet.Cells["A" + totalsIndex.ToString()].LoadFromText("TOTALS").Style.Font.Bold = true;
                    facilityWorksheet.Cells["B" + totalsIndex.ToString()].LoadFromText(firstTotalCount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["C" + totalsIndex.ToString()].LoadFromText(firstTotalValue.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["D" + totalsIndex.ToString()].LoadFromText(secondTotalCount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["E" + totalsIndex.ToString()].LoadFromText(secondTotalValue.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["F" + totalsIndex.ToString()].LoadFromText(thirdTotalCount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["G" + totalsIndex.ToString()].LoadFromText(thirdTotalValue.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["H" + totalsIndex.ToString()].LoadFromText(fourthTotalCount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["I" + totalsIndex.ToString()].LoadFromText(fourthTotalValue.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["J" + totalsIndex.ToString()].LoadFromText(TotalCount.ToString()).Style.Font.Bold = true;
                    facilityWorksheet.Cells["K" + totalsIndex.ToString()].LoadFromText(TotalValue.ToString()).Style.Font.Bold = true;
                    //facilityDetail.Cells.LoadFromDataTable(dataTable, true);

                    package.Save();

                    outputStream.Position = 0;

                    List<Stream> _attachements = new List<Stream>();
                    List<AttachmentModel> attachmentModels = new List<AttachmentModel>();
                    AttachmentModel attachment = new AttachmentModel();
                    attachment.Attachment = outputStream;
                    attachment.Name = "Summary GWP Report.xlsx";
                    attachmentModels.Add(attachment);

                    _attachements.Add(outputStream);
                    Debug.WriteLine("************Attached*************");

                    StringBuilder mailBody = new StringBuilder();
                    mailBody.AppendFormat("<div>Please Find Attached.</div>");

                    Debug.WriteLine("***********SendEmail**************");

                    string email = System.Configuration.ConfigurationManager.AppSettings["gwpemail"];

                   // email = "kindlebit.net@gmail.com";

                    Insurance.Service.EmailService objEmailService = new Insurance.Service.EmailService();
                    objEmailService.SendAttachedEmail(email, "", "", "Summary GWP Report - " + DateTime.Now.ToLongDateString(), mailBody.ToString(), attachmentModels);

                    Library.WriteErrorLog("Successfully Summary GWP Report - " + DateTime.Now.ToLongDateString());
                }



            }
            catch (Exception ex)
            {
                // Info.  
                Library.WriteErrorLog("Summary GWP Report - " + ex.Message);
            }
            finally
            {
                // Closing.  
                /*      sw.Flush();
                      sw.Dispose();
                      sw.Close();*/
            }
        }

        public List<GrossWrittenPremiumReportModels> getGWPData(string startDate, string endDate)
        {
            var ListGrossWrittenPremiumReport = new List<GrossWrittenPremiumReportModels>();
            try
            {
                

                int PayLater = 6;
                var query = " select PolicyDetail.PolicyNumber as Policy_Number, Customer.ALMId, case when Customer.ALMId is null  then  [dbo].fn_GetUserCallCenterAgent(SummaryDetail.CreatedBy) else [dbo].fn_GetUserALM(Customer.BranchId) end  as PolicyCreatedBy, Customer.FirstName + ' ' + Customer.LastName as Customer_Name,VehicleDetail.TransactionDate as Transaction_date, ";
                query += "  case when Customer.id=SummaryDetail.CreatedBy then [dbo].fn_GetUserBranch(Customer.id) else [dbo].fn_GetUserBranch(SummaryDetail.CreatedBy) end as BranchName, ";
                query += " VehicleDetail.CoverNote as CoverNoteNum, PaymentMethod.Name as Payment_Mode, PaymentTerm.Name as Payment_Term,CoverType.Name as CoverType, Currency.Name as Currency, ";
                query += " VehicleDetail.Premium + VehicleDetail.StampDuty + VehicleDetail.ZTSCLevy as Premium_due, VehicleDetail.StampDuty as Stamp_duty, VehicleDetail.ZTSCLevy as ZTSC_Levy, ";
                query += " cast(VehicleDetail.Premium * 30 / 100 as decimal(10, 2))    as Comission_Amount, VehicleDetail.IncludeRadioLicenseCost, ";
                query += " CASE WHEN IncludeRadioLicenseCost = 1 THEN VehicleDetail.RadioLicenseCost else 0 end as RadioLicenseCost, VehicleDetail.VehicleLicenceFee as Zinara_License_Fee, ";
                query += " VehicleDetail.RenewalDate as PolicyRenewalDate, VehicleDetail.IsActive, VehicleDetail.RenewPolicyNumber as RenewPolicyNumber, ";
                query += " VehicleDetail.BusinessSourceDetailId, SummaryDetail.id as SummaryDetailId, BusinessSource.Source as BusinessSourceName, SourceDetail.FirstName + ' ' + SourceDetail.LastName as SourceDetailName from PolicyDetail ";
                query += " join Customer on PolicyDetail.CustomerId = Customer.Id ";
                query += " join VehicleDetail on PolicyDetail.Id = VehicleDetail.PolicyId ";
                query += "join SummaryVehicleDetail on VehicleDetail.id = SummaryVehicleDetail.VehicleDetailsId ";
                query += " join SummaryDetail on SummaryDetail.id = SummaryVehicleDetail.SummaryDetailId ";
                //query += "  join PaymentInformation on SummaryDetail.Id=PaymentInformation.SummaryDetailId ";
                query += " join PaymentMethod on SummaryDetail.PaymentMethodId = PaymentMethod.Id ";
                query += "join PaymentTerm on VehicleDetail.PaymentTermId = PaymentTerm.Id ";
                query += " left join CoverType on VehicleDetail.CoverTypeId = CoverType.Id ";
                query += " left join Currency on VehicleDetail.CurrencyId = Currency.Id ";
                query += " left join BusinessSource on BusinessSource.Id = VehicleDetail.BusinessSourceDetailId ";
                query += " left   join SourceDetail on VehicleDetail.BusinessSourceDetailId = SourceDetail.Id join AspNetUsers on AspNetUsers.id=customer.UserID join AspNetUserRoles on AspNetUserRoles.UserId=AspNetUsers.Id ";
                query += " where (VehicleDetail.IsActive = 1 or VehicleDetail.IsActive = null) and SummaryDetail.isQuotation=0  and (  CONVERT(date, VehicleDetail.TransactionDate) >= convert(date, '" + startDate + "', 101)  and CONVERT(date, VehicleDetail.TransactionDate) <= convert(date, '" + endDate + "', 101))  order by  VehicleDetail.Id desc ";

                Library.WriteErrorLog(startDate + " " + endDate);

              //  Library.WriteErrorLog("query "+ query);

                DataTable table = new DataTable();
                string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand cmd = new SqlCommand(query, connection);
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.Fill(table);
                connection.Close();

                ListGrossWrittenPremiumReport = ConvertDataTable<GrossWrittenPremiumReportModels>(table);




                //ListGrossWrittenPremiumReport = InsuranceContext.Query(query).
                //Select(x => new GrossWrittenPremiumReportModels()
                //{

                //    Policy_Number = x.Policy_Number,
                //    BranchName = x.BranchName,
                //    PolicyCreatedBy = x.PolicyCreatedBy,
                //    Customer_Name = x.Customer_Name,
                //    Transaction_date = x.Transaction_date.ToShortDateString(),
                //    CoverNoteNum = x.CoverNoteNum,
                //    Payment_Mode = x.Payment_Mode,
                //    Payment_Term = x.Payment_Term,
                //    CoverType = x.CoverType,
                //    Currency = x.Currency,
                //    Premium_due = x.Premium_due,
                //    Stamp_duty = x.Stamp_duty,
                //    ZTSC_Levy = x.ZTSC_Levy,
                //    ALMId = x.ALMId,
                //    Comission_Amount = x.Comission_Amount,
                //    RadioLicenseCost = x.RadioLicenseCost,
                //    Zinara_License_Fee = x.Zinara_License_Fee,
                //    PolicyRenewalDate = x.PolicyRenewalDate,
                //    IsActive = x.IsActive,
                //    RenewPolicyNumber = x.RenewPolicyNumber,
                //}).ToList();
                //Debug.WriteLine("**************hdfhd***************");
                //Debug.WriteLine(ListGrossWrittenPremiumReport.Sum(x => x.Premium_due));
                //Debug.WriteLine("**************hdfhd***************");

            }
            catch (Exception ex)
            {

                //Debug.WriteLine(ex);

                Library.WriteErrorLog("getGWPData - " + ex.Message);
                return ListGrossWrittenPremiumReport;
            }

            return ListGrossWrittenPremiumReport;
        }



        public DataTable GetAllBranch()
        {

            var branchQuery = "select * from Branch";
            DataTable table = new DataTable();
            string connectionString = System.Configuration.ConfigurationManager.AppSettings["Insurance"].ToString();

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand(branchQuery, connection);
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(table);
            connection.Close();

            return table;
        }


       



    }
}
