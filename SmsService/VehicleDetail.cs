using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmsService
{
   public  class VehicleDetail
    {
        public int VehicalId { get; set; }
        public DateTime RenewalDate { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PolicyNumber { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string MakeDescription { get; set; }
        public string ModelDescription { get; set; }
        public string UserId { get; set; }
        public string Email { get; set; }
        public string PhoneNumber { get; set; }

        public string CountryCode { get; set; }

        public bool IsCustomEmail { get; set; }
    }

    public class GrossWrittenPremiumReportModels
    {
        public int? Id { get; set; }
        public string Customer_Name { get; set; }
        public string Policy_Number { get; set; }
        public string Policy_endate { get; set; }
        public string Policy_startdate { get; set; }
        public string Transaction_date { get; set; }
        public string Vehicle_makeandmodel { get; set; }
        public string Payment_Mode { get; set; }
        public string Payment_Term { get; set; }
        public decimal Annual_Premium { get; set; }
        public decimal Stamp_duty { get; set; }
        public decimal ZTSC_Levy { get; set; }
        public decimal? Net_Premium { get; set; }
        public decimal Premium_due { get; set; }
        public decimal Comission_percentage { get; set; }
        public decimal Comission_Amount { get; set; }
        public decimal Sum_Insured { get; set; }
        public decimal? RadioLicenseCost { get; set; }

        public string CoverType { get; set; }

        public decimal? Zinara_License_Fee { get; set; }

        public string PolicyCreatedBy { get; set; }

        public DateTime PolicyRenewalDate { get; set; }
        public bool? IsLapsed { get; set; }
        public bool? IsActive { get; set; }
        public string Currency { get; set; }

        public string RenewPolicyNumber { get; set; }

        public string ALMId { get; set; }

        public string CoverNoteNum { get; set; }

        public string BranchName { get; set; }

        public string BusinessSourceName { get; set; }

        public string SourceDetailName { get; set; }

        public int SummaryDetailId { get; set; }


    }

    public class ListGrossWrittenPremiumReportModels
    {
        public List<GrossWrittenPremiumReportModels> ListGrossWrittenPremiumReportdata { get; set; }
    }

    public class BranchModel
    {
        public int Id { get; set; }
       
        public string BranchName { get; set; }

        public string AlmId { get; set; }
    }


}
