using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SF_Lib
    {

    public class Import_HII_CSV : ICloneable
        {

        public string Sf_MemberID { get; set; }
        public string Sf_MemberID_Bkup { get; set; }
        public string Applicant_ID { get; set; }
        public string First_Name { get; set; }
        public string Last_Name { get; set; }
        public string Gender { get; set; }
        public string DOB { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Member_ID { get; set; }
        public string Product_Name { get; set; }
        public string Option { get; set; }
        public string Type { get; set; }
        public string Policy_Number { get; set; }
        public string Premium { get; set; }
        public string Enrollment_Fee { get; set; }
        public string Administration_Fee { get; set; }
        public string Myewellness_Fee { get; set; }
        public string TelaDoc_Fee { get; set; }
        public string Extra_Care_Package_Fee { get; set; }
        public string RxCard_Fee { get; set; }
        public string Utra_Care_Plus { get; set; }
        public string Provider_Fee { get; set; }
        public string RxAdvocacy_Fee { get; set; }
        public string Critical_Illness_Fee { get; set; }
        public string Multicare_Fee { get; set; }
        public string Med_Sense_Fee { get; set; }
        public string Savers_Package { get; set; }
        public string Dental_Benefit_Fee { get; set; }
        public string Freedom_Acc_Exp__Fee { get; set; }
        public string AD_D_AME_Starr_AD_D_Fee { get; set; }
        public string Starr_AD_Fee { get; set; }
        public string ING_ADD_Fee { get; set; }
        public string Chiro___Podiatry { get; set; }
        public string Care24x7 { get; set; }
        public string EyeMed { get; set; }
        public string HCC_Association_Fee { get; set; }
        public string Kare_360_Fee { get; set; }
        public string Total_Collected { get; set; }
        public string Coinsurance { get; set; }
        public string Coinsurance_Percentage { get; set; }
        public string Deductible { get; set; }
        public string Duration_Coverage { get; set; }
        public string Application_Date { get; set; }
        public string Effective_Date { get; set; }
        public string Termination_Date { get; set; }
        public string Payment_Status { get; set; }
        public string Payment_Status_Date { get; set; }
        public string Payment_Method { get; set; }
        public string Agent_First_Name { get; set; }
        public string Agent_Last_Name { get; set; }
        public string Agent_Code { get; set; }
        public string Cancellation_Code { get; set; }
        public string Agency_Company_Name { get; set; }
        public string OpportunityLine_id { get; set; }
        public string stringProduct_PlanType { get; set; }
        public string CurrentRuning_total { get; set; }
        public string Agent_Fullname { get; set; }
        public string Agent_id { get; set; }
        public string Opportunity_name { get; set; }
        public string Opportunity_id { get; set; }
        public string Agent_Status { get; set; }
        public string Agent_Commision { get; set; }
        public bool Verify { get; set; }
        public string Product_PlanType { get; set; }
        public bool MisMachEnrrol { get; set; }
        public DateTime Payroll_date { get; set; }
        public string Payroll_Type { get; set; }
        public bool IsUpdate { get; set; }
        public DateTime Ending_Date { get; set; }
        public bool IsManual { get; set; }
        public bool from_commission { get; set; }
        public string Agent__ProductProfile { get; set; }
        public string Oppline_Status__c { get; set; }
        public string Provider_Name { get; set; }
        public bool New_sales { get; set; }
        public bool Forward { get; set; }
        public object Clone()
            {
            return this.MemberwiseClone();
            }

        public Import_HII_CSV()
            {
            Applicant_ID = string.Empty;
            First_Name = string.Empty;
            Last_Name = string.Empty;
            Gender = string.Empty;
            DOB = string.Empty;
            Address = string.Empty;
            City = string.Empty;
            State = string.Empty;
            ZipCode = string.Empty;
            Phone = string.Empty;
            Email = string.Empty;
            Member_ID = string.Empty;
            Product_Name = string.Empty;
            Option = string.Empty;
            Type = string.Empty;
            Policy_Number = string.Empty;
            Premium = string.Empty;
            Enrollment_Fee = string.Empty;
            Administration_Fee = string.Empty;
            Myewellness_Fee = string.Empty;
            TelaDoc_Fee = string.Empty;
            Extra_Care_Package_Fee = string.Empty;
            RxCard_Fee = string.Empty;
            Utra_Care_Plus = string.Empty;
            Provider_Fee = string.Empty;
            RxAdvocacy_Fee = string.Empty;
            Critical_Illness_Fee = string.Empty;
            Multicare_Fee = string.Empty;
            Med_Sense_Fee = string.Empty;
            Savers_Package = string.Empty;
            Dental_Benefit_Fee = string.Empty;
            Freedom_Acc_Exp__Fee = string.Empty;
            AD_D_AME_Starr_AD_D_Fee = string.Empty;
            Starr_AD_Fee = string.Empty;
            ING_ADD_Fee = string.Empty;
            Chiro___Podiatry = string.Empty;
            Care24x7 = string.Empty;
            EyeMed = string.Empty;
            HCC_Association_Fee = string.Empty;
            Kare_360_Fee = string.Empty;
            Total_Collected = string.Empty;
            Coinsurance = string.Empty;
            Coinsurance_Percentage = string.Empty;
            Deductible = string.Empty;
            Duration_Coverage = string.Empty;
            Application_Date = string.Empty;
            Effective_Date = string.Empty;
            Termination_Date = string.Empty;
            Payment_Status = string.Empty;
            Payment_Status_Date = string.Empty;
            Payment_Method = string.Empty;
            Agent_First_Name = string.Empty;
            Agent_Last_Name = string.Empty;
            Agent_Code = string.Empty;
            Cancellation_Code = string.Empty;
            Agency_Company_Name = string.Empty;
            OpportunityLine_id = string.Empty;
            stringProduct_PlanType = string.Empty;
            CurrentRuning_total = string.Empty;
            Agent_Fullname = string.Empty;
            Agent_id = string.Empty;
            Opportunity_name = string.Empty;
            Opportunity_id = string.Empty;
            Agent_Status = string.Empty;
            Agent_Commision = string.Empty;
            Verify = false;
            Product_PlanType = string.Empty;
            MisMachEnrrol = true;
            Payroll_date = DateTime.MinValue;
            Payroll_Type = string.Empty;
            IsManual = false;
            from_commission = false;
            Agent__ProductProfile = "";
            Oppline_Status__c = "notFound";
            New_sales = false;

            }
        }


    public class Fidelity_file
        {
        public string Agent { get; set; }
        public string ReportDate { get; set; }
        public string Amount { get; set; }
        public string Category { get; set; }
        public string ReportingCategory { get; set; }
        public string PolicyNo { get; set; }
        public string State { get; set; }
        public string Insured { get; set; }
        public string WritingAgent { get; set; }
        public string WritingAgentName { get; set; }
        public string ProductGroup { get; set; }
        public string Contract { get; set; }
        public string CommissionRate { get; set; }
        public string SplitPct { get; set; }
        public string MonthsPaid { get; set; }
        public string Duration { get; set; }
        public string Premium { get; set; }
        public string PaidToDate { get; set; }
        public string EntryDate { get; set; }
        public string Other { get; set; }

        }

    public class GetMed_file
        {
        public string Member_ID { get; set; }
        public string IsNewPolicy { get; set; }
        public string StatusName { get; set; }
        public string CoverageType { get; set; }
        public string agent_id { get; set; }
        public string AgentFirstName { get; set; }
        public string FirstName { get; set; }
        public string MI { get; set; }
        public string LastName { get; set; }
        public string DOB { get; set; }
        public string Gender { get; set; }
        public string address1 { get; set; }
        public string zip { get; set; }
        public string Telephone { get; set; }
        public string Fax { get; set; }
        public string Email { get; set; }
        public string StartDate { get; set; }
        public string EffectiveDate { get; set; }
        public string TerminateDate { get; set; }
        public string GroupName { get; set; }
        public string AccountTypeName { get; set; }
        public string Textbox12 { get; set; }
        public string ccv { get; set; }
        public string Textbox4 { get; set; }
        public string ck_accntnumber { get; set; }
        public string ck_rtenumber { get; set; }
        public string acct_firstname { get; set; }
        public string acct_address1 { get; set; }
        public string acct_zip { get; set; }
        public string MainPackageName { get; set; }
        public string MembershipCost { get; set; }
        public string NextChargeDate { get; set; }


        }


    public class Enrollment123
        {

        public string Agent_ID { get; set; }
        public string Agent_Label { get; set; }
        public string Member_ID { get; set; }
        public string First_Name { get; set; }
        public string Last_Name { get; set; }
        public string Address { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zipcode { get; set; }
        public string Country { get; set; }
        public string County { get; set; }
        public string Other_Address { get; set; }
        public string Other_Address2 { get; set; }
        public string Other_City { get; set; }
        public string Other_State { get; set; }
        public string Other_Zipcode { get; set; }
        public string Phone_1 { get; set; }
        public string Phone_1_Ext { get; set; }
        public string Phone_2 { get; set; }
        public string Phone_2_Ext { get; set; }
        public string Phone_3 { get; set; }
        public string Phone_3_Ext { get; set; }
        public string Fax { get; set; }
        public string Fax_Ext { get; set; }
        public string Email { get; set; }
        public string Email2 { get; set; }
        public string DOB { get; set; }
        public string Gender { get; set; }
        public string SSN { get; set; }
        public string Drivers_License { get; set; }
        public string Company_Name { get; set; }
        public string Position { get; set; }
        public string Department { get; set; }
        public string Division { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string Internal_ID { get; set; }
        public string Referral { get; set; }
        public string Source { get; set; }
        public string Email_Opt_Out { get; set; }
        public string Language { get; set; }
        public string Type { get; set; }
        public string Type_2 { get; set; }
        public string Member_Created_Date { get; set; }
        public string Product_Created_Date { get; set; }
        public string Active_Date { get; set; }
        public string First_Billing_Date { get; set; }
        public string Next_Billing_Date { get; set; }
        public string Fulfillment_Date { get; set; }
        public string Inactive_Date { get; set; }
        public string Inactive_Reason { get; set; }
        public string Hold_Date { get; set; }
        public string Hold_Reason { get; set; }
        public string Hold_Return { get; set; }
        public string Hold_Amount { get; set; }
        public string Product_ID { get; set; }
        public string Product_Label { get; set; }
        public string Product_Benefit { get; set; }
        public string Period_Label { get; set; }
        public string Category_3 { get; set; }
        public string Category_4 { get; set; }
        public string Code { get; set; }
        public string Permanent_Bill_Day { get; set; }
        public string Product_Source { get; set; }
        public string Product_Source_Detail { get; set; }
        public string Product_Status { get; set; }
        public string Product_Stage { get; set; }
        public string Product_Number_Calls { get; set; }
        public string Product_Call_Status { get; set; }
        public string Product_Next_Step { get; set; }
        public string Product_Next_Step_Date { get; set; }
        public string Product_Estimated_Close_Date { get; set; }
        public string Billing_Fee { get; set; }
        public string DiabeticShipping_Fee { get; set; }
        public string Enrollment_Fee { get; set; }
        public string LifeBuyUp10k_Fee { get; set; }
        public string LifeBuyUp15k_Fee { get; set; }
        public string LifeBuyup25k_Fee { get; set; }
        public string NCEEnrollmentFee_Fee { get; set; }
        public string Processing_Fee { get; set; }
        public string Product_Fee { get; set; }
        public string Registration_Fee { get; set; }
        public string ShippingFee_Fee { get; set; }
        public string Contract_Number { get; set; }
        public string Contract_Length { get; set; }
        public string Enroller_ID { get; set; }
        public string Paytype { get; set; }
        public string Refund_Requested { get; set; }
        public string Refund_Requested_Date { get; set; }
        public string Refund_Provided { get; set; }
        public string Refund_Provided_Date { get; set; }
        public string Refund_Comment { get; set; }
        public string Prospect_Lead { get; set; }
        public string Product_Agent_ID { get; set; }
        public string Product_Agent_Label { get; set; }
        public string Paid { get; set; }
        public string Quantity { get; set; }
        public string Renewal_Date { get; set; }
        public string Sale_Date { get; set; }
        public string Next_Billing_Amount { get; set; }
        public string Assigned_Agents { get; set; }
        public string Last_Payment { get; set; }
        public string Source_Detail { get; set; }
        public string TPV_Code { get; set; }
        public string TPV_Date { get; set; }
        public string Do_Not_Call { get; set; }
        public string Email_3 { get; set; }
        public string Ethnicity { get; set; }
        public string Height { get; set; }
        public string Weight { get; set; }
        public string Disability { get; set; }
        public string Status { get; set; }
        public string Stage { get; set; }
        public string no_Calls { get; set; }
        public string Call_Status { get; set; }
        public string Best_Call_Time { get; set; }
        public string Next_Step { get; set; }
        public string Next_Step_Date { get; set; }
        public string Source_2 { get; set; }
        public string Probability { get; set; }
        public string Product_Code { get; set; }
        public string Agent_Code { get; set; }
        public string Agent_Code_2 { get; set; }
        public string Note_Date { get; set; }
        public string System_ID { get; set; }
        public string Underwriter { get; set; }
        public string Category { get; set; }

        }

    public class WPA
        {

        public string NO { get; set; }
        public string Created { get; set; }
        public string Agent { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Status { get; set; }
        public string Rank { get; set; }
        public string Flag { get; set; }
        public string Call_Status { get; set; }
        public string Next_Step { get; set; }
        public string Last_Trans_Amount { get; set; }
        public string Last_Trans_Method { get; set; }
        public string Last_Trans_Status { get; set; }
        public string Last_Trans_Date { get; set; }
        public string View { get; set; }
        public string Batch { get; set; }
        }

    public struct _Link
        {
        public string pagination;
        public string _href;
        public string method;//updateHistoryDisplay
        public string type;//index
        public string value;
        public string token;
        public string format;// json
            }




    }
