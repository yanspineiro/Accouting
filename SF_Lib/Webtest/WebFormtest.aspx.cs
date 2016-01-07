using SF_Lib;
using SF_Lib.com.salesforce.na5;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Webtest
    {
    public partial class WebFormtest : System.Web.UI.Page
        {
        protected void Page_Load(object sender, EventArgs e)
            {


            SF_Lib.SalesForce_Lib temp= new SalesForce_Lib();
            temp.showtext();
           /* CSV_Controller csvcontr;

            Salerforce_Interface sf_interface = new Salerforce_Interface(); ;
            List<Payment_Period__c> calendar_SF;
            calendar_SF = sf_interface.GetPaymentCalendarbyCompanny("HBC");

            List<Import_HII_CSV> HII_CSV_list;
            List<Import_HII_CSV> Verify_HII_List;
            List<Import_HII_CSV> MisMatch_HII_List;

            string adate = Convert.ToDateTime("02/21/2015").ToString("yyyy-MM-dd");

            List<Payment__c> payment_list = sf_interface.getpaymentbyPayday(adate);
            HII_CSV_list = new List<Import_HII_CSV>();
            Verify_HII_List = new List<Import_HII_CSV>();
            MisMatch_HII_List = new List<Import_HII_CSV>();
            Payment_Period__c calendar;
            
            for (int i = 0; i < payment_list.Count; i++)
                {
                Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();
                _Import_HII_CSV.Sf_MemberID = payment_list[i].Policy_Number__c;

                _Import_HII_CSV.Agent_id = payment_list[i].Agent__c;
                _Import_HII_CSV.Agent_Fullname = payment_list[i].Agent__r.First_Name__c + ' ' + payment_list[i].Agent__r.Last_Name__c;


                _Import_HII_CSV.Agent_Status = payment_list[i].Agent__r.Status__c;
                _Import_HII_CSV.Coinsurance = payment_list[i].Policy_Number__c;
                _Import_HII_CSV.Opportunity_id = payment_list[i].OpportunityLineItem_id__c;
                _Import_HII_CSV.Opportunity_id = payment_list[i].OpportunityLineItem_id__c;
                _Import_HII_CSV.Payroll_date = Convert.ToDateTime(payment_list[i].Payment_Date__c);
                _Import_HII_CSV.Payroll_Type = payment_list[i].Payment_Type__c;
                _Import_HII_CSV.Agent_Commision = payment_list[i].Payment_Value__c.ToString();
                _Import_HII_CSV.Verify = false;
                _Import_HII_CSV.IsUpdate = true;

                try
                    {
                    _Import_HII_CSV.Ending_Date = Convert.ToDateTime(payment_list[i].Agent__r.Ending_Date__c);
                    }
                catch (Exception ex)
                    {
                    _Import_HII_CSV.Ending_Date = DateTime.Today;
                    }

                Verify_HII_List.Add(_Import_HII_CSV);
                }

            try
                {


               calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(adate)).FirstOrDefault();

                csvcontr = new CSV_Controller();
                string userName = "mspiewak2014";// System.Configuration.ConfigurationSettings.AppSettings["username"];
                string password = "Matt1984";// System.Configuration.ConfigurationSettings.AppSettings["pwd_agent"];
                byte[] bytes = csvcontr.Getwebfile(Convert.ToDateTime(calendar.Start_Date__c).AddDays(-31), Convert.ToDateTime(calendar.Close_Date__c), userName, password);
                csvcontr.readFile(bytes);
                HII_CSV_list = csvcontr.HII_CSV_list;
                csvcontr.Calendar_Selected = calendar;

                }
            catch (Exception ex)
                {
                throw (ex);
                }

            try
                {

                csvcontr.matchSalerforce(Verify_HII_List, "HBC");
                MisMatch_HII_List = csvcontr.MisMatch_HII_List;
                }
            catch (Exception ex)
                {
                throw (ex);
                }



            /* payment_list = sf_interface.getpaymentbyPayday(adate);

             csvcontr = new CSV_Controller(); ;

            csvcontr.HII_CSV_list = HII_CSV_list;

           

                csvcontr.matchSalerforce(Verify_HII_List, "HBC");
                MisMatch_HII_List = csvcontr.MisMatch_HII_List;*/


            /*payment_list = sf_interface.getpaymentbyPayday(adate);
            calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(adate)).FirstOrDefault();

            csvcontr = new CSV_Controller();
            csvcontr.Calendar_Selected = calendar;
            csvcontr.HII_CSV_list = MisMatch_HII_List;
            var PolicyNumberList = (from z in Verify_HII_List
                                    select z.Sf_MemberID).Distinct();

            csvcontr.FillGrid(PolicyNumberList.ToList());
            MisMatch_HII_List = null;
            MisMatch_HII_List = csvcontr.HII_CSV_list;
            MisMatch_HII_List = csvcontr.HII_CSV_list.Where(c => c.Payroll_Type == "Termination" || c.Payroll_Type == "Chargeback" || c.Payroll_Type == "Commission").ToList();
            MisMatch_HII_List = MisMatch_HII_List.Where(c => c.Agent_id != null && c.Payroll_date >= Convert.ToDateTime(adate) && c.Payroll_date != DateTime.Today).ToList();
            */

              
            }
        }
    }