using SF_Lib.com.salesforce.na5;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq.Dynamic;
using HtmlAgilityPack;
using System.Net;

namespace SF_Lib
    {
    public class CSV_Controller
        {
        public List<Import_HII_CSV> HII_CSV_list { get; set; }
        public List<Import_HII_CSV> Verify_HII_List { get; set; }
        public List<Import_HII_CSV> MisMatch_HII_List { get; set; }

        public string Product_Profile { get; set; }
        private List<White_List_Resource__c> PayHoldPeriod { set; get; }
        List<string> cancellationStatus;

        public string linename { get; set; }
        public string message { get; set; }
        public Salerforce_Interface sf_interface { set; get; }
        public Payment_Period__c Calendar_Selected { set; get; }


        public CSV_Controller()
            {
            cancellationStatus = new List<string>();

            cancellationStatus.Add("Refunded");
            cancellationStatus.Add("Canceled");
            cancellationStatus.Add("ChargeBack");
            cancellationStatus.Add("Voided");
            cancellationStatus.Add("Terminated");
            cancellationStatus.Add("Chargeback");

            message = string.Empty;
            sf_interface = new Salerforce_Interface();
            Product_Profile = "HBC";
            }

        public void readFile(byte[] buffer)
            {

            MemoryStream ms = new MemoryStream(buffer);

            HII_CSV_list = new List<Import_HII_CSV>();
            StreamReader reader = new StreamReader(ms);
            List<string> index_val = new List<string>();
            while (!reader.EndOfStream)
                {
                var line = reader.ReadLine();
                var csvRecord = line.Replace("\",\"", ";");
                var csvRecordData = csvRecord.Split(';');
                Import_HII_CSV CSV_front_Obj = new Import_HII_CSV();


                if (line.Contains("Applicant ID"))
                    {
                    line = line.ToUpper();
                    csvRecord = line.Replace("\",\"", ";");
                    csvRecordData = csvRecord.Split(';');
                    index_val = new List<string>();
                    index_val = index_val.Concat(csvRecordData).ToList();
                    }

                int _index = 0;
                if (csvRecordData.Count() > 0)
                    {
                    try
                        {
                        CSV_front_Obj = new Import_HII_CSV();
                        _index = 0;// index_val.IndexOf("APPLICANT ID");
                        CSV_front_Obj.Applicant_ID = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("FIRST NAME");
                        CSV_front_Obj.First_Name = csvRecordData[_index].Replace("\"", "") + ' ' + csvRecordData[_index + 1].Replace("\"", "");
                        _index = index_val.IndexOf("LAST NAME");
                        CSV_front_Obj.Last_Name = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("GENDER");
                        CSV_front_Obj.Gender = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("DOB");
                        CSV_front_Obj.DOB = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("ADDRESS");
                        CSV_front_Obj.Address = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CITY");
                        CSV_front_Obj.City = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("STATE");
                        CSV_front_Obj.State = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("ZIPCODE");
                        CSV_front_Obj.ZipCode = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("PHONE");
                        CSV_front_Obj.Phone = CSV_Controller.FormatPhone(csvRecordData[_index].Replace("\"", ""));
                        _index = index_val.IndexOf("EMAIL");
                        CSV_front_Obj.Email = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("MEMBER ID");
                        CSV_front_Obj.Member_ID = csvRecordData[_index].Replace("\"", "");

                        if (CSV_front_Obj.Member_ID == "ACC08490200")
                            {
                            var c = "ACC08490200";
                            }

                        _index = index_val.IndexOf("PRODUCT NAME");
                        CSV_front_Obj.Product_Name = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("OPTION");
                        CSV_front_Obj.Option = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("TYPE");
                        CSV_front_Obj.Type = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("POLICY NUMBER");
                        CSV_front_Obj.Policy_Number = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("PREMIUM");
                        CSV_front_Obj.Premium = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("ENROLLMENT FEE");
                        CSV_front_Obj.Enrollment_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("ADMINISTRATION FEE");
                        CSV_front_Obj.Administration_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("MYEWELLNESS FEE");
                        CSV_front_Obj.Myewellness_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("TELADOC FEE");
                        CSV_front_Obj.TelaDoc_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("EXTRA CARE PACKAGE FEE");
                        CSV_front_Obj.Extra_Care_Package_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("RXCARD FEE");
                        CSV_front_Obj.RxCard_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("UTRA CARE PLUS");
                        CSV_front_Obj.Utra_Care_Plus = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("PROVIDER FEE");
                        CSV_front_Obj.Provider_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("RXADVOCACY FEE");
                        CSV_front_Obj.RxAdvocacy_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CRITICAL ILLNESS FEE");
                        CSV_front_Obj.Critical_Illness_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("MULTICARE FEE");
                        CSV_front_Obj.Multicare_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("MED SENSE FEE");
                        CSV_front_Obj.Med_Sense_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("SAVERS PACKAGE");
                        CSV_front_Obj.Savers_Package = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("DENTAL BENEFIT FEE");
                        CSV_front_Obj.Dental_Benefit_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("FREEDOM ACC.EXP. FEE");
                        CSV_front_Obj.Freedom_Acc_Exp__Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("AD&D AME/STARR AD&D FEE");
                        CSV_front_Obj.AD_D_AME_Starr_AD_D_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("STARR AD FEE");
                        CSV_front_Obj.Starr_AD_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("ING ADD FEE");
                        CSV_front_Obj.ING_ADD_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CHIRO & PODIATRY");
                        CSV_front_Obj.Chiro___Podiatry = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CARE24X7");
                        CSV_front_Obj.Care24x7 = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("EYEMED");
                        CSV_front_Obj.EyeMed = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CARRIER ASSOCIATION FEE");
                        CSV_front_Obj.HCC_Association_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("KARE 360 FEE");
                        CSV_front_Obj.Kare_360_Fee = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("TOTAL COLLECTED");
                        CSV_front_Obj.Total_Collected = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("COINSURANCE");
                        CSV_front_Obj.Coinsurance = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("COINSURANCE PERCENTAGE");
                        CSV_front_Obj.Coinsurance_Percentage = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("DEDUCTIBLE");
                        CSV_front_Obj.Deductible = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("DURATION COVERAGE");
                        CSV_front_Obj.Duration_Coverage = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("APPLICATION DATE");

                        try
                            {

                            CSV_front_Obj.Application_Date = Convert.ToDateTime(csvRecordData[_index].Replace("\"", "")).ToString("MM/dd/yyyy");
                            }

                        catch (Exception ex)
                            {

                            CSV_front_Obj.Application_Date = csvRecordData[_index].Replace("\"", "");
                            }
                        _index = index_val.IndexOf("EFFECTIVE DATE");
                        CSV_front_Obj.Effective_Date = csvRecordData[_index].Replace("\"", "");


                        try
                            {
                            _index = index_val.IndexOf("TERMINATION DATE");
                            CSV_front_Obj.Termination_Date = Convert.ToDateTime(csvRecordData[_index].Replace("\"", "")).ToString("MM/dd/yyyy");

                            if (CSV_front_Obj.Termination_Date != null)
                                {
                                _index = index_val.IndexOf("PAYMENT STATUS DATE");
                                CSV_front_Obj.Termination_Date = Convert.ToDateTime(csvRecordData[_index].Replace("\"", "")).ToString("MM/dd/yyyy");
                                CSV_front_Obj.Payment_Status = CSV_front_Obj.Payroll_Type = "Chargeback";
                                }

                            //CSV_front_Obj.Payment_Status = CSV_front_Obj.Payroll_Type = "Chargeback";
                            }

                        catch (Exception ex)
                            {
                            _index = index_val.IndexOf("TERMINATION DATE");
                            CSV_front_Obj.Termination_Date = csvRecordData[_index].Replace("\"", "");
                            }

                        _index = index_val.IndexOf("PAYMENT STATUS");
                        CSV_front_Obj.Payment_Status = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("PAYMENT STATUS DATE");
                        CSV_front_Obj.Payment_Status_Date = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("PAYMENT METHOD");

                        CSV_front_Obj.Payment_Method = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("AGENT FIRST NAME");
                        CSV_front_Obj.Agent_First_Name = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("AGENT LAST NAME");
                        CSV_front_Obj.Agent_Last_Name = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("AGENT CODE");
                        CSV_front_Obj.Agent_Code = csvRecordData[_index].Replace("\"", "");
                        _index = index_val.IndexOf("CANCELLATION CODE");
                        CSV_front_Obj.Cancellation_Code = csvRecordData[_index].Replace("\"", "");
                        _index = 60;// index_val.IndexOf("AGENCY/COMPANY NAME");
                        CSV_front_Obj.Agency_Company_Name = csvRecordData[_index].Replace("\"", "");

                        if (!CSV_front_Obj.Applicant_ID.Contains("Applicant ID"))
                            {
                            HII_CSV_list.Add(CSV_front_Obj);
                            }




                        }
                    catch (Exception ex)
                        {

                        }

                    }

                }

            }

        public byte[] Getwebfile(DateTime fromDate, DateTime toDate, string username, string password)
            {
            List<byte> responseArray = new List<byte>();
            if ((toDate - fromDate).TotalDays > 31)
                {
                int cant = (int)Math.Ceiling((toDate - fromDate).TotalDays / (double)2);

                responseArray = Getwebfile(fromDate, fromDate.AddDays(cant), username, password).ToList();
                responseArray = responseArray.Concat(Getwebfile(fromDate.AddDays(cant + 1), toDate, username, password).ToList()).ToList();

                }
            else
                {
                try
                    {

                    var client = new CookieAwareWebClient();
                    client.BaseAddress = @"https://www.hiiquote.com/data_admin/agent_login.php";
                    client.isfile = false;
                    var loginData = new NameValueCollection();
                    loginData.Add("username", username);//"mspiewak2014");
                    loginData.Add("pwd_agent", password);// "Matt1984");
                    loginData.Add("sFlag", "s");
                    client.UploadValues("agent_login_process.php", "POST", loginData);

                    string htmlSource = client.DownloadString("export_data.php");
                    if (htmlSource.Contains("<form name=\"frmExport\" method=\"post\" action=\"\">"))
                        {

                        string from = fromDate.ToString("MM/dd/yyyy").Replace('/', '-');
                        string to = toDate.ToString("MM/dd/yyyy").Replace('/', '-');

                        client.BaseAddress = @"https://www.hiiquote.com/data_admin/export_data.php";
                        client.isfile = true;
                        var searchvalues = new NameValueCollection();
                        searchvalues.Add("SearchPoduct", "");
                        searchvalues.Add("FromDate", from);
                        searchvalues.Add("ToDate", to);
                        searchvalues.Add("sFlag", "s");
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705;)");

                        responseArray = client.UploadValues("export_data.php", "POST", searchvalues).ToList();
                        //  string s = Encoding.ASCII.GetString(responseArray);


                        }
                    client.UploadValues("logout.php", "POST", loginData);
                    htmlSource = client.DownloadString("export_data.php");

                    }
                catch (Exception ex)
                    {
                    message = ex.Message;
                    }

                }




            return responseArray.ToArray();
            }


        public void FillGrid(List<string> no_PolicyNumber, string agent_id, string companny, DateTime _from, DateTime _to)
            {


            var PolicyNumberList = (from z in HII_CSV_list
                                    select z.Member_ID).Distinct();



            List<Opportunity> opportunity_list = new List<Opportunity>();
            List<OpportunityLineItem> opportunityLineBy = sf_interface.GetOpportunityLineBy_Policy_Number(PolicyNumberList.ToList(), agent_id, companny, _to);
            List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod = sf_interface.getChargeBackCancellationPeriod();
            List<Payment_Period__c> Calendar_List = sf_interface.GetPaymentCalendarbyCompanny(companny);
            List<White_List_Resource__c> PayHoldPeriod = sf_interface.getPayHoldPeriod();
            // List<Payment__c> Cancel_payroll_previus = sf_interface.getpaymentbyPolicyNumber(PolicyNumberList.ToList(), new List<string> { "Terminated", "Chargeback" });
            List<Payment__c> Commision_payroll_previus = sf_interface.getpaymentbyPolicyNumber(PolicyNumberList.ToList(), new List<string> { "Commission" });

            //OpportunityLineItem temp = opportunityLineBy.Where(c => c.Policy_Number__c == "CL3476600").FirstOrDefault();

            opportunityLineBy = opportunityLineBy.Where(c => Convert.ToDateTime(c.Enrollment_Date__c) > _from).ToList();
            Commision_payroll_previus = (from X in Commision_payroll_previus
                                         where !no_PolicyNumber.Contains(X.Policy_Number__c)
                                         select X).ToList();

            List<String> GI_Product = new List<string>();
            if (opportunityLineBy.Count > 0)
                {
                var opportunity_id = (from x in opportunityLineBy
                                      where x.OpportunityId != null && x.OpportunityId != ""
                                      select x.OpportunityId).Distinct();


                if (opportunity_id.ToList().Count() > 0)
                    {
                    //  opportunity_list = sf_interface.GetOpportunity_by_id(opportunity_id.ToList(), Product_Profile);
                    }



                for (int i = 0; i < HII_CSV_list.Count; i++)
                    {
                    var oppLine = (from x in opportunityLineBy
                                   where x.Policy_Number__c != null && x.Policy_Number__c.Equals(HII_CSV_list[i].Member_ID)
                                   select x).ToList().FirstOrDefault();

                    /* if (HII_CSV_list[i].Member_ID == "CL3476600")
                         {
                         var jj = "kk";
                         }*/
                    if (oppLine != null)
                        {
                        HII_CSV_list[i].OpportunityLine_id = ((OpportunityLineItem)oppLine).Id;
                        HII_CSV_list[i].Product_Name = ((OpportunityLineItem)oppLine).Product2.Name;
                        HII_CSV_list[i].Product_PlanType = ((OpportunityLineItem)oppLine).Product2.Plan_Type__c;
                        HII_CSV_list[i].CurrentRuning_total = ((OpportunityLineItem)oppLine).Product2.Current_Running_Totals_Category__c;
                        HII_CSV_list[i].MisMachEnrrol = false;
                        HII_CSV_list[i].Sf_MemberID = HII_CSV_list[i].Sf_MemberID_Bkup = ((OpportunityLineItem)oppLine).Policy_Number__c;
                        HII_CSV_list[i].Oppline_Status__c = ((OpportunityLineItem)oppLine).Status__c;
                        HII_CSV_list[i].Application_Date = Convert.ToDateTime(((OpportunityLineItem)oppLine).Enrollment_Date__c).Date.ToShortDateString();
                        HII_CSV_list[i].Effective_Date = Convert.ToDateTime(((OpportunityLineItem)oppLine).Effective_Date__c).Date.ToShortDateString();

                        /*var opp = (from x in opportunity_list
                                   where x.Id.Equals(((OpportunityLineItem)oppLine).OpportunityId) && x.GI_Product__r.Id.Equals(((OpportunityLineItem)oppLine).Product2.Id)
                                   select x).ToList().FirstOrDefault();*/



                        if (((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__c != null)//opp != null
                            {



                            HII_CSV_list[i].Agent_Fullname = ((OpportunityLineItem)oppLine).Opportunity.Fronter_s_Name__c == null ? "" : ((OpportunityLineItem)oppLine).Opportunity.Fronter_s_Name__c;// ((Opportunity)opp).Fronter_s_Name__c;
                            HII_CSV_list[i].Agent_id = ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__c == null ? "" : ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__c;// ((Opportunity)opp).Agent_ID_Lookup__c;
                            HII_CSV_list[i].Opportunity_name = ((OpportunityLineItem)oppLine).Opportunity.Name == null ? "" : ((OpportunityLineItem)oppLine).Opportunity.Name;  //((Opportunity)opp).Name;
                            HII_CSV_list[i].Opportunity_id = ((OpportunityLineItem)oppLine).OpportunityId == null ? "" : ((OpportunityLineItem)oppLine).OpportunityId;// ((Opportunity)opp).Id;
                            HII_CSV_list[i].Agent_Status = ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__c == null || ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.Status__c == null ? "" : ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.Status__c;//((Opportunity)opp).Agent_ID_Lookup__r.Status__c;
                            HII_CSV_list[i].Agent_Commision = ((OpportunityLineItem)oppLine).Commission_Payable_Number__c == null ? "0.00" : (((OpportunityLineItem)oppLine).Commission_Payable_Number__c).ToString();
                            HII_CSV_list[i].Verify = true;
                            HII_CSV_list[i].Agent__ProductProfile = ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__c == null || ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.product_profile__c == null ? "" : ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.product_profile__c;  //   ((Opportunity)opp).Agent_ID_Lookup__r.product_profile__c; ;




                            if (cancellationStatus.Contains(HII_CSV_list[i].Payment_Status))
                                {
                                HII_CSV_list[i].Agent_Commision = (Convert.ToDouble(HII_CSV_list[i].Agent_Commision) * -1).ToString("N");
                                string HII_enrollmentDate = DateTime.Today.Date.ToShortDateString();
                                try
                                    {
                                    HII_enrollmentDate = Convert.ToDateTime(HII_CSV_list[i].Application_Date).ToShortDateString();
                                    }
                                catch (Exception exx)
                                    {
                                    HII_enrollmentDate = DateTime.Today.Date.ToShortDateString();
                                    }

                                string HBCEnrollmentDate = Convert.ToDateTime(((OpportunityLineItem)oppLine).Enrollment_Date__c).Date.ToShortDateString();

                                if (HBCEnrollmentDate != HII_enrollmentDate)
                                    {
                                    HII_CSV_list[i].MisMachEnrrol = true;

                                    }
                                else
                                    {
                                    HII_CSV_list[i].Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                                    if (((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.Status__c == "Not Available")
                                        {
                                        DateTime payHolddate = Convert.ToDateTime(((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.Ending_Date__c);
                                        payHolddate = payHolddate.AddDays(searchPayHold(HII_CSV_list[i].CurrentRuning_total));

                                        Payment_Period__c new_Pariod = Calendar_List.Where(c => c.Start_Date__c <= payHolddate && c.Close_Date__c >= payHolddate && c.Product_Profile__c == ((OpportunityLineItem)oppLine).Opportunity.Agent_ID_Lookup__r.product_profile__c).FirstOrDefault();
                                        //  Opportunity.Agent_ID_Lookup__r.product_profile__c
                                        HII_CSV_list[i].Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                                        HII_CSV_list[i].Payroll_Type = HII_CSV_list[i].Payment_Status;


                                        /////**** cancellatio y chageback?

                                        DateTime Termination_date = DateTime.Today;
                                        DateTime EnrollmentDate = DateTime.Today;
                                        DateTime Efective_date = DateTime.Today;
                                        try
                                            {
                                            if (HII_CSV_list[i].Termination_Date == "" || HII_CSV_list[i].Termination_Date == null)
                                                {
                                                HII_CSV_list[i].Termination_Date = HII_CSV_list[i].Application_Date;
                                                }
                                            Termination_date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date);
                                            EnrollmentDate = Convert.ToDateTime(HII_CSV_list[i].Application_Date);
                                            Efective_date = Convert.ToDateTime(HII_CSV_list[i].Effective_Date);
                                            }
                                        catch (Exception exx)
                                            {
                                            HII_CSV_list[i].Verify = false;
                                            }



                                        //Clasify_ChargeBack_Terminated(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, HII_CSV_list[i]);
                                        Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, Efective_date, HII_CSV_list[i]);

                                        }
                                    else
                                        {
                                        DateTime Termination_date = DateTime.Today;
                                        DateTime EnrollmentDate = DateTime.Today;
                                        DateTime Efective_date = DateTime.Today;
                                        try
                                            {
                                            if (HII_CSV_list[i].Termination_Date == "" || HII_CSV_list[i].Termination_Date == null)
                                                {
                                                HII_CSV_list[i].Termination_Date = HII_CSV_list[i].Application_Date;
                                                }
                                            Termination_date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date);
                                            EnrollmentDate = Convert.ToDateTime(HII_CSV_list[i].Application_Date);
                                            Efective_date = Convert.ToDateTime(HII_CSV_list[i].Effective_Date);
                                            }
                                        catch (Exception exx)
                                            {
                                            HII_CSV_list[i].Verify = false;
                                            }

                                        //  Clasify_ChargeBack_Terminated(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, HII_CSV_list[i]);
                                        Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, Efective_date, HII_CSV_list[i]);

                                        }

                                    }





                                }
                            else
                                {
                                // commission
                                if (Commision_payroll_previus.Where(c => c.Policy_Number__c.Equals(HII_CSV_list[i].Member_ID) && (c.Payment_Type__c.Equals("Commission"))).ToList().Count == 0)
                                    {

                                    DateTime Termination_date = DateTime.Today;
                                    DateTime EnrollmentDate = DateTime.Today;
                                    try
                                        {
                                        // Termination_date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date);
                                        EnrollmentDate = Convert.ToDateTime(HII_CSV_list[i].Application_Date);
                                        }
                                    catch (Exception exx)
                                        {
                                        HII_CSV_list[i].Verify = false;
                                        }




                                    Payment_Period__c compare_calendar = Calendar_List.Where(c => c.Start_Date__c <= EnrollmentDate && EnrollmentDate <= c.Close_Date__c).FirstOrDefault();
                                    if (compare_calendar == null)
                                        {
                                        var rr = "";
                                        }

                                    HII_CSV_list[i].Payroll_Type = "Commission";
                                    HII_CSV_list[i].Payroll_date = Convert.ToDateTime(compare_calendar.Payment_Date__c);
                                    }

                                }


                            }
                        }
                    }

                }

            HII_CSV_list = HII_CSV_list.OrderBy(c => c.Application_Date).ToList(); ;

            //  HII_CSV_list= HII_CSV_list.Where(c => c.Payroll_Type == "Termination" || c.Payroll_Type == "Chargeback" || c.Payroll_Type == "Commission").ToList();
            //FilterList();


            }

        /*   public void Clasify_ChargeBack_Terminated(List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod, List<Payment_Period__c> Calendar_List, DateTime Termination_date, DateTime EnrollmentDate, Import_HII_CSV _HII_CSV)
               {

               if (_HII_CSV.Member_ID == "CL3971300")
                   {
                   var gg = "";
                   }
               Payment_Period__c compare_calendar = Calendar_List.Where(c => (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Close_Date__c)) && (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Payment_Date__c))).FirstOrDefault();

               if ((Convert.ToDateTime(Calendar_Selected.Start_Date__c) <= Termination_date && Termination_date < Convert.ToDateTime(Calendar_Selected.Payment_Date__c)) && Convert.ToDateTime(Calendar_Selected.Start_Date__c) <= EnrollmentDate && EnrollmentDate < Convert.ToDateTime(Calendar_Selected.Payment_Date__c))
                   {
                   _HII_CSV.Payroll_Type = "Terminated";
                   _HII_CSV.Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);

                   }
               else
                   {
                   White_List_Resource__c temp_whiteElement;

                   if (_HII_CSV.CurrentRuning_total != null)
                       {
                       temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Contains(_HII_CSV.CurrentRuning_total)).FirstOrDefault();
                       }
                   else
                       {
                       temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Equals("Chargeback Period")).FirstOrDefault();
                       }


                   if (temp_whiteElement == null)
                       {
                       temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Equals("Chargeback Period")).FirstOrDefault();

                       }

                   int period_value = 0;
                   if (temp_whiteElement != null)
                       {

                       period_value = Convert.ToInt32(temp_whiteElement.Value__c);
                       }



                   Payment_Period__c _new_Pariod = Calendar_List.Where(c => c.Pay_Period__c == Calendar_Selected.Pay_Period__c - period_value).FirstOrDefault();
                   if (_new_Pariod != null)
                       {
                       if (Convert.ToDateTime(_new_Pariod.Start_Date__c) <= Termination_date && Termination_date <= Convert.ToDateTime(Calendar_Selected.Payment_Date__c))
                           {
                           _HII_CSV.Payroll_Type = "Chargeback";
                           }
                       else
                           {

                           _new_Pariod = Calendar_List.Where(c => c.Pay_Period__c == Calendar_Selected.Pay_Period__c + period_value).FirstOrDefault();
                           if (_new_Pariod != null)
                               {
                               if (Convert.ToDateTime(Calendar_Selected.Start_Date__c) <= Termination_date && Termination_date <= Convert.ToDateTime(_new_Pariod.Payment_Date__c))
                                   {
                                   _HII_CSV.Payroll_Type = "Terminated";


                                   if (Calendar_Selected.Pay_Period__c - compare_calendar.Pay_Period__c > period_value)
                                       {
                                       _HII_CSV.Payroll_date = Convert.ToDateTime(compare_calendar.Payment_Date__c);
                                       }
                                   if (Calendar_Selected.Pay_Period__c != compare_calendar.Pay_Period__c)
                                       {
                                       _HII_CSV.Payroll_Type = "Chargeback";
                                       }

                                   }
                               else
                                   {
                                   _HII_CSV.Agent_Commision = "0.00";
                                   _HII_CSV.Payroll_Type = "Chargeback";
                                   }
                               }
                           else
                               {

                               compare_calendar = Calendar_List.Where(c => (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Close_Date__c)) && (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Payment_Date__c))).FirstOrDefault();
                               if (compare_calendar != null)
                                   {
                                   _HII_CSV.Payroll_Type = "Chargeback";
                                   _HII_CSV.Payroll_date = Convert.ToDateTime(compare_calendar.Payment_Date__c);
                                   }
                               else
                                   {
                                   _HII_CSV.Payroll_Type = "not Found";
                                   }

                               }


                           }
                       }
                   else
                       {
                       _HII_CSV.Payroll_Type = "not Found";
                       }
                   }
               }
         */

        public void Clasify_ChargeBack_Terminated_new(List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod, List<Payment_Period__c> Calendar_List, DateTime Termination_date, DateTime EnrollmentDate, DateTime effective_date, Import_HII_CSV _HII_CSV)
            {

            if (_HII_CSV.Sf_MemberID == "WS2594386")
                {
                var name = "";
                }

            Payment_Period__c compare_calendar = Calendar_List.Where(c => (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Close_Date__c)) && (Convert.ToDateTime(c.Start_Date__c) <= EnrollmentDate && EnrollmentDate <= Convert.ToDateTime(c.Payment_Date__c))).FirstOrDefault();



            if (Convert.ToDateTime(Calendar_Selected.Start_Date__c) <= EnrollmentDate && EnrollmentDate < Convert.ToDateTime(Calendar_Selected.Payment_Date__c))
                {
                if ((Convert.ToDateTime(Calendar_Selected.Start_Date__c) <= Termination_date && Termination_date < Convert.ToDateTime(Calendar_Selected.Payment_Date__c)))
                    {
                    _HII_CSV.Payroll_Type = "Terminated";
                    _HII_CSV.Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                    }
                else
                    {

                    int period_value = DayforPayment_chargeback(w_list_ChargeBackCancellationPeriod, _HII_CSV);

                    DateTime end_chargeback = effective_date.AddMonths(period_value);

                    if (effective_date <= Termination_date && Termination_date <= end_chargeback)
                        {
                        _HII_CSV.Payroll_Type = "Terminated";
                        _HII_CSV.Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                        }
                    else
                        {
                        _HII_CSV.Payroll_Type = "";
                        }
                    }
                }
            else
                {
                int period_value = DayforPayment_chargeback(w_list_ChargeBackCancellationPeriod, _HII_CSV);


                DateTime end_chargeback = effective_date.AddMonths(period_value);

                if (effective_date <= Termination_date && Termination_date <= end_chargeback)
                    {
                    _HII_CSV.Payroll_Type = "Chargeback";
                    _HII_CSV.Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                    }
                else
                    {
                    _HII_CSV.Payroll_Type = "";
                    }



                }
            }

        private static int DayforPayment_chargeback(List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod, Import_HII_CSV _HII_CSV)
            {
            White_List_Resource__c temp_whiteElement = new White_List_Resource__c();

            if (_HII_CSV.CurrentRuning_total != null)
                {
                temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Contains(_HII_CSV.CurrentRuning_total)).FirstOrDefault();
                }

            if (temp_whiteElement == null || temp_whiteElement.Value__c == null || temp_whiteElement.Value__c == "")
                {
                temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Equals("Chargeback Period Default")).FirstOrDefault();
                }




            if (temp_whiteElement == null || temp_whiteElement.Value__c == null || temp_whiteElement.Value__c == "")
                {

                temp_whiteElement = w_list_ChargeBackCancellationPeriod.Where(c => c.Resource_Type__c.Equals("Chargeback Period")).FirstOrDefault();

                }

            int period_value = 0;
            if (temp_whiteElement != null)
                {

                period_value = Convert.ToInt32(temp_whiteElement.Value__c);
                }
            return period_value;
            }

        public void matchSalerforce(List<Import_HII_CSV> sf_list, string _ProductProfile)
            {

            Verify_HII_List = new List<Import_HII_CSV>();
            MisMatch_HII_List = new List<Import_HII_CSV>();
            List<Payment_Period__c> Calendar_List = sf_interface.GetPaymentCalendarbyCompanny(_ProductProfile);
            List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod = sf_interface.getChargeBackCancellationPeriod();
            List<White_List_Resource__c> PayHoldPeriod = sf_interface.getPayHoldPeriod();

            List<Import_HII_CSV> tes = sf_list.Where(c => c.Sf_MemberID == "WS2594386").ToList();

            for (int i = 0; i < HII_CSV_list.Count; i++)
                {
                // List<Import_HII_CSV> _Import_HII_CSV_List_SFq = sf_list.Where(c => c.Sf_MemberID == null).ToList();
                List<Import_HII_CSV> _Import_HII_CSV_List_SF = sf_list.Where(c => c.Sf_MemberID.Equals(HII_CSV_list[i].Member_ID)).ToList();


                if (HII_CSV_list[i].Member_ID == "WS2594386")
                    {
                    var t = "t";
                    }

                if (HII_CSV_list[i].Sf_MemberID == "WS2594386")
                    {
                    var t = "t";
                    }

                if (_Import_HII_CSV_List_SF != null && _Import_HII_CSV_List_SF.Count > 0)
                    {

                    if (cancellationStatus.Contains(HII_CSV_list[i].Payment_Status))
                        {
                        if (_Import_HII_CSV_List_SF.Count == 1)
                            {

                            _Import_HII_CSV_List_SF[0].Verify = true;
                            _Import_HII_CSV_List_SF[0].IsUpdate = true;
                            _Import_HII_CSV_List_SF[0].Sf_MemberID_Bkup = _Import_HII_CSV_List_SF[0].Member_ID = _Import_HII_CSV_List_SF[0].Sf_MemberID;
                            _Import_HII_CSV_List_SF[0].First_Name = HII_CSV_list[i].First_Name;

                            if (_Import_HII_CSV_List_SF[0].Agent_Status == "Not Available")
                                {
                                DateTime payHolddate = _Import_HII_CSV_List_SF[0].Ending_Date;
                                payHolddate = payHolddate.AddDays(searchPayHold(""));

                                Payment_Period__c new_Pariod = Calendar_List.Where(c => c.Start_Date__c <= payHolddate && c.Close_Date__c >= payHolddate && c.Product_Profile__c == _ProductProfile).FirstOrDefault();

                                _Import_HII_CSV_List_SF[0].Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);


                                }

                            Import_HII_CSV new_Import_HII_CSV = (Import_HII_CSV)_Import_HII_CSV_List_SF[0].Clone();

                            new_Import_HII_CSV.Agent_Commision = new_Import_HII_CSV.Agent_Commision == null ? "0" : (Convert.ToDouble(new_Import_HII_CSV.Agent_Commision) * -1).ToString();
                            new_Import_HII_CSV.IsUpdate = false;

                            DateTime Termination_date = DateTime.Today;
                            DateTime EnrollmentDate = DateTime.Today;
                            DateTime Efective_date = DateTime.Today;
                            try
                                {
                                if (HII_CSV_list[i].Termination_Date == "" || HII_CSV_list[i].Termination_Date == null)
                                    {
                                    HII_CSV_list[i].Termination_Date = HII_CSV_list[i].Application_Date;
                                    }
                                Termination_date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date);
                                EnrollmentDate = Convert.ToDateTime(_Import_HII_CSV_List_SF[0].Application_Date);//HII_CSV_list[i].Application_Date
                                Efective_date = Convert.ToDateTime(HII_CSV_list[i].Effective_Date);
                                }
                            catch (Exception exx)
                                {
                                HII_CSV_list[i].Verify = false;
                                }



                            // Clasify_ChargeBack_Terminated(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, HII_CSV_list[i]);

                            Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, Efective_date, HII_CSV_list[i]);


                            new_Import_HII_CSV.Payroll_Type = HII_CSV_list[i].Payroll_Type;
                            new_Import_HII_CSV.Termination_Date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date).ToString("MM/dd/yyyy");
                            if (HII_CSV_list[i].Agent_Commision == "0.00")
                                {
                                new_Import_HII_CSV.Agent_Commision = "0.00";
                                }


                            sf_list.Add(new_Import_HII_CSV);
                            }
                        else
                            {


                            _Import_HII_CSV_List_SF[1].Verify = _Import_HII_CSV_List_SF[0].Verify = true;
                            _Import_HII_CSV_List_SF[1].IsUpdate = _Import_HII_CSV_List_SF[0].IsUpdate = true;
                            _Import_HII_CSV_List_SF[1].Member_ID = _Import_HII_CSV_List_SF[0].Member_ID = _Import_HII_CSV_List_SF[0].Sf_MemberID;
                            _Import_HII_CSV_List_SF[1].First_Name = _Import_HII_CSV_List_SF[0].First_Name = HII_CSV_list[i].First_Name + ' ' + HII_CSV_list[i].Last_Name;

                            if (_Import_HII_CSV_List_SF[0].Agent_Status == "Not Available")
                                {
                                DateTime payHolddate = _Import_HII_CSV_List_SF[0].Ending_Date;
                                payHolddate = payHolddate.AddDays(searchPayHold(HII_CSV_list[i].CurrentRuning_total));

                                Payment_Period__c new_Pariod = Calendar_List.Where(c => c.Start_Date__c <= payHolddate && c.Close_Date__c >= payHolddate && c.Product_Profile__c == _ProductProfile).FirstOrDefault();

                                _Import_HII_CSV_List_SF[1].Payroll_date = _Import_HII_CSV_List_SF[0].Payroll_date = _Import_HII_CSV_List_SF[0].Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);


                                }
                            var temp_HII = _Import_HII_CSV_List_SF.Where(c => cancellationStatus.Contains(c.Payment_Status) || cancellationStatus.Contains(c.Payroll_Type)).FirstOrDefault();
                            if (temp_HII != null)
                                {
                                DateTime Termination_date = DateTime.Today;
                                DateTime EnrollmentDate = DateTime.Today;
                                DateTime Efective_date = DateTime.Today;

                                try
                                    {
                                    if (temp_HII.Termination_Date == "" || temp_HII.Termination_Date == null)
                                        {
                                        temp_HII.Termination_Date = temp_HII.Application_Date;
                                        }
                                    Termination_date = Convert.ToDateTime(temp_HII.Termination_Date);
                                    EnrollmentDate = Convert.ToDateTime(temp_HII.Application_Date);
                                    Efective_date = Convert.ToDateTime(HII_CSV_list[i].Effective_Date);
                                    }
                                catch (Exception exx)
                                    {
                                    HII_CSV_list[i].Verify = false;
                                    }
                                // Clasify_ChargeBack_Terminated(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, temp_HII);
                                Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, Efective_date, HII_CSV_list[i]);
                                }
                            }




                        }
                    else
                        {
                        Import_HII_CSV _Import_HII_CSV_sf = (from X in _Import_HII_CSV_List_SF
                                                             where X.Payroll_Type == "Commission"
                                                             select X).FirstOrDefault();
                        if (_Import_HII_CSV_sf == null)
                            {
                            _Import_HII_CSV_sf = (from X in _Import_HII_CSV_List_SF
                                                  where cancellationStatus.Contains(X.Payroll_Type)
                                                  select X).FirstOrDefault();
                            }

                        if (_Import_HII_CSV_sf != null)
                            {
                            _Import_HII_CSV_sf.Verify = true;
                            _Import_HII_CSV_sf.IsUpdate = true;
                            _Import_HII_CSV_sf.Member_ID = _Import_HII_CSV_sf.Sf_MemberID;
                            _Import_HII_CSV_sf.First_Name = HII_CSV_list[i].First_Name + ' ' + HII_CSV_list[i].Last_Name;

                            //*****// verifico Agent Status para ver si hay q cambiar el payroll date
                            if (_Import_HII_CSV_sf.Agent_Status == "Not Available")
                                {
                                DateTime payHolddate = _Import_HII_CSV_sf.Ending_Date;
                                payHolddate = payHolddate.AddDays(searchPayHold(""));

                                Payment_Period__c new_Pariod = Calendar_List.Where(c => c.Start_Date__c <= payHolddate && c.Close_Date__c >= payHolddate && c.Product_Profile__c == _ProductProfile).FirstOrDefault();

                                _Import_HII_CSV_sf.Payroll_date = HII_CSV_list[i].Payroll_date = Convert.ToDateTime(Calendar_Selected.Payment_Date__c);
                                _Import_HII_CSV_sf.Payroll_Type = HII_CSV_list[i].Payroll_Type = HII_CSV_list[i].Payroll_Type;

                                }
                            }
                        }



                    }
                else
                    {
                    try
                        {
                        /*  if (HII_CSV_list[i].Payment_Status == "Chargeback")
                         { 
                           DateTime Termination_date = DateTime.Today;
                           DateTime EnrollmentDate = DateTime.Today;
                           DateTime Efective_date = DateTime.Today;
                           if (HII_CSV_list[i].Effective_Date == "" || (HII_CSV_list[i].Effective_Date == null))
                               {
                               HII_CSV_list[i].Effective_Date= HII_CSV_list[i].Application_Date;
                               }
                               EnrollmentDate = Convert.ToDateTime(HII_CSV_list[i].Application_Date);
                               Efective_date = Convert.ToDateTime(HII_CSV_list[i].Effective_Date);
                               Termination_date = Convert.ToDateTime(HII_CSV_list[i].Termination_Date);

                          Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_date, EnrollmentDate, Efective_date, HII_CSV_list[i]);

                         }*/

                        MisMatch_HII_List.Add(HII_CSV_list[i]);
                        }

                    catch (Exception ex)
                        {

                        throw;
                        }



                    }


                }


            }


        public int searchPayHold(string current_running_total)
            {
            try
                {
                var element = PayHoldPeriod.Where(c => c.Resource_Type__c.Contains(current_running_total)).FirstOrDefault();

                if (element == null)
                    {
                    element = PayHoldPeriod.Where(c => c.Resource_Type__c.Equals("Pay Hold")).FirstOrDefault();
                    }

                return Convert.ToInt32(((White_List_Resource__c)element).Value__c);
                }
            catch (Exception ex)
                {
                return 0;
                }


            }



        public List<Payment__c> Insert_HII()
            {
            List<Payment__c> duplicateEntry = new List<Payment__c>();
            List<Payment__c> ValidateListEntry = new List<Payment__c>();
            try
                {
                List<Import_HII_CSV> Imported_HII_temp = HII_CSV_list.Where(c => c.Payroll_date != null && c.Payroll_Type != null).ToList();

                List<Payment__c> newPaymentEntryList = new List<Payment__c>();
                var opplineid = (from x in Imported_HII_temp
                                 select x.OpportunityLine_id).ToList();
                List<Payment__c> result_list = new List<Payment__c>();


                if (opplineid != null)
                    {

                    duplicateEntry = sf_interface.getpaymentCancelbyOppLine(opplineid);

                    for (int i = 0; i < Imported_HII_temp.Count; i++)
                        {

                        Payment__c newPaymentEntry = new Payment__c();
                        newPaymentEntry.Agent__c = Imported_HII_temp[i].Agent_id;
                        newPaymentEntry.OpportunityLineItem_id__c = Imported_HII_temp[i].OpportunityLine_id;

                        newPaymentEntry.Payment_Date__c = Imported_HII_temp[i].Payroll_date.ToUniversalTime();
                        newPaymentEntry.Payment_Date__cSpecified = true;
                        newPaymentEntry.Payment_Type__c = Imported_HII_temp[i].Payroll_Type;

                        newPaymentEntry.Payment_Value__c = Convert.ToDouble(Imported_HII_temp[i].Agent_Commision);
                        newPaymentEntry.Payment_Value__cSpecified = true;
                        newPaymentEntry.Policy_Number__c = Imported_HII_temp[i].Member_ID;
                        newPaymentEntry.Verify__c = true;
                        newPaymentEntry.Verify__cSpecified = true;

                        if (duplicateEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c)).ToList().Count > 0)
                            {

                            var errorPayment = duplicateEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && c.Payment_Type__c.Equals("Commission") && newPaymentEntry.Payment_Type__c.Equals(c.Payment_Type__c)).FirstOrDefault();



                            if (errorPayment != null && errorPayment.Verify__c != true)
                                {
                                errorPayment.Verify__c = true;
                                errorPayment.Verify__cSpecified = true;
                                ValidateListEntry.Add(errorPayment);
                                }
                            else
                                {
                                if (errorPayment != null)
                                    {
                                    errorPayment.Description__c = "Error: Duplicate Entry";
                                    }
                                if (errorPayment == null && newPaymentEntry.Payment_Type__c.Equals("Commission"))
                                    {
                                    newPaymentEntryList.Add(newPaymentEntry);
                                    }
                                }

                            var errorPaymentCancel = duplicateEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && newPaymentEntry.Payment_Type__c.Equals(c.Payment_Type__c) && (c.Payment_Type__c.Equals("Terminated") || c.Payment_Type__c.Equals("Chargeback"))).FirstOrDefault();

                            if (errorPaymentCancel != null && errorPaymentCancel.Verify__c != true)
                                {
                                errorPaymentCancel.Verify__c = true;
                                errorPaymentCancel.Verify__cSpecified = true;
                                ValidateListEntry.Add(errorPaymentCancel);

                                errorPayment = duplicateEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && c.Payment_Type__c.Equals("Commission") && c.Verify__c == false).FirstOrDefault();

                                if (errorPayment != null)
                                    {
                                    if (ValidateListEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && c.Payment_Type__c.Equals("Commission")).ToList().Count == 0)
                                        {
                                        errorPayment.Verify__c = true;
                                        errorPayment.Verify__cSpecified = true;
                                        ValidateListEntry.Add(errorPayment);
                                        }
                                    }


                                }
                            else
                                {
                                if (errorPaymentCancel != null && errorPaymentCancel.Verify__c == true)
                                    {
                                    errorPaymentCancel.Description__c = "Error: Duplicate Entry";
                                    }
                                if (errorPaymentCancel == null && (newPaymentEntry.Payment_Type__c.Equals("Terminated") || newPaymentEntry.Payment_Type__c.Equals("Chargeback")))
                                    {
                                    newPaymentEntryList.Add(newPaymentEntry);

                                    errorPayment = duplicateEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && c.Payment_Type__c.Equals("Commission") && c.Verify__c == false).FirstOrDefault();

                                    if (errorPayment != null)
                                        {
                                        if (ValidateListEntry.Where(c => c.OpportunityLineItem_id__c.Equals(newPaymentEntry.OpportunityLineItem_id__c) && c.Payment_Type__c.Equals("Commission")).ToList().Count == 0)
                                            {
                                            errorPayment.Verify__c = true;
                                            errorPayment.Verify__cSpecified = true;
                                            ValidateListEntry.Add(errorPayment);
                                            }
                                        }

                                    }


                                }


                            //duplicateEntry.Add(newPaymentEntry);
                            }
                        else
                            {
                            if (newPaymentEntry.Payment_Type__c != "Commission")
                                {
                                Payment__c newPaymentCommission = new Payment__c();
                                newPaymentCommission.Agent__c = Imported_HII_temp[i].Agent_id;
                                newPaymentCommission.OpportunityLineItem_id__c = Imported_HII_temp[i].OpportunityLine_id;

                                newPaymentCommission.Payment_Date__c = Imported_HII_temp[i].Payroll_date.ToUniversalTime();
                                newPaymentCommission.Payment_Date__cSpecified = true;
                                newPaymentCommission.Payment_Type__c = "Commission";

                                newPaymentCommission.Payment_Value__c = -1 * Convert.ToDouble(Imported_HII_temp[i].Agent_Commision);
                                newPaymentCommission.Payment_Value__cSpecified = true;
                                newPaymentCommission.Policy_Number__c = Imported_HII_temp[i].Member_ID;
                                newPaymentCommission.Verify__c = true;
                                newPaymentCommission.Verify__cSpecified = true;
                                newPaymentEntryList.Add(newPaymentCommission);
                                }

                            newPaymentEntryList.Add(newPaymentEntry);
                            }



                        }

                    List<SaveResult> result = new List<SaveResult>();
                    if (newPaymentEntryList.Count > 0)
                        {

                        result = sf_interface.insertPayment(newPaymentEntryList);

                        }

                    if (result.Count > 0)
                        {
                        for (int i = 0; i < result.Count; i++)
                            {
                            if (result[i].errors != null)
                                {
                                string error_message = "";
                                string cond = "";
                                List<object> array = new List<object>();
                                for (int k = 0; k < result[i].errors.Count(); k++)
                                    {
                                    if (k == 0)
                                        {
                                        string[] _message = result[i].errors[0].message.Split(':');
                                        cond = result[i].errors[0].fields[0] + ".Equals(@0)";
                                        array.Add(_message[2].Trim());
                                        }
                                    else
                                        {
                                        string[] _message = result[i].errors[0].message.Split(':');
                                        cond = cond + " and " + result[i].errors[0].fields[0] + ".Equals(@" + k.ToString() + ")";
                                        array.Add(_message[2]);
                                        }

                                    }

                                var PaymentEntryError = newPaymentEntryList.AsQueryable().Where(cond, array.ToArray()).FirstOrDefault();
                                if (PaymentEntryError != null)
                                    {
                                    PaymentEntryError.Description__c = result[i].errors[0].message;
                                    duplicateEntry.Add(PaymentEntryError);
                                    }
                                }

                            }


                        }


                    //update
                    List<SaveResult> resultUpdate = new List<SaveResult>();
                    if (ValidateListEntry.Count > 0)
                        {

                        resultUpdate = sf_interface.UpdatePayment(ValidateListEntry);
                        }

                    if (resultUpdate.Count > 0)
                        {
                        for (int i = 0; i < resultUpdate.Count; i++)
                            {
                            if (resultUpdate[i].errors != null)
                                {
                                string error_message = "";
                                string cond = "";
                                List<object> array = new List<object>();
                                for (int k = 0; k < resultUpdate[i].errors.Count(); k++)
                                    {
                                    if (k == 0)
                                        {
                                        string[] _message = resultUpdate[i].errors[0].message.Split(':');
                                        cond = resultUpdate[i].errors[0].fields[0] + ".Equals(@0)";
                                        array.Add(_message[2].Trim());
                                        }
                                    else
                                        {
                                        string[] _message = resultUpdate[i].errors[0].message.Split(':');
                                        cond = cond + " and " + resultUpdate[i].errors[0].fields[0] + ".Equals(@" + k.ToString() + ")";
                                        array.Add(_message[2]);
                                        }

                                    }

                                var PaymentEntryError = ValidateListEntry.AsQueryable().Where(cond, array.ToArray()).FirstOrDefault();
                                if (PaymentEntryError != null)
                                    {
                                    PaymentEntryError.Description__c = resultUpdate[i].errors[0].message;
                                    duplicateEntry.Add(PaymentEntryError);
                                    }
                                }

                            }


                        }




                    }



                }
            catch (Exception ex)
                {
                return new List<Payment__c>();
                }

            return duplicateEntry;

            }

        public static string FormatPhone(string aphone)
            {
            string aformat = "";
            try
                {
                aformat = aphone.Replace("-", string.Empty).Replace("(", string.Empty).Replace(")", string.Empty).Replace(" ", string.Empty);
                aformat = String.Format("{0:(###) ###-####}", double.Parse(aformat));
                }
            catch (Exception ex)
                {
                aformat = aphone;
                }

            return aformat;
            }




        public List<Fidelity_file> Fidelity(string username, string password, DateTime fromDate)
            {
            List<byte> responseArray = new List<byte>();
            List<Fidelity_file> List_fidelity = new List<Fidelity_file>();
            try
                {


                var client = new CookieAwareWebClient();
                client.BaseAddress = "https://agents.fidelitylifeassociation.com/agentregister/";

                // System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3;

                var newweb = client.DownloadString("AgentLogin.aspx");
                string[] Set_Cookie = client.ResponseHeaders["Set-Cookie"].Split('=');
                string _session_id = Set_Cookie[1].Replace("; path", "");

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(newweb);

                string _VIEWSTATEGENERATOR = "";
                string _EVENTVALIDATION = "";
                string _VIEWSTATE = "";
                string _EVENTTARGET = "";
                string _EVENTARGUMENT = "";

                foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                    {
                    if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                        {
                        _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                        }

                    if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                        {
                        _EVENTVALIDATION = input.Attributes["value"].Value;
                        }
                    if (input.Attributes["name"].Value == "__VIEWSTATE")
                        {
                        _VIEWSTATE = input.Attributes["value"].Value;
                        }

                    if (input.Attributes["name"].Value == "__EVENTTARGET")
                        {
                        _EVENTTARGET = input.Attributes["value"].Value;
                        }
                    if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                        {
                        _EVENTARGUMENT = input.Attributes["value"].Value;
                        }


                    }

                string User_id = username;


                string pass = password;
                var loginData = new NameValueCollection();
                loginData.Add("Button1", "Login");
                loginData.Add("txtUserId", User_id);//"mspiewak2014");
                loginData.Add("txtPassword", pass);// "Matt1984");
                loginData.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);// "Matt1984");
                loginData.Add("__EVENTVALIDATION", _EVENTVALIDATION);
                loginData.Add("__VIEWSTATE", _VIEWSTATE);
                loginData.Add("__EVENTTARGET", _EVENTTARGET);
                loginData.Add("__EVENTARGUMENT", _EVENTARGUMENT);


                responseArray = client.UploadValues(@"AgentLogin.aspx", "Post", loginData).ToList();
                var he = client.ResponseHeaders;

                var resBody = Encoding.UTF8.GetString(responseArray.ToArray());

                var searchvalues = new NameValueCollection();
                searchvalues.Add("SessionId", "");
                searchvalues.Add("CaseId", "");
                searchvalues.Add("Message", "");
                searchvalues.Add("Vendor", "");

                responseArray = client.UploadValues("AgentInfo.aspx", "POST", searchvalues).ToList();

                string htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());

                if (htmlSource.Contains("Logout"))
                    {
                    ///***LOGIN**//

                    //**Commission click** POsT AgentInfo*///
                    NameValueCollection headerCollection = new NameValueCollection();
                    searchvalues = new NameValueCollection();

                    doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);

                    _VIEWSTATEGENERATOR = "";
                    _EVENTVALIDATION = "";
                    _VIEWSTATE = "";
                    _EVENTTARGET = "lbCommissionsWebSite";
                    _EVENTARGUMENT = "";

                    foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                        {
                        if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                            {
                            _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                            }

                        if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                            {
                            _EVENTVALIDATION = input.Attributes["value"].Value;
                            }
                        if (input.Attributes["name"].Value == "__VIEWSTATE")
                            {
                            _VIEWSTATE = input.Attributes["value"].Value;
                            }


                        if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                            {
                            _EVENTARGUMENT = input.Attributes["value"].Value;
                            }


                        }





                    //client.BaseAddress = @"https://agents.fidelitylifeassociation.com/AgentRegister/AgentInfo.aspx";
                    searchvalues.Add("__EVENTTARGET", _EVENTTARGET);
                    searchvalues.Add("__EVENTARGUMENT", _EVENTARGUMENT);
                    searchvalues.Add("__VIEWSTATE", _VIEWSTATE);
                    searchvalues.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);
                    searchvalues.Add("__EVENTVALIDATION", _EVENTVALIDATION);

                    responseArray = client.UploadValues("AgentInfo.aspx", "POST", searchvalues).ToList();
                    string myString = Encoding.UTF8.GetString(responseArray.ToArray());

                    searchvalues = new NameValueCollection();
                    searchvalues.Add("SessionId", _session_id + User_id);
                    searchvalues.Add("CaseId", "");
                    searchvalues.Add("Message", "");
                    searchvalues.Add("Vendor", "");

                    client.BaseAddress = @"https://agents.fidelitylifeassociation.com/prod/FLACommissions/";
                    responseArray = client.UploadValues("FLACommissionsLogin.aspx", "POST", searchvalues).ToList();

                    htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());

                    doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);

                    _VIEWSTATEGENERATOR = "";
                    _EVENTVALIDATION = "";
                    _VIEWSTATE = "";
                    _EVENTTARGET = "";
                    _EVENTARGUMENT = "";

                    foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                        {
                        if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                            {
                            _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                            }

                        if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                            {
                            _EVENTVALIDATION = input.Attributes["value"].Value;
                            }
                        if (input.Attributes["name"].Value == "__VIEWSTATE")
                            {
                            _VIEWSTATE = input.Attributes["value"].Value;
                            }


                        if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                            {
                            _EVENTARGUMENT = input.Attributes["value"].Value;
                            }


                        }



                    searchvalues = new NameValueCollection();

                    string __LASTFOCUS = "";
                    searchvalues.Add("__EVENTTARGET", _EVENTTARGET);
                    searchvalues.Add("__EVENTARGUMENT", _EVENTARGUMENT);
                    searchvalues.Add("__VIEWSTATE", _VIEWSTATE);
                    searchvalues.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);
                    searchvalues.Add("__EVENTVALIDATION", _EVENTVALIDATION);
                    searchvalues.Add("__LASTFOCUS", __LASTFOCUS);
                    searchvalues.Add("lstAgents", User_id);
                    searchvalues.Add("ddlStartDate", "Current Period");
                    searchvalues.Add("ddlEndDate", "Current Period");
                    searchvalues.Add("ddlCategory", "*");
                    searchvalues.Add("btnGenerateReport", "Statement Details");
                    searchvalues.Add("HiddenField1", "FormLoaded");
                    searchvalues.Add("LoggedInAgentNo", "");
                    searchvalues.Add("CallingProgram", "");
                    searchvalues.Add("DisplayAgent", "");
                    searchvalues.Add("HierarchyViewAgent", "");
                    searchvalues.Add("SelectedAgent", "");
                    searchvalues.Add("CallingURL", "");
                    // client.isfile = true;
                    client.BaseAddress = @"https://agents.fidelitylifeassociation.com/prod/FLACommissions/FLACommissions.aspx";
                    responseArray = client.UploadValues("FLACommissions.aspx", "POST", searchvalues).ToList();
                    htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());


                    /**EXcell*/


                    doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);

                    _VIEWSTATEGENERATOR = "";
                    _EVENTVALIDATION = "";
                    _VIEWSTATE = "";
                    _EVENTTARGET = "";
                    _EVENTARGUMENT = "";

                    foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                        {
                        if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                            {
                            _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                            }

                        if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                            {
                            _EVENTVALIDATION = input.Attributes["value"].Value;
                            }
                        if (input.Attributes["name"].Value == "__VIEWSTATE")
                            {
                            _VIEWSTATE = input.Attributes["value"].Value;
                            }


                        if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                            {
                            _EVENTARGUMENT = input.Attributes["value"].Value;
                            }


                        }






                    searchvalues = new NameValueCollection();

                    __LASTFOCUS = "";
                    searchvalues.Add("__EVENTTARGET", _EVENTTARGET);
                    searchvalues.Add("__EVENTARGUMENT", _EVENTARGUMENT);
                    searchvalues.Add("__VIEWSTATE", _VIEWSTATE);
                    searchvalues.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);
                    searchvalues.Add("__EVENTVALIDATION", _EVENTVALIDATION);
                    searchvalues.Add("__LASTFOCUS", __LASTFOCUS);
                    searchvalues.Add("lstAgents", User_id);
                    searchvalues.Add("ddlStartDate", "Current Period");
                    searchvalues.Add("ddlEndDate", "Current Period");
                    searchvalues.Add("ddlCategory", "*");
                    searchvalues.Add("btnExport", "Export Report to Excel");
                    searchvalues.Add("HiddenField1", "FormLoaded");
                    searchvalues.Add("LoggedInAgentNo", "");
                    searchvalues.Add("CallingProgram", "");
                    searchvalues.Add("DisplayAgent", "");
                    searchvalues.Add("HierarchyViewAgent", "");
                    searchvalues.Add("SelectedAgent", "");
                    searchvalues.Add("CallingURL", "");
                    client.isfile = true;
                    client.BaseAddress = @"https://agents.fidelitylifeassociation.com/prod/FLACommissions/FLACommissions.aspx";
                    responseArray = client.UploadValues("FLACommissions.aspx", "POST", searchvalues).ToList();
                    htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());



                    doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);


                    foreach (HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
                        {
                        int row_count = 0;
                        foreach (HtmlNode row in table.SelectNodes("tr"))
                            {

                            if (row_count > 0)
                                {
                                Fidelity_file _fidelity = new Fidelity_file();
                                _fidelity.Agent = row.SelectNodes("th|td")[0].InnerText.Replace("&nbsp;", "");
                                _fidelity.ReportDate = row.SelectNodes("th|td")[1].InnerText.Replace("&nbsp;", "");
                                _fidelity.Amount = row.SelectNodes("th|td")[2].InnerText.Replace("&nbsp;", "");
                                _fidelity.Category = row.SelectNodes("th|td")[3].InnerText.Replace("&nbsp;", "");
                                _fidelity.ReportingCategory = row.SelectNodes("th|td")[4].InnerText.Replace("&nbsp;", "");
                                _fidelity.PolicyNo = row.SelectNodes("th|td")[5].InnerText.Replace("&nbsp;", "").Split('P')[0];
                                _fidelity.State = row.SelectNodes("th|td")[6].InnerText.Replace("&nbsp;", "");
                                _fidelity.Insured = row.SelectNodes("th|td")[7].InnerText.Replace("&nbsp;", "");
                                _fidelity.WritingAgent = row.SelectNodes("th|td")[8].InnerText.Replace("&nbsp;", "");
                                _fidelity.WritingAgentName = row.SelectNodes("th|td")[9].InnerText.Replace("&nbsp;", "");
                                _fidelity.ProductGroup = row.SelectNodes("th|td")[10].InnerText.Replace("&nbsp;", "");
                                _fidelity.Contract = row.SelectNodes("th|td")[11].InnerText.Replace("&nbsp;", "");
                                _fidelity.CommissionRate = row.SelectNodes("th|td")[12].InnerText.Replace("&nbsp;", "");
                                _fidelity.SplitPct = row.SelectNodes("th|td")[13].InnerText.Replace("&nbsp;", "");
                                _fidelity.MonthsPaid = row.SelectNodes("th|td")[14].InnerText.Replace("&nbsp;", "");
                                _fidelity.Duration = row.SelectNodes("th|td")[15].InnerText.Replace("&nbsp;", "");
                                _fidelity.Premium = row.SelectNodes("th|td")[16].InnerText.Replace("&nbsp;", "");
                                _fidelity.PaidToDate = row.SelectNodes("th|td")[17].InnerText.Replace("&nbsp;", "");
                                _fidelity.EntryDate = row.SelectNodes("th|td")[18].InnerText.Replace("&nbsp;", "");
                                _fidelity.Other = row.SelectNodes("th|td")[19].InnerText.Replace("&nbsp;", "");

                                if (_fidelity.PolicyNo == "0100700380")
                                    {
                                    var vv = "";
                                    }
                                if (Convert.ToDateTime(_fidelity.EntryDate).Year < DateTime.Today.Year)
                                    {
                                    var vv = "";
                                    }
                                try
                                    {

                                    DateTime _create = Convert.ToDateTime(_fidelity.EntryDate);
                                    if (fromDate <= _create && List_fidelity.Where(c => c.PolicyNo == _fidelity.PolicyNo).ToList().Count == 0)
                                        {
                                        List_fidelity.Add(_fidelity);

                                        }

                                    }
                                catch (Exception)
                                    {
                                    if (List_fidelity.Where(c => c.PolicyNo == _fidelity.PolicyNo).ToList().Count == 0)
                                        {
                                        _fidelity.EntryDate = fromDate.ToString("MM/dd/yyyy");
                                        List_fidelity.Add(_fidelity);
                                        }
                                    throw;
                                    }


                                }
                            row_count++;
                            }
                        }


                    }
                }
            catch (Exception ex)
                {
                // message = ex.Message;
                }






            return List_fidelity;




            }

        public List<GetMed_file> GetMed(DateTime fromDate, DateTime toDate, string username, string password)
            {

            List<GetMed_file> list_GetMed_file = new List<GetMed_file>();
            List<byte> responseArray = new List<byte>();
            string _csv_file = "";
            try
                {


                var client = new CookieAwareWebClient();
                client.BaseAddress = "https://agent.mymemberinfo.com/";

                // System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3;

                var newweb = client.DownloadString("default.aspx");
                string[] Set_Cookie = client.ResponseHeaders["Set-Cookie"].Split('=');
                string _session_id = Set_Cookie[1].Replace("; path", "");

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(newweb);

                string _VIEWSTATEGENERATOR = "";
                string _EVENTVALIDATION = "";
                string _VIEWSTATE = "";
                string _EVENTTARGET = "";
                string _EVENTARGUMENT = "";
                string _txtUserName = "";
                string _txtPassword = "";
                string _btnLogin = "";

                foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                    {
                    if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                        {
                        _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                        }

                    if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                        {
                        _EVENTVALIDATION = input.Attributes["value"].Value;
                        }
                    if (input.Attributes["name"].Value == "__VIEWSTATE")
                        {
                        _VIEWSTATE = input.Attributes["value"].Value;
                        }

                    if (input.Attributes["name"].Value == "__EVENTTARGET")
                        {
                        _EVENTTARGET = input.Attributes["value"].Value;
                        }
                    if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                        {
                        _EVENTARGUMENT = input.Attributes["value"].Value;
                        }


                    if (input.Attributes["name"].Value.Contains("txtUserName"))
                        {
                        _txtUserName = input.Attributes["name"].Value;
                        }

                    if (input.Attributes["name"].Value.Contains("txtPassword"))
                        {
                        _txtPassword = input.Attributes["name"].Value;
                        }
                    if (input.Attributes["name"].Value.Contains("btnLogin"))
                        {
                        _btnLogin = input.Attributes["name"].Value;
                        }
                    }

                string User_id = username;


                string pass = password;
                var loginData = new NameValueCollection();
                loginData.Add(_btnLogin, User_id);
                loginData.Add(_txtUserName, User_id);//"mspiewak2014");
                loginData.Add(_txtPassword, pass);// "Matt1984");
                loginData.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);// "Matt1984");
                loginData.Add("__EVENTVALIDATION", _EVENTVALIDATION);
                loginData.Add("__VIEWSTATE", _VIEWSTATE);
                loginData.Add("__EVENTTARGET", _EVENTTARGET);
                loginData.Add("__EVENTARGUMENT", _EVENTARGUMENT);


                responseArray = client.UploadValues(@"default.aspx", "Post", loginData).ToList();


                string htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());

                //htmlSource = client.DownloadString("Agent/MyHome.aspx");
                if (htmlSource.Contains("Logout"))
                    {
                    string Source = htmlSource = client.DownloadString("https://agent.mymemberinfo.com/Report/Member/GenericMemberList.aspx");

                    doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);
                    string line = "";
                    foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//div[@class='art-content']"))
                        {
                        if (input != null)
                            {
                            var links = input.Descendants("script")
                                           .Select(a => a.InnerText)
                                           .ToList();

                            if (links != null)
                                {
                                for (int i = 0; i < links.Count; i++)
                                    {
                                    line = links[i].ToString();
                                    if (line.Contains("setTimeout(") && line.Contains("ReportSession") && line.Contains("ReportSession") && line.Contains("ControlID") && line.Contains("ReportStack") && line.Contains("OpType"))
                                        {
                                        break;
                                        }
                                    else
                                        {
                                        line = "";
                                        }
                                    }
                                if (line != "")
                                    {
                                    break;
                                    }
                                }
                            }


                        }
                    if (line != "")
                        {
                        string[] _array = line.Split(';');

                        string formatdate = DateTime.Now.ToUniversalTime().ToString("r");
                        _array = _array[0].Split('?');
                        _array = _array[1].Split('&');
                        string _CacheSeed = DateTime.Now.ToString("ddd MMM dd yyyy HH:mm:ss ';'zzzz") + " (Eastern Daylight Time)";
                        string gtm = _CacheSeed.Split(';')[1].Replace(":", string.Empty);
                        _CacheSeed = _CacheSeed.Split(';')[0] + "GMT" + gtm;
                        //formatdate = Uri.UnescapeDataString(_CacheSeed);

                        formatdate = System.Web.HttpUtility.UrlEncode(_CacheSeed);
                        formatdate = formatdate.Replace("+", "%20");
                        //_CacheSeed = _CacheSeed.Replace(" ", "%");
                        //_CacheSeed = Uri.EscapeUriString(_CacheSeed);//.Replace(":", "%").Replace(",", "&");
                        _CacheSeed = formatdate;
                        string _ReportSession = _array[0].Split('=')[1];
                        string _ControlID = _array[1].Split('=')[1];
                        string _Culture = _array[2].Split('=')[1];
                        string _UICulture = _array[3].Split('=')[1];
                        string _ReportStack = _array[4].Split('=')[1];
                        string _OpType = _array[5].Split('=')[1];
                        string _TimerMethod = _array[6].Split('=')[1];
                        string _urlGet_COntrol = "?ReportSession=" + _ReportSession + "&ControlID=" + _ControlID + "&Culture=" + _Culture + "&UICulture=" + _UICulture + "&ReportStack=" + _ReportStack + "&OpType=" + _OpType + "&TimerMethod=" + _TimerMethod + "&CacheSeed=" + _CacheSeed;
                        _urlGet_COntrol = "https://agent.mymemberinfo.com/Reserved.ReportViewerWebControl.axd" + _urlGet_COntrol;

                        htmlSource = client.DownloadString(_urlGet_COntrol);

                        /***eXECUTE rePORT*/

                        doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(Source);

                        _VIEWSTATEGENERATOR = "";
                        _EVENTVALIDATION = "";
                        _VIEWSTATE = "";
                        _EVENTTARGET = "";
                        _EVENTARGUMENT = "";
                        _txtUserName = "";
                        _txtPassword = "";

                        string _PREVIOUSPAGE = "";
                        string _ClientState_name = "";
                        string _ClientState_value = "";


                        string _ctl03_ddValue_value = "";
                        string _ctl03_ddValue_name = "";

                        string _txtSearchText_value = "";
                        string _txtSearchText_name = "";




                        string _ctl05_txtValue_value = "";
                        string _ctl05_txtValue_name = "";

                        string _ctl07_txtValue_value = "";
                        string _ctl07_txtValue_name = "";

                        string _ctl04_name = "";
                        string _ctl04_value = "";

                        string _ctl05_name = "";
                        string _ctl05_value = "";

                        string _ctl06_name = "";
                        string _ctl06_value = "";

                        string _ctl07_name = "";
                        string _ctl07_value = "";

                        string _ctl08_name = "";
                        string _ctl08_value = "";


                        string _ctl00_name = "";
                        string _ctl00_value = "";

                        foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//input"))
                            {
                            if (input.Attributes["name"].Value == "__VIEWSTATEGENERATOR")
                                {
                                _VIEWSTATEGENERATOR = input.Attributes["value"].Value;
                                }

                            if (input.Attributes["name"].Value == "__EVENTVALIDATION")
                                {
                                _EVENTVALIDATION = input.Attributes["value"].Value;
                                }
                            if (input.Attributes["name"].Value == "__VIEWSTATE")
                                {

                                _VIEWSTATE = input.Attributes["value"].Value;
                                }

                            if (input.Attributes["name"].Value == "__EVENTTARGET")
                                {
                                _EVENTTARGET = input.Attributes["value"].Value;
                                }
                            if (input.Attributes["name"].Value == "__EVENTARGUMENT")
                                {
                                _EVENTARGUMENT = input.Attributes["value"].Value;
                                }


                            if (input.Attributes["name"].Value.Contains("_PREVIOUSPAGE"))
                                {
                                _PREVIOUSPAGE = input.Attributes["value"].Value;
                                }

                            if (input.Attributes["name"].Value.Contains("TextBoxWatermarkExtender_ClientState"))
                                {

                                _ClientState_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ClientState_value = input.Attributes["value"].Value;
                                    }

                                }


                            if (input.Attributes["name"].Value.Contains("txtSearchText") && !input.Attributes["name"].Value.Contains("TextBoxWatermarkExtender_ClientState"))
                                {

                                _txtSearchText_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _txtSearchText_value = input.Attributes["value"].Value;
                                    }

                                }




                            if (input.Attributes["name"].Value.Contains("txtValue") && input.Attributes["name"].Value.Contains("ctl05"))
                                {
                                _ctl05_txtValue_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl05_txtValue_value = input.Attributes["value"].Value;
                                    }
                                else
                                    {
                                    _ctl05_txtValue_value = fromDate.ToString("yyyy-MM-dd");
                                    }
                                }
                            if (input.Attributes["name"].Value.Contains("txtValue") && input.Attributes["name"].Value.Contains("ctl07"))
                                {
                                _ctl07_txtValue_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl07_txtValue_value = input.Attributes["value"].Value;
                                    }
                                else
                                    {
                                    _ctl07_txtValue_value = toDate.ToString("yyyy-MM-dd");
                                    }
                                }

                            //ctl00$SheetContentPlaceHolder$rptMemberList$ctl04

                            if (input.Attributes["name"].Value.EndsWith("ctl04"))
                                {
                                _ctl04_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl04_value = input.Attributes["value"].Value;
                                    }
                                }
                            if (input.Attributes["name"].Value.EndsWith("ctl05"))
                                {
                                _ctl05_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl05_value = input.Attributes["value"].Value;
                                    }
                                }
                            if (input.Attributes["name"].Value.EndsWith("ctl06"))
                                {
                                _ctl06_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl06_value = input.Attributes["value"].Value;
                                    }
                                }

                            if (input.Attributes["name"].Value.EndsWith("ctl07"))
                                {
                                _ctl07_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl07_value = input.Attributes["value"].Value;
                                    }
                                }

                            if (input.Attributes["name"].Value.EndsWith("ctl08"))
                                {
                                _ctl08_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl08_value = input.Attributes["value"].Value;

                                    }

                                }

                            if (input.Attributes["name"].Value.EndsWith("ctl00"))
                                {
                                _ctl00_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl00_value = input.Attributes["value"].Value;

                                    }

                                }
                            }
                        foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//select"))
                            {
                            if (input.Attributes["name"].Value.Contains("ddValue") && input.Attributes["name"].Value.Contains("ctl03"))
                                {
                                _ctl03_ddValue_name = input.Attributes["name"].Value;
                                if (input.Attributes.Contains("value"))
                                    {
                                    _ctl03_ddValue_value = input.Attributes["value"].Value;
                                    }
                                else
                                    {
                                    _ctl03_ddValue_value = "1";
                                    }
                                }
                            }

                        _urlGet_COntrol = "";

                        foreach (HtmlNode input in doc.DocumentNode.SelectNodes("//iframe"))
                            {
                            if (input.Attributes["name"].Value.Contains("ReportFramectl00_SheetContentPlaceHolder_rptMemberList"))
                                {

                                if (input.Attributes.Contains("src"))
                                    {
                                    _urlGet_COntrol = input.Attributes["src"].Value;
                                    }

                                }
                            }

                        var searchvalues = new NameValueCollection();
                        searchvalues.Add("__EVENTTARGET", _EVENTTARGET);
                        searchvalues.Add("__EVENTARGUMENT", _EVENTARGUMENT);
                        searchvalues.Add("__VIEWSTATE", _VIEWSTATE.Trim());
                        searchvalues.Add("__VIEWSTATEGENERATOR", _VIEWSTATEGENERATOR);
                        searchvalues.Add("__PREVIOUSPAGE", _PREVIOUSPAGE);
                        searchvalues.Add("__EVENTVALIDATION", _EVENTVALIDATION);

                        searchvalues.Add(_ClientState_name, _ClientState_value);
                        searchvalues.Add(_txtSearchText_name, _txtSearchText_value);
                        searchvalues.Add(_ctl03_ddValue_name, _ctl03_ddValue_value);
                        searchvalues.Add(_ctl05_txtValue_name, _ctl05_txtValue_value);
                        searchvalues.Add(_ctl07_txtValue_name, _ctl07_txtValue_value);
                        searchvalues.Add(_ctl00_name, _ctl00_value);
                        searchvalues.Add(_ctl04_name, _ctl04_value);
                        searchvalues.Add(_ctl05_name, _ctl05_value);
                        searchvalues.Add(_ctl06_name, _ctl06_value);
                        searchvalues.Add(_ctl07_name, _ctl07_value);
                        searchvalues.Add(_ctl08_name, _ctl08_value);

                        client.BaseAddress = "https://agent.mymemberinfo.com/Report/Member/";
                        responseArray = client.UploadValues(@"GenericMemberList.aspx", "Post", searchvalues).ToList();
                        htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());




                        _urlGet_COntrol = System.Web.HttpUtility.UrlEncode("/Reserved.ReportViewerWebControl.axd?ReportSession=" + _ReportSession + "&ControlID=" + _ControlID + "&Culture=" + _Culture + "&UICulture=" + _UICulture + "&ReportStack=" + _ReportStack + "&OpType=ReportArea&Controller=ctl00_SheetContentPlaceHolder_rptMemberList&PageNumber=1&ZoomMode=Percent&ZoomPct=100&ReloadDocMap=true&SearchStartPage=0&LinkTarget=_top");
                        _urlGet_COntrol = _urlGet_COntrol.Replace("+", "%20");
                        _urlGet_COntrol = "?OpType=DocMapReport&ClientController=ctl00_SheetContentPlaceHolder_rptMemberList&ReportUrl=" + _urlGet_COntrol;

                        _urlGet_COntrol = @"https://agent.mymemberinfo.com/Reserved.ReportViewerWebControl.axd" + _urlGet_COntrol;

                        htmlSource = client.DownloadString(_urlGet_COntrol);
                        htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());
                        _urlGet_COntrol = "";

                        _urlGet_COntrol = "?ReportSession=" + _ReportSession + "&ControlID=" + _ControlID + "&Culture=" + _Culture + "&UICulture=" + _UICulture + "&ReportStack=" + _ReportStack + "&OpType=SessionKeepAlive&TimerMethod=" + _TimerMethod + "&CacheSeed=" + _CacheSeed;
                        htmlSource = client.DownloadString("https://agent.mymemberinfo.com/Reserved.ReportViewerWebControl.axd" + _urlGet_COntrol);
                        htmlSource = Encoding.UTF8.GetString(responseArray.ToArray());


                        _urlGet_COntrol = "";

                        _urlGet_COntrol = "?ReportSession=" + _ReportSession + "&ControlID=" + _ControlID + "&Culture=" + _Culture + "&UICulture=" + _UICulture + "&ReportStack=" + _ReportStack + "&OpType=ReportArea&Controller=ctl00_SheetContentPlaceHolder_rptMemberList&PageNumber=1&ZoomMode=Percent&ZoomPct=100&ReloadDocMap=true&SearchStartPage=0&LinkTarget=_top";
                        htmlSource = client.DownloadString("https://agent.mymemberinfo.com/Reserved.ReportViewerWebControl.axd" + _urlGet_COntrol);


                        _urlGet_COntrol = "https://agent.mymemberinfo.com/Reserved.ReportViewerWebControl.axd?ReportSession=" + _ReportSession + "&ControlID=" + _ControlID + "&Culture=" + _Culture + "&UICulture=" + _UICulture + "&ReportStack=" + _ReportStack + "&OpType=Export&FileName=Member List&ContentDisposition=OnlyHtmlInline&Format=CSV";
                        htmlSource = client.DownloadString(_urlGet_COntrol);


                        _csv_file = htmlSource;



                        }


                    }



                }

            catch (Exception ex)
                {

                }
            list_GetMed_file = new List<GetMed_file>();

            MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_csv_file));
            StreamReader reader = new StreamReader(ms);

            while (!reader.EndOfStream)
                {

                var line = reader.ReadLine();
                if (line.Contains("PGA9002752"))
                    {
                    var h = "k";
                    }

                var csvRecord = line.Replace("\",\"", ";");
                csvRecord = ParseCommasInQuotes(csvRecord);
                var csvRecordData = csvRecord.Split(',');
                GetMed_file CSV_front_Obj = new GetMed_file();


                if (csvRecordData.Count() > 0)
                    {

                    try
                        {


                        CSV_front_Obj = new GetMed_file();
                        CSV_front_Obj.Member_ID = csvRecordData[0];
                        CSV_front_Obj.IsNewPolicy = csvRecordData[1];
                        CSV_front_Obj.StatusName = csvRecordData[2];
                        CSV_front_Obj.CoverageType = csvRecordData[3];
                        CSV_front_Obj.agent_id = csvRecordData[4];
                        CSV_front_Obj.AgentFirstName = csvRecordData[5];
                        CSV_front_Obj.FirstName = csvRecordData[6];
                        CSV_front_Obj.MI = csvRecordData[7];
                        CSV_front_Obj.LastName = csvRecordData[8];
                        CSV_front_Obj.DOB = csvRecordData[9];
                        CSV_front_Obj.Gender = csvRecordData[10];
                        CSV_front_Obj.address1 = csvRecordData[11];
                        CSV_front_Obj.zip = csvRecordData[12];
                        CSV_front_Obj.Telephone = csvRecordData[13];
                        CSV_front_Obj.Fax = csvRecordData[14];
                        CSV_front_Obj.Email = csvRecordData[15];
                        CSV_front_Obj.StartDate = csvRecordData[16];
                        CSV_front_Obj.EffectiveDate = csvRecordData[17];
                        CSV_front_Obj.TerminateDate = csvRecordData[18];
                        CSV_front_Obj.GroupName = csvRecordData[19];
                        CSV_front_Obj.AccountTypeName = csvRecordData[20];
                        CSV_front_Obj.Textbox12 = csvRecordData[21];
                        CSV_front_Obj.ccv = csvRecordData[22];
                        CSV_front_Obj.Textbox4 = csvRecordData[23];
                        CSV_front_Obj.ck_accntnumber = csvRecordData[24];
                        CSV_front_Obj.ck_rtenumber = csvRecordData[25];
                        CSV_front_Obj.acct_firstname = csvRecordData[26];
                        CSV_front_Obj.acct_address1 = csvRecordData[27];
                        CSV_front_Obj.acct_zip = csvRecordData[28];
                        CSV_front_Obj.MainPackageName = csvRecordData[29];
                        CSV_front_Obj.MembershipCost = csvRecordData[30];
                        CSV_front_Obj.NextChargeDate = csvRecordData[31];

                        if (!CSV_front_Obj.Member_ID.Contains("Member_ID"))
                            list_GetMed_file.Add(CSV_front_Obj);
                        }
                    catch (Exception ex)
                    { }


                    }
                }


            return list_GetMed_file;

            }

        public List<Enrollment123> Getwebfileenrollment123(DateTime fromDate, DateTime toDate, string username, string password)
            {
            List<byte> responseArray = new List<byte>();
            string _csv_file = "";
            List<Enrollment123> Enrollment123_List = new List<Enrollment123>();
            try
                {


                var client = new CookieAwareWebClient();
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705;)");
                client.BaseAddress = @"https://www.enrollment123.com/manage/index.cfm";
                //client.isfile = false;
                var loginData = new NameValueCollection();
                loginData.Add("username", username);
                loginData.Add("password", password);//"mspiewak2014");
                loginData.Add("authenticate", "");// "Matt1984");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
         | SecurityProtocolType.Tls11
         | SecurityProtocolType.Tls12
         | SecurityProtocolType.Ssl3;


                client.UploadValues(@"index.cfm", "Post", loginData);

                string htmlSource = client.DownloadString(@"https://www.enrollment123.com/manage/index.cfm?");

                if (htmlSource.Contains("Log Out"))
                    {


                    htmlSource = client.DownloadString(@"https://www.enrollment123.com/manage/users/index.cfm?reportId=143&submit=View");
                    string CreatedStart = fromDate.ToString("MM/dd/yyyy");

                    var collection = maxpagination_WPA(htmlSource);
                    string token = "";
                    foreach (var item in collection)
                        {
                        if (item.token != null && item.token != "")
                            {
                            token = item.token;
                            break;
                            }
                        }

                    if (token != "")
                        {
                        var postData = new NameValueCollection();
                        postData.Add("method", "downloadFixed");
                        postData.Add("token", token);
                        postData.Add("command", "full");
                        postData.Add("returnFormat", "json");

                        client.Headers.Add("Host", "www.enrollment123.com");
                        client.Headers.Add("Accept", "application/json, text/javascript, */*");
                        client.Headers.Add("Origin", @"https://www.enrollment123.com");
                        client.Headers.Add("X-Requested-With", "XMLHttpRequest");
                        client.Headers.Add("Referer", @"https://www.enrollment123.com/manage/users/index.cfm?token=" + token);
                        client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36");

                        client.BaseAddress = @"https://www.enrollment123.com/includes/cfc/library/filter.cfc";
                        byte[] byteArray = client.UploadValues(@"filter.cfc", "Post", postData);
                        string result = System.Text.Encoding.UTF8.GetString(byteArray);
                        string[] str_array = result.Split(',');

                        string token_response = "";

                        foreach (var item in str_array)
                            {
                            if (item.ToUpper().Contains("DATA"))
                                {
                                token_response = item.Split(':')[1].Replace('}', ' ').Replace('\"', ' ');
                                break;
                                }
                            }

                        if (token_response != "")
                            {
                            string csv_url = @"https://www.enrollment123.com/manage/users/downloadCommandHandler.cfc?method=run&returnFormat=plain&command=full&token=" + token_response.Trim() + "&idtype=up_id";

                            client.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                            client.Headers.Add("Upgrade-Insecure-Requests", "1");
                            client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36");
                            client.Headers.Add("Referer", @"https://www.enrollment123.com/manage/users/index.cfm?token=" + token);
                            //   client.Headers.Add("Accept-Encoding", "csv");
                            client.isfile = true;

                            postData = new NameValueCollection();
                            postData.Add("method", "run");
                            postData.Add("returnFormat", "plain");
                            postData.Add("command", "full");
                            postData.Add("token", token_response.Trim());
                            postData.Add("idtype", "up_id");

                            client.BaseAddress = @"https://www.enrollment123.com/manage/users/downloadCommandHandler.cfc";
                            byteArray = client.UploadValues(@"downloadCommandHandler.cfc", "POST", postData);
                            _csv_file = System.Text.Encoding.UTF8.GetString(byteArray);

                            }


                        }


                    /*_csv_file = client.DownloadString(@"https://www.enrollment123.com/manage/users/downloadCommandHandler.cfc?method=run&returnFormat=plain&command=full&usertype=member&idType=up_id&location=member&qs=" + qs + "&urlVarList=" + urlVarList + "&output=full&lSelectedIDs=");

                    */


                    }


                }
            catch (Exception ex)
                {
                _csv_file = "";
                }

            if (_csv_file != "")
                {

                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_csv_file));
                StreamReader reader = new StreamReader(ms);

                while (!reader.EndOfStream)
                    {

                    var line = reader.ReadLine();
                    var csvRecord = line.Replace("\",\"", ";");
                    var csvRecordData = csvRecord.Split(';');



                    if (csvRecordData.Count() > 0)
                        {

                        try
                            {

                            Enrollment123 _Enrollment123 = new Enrollment123();

                            _Enrollment123.Agent_ID = csvRecordData[0];
                            _Enrollment123.Agent_Label = csvRecordData[1];
                            _Enrollment123.Member_ID = csvRecordData[2];
                            _Enrollment123.First_Name = csvRecordData[3];
                            _Enrollment123.Last_Name = csvRecordData[4];
                            _Enrollment123.Address = csvRecordData[5];
                            _Enrollment123.Address2 = csvRecordData[6];
                            _Enrollment123.City = csvRecordData[7];
                            _Enrollment123.State = csvRecordData[8];
                            _Enrollment123.Zipcode = csvRecordData[9];
                            _Enrollment123.Country = csvRecordData[10];
                            _Enrollment123.County = csvRecordData[11];
                            _Enrollment123.Other_Address = csvRecordData[12];
                            _Enrollment123.Other_Address2 = csvRecordData[13];
                            _Enrollment123.Other_City = csvRecordData[14];
                            _Enrollment123.Other_State = csvRecordData[15];
                            _Enrollment123.Other_Zipcode = csvRecordData[16];
                            _Enrollment123.Phone_1 = csvRecordData[17];
                            _Enrollment123.Phone_1_Ext = csvRecordData[18];

                            _Enrollment123.Phone_2 = csvRecordData[19];
                            _Enrollment123.Phone_2_Ext = csvRecordData[20];
                            _Enrollment123.Phone_3 = csvRecordData[21];
                            _Enrollment123.Phone_3_Ext = csvRecordData[22];
                            _Enrollment123.Fax = csvRecordData[23];
                            _Enrollment123.Fax_Ext = csvRecordData[24];
                            _Enrollment123.Email = csvRecordData[25];
                            _Enrollment123.Email2 = csvRecordData[26];
                            _Enrollment123.DOB = csvRecordData[27];
                            _Enrollment123.Gender = csvRecordData[28];
                            _Enrollment123.SSN = csvRecordData[29];
                            _Enrollment123.Drivers_License = csvRecordData[30];
                            _Enrollment123.Company_Name = csvRecordData[31];
                            _Enrollment123.Position = csvRecordData[32];
                            _Enrollment123.Department = csvRecordData[33];
                            _Enrollment123.Division = csvRecordData[34];
                            _Enrollment123.Username = csvRecordData[35];
                            _Enrollment123.Password = csvRecordData[36];
                            _Enrollment123.Internal_ID = csvRecordData[37];
                            _Enrollment123.Referral = csvRecordData[38];
                            _Enrollment123.Source = csvRecordData[39];
                            _Enrollment123.Email_Opt_Out = csvRecordData[40];
                            _Enrollment123.Language = csvRecordData[41];
                            _Enrollment123.Type = csvRecordData[42];
                            _Enrollment123.Type_2 = csvRecordData[43];
                            _Enrollment123.Member_Created_Date = csvRecordData[44];
                            _Enrollment123.Product_Created_Date = csvRecordData[45];
                            _Enrollment123.Active_Date = csvRecordData[46];
                            _Enrollment123.First_Billing_Date = csvRecordData[47];
                            _Enrollment123.Next_Billing_Date = csvRecordData[48];
                            _Enrollment123.Fulfillment_Date = csvRecordData[49];
                            _Enrollment123.Inactive_Date = csvRecordData[50];
                            _Enrollment123.Inactive_Reason = csvRecordData[51];
                            _Enrollment123.Hold_Date = csvRecordData[52];
                            _Enrollment123.Hold_Reason = csvRecordData[53];
                            _Enrollment123.Hold_Return = csvRecordData[54];
                            _Enrollment123.Hold_Amount = csvRecordData[55];
                            _Enrollment123.Product_ID = csvRecordData[56];
                            _Enrollment123.Product_Label = csvRecordData[57];
                            _Enrollment123.Product_Benefit = csvRecordData[58];
                            _Enrollment123.Period_Label = csvRecordData[59];
                            _Enrollment123.Category_3 = csvRecordData[60];
                            _Enrollment123.Category_4 = csvRecordData[61];
                            _Enrollment123.Code = csvRecordData[62];
                            _Enrollment123.Permanent_Bill_Day = csvRecordData[63];
                            _Enrollment123.Product_Source = csvRecordData[64];
                            _Enrollment123.Product_Source_Detail = csvRecordData[65];
                            _Enrollment123.Product_Status = csvRecordData[66];
                            _Enrollment123.Product_Stage = csvRecordData[67];
                            _Enrollment123.Product_Number_Calls = csvRecordData[68];
                            _Enrollment123.Product_Call_Status = csvRecordData[69];
                            _Enrollment123.Product_Next_Step = csvRecordData[70];
                            _Enrollment123.Product_Next_Step_Date = csvRecordData[71];
                            _Enrollment123.Product_Estimated_Close_Date = csvRecordData[72];
                            _Enrollment123.Billing_Fee = csvRecordData[73];
                            _Enrollment123.DiabeticShipping_Fee = csvRecordData[74];
                            _Enrollment123.Enrollment_Fee = csvRecordData[75];
                            _Enrollment123.LifeBuyUp10k_Fee = csvRecordData[76];
                            _Enrollment123.LifeBuyUp15k_Fee = csvRecordData[77];
                            _Enrollment123.LifeBuyup25k_Fee = csvRecordData[78];
                            _Enrollment123.NCEEnrollmentFee_Fee = csvRecordData[79];

                            _Enrollment123.Processing_Fee = csvRecordData[80];
                            _Enrollment123.Product_Fee = csvRecordData[81];
                            _Enrollment123.Registration_Fee = csvRecordData[82];
                            _Enrollment123.ShippingFee_Fee = csvRecordData[83];
                            _Enrollment123.Contract_Number = csvRecordData[84];
                            _Enrollment123.Contract_Length = csvRecordData[85];
                            _Enrollment123.Enroller_ID = csvRecordData[86];
                            _Enrollment123.Paytype = csvRecordData[87];
                            _Enrollment123.Refund_Requested = csvRecordData[88];
                            _Enrollment123.Refund_Requested_Date = csvRecordData[89];
                            _Enrollment123.Refund_Provided = csvRecordData[90];
                            _Enrollment123.Refund_Provided_Date = csvRecordData[91];
                            _Enrollment123.Refund_Comment = csvRecordData[92];
                            _Enrollment123.Prospect_Lead = csvRecordData[93];
                            _Enrollment123.Product_Agent_ID = csvRecordData[94];
                            _Enrollment123.Product_Agent_Label = csvRecordData[95];
                            _Enrollment123.Paid = csvRecordData[96];
                            _Enrollment123.Quantity = csvRecordData[97];
                            _Enrollment123.Renewal_Date = csvRecordData[98];
                            _Enrollment123.Sale_Date = csvRecordData[99];
                            _Enrollment123.Next_Billing_Amount = csvRecordData[100];
                            _Enrollment123.Assigned_Agents = csvRecordData[101];
                            _Enrollment123.Last_Payment = csvRecordData[102];
                            _Enrollment123.Source_Detail = csvRecordData[103];
                            _Enrollment123.TPV_Code = csvRecordData[104];
                            _Enrollment123.TPV_Date = csvRecordData[105];
                            _Enrollment123.Do_Not_Call = csvRecordData[106];
                            _Enrollment123.Email_3 = csvRecordData[106];
                            _Enrollment123.Ethnicity = csvRecordData[108];
                            _Enrollment123.Height = csvRecordData[109];
                            _Enrollment123.Weight = csvRecordData[110];
                            _Enrollment123.Disability = csvRecordData[111];
                            _Enrollment123.Status = csvRecordData[112];
                            _Enrollment123.Stage = csvRecordData[113];
                            _Enrollment123.no_Calls = csvRecordData[114];
                            _Enrollment123.Call_Status = csvRecordData[115];
                            _Enrollment123.Best_Call_Time = csvRecordData[116];
                            _Enrollment123.Next_Step = csvRecordData[117];
                            _Enrollment123.Next_Step_Date = csvRecordData[118];
                            _Enrollment123.Source_2 = csvRecordData[119];
                            _Enrollment123.Probability = csvRecordData[120];
                            _Enrollment123.Product_Code = csvRecordData[121];
                            _Enrollment123.Agent_Code = csvRecordData[122];
                            _Enrollment123.Agent_Code_2 = csvRecordData[123];
                            _Enrollment123.Note_Date = csvRecordData[124];
                            _Enrollment123.System_ID = csvRecordData[125];
                            _Enrollment123.Underwriter = csvRecordData[126];
                            _Enrollment123.Category = csvRecordData[127];

                            Enrollment123_List.Add(_Enrollment123);
                            }
                        catch (Exception ex)
                            {

                            }
                        }
                    }


                }



            return Enrollment123_List;
            }

        public List<WPA> GetWPA(string _username, string _password, DateTime fromDate, DateTime toDate, string brokerid)
            {
            List<byte> responseArray = new List<byte>();
            string _csv_file = "";
            List<WPA> WPA_List = new List<WPA>();
            int force_break = 0;
            try
                {



                var client = new CookieAwareWebClient();
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705;)");
                // client.Headers.Add("Connection", "keep-alive");
                //client.Headers.Add("Host", "www.mywellnessagent.com");
                //client.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");


                client.BaseAddress = @"https://www.mywellnessagent.com/manage/index.cfm";
                //client.isfile = false;
                var loginData = new NameValueCollection();
                loginData.Add("username", _username);// "spiewak");
                loginData.Add("password", _password);//"Insurance954");
                loginData.Add("authenticate", "");

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
         | SecurityProtocolType.Tls11
         | SecurityProtocolType.Tls12
         | SecurityProtocolType.Ssl3;

                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                client.UploadValues(@"index.cfm", "Post", loginData);

                string htmlSource = client.DownloadString(@"https://www.mywellnessagent.com/manage/");

                if (htmlSource.Contains("Log Out"))
                    {


                    // htmlSource = client.DownloadString(@"https://www.mywellnessagent.com/manage/users/index.cfm");
                    htmlSource = client.DownloadString(@"https://www.mywellnessagent.com/manage/users/index.cfm?reportId=2625&submit=View");


                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlSource);

                    bool end_pag = false;
                    string page_data = "";
                    bool reload = false;
                    List<_Link> pag = maxpagination_WPA(htmlSource);
                    List<_Link> pag_bkup = new List<_Link>();
                    int pos = 0;
                    WPA_List = WPA_List.Concat(buildList(htmlSource, fromDate, toDate)).ToList();
                    if (pag.Count > 0)
                        {
                        while (!end_pag)
                            {

                            if (pos == 5)
                                {
                                break;
                                }

                            if (pos < pag.Count)
                                {
                                if (pag[pos].pagination == ">>")
                                    {

                                    pag = maxpagination_WPA(page_data);
                                    pos = 0;
                                    reload = true;
                                    }
                                else
                                    {

                                    if (pag_bkup.Where(c => c.value == pag[pos].value).ToList().Count() == 0)
                                        {
                                        pag_bkup.Add(pag[pos]);
                                        var postData = new NameValueCollection();
                                        string pagination = "0";
                                        try
                                            {
                                            pagination = (Convert.ToInt16(pag[pos].value) - 1).ToString();
                                            }
                                        catch (Exception)
                                            {
                                            pagination = pag[pos].value;
                                            }
                                        postData.Add("method", "updateHistoryDisplay");
                                        postData.Add("type", "index");
                                        postData.Add("value", pagination);
                                        postData.Add("token", pag[pos].token);
                                        postData.Add("returnFormat", "json");
                                        postData.Add("brokerid", brokerid);


                                        client.BaseAddress = @"https://www.mywellnessagent.com/includes/cfc/library/filter.cfc";
                                        client.UploadValues(@"filter.cfc", "Post", postData);

                                        page_data = client.DownloadString(@"https://www.mywellnessagent.com/manage/users/index.cfm?token=" + pag[pos].token);

                                        List<WPA> list_new = buildList(page_data, fromDate, toDate);

                                        List<WPA> cancellationList = list_new.Where(c => c.Status != null && (c.Status.Trim().ToLower() == "cancel - accounting flag only")).ToList();
                                        foreach (var item in cancellationList)
                                            {
                                            string page_detail = client.DownloadString(@"https://www.mywellnessagent.com/manage/users/" + item.View);

                                            if (page_detail.Contains("<font class=\"inlineTitle\">Inactive:</font>"))
                                                {
                                                var block = page_detail.Replace("<font class=\"inlineTitle\">Inactive:</font>", "↑");
                                                var _array = block.Split('↑');
                                                block = _array[1].Replace("<strong>", "↑");
                                                _array = block.Split('↑');
                                                block = _array[1].Replace("</strong>", "↑");
                                                _array = block.Split('↑');
                                                DateTime cancel = Convert.ToDateTime(_array[0]);
                                                item.Next_Step = cancel.ToString("MM/dd/yyyy");
                                                item.Status = "Cancel";
                                                }
                                            /*else
                                            {
                                              if (page_detail.Contains("<font class=\"inlineTitle\">Hold:</font"))
                                                  {
                                                  var block = page_detail.Replace("<font class=\"inlineTitle\">Hold:</font>", "↑");
                                                  var _array = block.Split('↑');
                                                  block = _array[1].Replace("<strong>", "↑");
                                                  _array = block.Split('↑');
                                                  block = _array[1].Replace("</strong>", "↑");
                                                  _array = block.Split('↑');
                                                  DateTime cancel = Convert.ToDateTime(_array[0]);
                                                  item.Next_Step = cancel.ToString("MM/dd/yyyy");
                                                  }
                                              }*/


                                            }


                                        WPA_List = WPA_List.Concat(list_new).ToList();

                                        }
                                    }
                                }
                            else
                                {
                                end_pag = true;
                                }


                            if (!reload)
                                {
                                pos++;
                                }
                            else
                                {
                                reload = false;
                                }

                            force_break++;
                            if (force_break == 10000)
                                {
                                break;
                                }

                            }

                        }

                    }


                }
            catch (Exception ex)
                {
                _csv_file = "";
                }





            return WPA_List;
            }


        public List<_Link> maxpagination_WPA(string htmlSource)
            {

            List<_Link> pag = new List<_Link>();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlSource);
            string token = "";
            try
                {
                bool end_pag = false;
                foreach (HtmlNode _td in doc.DocumentNode.SelectNodes("//td"))
                    {

                    if (_td.Attributes.Contains("class"))
                        {
                        if (_td.Attributes["class"].Value == "pageText")
                            {
                            var e = _td.InnerHtml;
                            HtmlAgilityPack.HtmlDocument a_href = new HtmlAgilityPack.HtmlDocument();
                            a_href.LoadHtml(_td.InnerHtml);
                            bool noerror = false;
                            try
                                {
                                var cant_a = a_href.DocumentNode.SelectNodes("//a").Count();
                                noerror = true;
                                }
                            catch (Exception)
                                {

                                noerror = false;
                                }
                            if (noerror)
                                foreach (HtmlNode _a in a_href.DocumentNode.SelectNodes("//a"))
                                    {
                                    if (_a.Attributes.Contains("href"))
                                        {
                                        string[] _split = _a.Attributes["href"].Value.Split('=');

                                        //string pageIndex = _split.Where(c => c.Contains("pageIndex")).FirstOrDefault();



                                        if (_a.InnerHtml != ">>")
                                            {
                                            _Link alink = new _Link();
                                            alink._href = _a.Attributes["href"].Value;
                                            alink.value = _a.InnerHtml;
                                            //alink.pagination = pageIndex.Split('=')[1];

                                            pag.Add(alink);
                                            end_pag = true;

                                            }
                                        else
                                            {
                                            if (_a.InnerHtml == ">>" && pag.Count() > 0)
                                                {
                                                _Link alink = new _Link();
                                                alink._href = ">>";
                                                alink.pagination = ">>";
                                                alink.token = _split[1].Trim();
                                                token = alink.token;
                                                pag.Add(alink);
                                                end_pag = true;

                                                break;
                                                }

                                            }


                                        }

                                    }



                            }
                        }
                    if (end_pag)
                        {
                        break;
                        }
                    }
                }
            catch (Exception ex)
                {

                throw;
                }
            List<_Link> pag2 = new List<_Link>();
            for (int i = 0; i < pag.Count; i++)
                {

                _Link alink = new _Link();
                alink._href = pag[i]._href;
                alink.pagination = pag[i].pagination;
                alink.token = token;
                alink.value = pag[i].value;
                pag2.Add(alink);

                }
            return pag2;
            }
        public List<_Link> maxpagination(string htmlSource)
            {

            List<_Link> pag = new List<_Link>();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlSource);

            try
                {
                bool end_pag = false;
                foreach (HtmlNode _td in doc.DocumentNode.SelectNodes("//td"))
                    {

                    if (_td.Attributes.Contains("class"))
                        {
                        if (_td.Attributes["class"].Value == "pageText")
                            {
                            var e = _td.InnerHtml;
                            HtmlAgilityPack.HtmlDocument a_href = new HtmlAgilityPack.HtmlDocument();
                            a_href.LoadHtml(_td.InnerHtml);
                            bool noerror = false;
                            try
                                {
                                var cant_a = a_href.DocumentNode.SelectNodes("//a").Count();
                                noerror = true;
                                }
                            catch (Exception)
                                {

                                noerror = false;
                                }
                            if (noerror)
                                foreach (HtmlNode _a in a_href.DocumentNode.SelectNodes("//a"))
                                    {
                                    if (_a.Attributes.Contains("href"))
                                        {
                                        string[] _split = _a.Attributes["href"].Value.Split('=');

                                        //string pageIndex = _split.Where(c => c.Contains("pageIndex")).FirstOrDefault();

                                        if (_split.Count() > 1 && _split[0].Contains("token") && _split[1].Trim() != "")

                                            if (_a.InnerHtml != ">>")
                                                {
                                                _Link alink = new _Link();
                                                alink._href = _a.Attributes["href"].Value;
                                                alink.value = _a.InnerHtml;
                                                //alink.pagination = pageIndex.Split('=')[1];
                                                alink.token = _split[1].Trim();
                                                pag.Add(alink);
                                                end_pag = true;

                                                }
                                            else
                                                {
                                                if (_a.InnerHtml == ">>" && pag.Count() > 0)
                                                    {
                                                    _Link alink = new _Link();
                                                    alink._href = ">>";
                                                    alink.pagination = ">>";
                                                    pag.Add(alink);
                                                    end_pag = true;
                                                    break;
                                                    }

                                                }


                                        }

                                    }



                            }
                        }
                    if (end_pag)
                        {
                        break;
                        }
                    }
                }
            catch (Exception ex)
                {

                throw;
                }


            return pag;
            }

        public List<WPA> buildList(string htmlSource, DateTime fromDate, DateTime toDate)
            {

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlSource);

            List<WPA> List_WPA = new List<WPA>();
            foreach (HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
                {

                if (table.Attributes.Contains("class") && table.Attributes["class"].Value == "detail_border")
                    {

                    int row_count = 0;
                    foreach (HtmlNode row in table.SelectNodes("tr"))
                        {

                        if (row_count > 0)
                            {
                            WPA _WPA = new WPA();
                            _WPA.NO = row.SelectNodes("th|td")[0].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Created = row.SelectNodes("th|td")[1].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Agent = row.SelectNodes("th|td")[2].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Id = row.SelectNodes("th|td")[3].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Name = row.SelectNodes("th|td")[4].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Address = row.SelectNodes("th|td")[5].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Email = row.SelectNodes("th|td")[6].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Phone = row.SelectNodes("th|td")[7].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Status = row.SelectNodes("th|td")[8].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Rank = row.SelectNodes("th|td")[9].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Flag = row.SelectNodes("th|td")[10].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Call_Status = row.SelectNodes("th|td")[11].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Next_Step = row.SelectNodes("th|td")[12].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Last_Trans_Amount = row.SelectNodes("th|td")[13].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Last_Trans_Method = row.SelectNodes("th|td")[14].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Last_Trans_Status = row.SelectNodes("th|td")[15].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.Last_Trans_Date = row.SelectNodes("th|td")[16].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            _WPA.View = row.SelectNodes("th|td")[17].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            //_WPA.Batch = row.SelectNodes("th|td")[18].InnerText.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "");
                            //_WPA.View = row.SelectNodes("th|td")[17].InnerHtml.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Replace("&nbsp;", "").Split('>')[0].Replace("<a href=", "").Replace("\"", "");

                            try
                                {

                                DateTime _create = Convert.ToDateTime(_WPA.Created.Split('a')[0]);
                                if (fromDate <= _create && List_WPA.Where(c => c.Id == _WPA.Id).ToList().Count == 0)
                                    {
                                    List_WPA.Add(_WPA);

                                    }

                                }
                            catch (Exception)
                                {
                                if (List_WPA.Where(c => c.Id == _WPA.Id).ToList().Count == 0)
                                    {
                                    _WPA.Created = fromDate.ToString("MM/dd/yyyy");
                                    List_WPA.Add(_WPA);
                                    }
                                throw;
                                }

                            }
                        row_count++;
                        }
                    }
                }

            return List_WPA;
            }






        private string ParseCommasInQuotes(string arg)
            {

            bool foundEndQuote = false;
            bool foundStartQuote = false;
            StringBuilder output = new StringBuilder();

            //44 = comma
            //34 = double quote

            foreach (char element in arg)
                {
                if (foundEndQuote)
                    {
                    foundStartQuote = false;
                    foundEndQuote = false;
                    }

                if (element.Equals((Char)34) & (!foundEndQuote) & foundStartQuote)
                    {
                    foundEndQuote = true;
                    continue;
                    }

                if (element.Equals((Char)34) & !foundStartQuote)
                    {
                    foundStartQuote = true;
                    continue;
                    }

                if ((element.Equals((Char)44) & foundStartQuote))
                    {
                    //skip the comma...its between double quotes
                    }
                else
                    {
                    output.Append(element);
                    }
                }
            return output.ToString();
            }





        }
    }