
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SF_Lib.com.salesforce.na5;

namespace SF_Lib
    {
    public class Salerforce_Interface
        {
        SforceService SfdcBinding;
        LoginResult CurrentLoginResult;
        public Salerforce_Interface()
            {
            SalesForceConecction();
            }

        private void SalesForceConecction()
            {
            string userName = string.Empty;
            string password = string.Empty;
            userName = "external@hbcinsure.com";//System.Configuration.ConfigurationSettings.AppSettings["sfdcUserName"];
            password = "Insurance954KsxCKalIu6gvpAIsSQ8lrOcT7";//System.Configuration.ConfigurationSettings.AppSettings["sfdcPassword"] + System.Configuration.ConfigurationSettings.AppSettings["sfdcToken"];

            SfdcBinding = null;
            CurrentLoginResult = null;
            SfdcBinding = new SforceService();

            try
                {
                CurrentLoginResult = SfdcBinding.login(userName, password);
                }
            catch (System.Web.Services.Protocols.SoapException ex)
                {
                // This is likley to be caused by bad username or password
                SfdcBinding = null;
                throw (ex);
                }
            catch (Exception ex)
                {
                // This is something else, probably comminication
                SfdcBinding = null;
                throw (ex);
                }
            }

        public void closeService()
            {
            SfdcBinding.logout();

            }
        public List<Payment_Period__c> GetPaymentCalendar()
            {
            List<Payment_Period__c> PaymentCalendar = new List<Payment_Period__c>();

            string soql = "select id,Start_Date__c,Close_Date__c,Pay_Period__c,Close_Period__c,Payment_Date__c,Product_Profile__c  from Payment_Period__c where Close_Date__c<=" + DateTime.Today.ToString("yyyy-MM-dd") + " and   Close_Period__c=false  order by Payment_Date__c";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned


                for (int i = 0; i < queryResult.size; i++)
                    {
                    PaymentCalendar.Add((Payment_Period__c)queryResult.records[i]);


                    }
                }

            return PaymentCalendar;
            }

        public List<Payment_Period__c> GetAllPaymentCalendar()
            {
            List<Payment_Period__c> PaymentCalendar = new List<Payment_Period__c>();

            string soql = "select id,Start_Date__c,Close_Date__c,Pay_Period__c,Close_Period__c,Payment_Date__c,Product_Profile__c  from Payment_Period__c where    Close_Period__c=false  order by Payment_Date__c";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned


                for (int i = 0; i < queryResult.size; i++)
                    {
                    PaymentCalendar.Add((Payment_Period__c)queryResult.records[i]);


                    }
                }

            return PaymentCalendar;
            }

        public List<Payment_Period__c> GetAllPaymentCalendarByProductProfile(string _productProfile)
            {
            List<Payment_Period__c> PaymentCalendar = new List<Payment_Period__c>();


            string soql = "select id,Start_Date__c,Close_Date__c,Pay_Period__c,Close_Period__c,Payment_Date__c,Product_Profile__c  from Payment_Period__c where    Close_Period__c=false and  Product_Profile__c = '" + _productProfile +"' order by Payment_Date__c";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned


                for (int i = 0; i < queryResult.size; i++)
                    {
                    PaymentCalendar.Add((Payment_Period__c)queryResult.records[i]);


                    }
                }

            return PaymentCalendar;
            }
        public List<OpportunityLineItem> GetOpportunityLineBy_Policy_Number(List<String> policy_number_list, string agent_id, string companny, DateTime _to)
            {


            List<OpportunityLineItem> oppList = new List<OpportunityLineItem>();
            if (policy_number_list.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)policy_number_list.Count / (double)500);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(policy_number_list, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<OpportunityLineItem> temp_opp = GetOpportunityLineBy_Policy_Number(temp, agent_id, companny, _to);
                    if (temp_opp.Count > 0)
                        oppList = oppList.Concat(temp_opp).ToList();
                    }



                }
            else
                {


                string agent_filter = " ";
                if (agent_id != "All")
                    {
                    agent_filter = "  And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                    }

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;
                string temp = policy_number_list.Where(c =>c == "SXH003418").FirstOrDefault();
                if (temp!= null )
                    {
                    temp = "this";
                    }
                string PolicyFormated = formatString(policy_number_list.ToList());

                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string SOQL = "Select id ,Status__c, Policy_Number__c ,Enrollment_Date__c,Effective_Date__c,Opportunityid,Product2.id, Product2.Name, Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C , Product2.Plan_Type__c, Commission_Payable_Number__c ,	Commission_Payable__c ,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Opportunity.Fronter_s_Name__c,  Opportunity.Agent_ID_Lookup__r.Ending_Date__c from opportunitylineitem where   Opportunity.Agent_ID_Lookup__r.product_profile__c='" + companny + "' and Policy_Number__c in (" + PolicyFormated + ") and Status__c = 'Processed'  and Commission_Payable__c > 0   AND Enrollment_Date__c <= " + _to.ToString("yyyy-MM-dd") + " " + agent_filter;

               // and(Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )

                queryResult = SfdcBinding.query(SOQL);
                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            oppList.Add((OpportunityLineItem)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }

                }


            
            return oppList;
            }


        public List<OpportunityLineItem> GetOpportunityLineBy_Policy_Number_Enroll(List<String> policy_number_list, DateTime from, DateTime to, string agent_id,  string companny)
            {
            List<OpportunityLineItem> oppList = new List<OpportunityLineItem>();
            if (policy_number_list.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)policy_number_list.Count / (double)1000);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(policy_number_list, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<OpportunityLineItem> temp_opp = GetOpportunityLineBy_Policy_Number_Enroll(temp, from, to, agent_id, companny);
                    if (temp_opp.Count > 0)
                        oppList = oppList.Concat(temp_opp).ToList();
                    }



                }
            else
                {

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;

                string PolicyFormated = formatString(policy_number_list.ToList());


                string agent_filter = " ";
                if (agent_id != "All")
                    {
                    agent_filter = "  And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                    }


                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string SOQL = "Select id , Policy_Number__c ,Enrollment_Date__c,Effective_Date__c,Opportunityid,Product2.id, Product2.Name, Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C , Product2.Plan_Type__c, Commission_Payable__c,Commission_Payable_Number__c    from opportunitylineitem where  Opportunity.Agent_ID_Lookup__r.product_profile__c='" + companny + "' and Policy_Number__c in (" + PolicyFormated + ")  and Enrollment_Date__c >=" + from.ToUniversalTime().ToString("yyyy-MM-dd") + "and Enrollment_Date__c <=" + to.ToUniversalTime().ToString("yyyy-MM-dd") + agent_filter;
                //(Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )
                queryResult = SfdcBinding.query(SOQL);
                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            oppList.Add((OpportunityLineItem)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }

                }



            return oppList;
            }




        public List<Opportunity> GetOpportunity_by_id(List<String> id_list, string product_Profile)
            {
            List<Opportunity> oppList = new List<Opportunity>();
            if (id_list.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)id_list.Count / (double)1000);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(id_list, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<Opportunity> temp_opp = GetOpportunity_by_id(temp, product_Profile);
                    if (temp_opp.Count > 0)
                        oppList = oppList.Concat(temp_opp).ToList();
                    }
                Opportunity opportunity = new Opportunity();
                
                 

                
                }
            else
                {

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;



                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string opp_id_format = formatString(id_list);
                string SOQL = "select id, name, Agent_ID_Lookup__c,Fronter_s_Name__c,GI_Product__c,GI_Product__r.id, Commission_Payable__c,Agent_ID_Lookup__r.Status__c, Agent_ID_Lookup__r.Ending_Date__c , Agent_ID_Lookup__r.Product_Profile__c from opportunity where Agent_ID_Lookup__r.Product_Profile__c ='" + product_Profile + "' and id in (" + opp_id_format + ")";
                //(GI_Product__r.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or GI_Product__r.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )";

                queryResult = SfdcBinding.query(SOQL);
                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            oppList.Add((Opportunity)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }

                }



            return oppList;
            }

        public String formatString(List<string> PolicyNumberList)
            {
            string policyformated = "";
            for (int i = 0; i < PolicyNumberList.Count; i++)
                {
                if (i + 1 == PolicyNumberList.Count)
                    {
                    policyformated += "'" + PolicyNumberList[i] + "'";
                    }
                else
                    {
                    policyformated += "'" + PolicyNumberList[i] + "',";
                    }

                }
            return policyformated;

            }

        public List<White_List_Resource__c> getChargeBackCancellationPeriod()
            {
            List<White_List_Resource__c> white_List = new List<White_List_Resource__c>();

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;



            // PolicyNumberList = PolicyNumberList.ToList().Distinct();
            string SOQL = "Select Resource_Type__c,value__c from White_List_Resource__c where  Resource_Type__c like '%Chargeback Period%'";

            queryResult = SfdcBinding.query(SOQL);
            if (queryResult.size > 0)
                {
                for (int i = 0; i < queryResult.size; i++)
                    {
                    white_List.Add((White_List_Resource__c)queryResult.records[i]);
                    }
                }

            return white_List;

            }

        public List<White_List_Resource__c> getPayHoldPeriod()
            {
            List<White_List_Resource__c> white_List = new List<White_List_Resource__c>();

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;



            // PolicyNumberList = PolicyNumberList.ToList().Distinct();
            string SOQL = "Select Resource_Type__c,value__c from White_List_Resource__c where  Resource_Type__c like '%Pay Hold%'";

            queryResult = SfdcBinding.query(SOQL);
            if (queryResult.size > 0)
                {
                for (int i = 0; i < queryResult.size; i++)
                    {
                    white_List.Add((White_List_Resource__c)queryResult.records[i]);
                    }
                }

            return white_List;

            }

        public List<Payment__c> getpaymentCancelbyOppLine(List<string> oppLines_id)
            {
            List<Payment__c> opplines = new List<Payment__c>();

            if (oppLines_id.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)oppLines_id.Count / (double)1000);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(oppLines_id, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<Payment__c> temp_opp = getpaymentCancelbyOppLine(temp);
                    if (temp_opp.Count > 0)
                        opplines = opplines.Concat(temp_opp).ToList();
                    }
                }
            else
                {

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;



                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string opp_id_format = formatString(oppLines_id);
                string SOQL = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Verify__c ,Cancel_Date__c from Payment__c where OpportunityLineItem_id__c in (" + opp_id_format + ") and  Payment_Type__c in ( 'Chargeback', 'Terminated','Commission')";

                queryResult = SfdcBinding.query(SOQL);


                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            opplines.Add((Payment__c)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }




                }



            return opplines;


            }

        public List<Payment__c> getpaymentbyPolicyNumber(List<string> _policyNumber)
            {
            List<Payment__c> opplines = new List<Payment__c>();

            if (_policyNumber.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)_policyNumber.Count / (double)1000);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(_policyNumber, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<Payment__c> temp_opp = getpaymentbyPolicyNumber(temp);
                    if (temp_opp.Count > 0)
                        opplines = opplines.Concat(temp_opp).ToList();
                    }
                }
            else
                {

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;



                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string _policyNumber_format = formatString(_policyNumber);
                string SOQL = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Verify__c ,Cancel_Date__c from Payment__c where Policy_Number__c in (" + _policyNumber_format + ") and  Payment_Type__c in ( 'Chargeback', 'Terminated','Commission')";

                queryResult = SfdcBinding.query(SOQL);

                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            opplines.Add((Payment__c)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }


                }



            return opplines;


            }




        public List<Payment__c> getpaymentbyPolicyNumber(List<string> oppLines_PolicyNumber, List<string> trans_Type_list)
            {
            List<Payment__c> opplines = new List<Payment__c>();

            if (oppLines_PolicyNumber.Count > 1000)
                {

                int cant = (int)Math.Ceiling((double)oppLines_PolicyNumber.Count / (double)1000);
                IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(oppLines_PolicyNumber, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<string> temp = matrix.ElementAt(i).ToList();

                    List<Payment__c> temp_opp = getpaymentbyPolicyNumber(temp, trans_Type_list);
                    if (temp_opp.Count > 0)
                        opplines = opplines.Concat(temp_opp).ToList();
                    }
                }
            else
                {

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;



                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string opp_id_format = formatString(oppLines_PolicyNumber);

                string tras_type_str = formatString(trans_Type_list);
                string SOQL = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Verify__c ,Cancel_Date__c from Payment__c where Policy_Number__c in (" + opp_id_format + ")  and Payment_Type__c in (" + tras_type_str + ") ";

                queryResult = SfdcBinding.query(SOQL);

                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            opplines.Add((Payment__c)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }





                }



            return opplines;


            }


        public List<Payment__c> getpaymentbyPayday(string adate, string agent_id, string product_profile)
            {
            List<Payment__c> opplines = new List<Payment__c>();



            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            string agent_filter="";
            if (agent_id!="All")
                {
                agent_filter = "And  Agent__c='" + agent_id + "'";
                }

            // PolicyNumberList = PolicyNumberList.ToList().Distinct();

            string SOQL = "select id, Agent__c,	Agent__r.Status__c,Agent__r.Last_Name__c,Agent__r.First_Name__c,Agent__r.Username__c,Agent__r.Ending_Date__c,Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Verify__c ,Agent__r.product_profile__c , Cancel_Date__c from Payment__c where Payment_Date__c = " + adate + " and  Payment_Type__c in ( 'Chargeback', 'Terminated','Commission')" + agent_filter + "and Agent__r.product_profile__c = '" + product_profile + "'";

            queryResult = SfdcBinding.query(SOQL);

            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        opplines.Add((Payment__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }






            return opplines;


            }


        public List<Payment__c> getPayment()
            {
            List<Payment__c> result = new List<Payment__c>();
            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;



            // PolicyNumberList = PolicyNumberList.ToList().Distinct();

            string SOQL = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c ,Cancel_Date__c from Payment__c  order by Agent__c, Payment_Date__c";

            queryResult = SfdcBinding.query(SOQL);

            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        result.Add((Payment__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }





            return result;
            }

        public string getCompannyProfile()
            {
            string white_List = "";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;



            // PolicyNumberList = PolicyNumberList.ToList().Distinct();
            string SOQL = "Select Resource_Type__c,value__c from White_List_Resource__c where  Resource_Type__c = 'Companny Profile'";

            queryResult = SfdcBinding.query(SOQL);
            if (queryResult.size > 0)
                {
                white_List = ((White_List_Resource__c)queryResult.records[0]).Value__c.ToString();

                }

            return white_List;

            }


        public List<Payment_Period__c> GetPaymentCalendarbyCompanny(string product_Profile)
            {
            List<Payment_Period__c> PaymentCalendar = new List<Payment_Period__c>();

            string soql = "select id,Start_Date__c,Close_Date__c,Pay_Period__c,Close_Period__c,Payment_Date__c,Product_Profile__c  from Payment_Period__c where Product_Profile__c='" + product_Profile + "'  order by Payment_Date__c ";//Close_Date__c<=" + DateTime.Today.ToString("yyyy-MM-dd") + " and and    Close_Period__c=false

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned


                for (int i = 0; i < queryResult.size; i++)
                    {
                    PaymentCalendar.Add((Payment_Period__c)queryResult.records[i]);


                    }
                }

            return PaymentCalendar;
            }


        public List<SaveResult> insertPayment(List<Payment__c> newEntrys)
            {
            
            List<SaveResult> ResultList = new List<SaveResult>();

            if (newEntrys.Count > 100)
                {

                int cant = (int)Math.Ceiling((double)newEntrys.Count / (double)100);
                IEnumerable<IEnumerable<Payment__c>> matrix = LinqExtensions.Split(newEntrys, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<Payment__c> temp = matrix.ElementAt(i).ToList();

                    List<SaveResult> temp_opp = insertPayment(temp);
                    if (temp_opp.Count > 0)
                        ResultList = ResultList.Concat(temp_opp).ToList();
                    }

                }
            else
                {
                if (newEntrys.Count > 0)
                    {
                    SfdcBinding.Url = CurrentLoginResult.serverUrl;

                    //Create a new session header object and set the session id to that returned by the login
                    SfdcBinding.SessionHeaderValue = new SessionHeader();
                    SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;


                    sObject[] mycusotm = newEntrys.ToArray();

                    SaveResult[] saveResults = SfdcBinding.create(mycusotm);

                    for (int i = 0; i < saveResults.Count(); i++)
                        {
                        ResultList.Add(saveResults[i]);
                        }

                    }
                }

            return ResultList;



            }


        public List<SaveResult> UpdatePayment(List<Payment__c> newEntrys)
            {

            List<SaveResult> ResultList = new List<SaveResult>();

            if (newEntrys.Count > 100)
                {

                int cant = (int)Math.Ceiling((double)newEntrys.Count / (double)100);
                IEnumerable<IEnumerable<Payment__c>> matrix = LinqExtensions.Split(newEntrys, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<Payment__c> temp = matrix.ElementAt(i).ToList();

                    List<SaveResult> temp_opp = UpdatePayment(temp);
                    if (temp_opp.Count > 0)
                        ResultList = ResultList.Concat(temp_opp).ToList();
                    }

                }
            else
                {
                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

                sObject[] mycusotm = newEntrys.ToArray();

                SaveResult[] saveResults = SfdcBinding.update(mycusotm);

                for (int i = 0; i < saveResults.Count(); i++)
                    {
                    ResultList.Add(saveResults[i]);
                    }

                }

            return ResultList;

            }

        public List<Agent__c> getAgentbyProfile(string _profile, bool available)
            {
            string agents = " Force_com_Enabled__c = false";
            if (available)
                {
                agents = "Force_com_Enabled__c = true";
                }
          
            List<Agent__c> agentList = new List<Agent__c>();

            string soql = "Select id,First_Name__c,Last_Name__c ,Status__c,	Username__c ,product_profile__c	,Ending_Date__c,Force_com_Enabled__c from Agent__c where product_profile__c='" + _profile + "' and " + agents;
            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        agentList.Add((Agent__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }






            return agentList;
            }


        public List<Agent__c> getAgentbyid(List<string> list_id)
            {
           

            List<Agent__c> agentList = new List<Agent__c>();
            string agent_id_format = formatString(list_id);
            string soql = "Select id,First_Name__c,Last_Name__c ,Status__c,	Username__c ,product_profile__c	,Ending_Date__c,Force_com_Enabled__c from Agent__c where id in(" + agent_id_format + " ) ";
            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        agentList.Add((Agent__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }






            return agentList;
            }




        public List<OpportunityLineItem> Commissios_for_Pay(DateTime _from, DateTime _to, string agent_id, string companny)
            {
            string agent_filter = " ";
            if (agent_id != "All")
                {
                agent_filter = "And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                }

            List<OpportunityLineItem> lines = new List<OpportunityLineItem>(); ;
            string soql = "SELECT Id,Status__c, Product2.Name, Name, OpportunityID,Opportunity.name ,Product_Family__c, Plan_Type__c, Commission_Payable_Number__c, Monthly_Rate__c, Commission_Payable__c,Enrollment_Date__c,Effective_Date__c,Policy_Number__c,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Product2.Additional_Product_1__c,Product2.Additional_Product_2__c,Product2.Additional_Product_3__c,Product2.Additional_Product_4__c,Product2Id FROM OpportunityLineItem WHERE  Opportunity.Agent_ID_Lookup__r.product_profile__c='" + companny + "' and  Enrollment_Date__c >= " + _from.ToString("yyyy-MM-dd") + " AND Enrollment_Date__c <= " + _to.ToString("yyyy-MM-dd") + " AND Policy_Number__c != ''  and Commission_Payable_Number__c > 0" + agent_filter + "    ORDER BY OpportunityID, Plan_Type__c";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        lines.Add((OpportunityLineItem)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }
          
            return lines;

            }





        public List<DeleteResult> deletePayment(DateTime Payroll_date)
            {


            List<string> ids_list = new List<string>();
            List<DeleteResult> _DeleteResult = new List<DeleteResult>();

            string result = "Empty";
            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

            string soql = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c  from Payment__c  where Close_Period__c =false and Payment_Date__c= " + Payroll_date.ToString("yyyy-MM-dd");

            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        ids_list.Add(((Payment__c)queryResult.records[i]).Id);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }





            if (ids_list.Count > 0)
                {

                DeleteResult[] deleteResults = SfdcBinding.delete(ids_list.ToArray());

                DeleteResult deleteResult = deleteResults[0];

                _DeleteResult = deleteResults.ToList();

                }
            return _DeleteResult;

            }


        public List<DeleteResult> deletePayment_byId(string Payment_Id)
            {


            List<string> ids_list = new List<string>();
            List<DeleteResult> _DeleteResult = new List<DeleteResult>();

            string result = "Empty";
            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

        
            if (Payment_Id!= "" && Payment_Id!= null)
                {

                ids_list.Add(Payment_Id);
                DeleteResult[] deleteResults = SfdcBinding.delete(ids_list.ToArray());

                DeleteResult deleteResult = deleteResults[0];

                _DeleteResult = deleteResults.ToList();

                }

            return _DeleteResult;

            }



        public List<Payment__c> Payment_for_Print(DateTime Payroll_date, string _productProfile, string agent_id)
            {
            List<Payment__c> _Payment__c = new List<Payment__c>();
           
            string agent_filter = " ";
            if (agent_id != "All" && agent_id != "")
                {
                agent_filter = "  And Agent__c='" + agent_id + "'";
                }

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

            string soql = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Agent__r.First_Name__c,  Agent__r.Last_Name__c  , Cancel_Date__c,Agent__r.Payment_company__c, CreatedDate, Agent__r.Time_Shift__c,Agent__r.product_profile__c  from Payment__c  where  Payment_Date__c= " + Payroll_date.ToString("yyyy-MM-dd") + " and Agent__r.product_profile__c= '" + _productProfile + "'" + agent_filter;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        ((Payment__c)queryResult.records[i]).Close_Period__c = true;
                        ((Payment__c)queryResult.records[i]).Close_Period__cSpecified = true;
                        _Payment__c.Add((Payment__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }
                }
            return _Payment__c;



            }

        public List<Payment__c> Payment_for_PrintGetALLfromdate(DateTime Payroll_date, string _productProfile, string agent_id)
            {
            List<Payment__c> _Payment__c = new List<Payment__c>();

            string agent_filter = " ";
            if (agent_id != "All" && agent_id != "")
                {
                agent_filter = "  And Agent__c='" + agent_id + "'";
                }

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

            string soql = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c,Agent__r.First_Name__c,  Agent__r.Last_Name__c  , Cancel_Date__c,Agent__r.Payment_company__c  from Payment__c  where  Payment_Date__c >= " + Payroll_date.ToString("yyyy-MM-dd") + " and Agent__r.product_profile__c= '" + _productProfile + "'" + agent_filter;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        ((Payment__c)queryResult.records[i]).Close_Period__c = true;
                        ((Payment__c)queryResult.records[i]).Close_Period__cSpecified = true;
                        _Payment__c.Add((Payment__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }
                }
            return _Payment__c;



            }

        public List<OpportunityLineItem> GetOpportunityLineBy_Policy_NumberforPrint(List<String> policy_number_list, string agent_id, string  companny)
           {


           List<OpportunityLineItem> oppList = new List<OpportunityLineItem>();
           if (policy_number_list.Count > 1000)
               {

               int cant = (int)Math.Ceiling((double)policy_number_list.Count / (double)500);
               IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(policy_number_list, cant + 1);

               for (int i = 0; i < matrix.Count(); i++)
                   {
                   List<string> temp = matrix.ElementAt(i).ToList();

                   List<OpportunityLineItem> temp_opp = GetOpportunityLineBy_Policy_NumberforPrint(temp, agent_id, companny);
                   if (temp_opp.Count > 0)
                       oppList = oppList.Concat(temp_opp).ToList();
                   }



               }
           else
               {


               string agent_filter = " ";
               if (agent_id != "All" && agent_id!="")
                   {
                   agent_filter = "  And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                   }

               SfdcBinding.Url = CurrentLoginResult.serverUrl;

               //Create a new session header object and set the session id to that returned by the login
               SfdcBinding.SessionHeaderValue = new SessionHeader();
               SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
               QueryResult queryResult = null;
   /*         string temp = policy_number_list.Where(c => c == "SXH003418").FirstOrDefault();
               if (temp != null)
                   {
                   temp = "this";
                   }*/
               string PolicyFormated = formatString(policy_number_list.ToList());

               // PolicyNumberList = PolicyNumberList.ToList().Distinct();
               string SOQL = "Select id ,Status__c, Policy_Number__c ,Enrollment_Date__c,Effective_Date__c,Opportunityid,Product2.id, Product2.Name, Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C , Product2.Plan_Type__c, Commission_Payable_Number__c ,	Commission_Payable__c ,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Opportunity.Fronter_s_Name__c,  Opportunity.Agent_ID_Lookup__r.Ending_Date__c,Product2.Sales_Board_Category__c,Product2.eApp_Display_Name__c from opportunitylineitem where  Opportunity.Agent_ID_Lookup__r.product_profile__c='" + companny + "' and Policy_Number__c in (" + PolicyFormated + ")  " + agent_filter;

               // and(Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )

               queryResult = SfdcBinding.query(SOQL);
               int max_value = 0;
               int Increment_value = 0;
               if (queryResult.size > 0)
                   {
                   max_value = queryResult.size;
                   //put some code in here to handle the records being returned
                   bool _done = false;
                   while (!_done)
                       {
                       for (int i = 0; i < queryResult.records.Count(); i++)
                           {
                           Increment_value++;
                           oppList.Add((OpportunityLineItem)queryResult.records[i]);

                           }
                       if (max_value == Increment_value)
                           {
                           _done = true;
                           }
                       else
                           {
                           queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                           }
                       }




                   }

               }



           return oppList;
           }




       public List<OpportunityLineItem> GetOpportunityLineBy_OppLineId_forPrint(List<String> OppLineId_list, string agent_id, string companny)
           {

        
           List<OpportunityLineItem> oppList = new List<OpportunityLineItem>();
           if (OppLineId_list.Count > 700)
               {

               int cant = (int)Math.Ceiling((double)OppLineId_list.Count / (double)500);
               IEnumerable<IEnumerable<string>> matrix = LinqExtensions.Split(OppLineId_list, cant + 1);

               for (int i = 0; i < matrix.Count(); i++)
                   {
                   List<string> temp = matrix.ElementAt(i).ToList();

                   List<OpportunityLineItem> temp_opp = GetOpportunityLineBy_OppLineId_forPrint(temp, agent_id, companny);
                   if (temp_opp.Count > 0)
                       oppList = oppList.Concat(temp_opp).ToList();
                   }



               }
           else
               {


               string agent_filter = " ";
               if (agent_id != "All" && agent_id != "")
                   {
                   agent_filter = "  And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                   }

               SfdcBinding.Url = CurrentLoginResult.serverUrl;

               //Create a new session header object and set the session id to that returned by the login
               SfdcBinding.SessionHeaderValue = new SessionHeader();
               SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
               QueryResult queryResult = null;
               /*         string temp = policy_number_list.Where(c => c == "SXH003418").FirstOrDefault();
                           if (temp != null)
                               {
                               temp = "this";
                               }*/
               string OppLineId = formatString(OppLineId_list.ToList());

               // PolicyNumberList = PolicyNumberList.ToList().Distinct();
               string SOQL = "Select id ,Status__c, Policy_Number__c ,Enrollment_Date__c,Effective_Date__c,Opportunityid,Product2.id, Product2.Name, Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C , Product2.Plan_Type__c, Commission_Payable_Number__c ,	Commission_Payable__c ,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Opportunity.Fronter_s_Name__c,  Opportunity.Agent_ID_Lookup__r.Ending_Date__c,Product2.Sales_Board_Category__c,Product2.eApp_Display_Name__c from opportunitylineitem where  Opportunity.Agent_ID_Lookup__r.product_profile__c='" + companny + "' and id in (" + OppLineId + ")  " + agent_filter;

               // and(Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )

               queryResult = SfdcBinding.query(SOQL);
               int max_value = 0;
               int Increment_value = 0;
               if (queryResult.size > 0)
                   {
                   max_value = queryResult.size;
                   //put some code in here to handle the records being returned
                   bool _done = false;
                   while (!_done)
                       {
                       for (int i = 0; i < queryResult.records.Count(); i++)
                           {
                           Increment_value++;
                           oppList.Add((OpportunityLineItem)queryResult.records[i]);

                           }
                       if (max_value == Increment_value)
                           {
                           _done = true;
                           }
                       else
                           {
                           queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                           }
                       }




                   }

               }



           return oppList;
           }





       public List<SalesBoard__c> get_salesborad(string companny, DateTime from, DateTime to, string agent_id)
           {
           List<SalesBoard__c> list_aggregate = new List<SalesBoard__c>();
            
            string agent_filter = "  ";
            if (agent_id != "All" && agent_id != "")
                {
                agent_filter = "  And Agent__c='" + agent_id + "'";
                }

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;


            string soql = "Select Agent__c, sum( Amount__c) amount from SalesBoard__c where Date__c >= " + from.ToString("yyyy-MM-dd") + " and Date__c <=" + to.ToString("yyyy-MM-dd") + " and  Company__c ='" + companny + "'" + agent_filter + " group by Agent__c";

            
           QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        AggregateResult _AggregateResult = (AggregateResult)queryResult.records[i];

                        double value = 0;
                        try
                            {
                            value = Convert.ToDouble(_AggregateResult.Any[1].InnerText);
                            }
                        catch (Exception)
                            {
                            
                            value = 0;
                            }
                        SalesBoard__c _SalesBoard__c = new SalesBoard__c() { Agent__c = _AggregateResult.Any[0].InnerText, Amount__c = value };


                        list_aggregate.Add(_SalesBoard__c);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }

                }

            return list_aggregate;
           }


       public List<Payment__c> Payment_for_Print_agregation(DateTime Payroll_date, string _productProfile, string agent_id)
           {
           List<Payment__c> _Payment__c = new List<Payment__c>();

           string agent_filter = " ";
           if (agent_id != "All" && agent_id != "")
               {
               agent_filter = "  And Agent__c='" + agent_id + "'";
               }

           SfdcBinding.Url = CurrentLoginResult.serverUrl;

           //Create a new session header object and set the session id to that returned by the login
           SfdcBinding.SessionHeaderValue = new SessionHeader();
           SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

           string soql = "select  Agent__c,Payment_Type__c,sum(Payment_Value__c)    from Payment__c  where   Payment_Date__c= " + Payroll_date.ToString("yyyy-MM-dd") + " and Agent__r.product_profile__c= '" + _productProfile + "'" + agent_filter + " group by Agent__c, Payment_Type__c ";
           QueryResult queryResult = null;
           queryResult = SfdcBinding.query(soql);
           int max_value = 0;
           int Increment_value = 0;
           if (queryResult.size > 0)
               {
               //put some code in here to handle the records being returned
               max_value = queryResult.size;
               //put some code in here to handle the records being returned
               bool _done = false;
               while (!_done)
                   {
                   for (int i = 0; i < queryResult.records.Count(); i++)
                       {
                       Increment_value++;
                       AggregateResult _AggregateResult = (AggregateResult)queryResult.records[i];

                       double value = 0;
                       try
                           {
                           value = Convert.ToDouble(_AggregateResult.Any[2].InnerText);
                           }
                       catch (Exception)
                           {

                           value = 0;
                           }

                       Payment__c _item = new Payment__c() { Agent__c = _AggregateResult.Any[0].InnerText, Payment_Value__c = value, Payment_Type__c = _AggregateResult.Any[1].InnerText };
                       _Payment__c.Add(_item);

                       }
                   if (max_value == Increment_value)
                       {
                       _done = true;
                       }
                   else
                       {
                       queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                       }
                   }
               }
           return _Payment__c;



           }


        public List<Payment__c> closePerirod(DateTime Payroll_date, string _productProfile)
            {
            SfdcBinding.Url = CurrentLoginResult.serverUrl;
            List<Payment__c> _Payment__c = new List<Payment__c>();
            List<SaveResult> _SaveResult = new List<SaveResult>();
            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;

            string soql = "select id, Agent__c,	Close_Period__c,Description__c,OpportunityLineItem_id__c,Payment_Date__c,Payment_Type__c,Payment_Value__c,Policy_Number__c , Cancel_Date__c from Payment__c  where Close_Period__c =false and Payment_Date__c= " + Payroll_date.ToString("yyyy-MM-dd") + " and Agent__r.product_profile__c= '" + _productProfile + "'";
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        ((Payment__c)queryResult.records[i]).Close_Period__c = true;
                        ((Payment__c)queryResult.records[i]).Close_Period__cSpecified = true;
                        _Payment__c.Add((Payment__c)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }










                _SaveResult = UpdatePayment(_Payment__c);


                if (_SaveResult.Where(c => c.errors != null).ToList().Count == 0)
                    {


                    soql = " select id,Start_Date__c,Close_Date__c,Pay_Period__c,Close_Period__c,Payment_Date__c,Product_Profile__c  from Payment_Period__c where    Close_Period__c=false and  Product_Profile__c = '" + _productProfile + "'  and Payment_Date__c= " + Payroll_date.ToString("yyyy-MM-dd");
                    queryResult = null;
                    queryResult = SfdcBinding.query(soql);
                    List<Payment_Period__c> Payment_Period__c = new List<Payment_Period__c>();
                    if (queryResult.size > 0)
                        {
                        //put some code in here to handle the records being returned


                        for (int i = 0; i < queryResult.size; i++)
                            {
                            ((Payment_Period__c)queryResult.records[i]).Close_Period__c = true;
                            ((Payment_Period__c)queryResult.records[i]).Close_Period__cSpecified = true;
                            Payment_Period__c.Add((Payment_Period__c)queryResult.records[i]);
                            }

                        sObject[] mycusotm = Payment_Period__c.ToArray();

                        SaveResult[] saveResults = SfdcBinding.update(mycusotm);

                        for (int i = 0; i < saveResults.Count(); i++)
                            {
                            _SaveResult.Add(saveResults[i]);
                            }


                        }
                    }

                }

            if (_SaveResult.Count > 0 && _Payment__c.Count > 0)
                {
                for (int x = 0; x < _Payment__c.Count; x++)
                    {

                    var results = _SaveResult.Where(c => c.id == _Payment__c[x].Id).FirstOrDefault();
                    if (results != null)
                        {
                        if (results.success)
                            _Payment__c[x].Description__c = "Success";
                        }
                    else
                        {
                        _Payment__c[x].Description__c = "Error";

                        }
                    }

               
                }
            return _Payment__c;
            }


        public int newSales()
            {

            string soql = "SELECT eAppForm__c, CreatedDate  from opportunity where CreatedDate > 2015-05-15T00:00:00Z   order by CreatedDate ";
            List<string> ideapp = new List<string>();
            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            if (queryResult.size > 0)
                {
                //put some code in here to handle the records being returned


                for (int i = 0; i < queryResult.size; i++)
                    {
                    ideapp.Add(((Opportunity)queryResult.records[i]).eAppForm__c);


                    }
                }




            soql = "SELECT lead_id__C from eAppForm__c where lead_id__C != null AND   id in (" + formatString(ideapp.ToList()) + ")";
           
            queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int count = 0;
            if (queryResult.size > 0)
                {
                ideapp = new List<string>();
                for (int i = 0; i < queryResult.size; i++)
                    {
                    ideapp.Add(((eAppForm__c)queryResult.records[i]).Lead_Id__c);
                    count++;

                    }
                }

            return count;

            }



        /***AUDIT**/
        public List<OpportunityLineItem> AUDIT_Commissios_for_Pay(DateTime _from, DateTime _to, string agent_id)
            {

            string agent_filter = " ";
            if (agent_id != "All" && agent_id != "")
                {
                agent_filter = "  And  Opportunity.Agent_ID_Lookup__c='" + agent_id + "'";
                }


            List<OpportunityLineItem> lines = new List<OpportunityLineItem>(); ;
            string soql = "SELECT Id,Status__c, Product2.Name, Name, OpportunityID,Opportunity.name ,Product_Family__c, Plan_Type__c, Commission_Payable_Number__c, Monthly_Rate__c, Commission_Payable__c,Enrollment_Date__c,Policy_Number__c,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Product2.Additional_Product_1__c,Product2.Additional_Product_2__c,Product2.Additional_Product_3__c,Product2.Additional_Product_4__c,Product2Id FROM OpportunityLineItem WHERE  Enrollment_Date__c >= " + _from.ToString("yyyy-MM-dd") + " AND Enrollment_Date__c <= " + _to.ToString("yyyy-MM-dd") + " AND Policy_Number__c != ''     ORDER BY OpportunityID, Plan_Type__c";

            SfdcBinding.Url = CurrentLoginResult.serverUrl;

            //Create a new session header object and set the session id to that returned by the login
            SfdcBinding.SessionHeaderValue = new SessionHeader();
            SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            QueryResult queryResult = null;
            queryResult = SfdcBinding.query(soql);
            int max_value = 0;
            int Increment_value = 0;
            if (queryResult.size > 0)
                {
                max_value = queryResult.size;
                //put some code in here to handle the records being returned
                bool _done = false;
                while (!_done)
                    {
                    for (int i = 0; i < queryResult.records.Count(); i++)
                        {
                        Increment_value++;
                        lines.Add((OpportunityLineItem)queryResult.records[i]);

                        }
                    if (max_value == Increment_value)
                        {
                        _done = true;
                        }
                    else
                        {
                        queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                        }
                    }




                }

            return lines;

            }

        public List<OpportunityLineItem> GetOpportunityLineBy_changePolicy(List<Oppline_policy> OppLineId_list)
            {


            List<OpportunityLineItem> oppList = new List<OpportunityLineItem>();
            if (OppLineId_list.Count > 700)
                {

                int cant = (int)Math.Ceiling((double)OppLineId_list.Count / (double)500);
                IEnumerable<IEnumerable<Oppline_policy>> matrix = LinqExtensions.Split(OppLineId_list, cant + 1);

                for (int i = 0; i < matrix.Count(); i++)
                    {
                    List<Oppline_policy> temp = matrix.ElementAt(i).ToList();

                    List<OpportunityLineItem> temp_opp = GetOpportunityLineBy_changePolicy(temp);
                    if (temp_opp.Count > 0)
                        oppList = oppList.Concat(temp_opp).ToList();
                    }



                }
            else
                {


                
              

                SfdcBinding.Url = CurrentLoginResult.serverUrl;

                //Create a new session header object and set the session id to that returned by the login
                SfdcBinding.SessionHeaderValue = new SessionHeader();
                SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
                QueryResult queryResult = null;
                /*         string temp = policy_number_list.Where(c => c == "SXH003418").FirstOrDefault();
                            if (temp != null)
                                {
                                temp = "this";
                                }*/
                List<string> OppLine_list = new List<string>();
                List<string> policy_list = new List<string>();
                foreach (var item in OppLineId_list)
                    {
                    OppLine_list.Add(item.oppline_id);
                    policy_list.Add(item.policy);
                    }
                string OppLineId = formatString(OppLine_list.ToList());
                string policyId = formatString(policy_list.ToList());
                // PolicyNumberList = PolicyNumberList.ToList().Distinct();
                string SOQL = "Select id ,Status__c, Policy_Number__c ,Enrollment_Date__c,Effective_Date__c,Opportunityid,Product2.id, Product2.Name, Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C , Product2.Plan_Type__c, Commission_Payable_Number__c ,	Commission_Payable__c ,Opportunity.Agent_ID_Lookup__c,Opportunity.Agent_ID_Lookup__r.First_Name__c,Opportunity.Agent_ID_Lookup__r.Last_Name__c ,Opportunity.Agent_ID_Lookup__r.Status__c, Opportunity.Agent_ID_Lookup__r.Username__c ,Opportunity.Account.FirstName,Opportunity.Account.LastName,  Opportunity.Account.Phone, Opportunity.Agent_ID_Lookup__r.product_profile__c, Opportunity.Fronter_s_Name__c,  Opportunity.Agent_ID_Lookup__r.Ending_Date__c,Product2.Sales_Board_Category__c,Product2.eApp_Display_Name__c from opportunitylineitem where  id in (" + OppLineId + ")  and Policy_Number__c not in (" + policyId + ")";

                // and(Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C in ('1st Med STM','HealtheMed STM') or Product2.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )

                queryResult = SfdcBinding.query(SOQL);
                int max_value = 0;
                int Increment_value = 0;
                if (queryResult.size > 0)
                    {
                    max_value = queryResult.size;
                    //put some code in here to handle the records being returned
                    bool _done = false;
                    while (!_done)
                        {
                        for (int i = 0; i < queryResult.records.Count(); i++)
                            {
                            Increment_value++;
                            oppList.Add((OpportunityLineItem)queryResult.records[i]);

                            }
                        if (max_value == Increment_value)
                            {
                            _done = true;
                            }
                        else
                            {
                            queryResult = SfdcBinding.queryMore(queryResult.queryLocator);
                            }
                        }




                    }

                }



            return oppList;
            }

        public struct Oppline_policy
            {
            public string oppline_id;
            public string policy;
            }




        }
    }
