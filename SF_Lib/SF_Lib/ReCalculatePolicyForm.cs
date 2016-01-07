using SF_Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SF_Lib.com.salesforce.na5;

namespace winform_SF_Lib
    {
    public partial class ReCalculatePolicyForm : Form
        {
        public List<Import_HII_CSV> HII_CSV_list { get; set; }
        public List<Import_HII_CSV> update_CSV_list { get; set; }
        FormProgress Proggre;
        public List<string> old_policys;
        public bool Isupdate;
        public List<SaveResult> UpdateResult;

        public ReCalculatePolicyForm()
            {
            InitializeComponent();
            HII_CSV_list = new List<Import_HII_CSV>();
            update_CSV_list = new List<Import_HII_CSV>();
            }

        private void ReCalculatePolicyForm_Load(object sender, EventArgs e)
            {
            old_policys = new List<string>();
            Isupdate = false;
            dataGridView_HII.AutoGenerateColumns = false;
            UpdateResult = new List<SaveResult>();

            dataGridView_HII.DataSource = HII_CSV_list.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
            }
        public static bool HasSpecialCharacters(string str)
            {
            string specialCharacters = @"%!@#$%^&*()?/>.<,:;'\|}]{[_~`+=-" + "\"";
            char[] specialCharactersArray = specialCharacters.ToCharArray();

            int index = str.IndexOfAny(specialCharactersArray);
            //index == -1 no special characters
            if (index == -1)
                return false;
            else
                return true;
            }

        private void backgroundWorker_Search_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                List<SF_Lib.Salerforce_Interface.Oppline_policy> list_new = new List<Salerforce_Interface.Oppline_policy>();
                update_CSV_list = new List<Import_HII_CSV>();
                foreach (var item in HII_CSV_list)
                    {
                    if (!ReCalculatePolicyForm.HasSpecialCharacters(item.Sf_MemberID))
                        {
                        SF_Lib.Salerforce_Interface.Oppline_policy opp_policy = new Salerforce_Interface.Oppline_policy();
                        opp_policy.policy = item.Sf_MemberID;
                        opp_policy.oppline_id = item.OpportunityLine_id;
                        list_new.Add(opp_policy);
                        }

                    }

                Salerforce_Interface sf_interface = new Salerforce_Interface();
                List<OpportunityLineItem> opp_list = sf_interface.GetOpportunityLineBy_changePolicy(list_new);
                if (opp_list.Count > 0)
                    {
                    var PolicyNumberList = (from z in opp_list
                                            where z.Policy_Number__c != null && z.Policy_Number__c != ""
                                            select z.Policy_Number__c).Distinct().ToList();

                    List<Payment__c> paymentbyPolicyNumber = sf_interface.getpaymentbyPolicyNumber((List<string>)PolicyNumberList);
                    PolicyNumberList = (from z in paymentbyPolicyNumber
                                        where z.Policy_Number__c != null && z.Policy_Number__c != ""
                                        select z.Policy_Number__c).Distinct().ToList();

                    List<OpportunityLineItem> except = opp_list.Where(c => PolicyNumberList.Contains(c.Policy_Number__c)).ToList();
                    opp_list = opp_list.Except(except).ToList();
                    foreach (var item in opp_list)
                        {
                        List<Import_HII_CSV> opp_new = HII_CSV_list.Where(c => c.OpportunityLine_id.Equals(item.Id)).ToList();
                        foreach (var item2 in opp_new)
                            {
                            old_policys.Add(item2.Sf_MemberID);
                            Import_HII_CSV opp = (Import_HII_CSV)item2.Clone();
                            opp.Sf_MemberID = item.Policy_Number__c;
                            update_CSV_list.Add(opp);
                            opp.Product_Name = item.Product2.Name;
                            opp.Agent_Fullname = item.Opportunity.Fronter_s_Name__c;
                            }
                        }

                    }

                e.Result = "done";
                }
            catch (Exception ex)
                {

                e.Result = ex.Message;
                }


            }

        private void backgroundWorker_Search_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            this.Focus();
            string result = (string)e.Result;
            if (result != "done")
                {
                MessageBox.Show(result);
                }
            dataGridViewupdate.AutoGenerateColumns = false;
            dataGridViewupdate.DataSource = null;
            dataGridViewupdate.Refresh();
            dataGridViewupdate.DataSource = update_CSV_list.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();

            }

        private void button3_Click(object sender, EventArgs e)
            {
            if (HII_CSV_list.Count > 0)
                {
                backgroundWorker_Search.RunWorkerAsync();
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Search with Saleforce";
                Proggre.Show();
                }

            }

        private void button1_Click(object sender, EventArgs e)
            {
            Isupdate = true;
            Salerforce_Interface sf_interface = new Salerforce_Interface();
            List<Payment__c> paymentbyPolicyNumber = sf_interface.getpaymentbyPolicyNumber(old_policys);
            foreach (var item in update_CSV_list)
                {
                Payment__c newPaymentEntry = paymentbyPolicyNumber.Where(c => c.OpportunityLineItem_id__c == item.OpportunityLine_id).FirstOrDefault();
                if (newPaymentEntry != null)
                    {
                    newPaymentEntry.Policy_Number__c = item.Sf_MemberID;
                    }


                }
            UpdateResult = sf_interface.UpdatePayment(paymentbyPolicyNumber);
            this.Close();
            }

        private void button2_Click(object sender, EventArgs e)
            {
            this.Close();
            }


        }
    }
