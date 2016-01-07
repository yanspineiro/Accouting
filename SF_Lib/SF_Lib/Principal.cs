

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq.Dynamic;
using SF_Lib.com.salesforce.na5;
using SF_Lib;
using System.Text.RegularExpressions;
using winform_SF_Lib;

namespace SF_Lib
    {
    public partial class Principal : Form
        {
        //public List<Payment_Period__c> listPayment_Period;
        FormProgress Proggre;
        CSV_Controller csvcontr;

        Salerforce_Interface sf_interface;
        List<Payment_Period__c> calendar_SF;

        public List<Import_HII_CSV> HII_CSV_list { get; set; }
        public List<Import_HII_CSV> Verify_HII_List { get; set; }
        public List<Import_HII_CSV> MisMatch_HII_List { get; set; }
        public List<Import_HII_CSV> forward_List { get; set; }
        public List<Agent__c> ListAgent { get; set; }
        private bool _from_rigthclick { get; set; }
        string compannys = "";
        private string Policy_number_Principal { get; set; }
        private string Agent_id_Principal { get; set; }
        List<Payment__c> payment_list;
        DateTime from_Audit;
        DateTime To_audit;
        public Boolean run_all;
        //List<OpportunityLineItem> ListAdditionals;
        public Principal()
            {
            InitializeComponent();
            HII_CSV_list = new List<Import_HII_CSV>();
            _from_rigthclick = false;
            Policy_number_Principal = "";
            Agent_id_Principal = "";
            //ListAdditionals = new List<OpportunityLineItem>();
            run_all = false;
            }

        private void btn_import_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerSeachPayment.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with Saleforce and HII";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void Principal_Load(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            backgroundWorkerProfile.RunWorkerAsync();
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Conecting with Saleforce";
            Proggre.Show();
            forward_List = new List<Import_HII_CSV>();

            }


        private void cb_Companny_SelectedIndexChanged(object sender, EventArgs e)
            {
            if (cb_Profile.Text != string.Empty)
                {
                backgroundWorkerCalendar.RunWorkerAsync(cb_Profile.Text);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with Saleforce";
                Proggre.Show();

                }

            }

        private void backgroundWorkerCalendar_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                string companny = (string)e.Argument;
                List<Payment_Period__c> calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(companny);


                ListAgent = sf_interface.getAgentbyProfile(companny, checkBoxAvailable.Checked);

                e.Result = calendar_SF;

                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }
            }

        private void backgroundWorkerCalendar_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            this.Focus();
            cb_payment_period.Items.Clear();
            cb_payment_period.Text = "";
            List<Payment_Period__c> calendar_SF = (List<Payment_Period__c>)e.Result;

            if (calendar_SF.Count > 0)
                {
                for (int i = 0; i < calendar_SF.Count; i++)
                    if (Convert.ToDateTime(calendar_SF[i].Close_Date__c) <= DateTime.Today)//c
                        {
                        if (checkBoxAvailable.Checked && calendar_SF[i].Close_Period__c == false)
                            {
                            cb_payment_period.Items.Add(calendar_SF[i]);
                            }
                        else
                            {
                            if (!checkBoxAvailable.Checked)
                                {
                                cb_payment_period.Items.Add(calendar_SF[i]);
                                }

                            }

                        }
                }

            comboBox_Agent.Items.Clear();
            comboBox_Agent.Items.Add("All");
            ListAgent = ListAgent.OrderBy(c => c.Username__c).ToList();
            if (ListAgent.Count > 0)
                {
                for (int i = 0; i < ListAgent.Count; i++)
                    comboBox_Agent.Items.Add(ListAgent[i].Username__c);
                }
            comboBox_Agent.SelectedItem = "All";

            Verify_HII_List = new List<Import_HII_CSV>();

            dataGridView_HII.DataSource = null;
            dataGridView_HII.Refresh();
            dataGridView_HII.AutoGenerateColumns = false;
            dataGridView_HII.DataSource = Verify_HII_List;

            }

        private void backgroundWorkerProfile_DoWork(object sender, DoWorkEventArgs e)
            {

            compannys = "";
            try
                {


                compannys = sf_interface.getCompannyProfile();

                }
            catch (Exception ex)
                {
                throw (ex);
                }

            e.Result = compannys;

            }

        private void backgroundWorkerProfile_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            cb_payment_period.Items.Clear();
            string compannys = (string)e.Result;
            string[] companny_array = compannys.Split(',');

            if (companny_array.Count() > 0)
                {
                for (int i = 0; i < companny_array.Count(); i++)
                    cb_Profile.Items.Add(companny_array[i]);
                }

            }

        private void backgroundWorkerSeachPayment_DoWork(object sender, DoWorkEventArgs e)
            {
            string[] avalues = (string[])e.Argument;

            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
            calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
            Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();

            string agent_filter = "All";
            if (avalues.Count() == 3)
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }


            payment_list = sf_interface.getpaymentbyPayday(adate, agent_filter, avalues[1]);


            // payment_list = sf_interface.getpaymentbyPayday(adate, agent_filter);
            List<OpportunityLineItem> oppline = sf_interface.Commissios_for_Pay(Convert.ToDateTime(calendar.Start_Date__c), Convert.ToDateTime(calendar.Close_Date__c), agent_filter, avalues[1]);
            //List<String> opp_NoCommsionable = (from x in oppline
            //                                   where x.Commission_Payable__c == null || x.Commission_Payable__c == 0
            //                                   select x.Policy_Number__c).ToList();

            HII_CSV_list = new List<Import_HII_CSV>();
            Verify_HII_List = new List<Import_HII_CSV>();
            MisMatch_HII_List = new List<Import_HII_CSV>();
            CSV_Controller csvcontr;
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
                _Import_HII_CSV.Agent__ProductProfile = payment_list[i].Agent__r.product_profile__c;

                // Proggre.Text_value = "Serching in HII";
                // Proggre.Refresh();
                try
                    {
                    _Import_HII_CSV.Ending_Date = Convert.ToDateTime(payment_list[i].Agent__r.Ending_Date__c);
                    }
                catch (Exception ex)
                    {
                    // _Import_HII_CSV.Ending_Date = DateTime.Today;
                    }


                OpportunityLineItem line = oppline.Where(c => c.Policy_Number__c == _Import_HII_CSV.Sf_MemberID).FirstOrDefault();
                if (line != null)
                    {

                    _Import_HII_CSV.Application_Date = Convert.ToDateTime(line.Enrollment_Date__c).ToString("MM/dd/yyyy");
                    _Import_HII_CSV.Product_Name = line.Product2.Name;
                    _Import_HII_CSV.Opportunity_name = line.Opportunity.Name;
                    _Import_HII_CSV.First_Name = line.Opportunity.Account.FirstName + ' ' + line.Opportunity.Account.LastName;
                    _Import_HII_CSV.Last_Name = line.Opportunity.Account.LastName;
                    _Import_HII_CSV.Phone = CSV_Controller.FormatPhone(line.Opportunity.Account.Phone);
                    _Import_HII_CSV.Product_Name = line.Product2.Name;
                    _Import_HII_CSV.Agent__ProductProfile = line.Opportunity.Agent_ID_Lookup__r.product_profile__c;
                    _Import_HII_CSV.Oppline_Status__c = line.Status__c;

                    oppline.Remove(line);
                    }


                Verify_HII_List.Add(_Import_HII_CSV);
                }

            for (int i = 0; i < oppline.Count; i++)
                {
                Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();
                _Import_HII_CSV.Sf_MemberID = oppline[i].Policy_Number__c;

                _Import_HII_CSV.Agent_id = oppline[i].Opportunity.Agent_ID_Lookup__c;
                _Import_HII_CSV.Agent_Fullname = oppline[i].Opportunity.Agent_ID_Lookup__r.First_Name__c + ' ' + oppline[i].Opportunity.Agent_ID_Lookup__r.Last_Name__c;


                _Import_HII_CSV.Agent_Status = oppline[i].Opportunity.Agent_ID_Lookup__r.Status__c;
                _Import_HII_CSV.Sf_MemberID_Bkup = oppline[i].Policy_Number__c;
                _Import_HII_CSV.Opportunity_id = oppline[i].OpportunityId;
                _Import_HII_CSV.OpportunityLine_id = oppline[i].Id;
                _Import_HII_CSV.Payroll_date = Convert.ToDateTime(avalues[0]);
                _Import_HII_CSV.Payroll_Type = "Commission";
                _Import_HII_CSV.Agent_Commision = oppline[i].Commission_Payable_Number__c == null ? "0" : oppline[i].Commission_Payable_Number__c.ToString();
                _Import_HII_CSV.Verify = false;
                _Import_HII_CSV.IsUpdate = false;
                _Import_HII_CSV.IsManual = false;
                _Import_HII_CSV.from_commission = true;
                _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                _Import_HII_CSV.Opportunity_name = oppline[i].Opportunity.Name;
                _Import_HII_CSV.Application_Date = Convert.ToDateTime(oppline[i].Enrollment_Date__c).ToString("MM/dd/yyyy");
                _Import_HII_CSV.First_Name = oppline[i].Opportunity.Account.FirstName + ' ' + oppline[i].Opportunity.Account.LastName;
                _Import_HII_CSV.Last_Name = oppline[i].Opportunity.Account.LastName;
                _Import_HII_CSV.Phone = CSV_Controller.FormatPhone(oppline[i].Opportunity.Account.Phone);
                _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                _Import_HII_CSV.Agent__ProductProfile = oppline[i].Opportunity.Agent_ID_Lookup__r.product_profile__c;
                _Import_HII_CSV.Oppline_Status__c = oppline[i].Status__c;


                try
                    {
                    _Import_HII_CSV.Ending_Date = Convert.ToDateTime(oppline[i].Opportunity.Agent_ID_Lookup__r.Ending_Date__c);
                    }
                catch (Exception ex)
                    {

                    }

                Verify_HII_List.Add(_Import_HII_CSV);

                }


            try
                {




                csvcontr = new CSV_Controller();
                string userName = "mspiewak2014";// System.Configuration.ConfigurationSettings.AppSettings["username"];
                string password = "Insurance780";// System.Configuration.ConfigurationSettings.AppSettings["pwd_agent"];
                byte[] bytes = csvcontr.Getwebfile(Convert.ToDateTime(calendar.Start_Date__c).AddDays(-31), Convert.ToDateTime(calendar.Close_Date__c), userName, password);
                csvcontr.readFile(bytes);
                csvcontr.HII_CSV_list.RemoveAll(c => c.Termination_Date == "" && Convert.ToDateTime(c.Application_Date) < calendar.Start_Date__c && calendar.Close_Date__c > Convert.ToDateTime(c.Application_Date));
                HII_CSV_list = csvcontr.HII_CSV_list;
                csvcontr.Calendar_Selected = calendar;

                }
            catch (Exception ex)
                {
                throw (ex);
                }

            try
                {

                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);
                MisMatch_HII_List = csvcontr.MisMatch_HII_List;

                DeppMatch(avalues, adate, agent_filter);

                foreach (var item in Verify_HII_List)
                    {
                    if (item.Verify)
                        {
                        item.Provider_Name = "HII";
                        }
                    }

                List<Import_HII_CSV> test = Verify_HII_List.Where(c => (c.Oppline_Status__c != "Processed" && c.Oppline_Status__c != "notFound") && (c.Sf_MemberID == c.Member_ID)).ToList();
                Verify_HII_List = Verify_HII_List.Where(c => c.Agent__ProductProfile == avalues[1] && Convert.ToDouble(c.Agent_Commision) != 0 && (c.Oppline_Status__c == "Processed" || c.Oppline_Status__c == "notFound")).ToList();




                //Verify_HII_List = Verify_HII_List.Where(c => !opp_NoCommsionable.Contains(c.Policy_Number)).ToList();
                // HII_CSV_list.RemoveAll(c => c.Termination_Date == "" && Convert.ToDateTime(c.Application_Date) < calendar.Start_Date__c && calendar.Close_Date__c > Convert.ToDateTime(c.Application_Date));
                // MisMatch_HII_List.RemoveAll(c => c.Termination_Date == "" && Convert.ToDateTime(c.Application_Date) < calendar.Start_Date__c && calendar.Close_Date__c > Convert.ToDateTime(c.Application_Date));

                }
            catch (Exception ex)
                {
                throw (ex);
                }
            }

        private void backgroundWorkerSeachPayment_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {

            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }




            }

        private void button1_Click(object sender, EventArgs e)
            {
            this.Close();
            }

        private void backgroundWorkerReMatch_DoWork(object sender, DoWorkEventArgs e)
            {

            string[] avalues = (string[])e.Argument;

            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
            string agent_filter = "All";
            if (avalues.Count() == 3)
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }


            payment_list = sf_interface.getpaymentbyPayday(adate, agent_filter, avalues[1]);

            CSV_Controller csvcontr = new CSV_Controller(); ;

            csvcontr.HII_CSV_list = MisMatch_HII_List;

            try
                {
                csvcontr.Calendar_Selected = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(adate)).FirstOrDefault();
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);
                MisMatch_HII_List = csvcontr.MisMatch_HII_List;


                }
            catch (Exception ex)
                {
                throw (ex);
                }


            }

        private void backgroundWorkerReMatch_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {

            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();

                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();
                rb_currentPeriod.Checked = true;
                }

            }

        private void btn_MatchAgain_Click(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
            backgroundWorkerReMatch.RunWorkerAsync(_adata);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Match Again";
            Proggre.Show();
            }

        private void backgroundWorkerDeepMatch_DoWork(object sender, DoWorkEventArgs e)
            {

            string[] avalues = (string[])e.Argument;
            string agent_filter = "All";
            if (avalues.Count() == 3)
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }
            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
            DeppMatch(avalues, adate, agent_filter);

            }

        private void DeppMatch(string[] avalues, string adate, string agent_id)
            {


            string agent_filter = "All";
            if (avalues.Count() == 3)
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }

            List<Payment__c> payment_list = sf_interface.getpaymentbyPayday(adate, agent_filter, avalues[1]);
            Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
            Payment_Period__c calendar_old = calendar_SF.Where(c => c.Pay_Period__c + 3 == calendar.Pay_Period__c).FirstOrDefault();
            csvcontr = new CSV_Controller();
            csvcontr.Calendar_Selected = calendar;
            csvcontr.HII_CSV_list = MisMatch_HII_List;
            var PolicyNumberList = (from z in MisMatch_HII_List
                                    where z.Member_ID != null && z.Member_ID != ""
                                    select z.Member_ID).Distinct();
            csvcontr.FillGrid(PolicyNumberList.ToList(), agent_id, avalues[1], Convert.ToDateTime(calendar_old.Start_Date__c), Convert.ToDateTime(calendar.Close_Date__c));
            MisMatch_HII_List = null;
            MisMatch_HII_List = csvcontr.HII_CSV_list;

            List<Import_HII_CSV> temp_HII_List = new List<Import_HII_CSV>();
            if (MisMatch_HII_List != null)
                {
                // dataGridMismatch.AutoGenerateColumns = false;
                temp_HII_List = MisMatch_HII_List.Where(c => c.Payroll_date >= Convert.ToDateTime(avalues[0])).ToList();

                MisMatch_HII_List.RemoveAll(c => temp_HII_List.Contains(c));

                }

            if (Verify_HII_List != null)
                {
                if (temp_HII_List != null && temp_HII_List.Count > 0)
                    Verify_HII_List = Verify_HII_List.Concat(temp_HII_List).ToList();

                }
            else
                {
                Verify_HII_List = temp_HII_List;
                }

            temp_HII_List = new List<Import_HII_CSV>();

            foreach (var item in Verify_HII_List)
                {
                if (temp_HII_List.Where(c => c.Sf_MemberID == item.Sf_MemberID && c.Payroll_Type == item.Payroll_Type).ToList().Count == 0)
                    {
                    temp_HII_List.Add(item);
                    }

                }

            Verify_HII_List = temp_HII_List;

            }

        private void backgroundWorkerDeepMatch_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";



            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();
                }
            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;



                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();
                }




            }

        private void btn_deepMatch_Click(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text };
            backgroundWorkerDeepMatch.RunWorkerAsync(_adata);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Searching ...";
            Proggre.Show();
            }

        private void button2_Click(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            backgroundWorkerManual.RunWorkerAsync(cb_Profile.Text);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Getting salesforce Data...";
            Proggre.Show();

            }

        private void backgroundWorkerManual_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                string aprofile = (string)e.Argument;
                if (ListAgent == null || ListAgent.Count == 0)
                    {
                    ListAgent = sf_interface.getAgentbyProfile(aprofile, checkBoxAvailable.Checked);
                    }


                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(aprofile);
                    }

                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }


            }

        private void backgroundWorkerManual_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            Form_Manual_Insert _Form_Manual_Insert = new Form_Manual_Insert();
            if (Verify_HII_List != null)
                _Form_Manual_Insert.HII_CSV_list = Verify_HII_List;
            else
                _Form_Manual_Insert.HII_CSV_list = new List<Import_HII_CSV>();

            _Form_Manual_Insert.calendar_SF = calendar_SF;
            _Form_Manual_Insert.ListAgent = ListAgent;
            if (_from_rigthclick)
                {

                if (dataGridView_HII.SelectedRows.Count != 0)
                    {
                    DataGridViewRow row_linked = this.dataGridView_HII.SelectedRows[0];
                    _Form_Manual_Insert.Policy_number_Principal = row_linked.Cells["ColumnSf_MemberID"].Value.ToString();
                    _Form_Manual_Insert.Agent_id_Principal = row_linked.Cells["Agent_id_c"].Value.ToString();
                    _Form_Manual_Insert.Amount = row_linked.Cells["ColumnAgent_Commision"].Value.ToString();

                    }
                }
            _Form_Manual_Insert.ShowDialog();
            Verify_HII_List = _Form_Manual_Insert.HII_CSV_list;

            dataGridView_HII.AutoGenerateColumns = false;
            dataGridView_HII.DataSource = null;

            dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
            _from_rigthclick = false;
            Policy_number_Principal = "";
            Agent_id_Principal = "";

            }

        private void backgroundWorkerInsert_DoWork(object sender, DoWorkEventArgs e)
            {

            List<Payment__c> duplicateEntry = new List<Payment__c>();
            List<Payment__c> ValidateListEntry = new List<Payment__c>();

            try
                {

                List<Import_HII_CSV> Imported_HII_temp = Verify_HII_List.Where(c => c.Payroll_date != null && c.Payroll_Type != null && c.Verify).ToList();

                List<Payment__c> newPaymentEntryList = new List<Payment__c>();
                var _policy_Number = (from x in Imported_HII_temp
                                      select x.Sf_MemberID).ToList();
                List<Payment__c> result_list = new List<Payment__c>();
                if (_policy_Number != null)
                    {

                    duplicateEntry = sf_interface.getpaymentbyPolicyNumber(_policy_Number);

                    for (int i = 0; i < Imported_HII_temp.Count; i++)
                        {

                        Payment__c newPaymentEntry = new Payment__c();
                        newPaymentEntry.Agent__c = Imported_HII_temp[i].Agent_id;
                        newPaymentEntry.OpportunityLineItem_id__c = Imported_HII_temp[i].OpportunityLine_id;

                        if (Imported_HII_temp[i].OpportunityLine_id == null || Imported_HII_temp[i].OpportunityLine_id == "")
                            {
                            Import_HII_CSV _Import_HII_CSV = Imported_HII_temp.Where(c => c.OpportunityLine_id != "" && c.OpportunityLine_id != null && c.Sf_MemberID == Imported_HII_temp[i].Sf_MemberID).FirstOrDefault();
                            if (_Import_HII_CSV != null)
                                {
                                newPaymentEntry.OpportunityLineItem_id__c = _Import_HII_CSV.OpportunityLine_id;
                                }
                            }

                        newPaymentEntry.Payment_Date__c = Imported_HII_temp[i].Payroll_date.ToUniversalTime();
                        newPaymentEntry.Payment_Date__cSpecified = true;
                        newPaymentEntry.Payment_Type__c = Imported_HII_temp[i].Payroll_Type;

                        newPaymentEntry.Payment_Value__c = Convert.ToDouble(Imported_HII_temp[i].Agent_Commision);
                        newPaymentEntry.Payment_Value__cSpecified = true;
                        newPaymentEntry.Policy_Number__c = Imported_HII_temp[i].Sf_MemberID;
                        newPaymentEntry.Verify__c = Imported_HII_temp[i].Verify;
                        newPaymentEntry.Verify__cSpecified = true;


                        if (Imported_HII_temp[i].Payroll_Type.Equals("Terminated") || Imported_HII_temp[i].Payroll_Type.Equals("Chargeback"))
                            {
                            newPaymentEntry.Cancel_Date__c = Convert.ToDateTime(Imported_HII_temp[i].Termination_Date).ToUniversalTime();
                            newPaymentEntry.Cancel_Date__cSpecified = true;
                            }

                        if (duplicateEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c)).ToList().Count > 0)
                            {

                            var errorPayment = duplicateEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) && c.Payment_Type__c.Equals("Commission") && newPaymentEntry.Payment_Type__c.Equals(c.Payment_Type__c)).FirstOrDefault();


                            if (errorPayment != null && errorPayment.Verify__c != true)
                                {
                                errorPayment.Verify__c = Imported_HII_temp[i].Verify;
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
                                    if (newPaymentEntryList.Where(c => c.Policy_Number__c == newPaymentEntry.Policy_Number__c && c.Payment_Type__c == newPaymentEntry.Payment_Type__c).ToList().Count == 0)
                                        {
                                        newPaymentEntryList.Add(newPaymentEntry);
                                        }
                                    }
                                }

                            if (newPaymentEntry.Payment_Type__c.Equals("Terminated") || newPaymentEntry.Payment_Type__c.Equals("Chargeback"))
                                {
                                var errorPaymentCancel = duplicateEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) && (c.Payment_Type__c.Equals("Terminated") || c.Payment_Type__c.Equals("Chargeback"))).FirstOrDefault();

                                if (errorPaymentCancel != null && errorPaymentCancel.Verify__c != true)
                                    {
                                    errorPaymentCancel.Verify__c = Imported_HII_temp[i].Verify;
                                    errorPaymentCancel.Verify__cSpecified = true;
                                    ValidateListEntry.Add(errorPaymentCancel);

                                    // errorPayment = duplicateEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) &&  (c.Payment_Type__c.Equals("Terminated") || c.Payment_Type__c.Equals("Chargeback")) && c.Verify__c == false).FirstOrDefault();

                                    /* if (errorPayment != null)
                                            {
                                            if (ValidateListEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) && c.Payment_Type__c.Equals("Commission")).ToList().Count == 0)
                                                {
                                                errorPayment.Verify__c = Imported_HII_temp[i].Verify;
                                                errorPayment.Verify__cSpecified = true;
                                                ValidateListEntry.Add(errorPayment);
                                                }
                                            }*/


                                    }
                                else
                                    {
                                    if (errorPaymentCancel != null && errorPaymentCancel.Verify__c == true)//&& errorPaymentCancel.Payment_Type__c == newPaymentEntry.Payment_Type__c
                                        {
                                        errorPaymentCancel.Description__c = "Error: Duplicate Entry";
                                        }
                                    if (errorPaymentCancel == null && (newPaymentEntry.Payment_Type__c.Equals("Terminated") || newPaymentEntry.Payment_Type__c.Equals("Chargeback")))
                                        {
                                        if (newPaymentEntryList.Where(c => c.Policy_Number__c == newPaymentEntry.Policy_Number__c && (c.Payment_Type__c == "Terminated" || c.Payment_Type__c == "Chargeback")).ToList().Count == 0)
                                            {
                                            newPaymentEntryList.Add(newPaymentEntry);
                                            }
                                        errorPayment = duplicateEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) && c.Payment_Type__c.Equals("Commission") && c.Verify__c == false).FirstOrDefault();

                                        if (errorPayment != null)
                                            {
                                            if (ValidateListEntry.Where(c => c.Policy_Number__c.Equals(newPaymentEntry.Policy_Number__c) && c.Payment_Type__c.Equals("Commission")).ToList().Count == 0)
                                                {
                                                errorPayment.Verify__c = Imported_HII_temp[i].Verify;
                                                errorPayment.Verify__cSpecified = true;
                                                ValidateListEntry.Add(errorPayment);
                                                }
                                            }

                                        }


                                    }

                                }



                            }
                        else
                            {



                            if (newPaymentEntry.Payment_Type__c != "Commission" && !Imported_HII_temp[i].IsManual)
                                {
                                if (newPaymentEntryList.Where(c => c.Policy_Number__c == Imported_HII_temp[i].Member_ID && c.Payment_Type__c == "Commission").ToList().Count == 0)
                                    {
                                    Payment__c newPaymentCommission = new Payment__c();
                                    newPaymentCommission.Agent__c = Imported_HII_temp[i].Agent_id;
                                    newPaymentCommission.OpportunityLineItem_id__c = Imported_HII_temp[i].OpportunityLine_id;

                                    DateTime payroll_date = Imported_HII_temp[i].Payroll_date;
                                    if (Imported_HII_temp[i].Application_Date != null)
                                        {
                                        Payment_Period__c Payment_Period__c = calendar_SF.Where(c => Convert.ToDateTime(c.Start_Date__c) <= Convert.ToDateTime(Imported_HII_temp[i].Application_Date) && Convert.ToDateTime(Imported_HII_temp[i].Application_Date) <= Convert.ToDateTime(c.Close_Date__c)).FirstOrDefault();
                                        if (Payment_Period__c != null)
                                            {
                                            payroll_date = Convert.ToDateTime(Payment_Period__c.Payment_Date__c);
                                            }
                                        }
                                    newPaymentCommission.Payment_Date__c = payroll_date.ToUniversalTime();
                                    newPaymentCommission.Payment_Date__cSpecified = true;
                                    newPaymentCommission.Payment_Type__c = "Commission";

                                    newPaymentCommission.Payment_Value__c = -1 * Convert.ToDouble(Imported_HII_temp[i].Agent_Commision);
                                    newPaymentCommission.Payment_Value__cSpecified = true;
                                    newPaymentCommission.Policy_Number__c = Imported_HII_temp[i].Sf_MemberID;
                                    newPaymentCommission.Verify__c = Imported_HII_temp[i].Verify;
                                    newPaymentCommission.Verify__cSpecified = true;
                                    newPaymentEntryList.Add(newPaymentCommission);
                                    }
                                }
                            if (newPaymentEntryList.Where(c => c.Policy_Number__c == newPaymentEntry.Policy_Number__c && c.Payment_Type__c == newPaymentEntry.Payment_Type__c).ToList().Count == 0)
                                {
                                newPaymentEntryList.Add(newPaymentEntry);
                                }
                            }


                        }

                    /*insert**/
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



                    e.Result = duplicateEntry;


                    }
                }

            catch (Exception ex)
                {
                Proggre.Close();
                MessageBox.Show(ex.Message);
                }


            }

        private void backgroundWorkerInsert_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {

            List<Payment__c> duplicateEntry = (List<Payment__c>)e.Result;

            Proggre.Close();
            this.Enabled = true;

            if (duplicateEntry == null || duplicateEntry.Count == 0)
                {
                MessageBox.Show("The data was successfully save");
                }
            else
                {

                MessageBox.Show(duplicateEntry.Count.ToString() + " error Found");
                tabPage4.Text = "Error (" + duplicateEntry.Count.ToString() + ")";
                dataGridViewError.AutoGenerateColumns = false;
                dataGridViewError.DataSource = null;
                dataGridViewError.Refresh();
                dataGridViewError.DataSource = duplicateEntry.Where(c => c.Description__c != null && c.Description__c != "").ToList();//.Where(c => c.Description__c.Trim() != "").ToList();
                dataGridViewError.Refresh();
                LabelErrorTotal.Text = duplicateEntry.Count.ToString();
                }

            }

        private void btn_insert_Click(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            backgroundWorkerInsert.RunWorkerAsync();
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Show();
            }

        private void button3_Click(object sender, EventArgs e)
            {
            sf_interface = new Salerforce_Interface();

            if (calendar_SF == null || calendar_SF.Count == 0)
                {
                calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(cb_Profile.Text);
                }
            Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(cb_payment_period.Text)).FirstOrDefault();
            // var xxx = sf_interface.Commissios_for_Pay(Convert.ToDateTime(calendar.Start_Date__c), Convert.ToDateTime(calendar.Close_Date__c), cb_Profile.Text);

            }

        private void tabPage2_Click(object sender, EventArgs e)
            {

            }


        public void Filterlist()
            {
            if (rb_AllStatus.Checked && rbt_alltype.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Count.ToString();
                }

            if (rb_AllStatus.Checked && rbn_commission.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type == "Commission").OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type == "Commission").ToList().Count.ToString();
                }


            if (rb_AllStatus.Checked && rbn_cancelled.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type != "Commission").OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type != "Commission").ToList().Count.ToString();
                }
            /** rbn_Verify ***/

            if (rbn_Verify.Checked && rbt_alltype.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Verify).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Verify).ToList().Count.ToString();
                }

            if (rbn_Verify.Checked && rbn_commission.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type == "Commission" && c.Verify).OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type == "Commission" && c.Verify).ToList().Count.ToString();
                }


            if (rbn_Verify.Checked && rbn_cancelled.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type != "Commission" && c.Verify).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type != "Commission" && c.Verify).ToList().Count.ToString();
                }

            /** rNot_Verify ***/

            if (rb_NotVerify.Checked && rbt_alltype.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => !c.Verify).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => !c.Verify).ToList().Count.ToString();
                }

            if (rb_NotVerify.Checked && rbn_commission.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type == "Commission" && !c.Verify).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type == "Commission" && !c.Verify).ToList().Count.ToString();
                }


            if (rb_NotVerify.Checked && rbn_cancelled.Checked)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = Verify_HII_List.Where(c => c.Payroll_Type != "Commission" && !c.Verify).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Payroll_Type != "Commission" && !c.Verify).ToList().Count.ToString();


                }

            }

        private void rb_AllStatus_CheckedChanged(object sender, EventArgs e)
            {
            if (rb_AllStatus.Checked)
                Filterlist();
            }

        private void rbn_Verify_CheckedChanged(object sender, EventArgs e)
            {
            if (rbn_Verify.Checked)
                Filterlist();
            }

        private void rb_NotVerify_CheckedChanged(object sender, EventArgs e)
            {
            if (rb_NotVerify.Checked)
                Filterlist();
            }

        private void rbt_alltype_CheckedChanged(object sender, EventArgs e)
            {
            if (rbt_alltype.Checked)
                Filterlist();
            }

        private void rbn_commission_CheckedChanged(object sender, EventArgs e)
            {
            if (rbn_commission.Checked)
                Filterlist();
            }

        private void rbn_cancelled_CheckedChanged(object sender, EventArgs e)
            {
            if (rbn_cancelled.Checked)
                Filterlist();
            }

        private void btn_Link_Click(object sender, EventArgs e)
            {
            backgroundWorker1.RunWorkerAsync();
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.Show();
            }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
            {

            }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            sf_interface = new Salerforce_Interface();
            CSV_Controller csvcont = new CSV_Controller();
            List<Payment_Period__c> Calendar_List = sf_interface.GetPaymentCalendarbyCompanny(cb_Profile.Text);
            List<White_List_Resource__c> w_list_ChargeBackCancellationPeriod = sf_interface.getChargeBackCancellationPeriod();

            FormMatch _FormMatch = new FormMatch();
            _FormMatch.Saleforce_list = Verify_HII_List.Where(c => c.Verify == false).ToList();
            _FormMatch.CarrierList = MisMatch_HII_List;
            _FormMatch.ShowDialog();

            for (int i = 0; i < _FormMatch.Linked_List.Count; i++)
                {
                Import_HII_CSV _SF_value = Verify_HII_List.Where(c => c.Sf_MemberID == _FormMatch.Linked_List[i].Sf_MemberID).FirstOrDefault();
                Import_HII_CSV _Carrier_value = _FormMatch.SF_HII_list.Where(c => c.Member_ID == _FormMatch.Linked_List[i].Carrier_Policynumber).FirstOrDefault();

                if (_SF_value.Payroll_Type == _Carrier_value.Payroll_Type && _SF_value.Payroll_Type == "Commission")
                    {
                    _SF_value.Sf_MemberID = _SF_value.Member_ID = _Carrier_value.Member_ID;
                    _SF_value.Verify = true;
                    }
                else
                    {
                    if (_Carrier_value.Termination_Date != String.Empty)
                        {

                        _SF_value.Sf_MemberID = _SF_value.Member_ID = _Carrier_value.Member_ID;
                        _SF_value.Verify = true;
                        _Carrier_value.Payroll_date = _SF_value.Payroll_date;
                        _Carrier_value.Sf_MemberID = _Carrier_value.Sf_MemberID_Bkup = _Carrier_value.Member_ID;
                        _Carrier_value.Agent_Commision = (Convert.ToDouble(_SF_value.Agent_Commision) * -1).ToString();
                        DateTime Termination_Date = DateTime.MinValue;
                        try
                            {
                            Termination_Date = Convert.ToDateTime(_Carrier_value.Termination_Date);
                            }
                        catch (Exception ex)
                            {
                            Termination_Date = DateTime.MinValue;
                            }
                        if (Termination_Date != DateTime.MinValue)
                            {
                            Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(cb_payment_period.Text)).FirstOrDefault();
                            csvcont.Calendar_Selected = calendar;
                            //csvcont.Clasify_ChargeBack_Terminated(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_Date, Convert.ToDateTime(_SF_value.Application_Date), _Carrier_value);

                            csvcont.Clasify_ChargeBack_Terminated_new(w_list_ChargeBackCancellationPeriod, Calendar_List, Termination_Date, Convert.ToDateTime(_SF_value.Application_Date), Convert.ToDateTime(_SF_value.Effective_Date), _Carrier_value);
                            }
                        _Carrier_value.Verify = true;
                        Verify_HII_List.Add(_Carrier_value);

                        }
                    else
                        {
                        _SF_value.Sf_MemberID = _SF_value.Member_ID = _Carrier_value.Member_ID;
                        _SF_value.Verify = true;
                        }

                    }

                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();

                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();

                }

            Proggre.Close();
            this.Enabled = true;


            }

        private void btn_ExportReady_Click(object sender, EventArgs e)
            {
            /*SimpleExcelInterface _excel = new SimpleExcelInterface(false);
            _excel.generateExcelFromDataGridView(dataGridView_HII, "Saleforce_data");*/

            backgroundWorkerFillExcel.RunWorkerAsync(dataGridView_HII);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Show();
            }

        private void btn_exportMismached_Click(object sender, EventArgs e)
            {
            /*SimpleExcelInterface _excel = new SimpleExcelInterface(false);
            _excel.generateExcelFromDataGridView(dataGridMismatch, "Provider_data");*/
            backgroundWorkerFillExcel.RunWorkerAsync(dataGridMismatch);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Show();
            }

        private void button3_Click_1(object sender, EventArgs e)
            {
            /* SimpleExcelInterface _excel = new SimpleExcelInterface(false);
             _excel.generateExcelFromDataGridView(dataGridViewError, "Provider_data");*/

            backgroundWorkerFillExcel.RunWorkerAsync(dataGridViewError);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Show();
            }

        private void textBox_SF_TextChanged(object sender, EventArgs e)
            {
            rb_AllStatus.Checked = true;
            rbt_alltype.Checked = true;
            if (textBox_SF.Text != "")
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                var temp = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Product_Name.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Agent_Fullname.ToUpper().Contains(textBox_SF.Text.ToUpper())).ToList();
                dataGridView_HII.DataSource = temp.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                dataGridView_HII.Refresh();
                labelTotal.Text = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Product_Name.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Agent_Fullname.ToUpper().Contains(textBox_SF.Text.ToUpper())).ToList().Count().ToString();
                }
            else
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();

                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                labelTotal.Text = Verify_HII_List.Count().ToString();
                dataGridView_HII.Refresh();
                }
            }

        private void textBox_Carrier_TextChanged(object sender, EventArgs e)
            {
            rb_allPeriod.Checked = true;
            if (textBox_Carrier.Text != "")
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                dataGridMismatch.DataSource = MisMatch_HII_List.Where(c => c.Member_ID.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Product_Name.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Agent_Fullname.ToUpper().Contains(textBox_Carrier.Text.ToUpper())).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList().OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();

                dataGridMismatch.Refresh();
                Lb_MistMach.Text = MisMatch_HII_List.Where(c => c.Member_ID.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Product_Name.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Agent_Fullname.ToUpper().Contains(textBox_Carrier.Text.ToUpper())).Count().ToString();
                }
            else
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count().ToString();
                dataGridMismatch.Refresh();
                }
            }

        public void FilterUnmatchPeriod()
            {

            if (rb_allPeriod.Checked)
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                dataGridMismatch.Refresh();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();
                }

            if (rb_currentPeriod.Checked)
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(cb_payment_period.Text)).FirstOrDefault();
                dataGridMismatch.DataSource = MisMatch_HII_List.Where(c => Convert.ToDateTime(c.Application_Date) >= Convert.ToDateTime(calendar.Start_Date__c) && Convert.ToDateTime(c.Application_Date) <= Convert.ToDateTime(calendar.Close_Date__c)).ToList().OrderBy(x => x.Application_Date).ThenByDescending(x => x.First_Name).ToList();//
                dataGridMismatch.Refresh();
                Lb_MistMach.Text = MisMatch_HII_List.Where(c => Convert.ToDateTime(c.Application_Date) >= Convert.ToDateTime(calendar.Start_Date__c) && Convert.ToDateTime(c.Application_Date) <= Convert.ToDateTime(calendar.Close_Date__c)).ToList().Count.ToString();
                }


            if (rb_outPeriod.Checked)
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(cb_payment_period.Text)).FirstOrDefault();
                dataGridMismatch.DataSource = MisMatch_HII_List.Where(c => Convert.ToDateTime(c.Application_Date) < Convert.ToDateTime(calendar.Start_Date__c) || Convert.ToDateTime(c.Application_Date) > Convert.ToDateTime(calendar.Close_Date__c)).ToList().OrderBy(x => x.Application_Date).ThenByDescending(x => x.First_Name).ToList();//(x => x.Application_Date).ThenByDescending(
                dataGridMismatch.Refresh();
                Lb_MistMach.Text = MisMatch_HII_List.Where(c => Convert.ToDateTime(c.Application_Date) < Convert.ToDateTime(calendar.Start_Date__c) || Convert.ToDateTime(c.Application_Date) > Convert.ToDateTime(calendar.Close_Date__c)).ToList().Count.ToString();
                }

            }

        private void rb_allPeriod_CheckedChanged(object sender, EventArgs e)
            {
            if (rb_allPeriod.Checked)
                {
                FilterUnmatchPeriod();
                }
            }

        private void rb_currentPeriod_CheckedChanged(object sender, EventArgs e)
            {
            if (rb_currentPeriod.Checked)
                {
                FilterUnmatchPeriod();
                }
            }

        private void rb_outPeriod_CheckedChanged(object sender, EventArgs e)
            {
            if (rb_outPeriod.Checked)
                {
                FilterUnmatchPeriod();
                }
            }

        private void backgroundWorkerFillExcel_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                DataGridView dataGridView = (DataGridView)e.Argument;
                SimpleExcelInterface _excel = new SimpleExcelInterface(false);
                _excel.generateExcelFromDataGridView(dataGridView, "Data");
                }
            catch (Exception ex)
                {
                throw (ex);
                }

            }

        private void backgroundWorkerFillExcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            }

        private void button5_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text };
                switch (MessageBox.Show("Do you want to close this Payroll(" + cb_payment_period.Text + ") Period?",
                 "Question",
                 MessageBoxButtons.OKCancel,
                 MessageBoxIcon.Question))
                    {
                    case DialogResult.OK:
                            {
                            backgroundWorkerClosePerirod.RunWorkerAsync(_adata);
                            this.Enabled = false;
                            Proggre = new FormProgress();
                            Proggre.TopLevel = true;
                            Proggre.TopMost = true;
                            Proggre.Text_value = "Closing Payroll Period...";
                            Proggre.Show();
                            break;
                            }



                    case DialogResult.Cancel:
                        // "Cancel" processing
                        break;
                    }
                }

            }

        private void backgroundWorkerClosePerirod_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                sf_interface = new Salerforce_Interface();
                string[] avalues = (string[])e.Argument;

                DateTime adate = Convert.ToDateTime(avalues[0]);
                List<Payment__c> list_SaveResult = sf_interface.closePerirod(adate, avalues[1].ToString());
                e.Result = list_SaveResult;

                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }




            }

        private void backgroundWorkerClosePerirod_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            List<Payment__c> list_SaveResult = (List<Payment__c>)e.Result;

            if (list_SaveResult != null)
                {
                dataGridViewError.AutoGenerateColumns = false;
                dataGridViewError.DataSource = null;
                dataGridViewError.Refresh();
                dataGridViewError.DataSource = list_SaveResult;//.Where(c => c.Description__c.Trim() != "").ToList();
                dataGridViewError.Refresh();
                LabelErrorTotal.Text = list_SaveResult.Count.ToString();
                }
            }

        private void backgroundWorkerFidelity_DoWork(object sender, DoWorkEventArgs e)
            {

            try
                {

                string[] avalues = (string[])e.Argument;


                string agent_filter = "All";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }
                    }


                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
                    }

                var this_calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
                var old_calendar = calendar_SF.Where(c => c.Pay_Period__c + 3 == this_calendar.Pay_Period__c).FirstOrDefault();
                csvcontr = new CSV_Controller();

                List<Fidelity_file> _Fidelity_list = csvcontr.Fidelity("16633882", "Insurance6", Convert.ToDateTime(old_calendar.Start_Date__c));
                _Fidelity_list = _Fidelity_list.Concat(csvcontr.Fidelity("17083164", "Insurance2", Convert.ToDateTime(old_calendar.Start_Date__c))).ToList();
                _Fidelity_list = _Fidelity_list.Concat(csvcontr.Fidelity("17034306", "1984Matt", Convert.ToDateTime(old_calendar.Start_Date__c))).ToList();

                _Fidelity_list = _Fidelity_list.Where(c => Convert.ToDateTime(c.EntryDate) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.EntryDate) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();

                List<Import_HII_CSV> _listFidelity = new List<Import_HII_CSV>();

                foreach (var Item in _Fidelity_list)
                    {
                    if (Item.PolicyNo == "0100700380")
                        {
                        var vv = "";
                        }
                    Import_HII_CSV _fidelity = new Import_HII_CSV();
                    _fidelity.Member_ID = Item.PolicyNo;

                    _fidelity.Agent_id = Item.Agent;
                    _fidelity.Agent_Fullname = Item.WritingAgentName;


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _fidelity.Payroll_Type = Item.ReportingCategory;

                    _fidelity.Verify = false;
                    _fidelity.IsUpdate = false;
                    _fidelity.IsManual = false;
                    _fidelity.from_commission = true;
                    _fidelity.Product_Name = "Fidelity";

                    _fidelity.Application_Date = Convert.ToDateTime(Item.EntryDate).ToString("MM/dd/yyyy");
                    _fidelity.First_Name = Item.Insured.Split(',')[1] + ' ' + Item.Insured.Split(',')[0];
                    _fidelity.Last_Name = Item.Insured.Split(',')[0];
                    _fidelity.Phone = "";



                    _fidelity.Agent__ProductProfile = avalues[1];

                    if (_fidelity.Payroll_Type.ToUpper() == "CANCELLATION" || _fidelity.Payroll_Type.ToUpper() == "TERMINATED" || _fidelity.Payroll_Type.ToUpper() == "CHARGEBACK")
                        {
                        _fidelity.Termination_Date = _fidelity.Application_Date;//  fichero no trae termination date

                        }
                    _fidelity.Payment_Status = _fidelity.Payroll_Type;

                    if (_fidelity.Termination_Date != "")
                        {
                        _fidelity.Payment_Status = _fidelity.Payroll_Type = "Chargeback";
                        }

                    if (Convert.ToDateTime(Item.EntryDate).Year < DateTime.Today.Year)
                        {
                        var vv = "";
                        }

                    var _val = _listFidelity.Where(c => c.First_Name == _fidelity.First_Name && c.Phone == _fidelity.Phone && c.Payroll_Type == _fidelity.Payroll_Type && c.Member_ID == _fidelity.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {

                        _listFidelity.Add(_fidelity);
                        }
                    }

                HII_CSV_list = HII_CSV_list.Concat(_listFidelity).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                csvcontr.HII_CSV_list = _listFidelity;
                csvcontr.Calendar_Selected = this_calendar;
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);
                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();

                DeppMatch(avalues, adate, agent_filter);
                Verify_HII_List = Verify_HII_List.Where(c => c.Payroll_Type != "").ToList();

                }
            catch (Exception ex)
                {
                e.Result = new List<Import_HII_CSV>();
                }



            }

        private void backgroundWorkerFidelity_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            List<Import_HII_CSV> list_SaveResult = (List<Import_HII_CSV>)e.Result;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;
                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;

                labelTotal.Text = Verify_HII_List.Count.ToString();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }
            if (run_all)
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerGetMed.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from GetMed";
                Proggre.Show();
                }

            }

        private void btn_fidelity_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerFidelity.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from Fidelity";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void btn_GetMed_Click(object sender, EventArgs e)
            {

            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerGetMed.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from GetMed";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }

            }

        private void backgroundWorkerGetMed_DoWork(object sender, DoWorkEventArgs e)
            {
            int cant = 0;
            try
                {

                string[] avalues = (string[])e.Argument;

                string agent_filter = "All";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }
                    }


                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
                    }

                var this_calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
                var old_calendar = calendar_SF.Where(c => c.Pay_Period__c + 2 == this_calendar.Pay_Period__c).FirstOrDefault();
                csvcontr = new CSV_Controller();

                List<GetMed_file> _GetMed_list = csvcontr.GetMed(Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "mspiewak", "6178");
                List<GetMed_file> Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();
                _GetMed_list = _GetMed_list.Concat(csvcontr.GetMed(Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "BSWAIN", "3550")).ToList();
                Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();
                _GetMed_list = _GetMed_list.Concat(csvcontr.GetMed(Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "SGRANT", "3546")).ToList();
                Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();

                //  var test = _GetMed_list.Where(c => c.EffectiveDate == null || c.EffectiveDate == "");



                //_GetMed_list = _GetMed_list.Where(c => Convert.ToDateTime(c.EffectiveDate) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.EffectiveDate) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();

                List<Import_HII_CSV> _listGetMed = new List<Import_HII_CSV>();

                foreach (var Item in _GetMed_list)
                    {
                    cant++;
                    Import_HII_CSV _GetMed = new Import_HII_CSV();
                    _GetMed.Member_ID = Item.Member_ID;
                    if (Item.Member_ID == "PGA9003159")
                        {
                        var jj = "kk";
                        }
                    _GetMed.Agent_id = "";//Item.Agent;
                    _GetMed.Agent_Fullname = "";


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _GetMed.Payroll_Type = Item.StatusName;
                    if (Item.StatusName == "Active")
                        {
                        _GetMed.Payroll_Type = "Commission";

                        }
                    if (Item.StatusName == "Pending Term")
                        {
                        _GetMed.Payroll_Type = "Terminated";
                        }

                    _GetMed.Verify = false;
                    _GetMed.IsUpdate = false;
                    _GetMed.IsManual = false;
                    _GetMed.from_commission = false;
                    _GetMed.Product_Name = "GetMed";

                    _GetMed.Application_Date = Convert.ToDateTime(Item.EffectiveDate).ToString("MM/dd/yyyy");
                    _GetMed.First_Name = Item.FirstName + ' ' + Item.LastName;
                    _GetMed.Last_Name = Item.LastName;
                    _GetMed.Phone = Item.Telephone;
                    _GetMed.Termination_Date = Item.TerminateDate;

                    if (Item.TerminateDate != "")
                        {
                        _GetMed.Payroll_Type = "Chargeback";
                        }
                    _GetMed.Payment_Status = _GetMed.Payroll_Type;
                    _GetMed.Agent__ProductProfile = avalues[1];

                    var _val = _listGetMed.Where(c => c.First_Name == _GetMed.First_Name && c.Phone == _GetMed.Phone && c.Payroll_Type == _GetMed.Payroll_Type && c.Member_ID == _GetMed.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _listGetMed.Add(_GetMed);
                        }

                    }
                // _listGetMed = _listGetMed.Where(c => Convert.ToDateTime(c.Application_Date) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.Application_Date) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();
                HII_CSV_list = HII_CSV_list.Concat(_listGetMed).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                csvcontr.HII_CSV_list = _listGetMed;
                csvcontr.Calendar_Selected = this_calendar;
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);

                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();

                DeppMatch(avalues, adate, agent_filter);
                Verify_HII_List = Verify_HII_List.Where(c => c.Payroll_Type != "").ToList();
                /* csvcontr = new CSV_Controller();
                 csvcontr.Calendar_Selected = this_calendar;
                 csvcontr.HII_CSV_list = _listGetMed.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                 var PolicyNumberList = (from z in MisMatch_HII_List
                                         where z.Member_ID != null && z.Member_ID != ""
                                         select z.Member_ID).Distinct();
                 csvcontr.FillGrid(PolicyNumberList.ToList(), agent_filter);
                 //MisMatch_HII_List = null;
                 MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();
                 */


                // e.Result = _listFidelity;
                }
            catch (Exception ex)
                {
                e.Result = new List<Import_HII_CSV>();
                }


            }

        private void backgroundWorkerGetMed_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            List<Import_HII_CSV> list_SaveResult = (List<Import_HII_CSV>)e.Result;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;
                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;

                labelTotal.Text = Verify_HII_List.Count.ToString();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }

            if (run_all)
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerWPA.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from Unified";
                Proggre.Show();
                }

            }

        private void btnWPA_BKUP_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerWPA.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from Unified";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void backgroundWorkerWPA_DoWork(object sender, DoWorkEventArgs e)
            {

            try
                {

                string[] avalues = (string[])e.Argument;
                string agent_filter = "All";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }
                    }


                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
                    }

                var this_calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
                var old_calendar = calendar_SF.Where(c => c.Pay_Period__c + 2 == this_calendar.Pay_Period__c).FirstOrDefault();
                csvcontr = new CSV_Controller();

                List<Enrollment123> _enrollment_list = csvcontr.Getwebfileenrollment123(Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "HBCInsure954", "Insurance954");


                //  var test = _GetMed_list.Where(c => c.EffectiveDate == null || c.EffectiveDate == "");
                _enrollment_list = _enrollment_list.Where(c => Convert.ToDateTime(c.Member_Created_Date) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.Member_Created_Date) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();

                List<Import_HII_CSV> _listUnified = new List<Import_HII_CSV>();

                foreach (var Item in _enrollment_list)
                    {

                    Import_HII_CSV _enrollmentMed = new Import_HII_CSV();
                    _enrollmentMed.Member_ID = Item.Member_ID;

                    _enrollmentMed.Agent_id = "";//Item.Agent;
                    _enrollmentMed.Agent_Fullname = "";


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _enrollmentMed.Payroll_Type = "Commission";

                    _enrollmentMed.Verify = false;
                    _enrollmentMed.IsUpdate = false;
                    _enrollmentMed.IsManual = false;
                    _enrollmentMed.from_commission = false;
                    _enrollmentMed.Product_Name = "Unifield";

                    _enrollmentMed.Application_Date = Convert.ToDateTime(Item.Member_Created_Date).ToString("MM/dd/yyyy");
                    _enrollmentMed.First_Name = Item.First_Name + ' ' + Item.Last_Name;
                    _enrollmentMed.Last_Name = Item.Last_Name;
                    _enrollmentMed.Phone = Item.Phone_1;

                    if (Item.Hold_Reason.Contains("Cancellation"))
                        {
                        _enrollmentMed.Termination_Date = Item.Hold_Date;
                        _enrollmentMed.Payroll_Type = "Cancellation";
                        }

                    if (_enrollmentMed.Termination_Date != "")
                        {
                        _enrollmentMed.Payment_Status = _enrollmentMed.Payroll_Type = "Chargeback";
                        }

                    _enrollmentMed.Payment_Status = _enrollmentMed.Payroll_Type;
                    _enrollmentMed.Agent__ProductProfile = avalues[1];

                    var _val = _listUnified.Where(c => c.First_Name == _enrollmentMed.First_Name && c.Phone == _enrollmentMed.Phone && c.Payroll_Type == _enrollmentMed.Payroll_Type && c.Member_ID == _enrollmentMed.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _listUnified.Add(_enrollmentMed);
                        }

                    }

                HII_CSV_list = HII_CSV_list.Concat(_listUnified).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();

                csvcontr.HII_CSV_list = _listUnified;
                csvcontr.Calendar_Selected = this_calendar;
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);


                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();

                DeppMatch(avalues, adate, agent_filter);
                Verify_HII_List = Verify_HII_List.Where(c => c.Payroll_Type != "").ToList();

                // csvcontr = new CSV_Controller();
                /*  csvcontr.Calendar_Selected = this_calendar;
                  csvcontr.HII_CSV_list = _listUnified.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();;
                  var PolicyNumberList = (from z in MisMatch_HII_List
                                          where z.Member_ID != null && z.Member_ID != ""
                                          select z.Member_ID).Distinct();
                  csvcontr.FillGrid(PolicyNumberList.ToList(), agent_filter);
                  //MisMatch_HII_List = null;
                  //MisMatch_HII_List = csvcontr.HII_CSV_list;
                  MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();

                  */
                // e.Result = _listFidelity;
                }
            catch (Exception ex)
                {

                }

            }

        private void backgroundWorkerWPA_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;

            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;
                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;

                labelTotal.Text = Verify_HII_List.Count.ToString();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }

            if (run_all)
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorker_Wellness.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from WPA";
                Proggre.Show();
                }


            }

        private void unbindToolStripMenuItem_Click(object sender, EventArgs e)
            {
            Policy_number_Principal = "";
            Agent_id_Principal = "";
            _from_rigthclick = true;

            /* if (dataGridView_HII.SelectedRows.Count != 0)
                  {
                  DataGridViewRow row_linked = this.dataGridView_HII.SelectedRows[0];
                  string Sf_MemberID = row_linked.Cells["Sf_MemberID"].Value.ToString();
                  string Carrier_Policynumber = row_linked.Cells["Member_ID"].Value.ToString();

                  }*/

            sf_interface = new Salerforce_Interface();
            backgroundWorkerManual.RunWorkerAsync(cb_Profile.Text);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Getting salesforce Data...";
            Proggre.Show();
            }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
            {

            }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
            {


            if (dataGridView_HII.SelectedRows.Count != 0)
                {

                DataGridViewRow row_linked = this.dataGridView_HII.SelectedRows[0];
                string Sf_MemberID = row_linked.Cells["ColumnSf_MemberID"].Value.ToString();
                string Agent_id = row_linked.Cells["Agent_id_c"].Value.ToString();
                string Trans_Type = row_linked.Cells["ColumnTrans_Type"].Value.ToString();
                string Agent_Commision = row_linked.Cells["ColumnAgent_Commision"].Value.ToString();

                DialogResult dr = MessageBox.Show("Are you want to delete this commission (" + Sf_MemberID + ")   now?", "Delete Record", MessageBoxButtons.YesNo);
                switch (dr)
                    {
                    case DialogResult.Yes:
                            {
                            Import_HII_CSV _Import_HII_CSV = Verify_HII_List.Where(c => c.Sf_MemberID == Sf_MemberID && c.Agent_id == Agent_id && c.Payroll_Type == Trans_Type && c.Agent_Commision == Agent_Commision).FirstOrDefault();
                            if (_Import_HII_CSV != null)
                                {
                                Verify_HII_List.Remove(_Import_HII_CSV);

                                if (payment_list != null && payment_list.Count > 0 && _Import_HII_CSV.Sf_MemberID != null && _Import_HII_CSV.Sf_MemberID != "")
                                    {

                                    Payment__c _Payment__c = payment_list.Where(c => c.Agent__c == Agent_id && c.Payment_Type__c == Trans_Type && c.Policy_Number__c == _Import_HII_CSV.Sf_MemberID).FirstOrDefault();

                                    if (_Payment__c != null)
                                        {

                                        sf_interface = new Salerforce_Interface();

                                        backgroundWorkerdeleteRecord.RunWorkerAsync(_Payment__c);
                                        this.Enabled = false;
                                        Proggre = new FormProgress();
                                        Proggre.TopLevel = true;
                                        Proggre.TopMost = true;
                                        Proggre.Text_value = "Payroll  will delete from Salesforce";
                                        Proggre.Show();



                                        }

                                    }

                                dataGridView_HII.DataSource = null;
                                dataGridView_HII.Refresh();
                                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                                dataGridView_HII.Refresh();
                                labelTotal.Text = Verify_HII_List.Count.ToString();
                                }

                            break;
                            }
                    case DialogResult.No:
                            {

                            break;
                            }
                    }








                }



            }

        private void backgroundWorkerdeleteRecord_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                Payment__c _Payment__c = (Payment__c)e.Argument;
                if (_Payment__c != null)
                    {
                    var _result = sf_interface.deletePayment_byId(_Payment__c.Id);
                    }

                }
            catch (Exception ex)
                {
                throw ex;
                }

            }

        private void backgroundWorkerdeleteRecord_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            }


        private void button6_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorker_Search_SF.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with Saleforce";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }


        private void backgroundWorker_Search_SF_DoWork(object sender, DoWorkEventArgs e)
            {
            string[] avalues = (string[])e.Argument;

            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");

            calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);


            List<Payment__c> duplicateEntry = new List<Payment__c>();

            Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
            Payment_Period__c calendar_next = calendar_SF.Where(c => c.Pay_Period__c == calendar.Pay_Period__c + 1).FirstOrDefault();



            string agent_filter = "All";
            if (avalues.Count() == 3)
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }


            payment_list = sf_interface.getpaymentbyPayday(adate, agent_filter, avalues[1]);
            List<OpportunityLineItem> oppline = sf_interface.Commissios_for_Pay(Convert.ToDateTime(calendar.Start_Date__c), Convert.ToDateTime(calendar.Close_Date__c), agent_filter, avalues[1]);

            List<Payment__c> ListPaymnt_next = sf_interface.Payment_for_Print(Convert.ToDateTime(calendar_next.Payment_Date__c), avalues[1], agent_filter);


            if (!checkBoxMemory.Checked)
                {
                HII_CSV_list = new List<Import_HII_CSV>();
                }


            Verify_HII_List = new List<Import_HII_CSV>();
            MisMatch_HII_List = new List<Import_HII_CSV>();
            CSV_Controller csvcontr;
            List<String> _policy_number_List = new List<string>();
            #region loop payment
            for (int i = 0; i < payment_list.Count; i++)
                {
                Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();
                _Import_HII_CSV.Sf_MemberID = payment_list[i].Policy_Number__c;

                _Import_HII_CSV.Agent_id = payment_list[i].Agent__c;
                if (payment_list[i].Agent__c != null)
                    {
                    _Import_HII_CSV.Agent_Fullname = payment_list[i].Agent__r.First_Name__c + ' ' + payment_list[i].Agent__r.Last_Name__c;
                    _Import_HII_CSV.Agent_Status = payment_list[i].Agent__r.Status__c;
                    _Import_HII_CSV.Agent__ProductProfile = payment_list[i].Agent__r.product_profile__c;

                    }

                _Import_HII_CSV.Coinsurance = payment_list[i].Policy_Number__c;
                //_Import_HII_CSV.Opportunity_id = payment_list[i].;
                _Import_HII_CSV.OpportunityLine_id = payment_list[i].OpportunityLineItem_id__c;
                _Import_HII_CSV.Payroll_date = Convert.ToDateTime(payment_list[i].Payment_Date__c);
                _Import_HII_CSV.Payroll_Type = payment_list[i].Payment_Type__c;
                _Import_HII_CSV.Agent_Commision = payment_list[i].Payment_Value__c.ToString();
                _Import_HII_CSV.Verify = Convert.ToBoolean(payment_list[i].Verify__c);
                _Import_HII_CSV.IsUpdate = true;


                if (payment_list[i].Cancel_Date__c != null)
                    {
                    _Import_HII_CSV.Termination_Date = Convert.ToDateTime(payment_list[i].Cancel_Date__c).ToString("MM/dd/yyy");
                    }


                try
                    {
                    _Import_HII_CSV.Ending_Date = Convert.ToDateTime(payment_list[i].Agent__r.Ending_Date__c);
                    }
                catch (Exception ex)
                    {

                    }


                List<OpportunityLineItem> line_list = oppline.Where(c => c.Policy_Number__c == _Import_HII_CSV.Sf_MemberID).ToList();

                if (line_list.Count == 0)
                    {
                    line_list = oppline.Where(c => c.Id == _Import_HII_CSV.OpportunityLine_id).ToList();
                    }

                OpportunityLineItem line = null;

                if (line_list.Count == 1)
                    {
                    line = line_list[0];
                    }
                else if (line_list.Count > 1)
                    {
                    line = line_list[0];


                    }

                if (line != null)
                    {

                    _Import_HII_CSV.Application_Date = Convert.ToDateTime(line.Enrollment_Date__c).ToString("MM/dd/yyyy");
                    _Import_HII_CSV.Effective_Date = Convert.ToDateTime(line.Effective_Date__c).ToString("MM/dd/yyyy");
                    _Import_HII_CSV.Product_Name = line.Product2.Name;
                    _Import_HII_CSV.Opportunity_name = line.Opportunity.Name;
                    _Import_HII_CSV.First_Name = line.Opportunity.Account.FirstName + ' ' + line.Opportunity.Account.LastName;
                    _Import_HII_CSV.Last_Name = line.Opportunity.Account.LastName;
                    _Import_HII_CSV.Phone = CSV_Controller.FormatPhone(line.Opportunity.Account.Phone);
                    _Import_HII_CSV.Product_Name = line.Product2.Name;
                    _Import_HII_CSV.Agent__ProductProfile = line.Opportunity.Agent_ID_Lookup__r.product_profile__c;
                    _Import_HII_CSV.Oppline_Status__c = line.Status__c;

                    oppline.Remove(line);
                    }
                else
                    {
                    if (_Import_HII_CSV.Sf_MemberID != null && _Import_HII_CSV.Sf_MemberID != "")
                        {
                        _policy_number_List.Add(_Import_HII_CSV.Sf_MemberID);
                        }

                    }

                Verify_HII_List.Add(_Import_HII_CSV);

                }
            #endregion
            if (_policy_number_List.Count > 0)
                {
                List<OpportunityLineItem> oppline_out = sf_interface.GetOpportunityLineBy_Policy_NumberforPrint(_policy_number_List, agent_filter, avalues[1]);
                // List<OpportunityLineItem> t = oppline_out.Where(c => c.Policy_Number__c == "SCI00116300").ToList();
                if (oppline_out.Count > 0)
                    {

                    for (int i = 0; i < Verify_HII_List.Count; i++)
                        {
                        if ((Verify_HII_List[i].First_Name == null || Verify_HII_List[i].First_Name == "") && (Verify_HII_List[i].Last_Name == null || Verify_HII_List[i].Last_Name == ""))
                            {
                            List<OpportunityLineItem> line_list = oppline_out.Where(c => c.Policy_Number__c.Equals(Verify_HII_List[i].Sf_MemberID)).ToList();
                            if (line_list.Count == 1)
                                {

                                Verify_HII_List[i].Application_Date = Convert.ToDateTime(line_list[0].Enrollment_Date__c).ToString("MM/dd/yyyy");
                                Verify_HII_List[i].Effective_Date = Convert.ToDateTime(line_list[0].Effective_Date__c).ToString("MM/dd/yyyy");
                                Verify_HII_List[i].Product_Name = line_list[0].Product2.Name;
                                Verify_HII_List[i].Opportunity_name = line_list[0].Opportunity.Name;
                                Verify_HII_List[i].First_Name = line_list[0].Opportunity.Account.FirstName + ' ' + line_list[0].Opportunity.Account.LastName;
                                Verify_HII_List[i].Last_Name = line_list[0].Opportunity.Account.LastName;
                                Verify_HII_List[i].Phone = CSV_Controller.FormatPhone(line_list[0].Opportunity.Account.Phone);
                                Verify_HII_List[i].Product_Name = line_list[0].Product2.Name;
                                Verify_HII_List[i].Agent__ProductProfile = line_list[0].Opportunity.Agent_ID_Lookup__r.product_profile__c;
                                Verify_HII_List[i].Oppline_Status__c = line_list[0].Status__c;

                                }

                            }

                        }


                    }

                }

            for (int i = 0; i < oppline.Count; i++)
                {


                if (ListPaymnt_next.Where(c => c.Policy_Number__c == oppline[i].Policy_Number__c).ToList().Count == 0)
                    {

                    Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();
                    _Import_HII_CSV.Sf_MemberID = oppline[i].Policy_Number__c;

                    _Import_HII_CSV.Agent_id = oppline[i].Opportunity.Agent_ID_Lookup__c;
                    _Import_HII_CSV.Agent_Fullname = oppline[i].Opportunity.Agent_ID_Lookup__r.First_Name__c + ' ' + oppline[i].Opportunity.Agent_ID_Lookup__r.Last_Name__c;


                    _Import_HII_CSV.Agent_Status = oppline[i].Opportunity.Agent_ID_Lookup__r.Status__c;
                    _Import_HII_CSV.Sf_MemberID_Bkup = oppline[i].Policy_Number__c;
                    _Import_HII_CSV.Opportunity_id = oppline[i].OpportunityId;
                    _Import_HII_CSV.OpportunityLine_id = oppline[i].Id;
                    _Import_HII_CSV.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _Import_HII_CSV.Payroll_Type = "Commission";

                    // _Import_HII_CSV.Agent_Commision = oppline[i].Commission_Payable__c == null ? "0" : oppline[i].Commission_Payable__c.ToString();
                    _Import_HII_CSV.Agent_Commision = oppline[i].Commission_Payable_Number__c == null ? "0" : oppline[i].Commission_Payable_Number__c.ToString();


                    _Import_HII_CSV.Verify = false;
                    _Import_HII_CSV.IsUpdate = false;
                    _Import_HII_CSV.IsManual = false;
                    _Import_HII_CSV.from_commission = true;
                    _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                    _Import_HII_CSV.Opportunity_name = oppline[i].Opportunity.Name;
                    _Import_HII_CSV.Application_Date = Convert.ToDateTime(oppline[i].Enrollment_Date__c).ToString("MM/dd/yyyy");
                    _Import_HII_CSV.Effective_Date = Convert.ToDateTime(oppline[i].Effective_Date__c).ToString("MM/dd/yyyy");
                    _Import_HII_CSV.First_Name = oppline[i].Opportunity.Account.FirstName + ' ' + oppline[i].Opportunity.Account.LastName;
                    _Import_HII_CSV.Last_Name = oppline[i].Opportunity.Account.LastName;
                    _Import_HII_CSV.Phone = CSV_Controller.FormatPhone(oppline[i].Opportunity.Account.Phone);
                    _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                    _Import_HII_CSV.Agent__ProductProfile = oppline[i].Opportunity.Agent_ID_Lookup__r.product_profile__c;
                    _Import_HII_CSV.Oppline_Status__c = oppline[i].Status__c;
                    _Import_HII_CSV.New_sales = true;

                    try
                        {
                        _Import_HII_CSV.Ending_Date = Convert.ToDateTime(oppline[i].Opportunity.Agent_ID_Lookup__r.Ending_Date__c);
                        }
                    catch (Exception ex)
                        {

                        }

                    if (Verify_HII_List.Where(c => c.Sf_MemberID == _Import_HII_CSV.Sf_MemberID && c.OpportunityLine_id != _Import_HII_CSV.OpportunityLine_id).ToList().Count == 0)
                        {
                        Verify_HII_List.Add(_Import_HII_CSV);
                        }
                    else
                        {


                        Payment__c _Payment = new Payment__c();

                        _Payment.Agent__c = _Import_HII_CSV.Agent_id;
                        _Payment.OpportunityLineItem_id__c = _Import_HII_CSV.Opportunity_id;
                        _Payment.Policy_Number__c = _Import_HII_CSV.Sf_MemberID;
                        _Payment.Description__c = "Agent :" + _Import_HII_CSV.Agent_Fullname + " Account:" + _Import_HII_CSV.First_Name;

                        duplicateEntry.Add(_Payment);
                        _Payment = new Payment__c();

                        var entry = Verify_HII_List.Where(c => c.Sf_MemberID == _Import_HII_CSV.Sf_MemberID && c.OpportunityLine_id != _Import_HII_CSV.OpportunityLine_id).FirstOrDefault();

                        _Payment.Agent__c = entry.Agent_id;
                        _Payment.OpportunityLineItem_id__c = entry.Opportunity_id;
                        _Payment.Policy_Number__c = entry.Sf_MemberID;
                        _Payment.Description__c = "Agent :" + entry.Agent_Fullname + " Account:" + entry.First_Name;
                        duplicateEntry.Add(_Payment);
                        }

                    }


                }


            if (checkBoxMemory.Checked)
                {
                csvcontr = new CSV_Controller();
                csvcontr.Calendar_Selected = calendar;
                csvcontr.HII_CSV_list = HII_CSV_list;
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);
                MisMatch_HII_List = csvcontr.MisMatch_HII_List;

                DeppMatch(avalues, adate, agent_filter);
                }
            e.Result = duplicateEntry;
            }

        private void backgroundWorker_Search_SF_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {

            Proggre.Close();
            this.Enabled = true;
            this.Focus();
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                labelTotal.Text = Verify_HII_List.Count.ToString();

                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }

            List<Payment__c> duplicateEntry = (List<Payment__c>)e.Result;
            dataGridViewError.DataSource = null;
            dataGridViewError.Refresh();

            if (duplicateEntry != null && duplicateEntry.Count > 0)
                {


                MessageBox.Show(duplicateEntry.Count.ToString() + " error Found");
                tabPage4.Text = "Error (" + duplicateEntry.Count.ToString() + ")";
                dataGridViewError.AutoGenerateColumns = false;
                dataGridViewError.DataSource = null;
                dataGridViewError.Refresh();
                dataGridViewError.DataSource = duplicateEntry.Where(c => c.Description__c != null && c.Description__c != "").ToList();//.Where(c => c.Description__c.Trim() != "").ToList();
                dataGridViewError.Refresh();
                LabelErrorTotal.Text = duplicateEntry.Count.ToString();
                }

            }

        private void backgroundWorker_HII_DoWork(object sender, DoWorkEventArgs e)
            {
            // List<Import_HII_CSV> tes = Verify_HII_List.Where(c => c.Agent_Commision == null || c.Agent_Commision=="").ToList();




            string[] avalues = (string[])e.Argument;
            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
            string agent_filter = "All";
            try
                {


                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }
                    }


                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
                    }

                Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();


                csvcontr = new CSV_Controller();
                string userName = "mspiewak2014";
                string password = "zhVwIU52b2EZCx!a";
                byte[] bytes = csvcontr.Getwebfile(Convert.ToDateTime(calendar.Start_Date__c).AddDays(-31), Convert.ToDateTime(calendar.Close_Date__c), userName, password);
                csvcontr.readFile(bytes);

                csvcontr.HII_CSV_list = csvcontr.HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                HII_CSV_list = HII_CSV_list.Concat(csvcontr.HII_CSV_list).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                csvcontr.Calendar_Selected = calendar;


                    {
                    var xx = "";
                    }
                }
            catch (Exception ex)
                {
                throw (ex);
                }

            try
                {
                //List<Import_HII_CSV> tes = Verify_HII_List.Where(c => c.Sf_MemberID == "DTX0488000").ToList();
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);
                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();
                //tes = Verify_HII_List.Where(c => c.Sf_MemberID == "DTX0488000").ToList();


                DeppMatch(avalues, adate, agent_filter);
                //List<Import_HII_CSV> tes = Verify_HII_List.Where(c => c.Agent_Commision == null || c.Agent_Commision == "").ToList();

                foreach (var item in Verify_HII_List)
                    {
                    if (item.Verify)
                        {
                        item.Provider_Name = "HII";
                        }
                    }

                Verify_HII_List = Verify_HII_List.Where(c => c.Agent__ProductProfile == avalues[1] && Convert.ToDouble(c.Agent_Commision) != 0 && (c.Oppline_Status__c == "Processed" || c.Oppline_Status__c == "notFound")).ToList();
                Verify_HII_List = Verify_HII_List.Where(c => c.Payroll_Type != "").ToList();

                }
            catch (Exception ex)
                {
                throw (ex);
                }
            }

        private void backgroundWorker_HII_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();
                dataGridView_HII.Refresh();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();
                dataGridMismatch.Refresh();

                rb_currentPeriod.Checked = true;
                }

            if (run_all)
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                //backgroundWorkerFidelity.RunWorkerAsync(_adata);
                backgroundWorkerGetMed.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from GetMed";
                Proggre.Show();
                }


            }

        private void btn_HII_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorker_HII.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with HII";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void checkBoxMemory_CheckedChanged(object sender, EventArgs e)
            {
            if (checkBoxMemory.Checked)
                {

                if (HII_CSV_list != null && HII_CSV_list.Count > 0)
                    {

                    btn_HII.Enabled = false;
                    btn_fidelity.Enabled = false;
                    btn_GetMed.Enabled = false;
                    btnWPA_BKUP.Enabled = false;


                    }
                else
                    {

                    MessageBox.Show("System doesn't found any provider information in memory");
                    btn_HII.Enabled = true;
                    btn_fidelity.Enabled = true;
                    btn_GetMed.Enabled = true;
                    btnWPA_BKUP.Enabled = true;
                    checkBoxMemory.Checked = false;

                    }

                }
            else
                {
                btn_HII.Enabled = true;
                btn_fidelity.Enabled = true;
                btn_GetMed.Enabled = true;
                btnWPA_BKUP.Enabled = true;

                }

            }

        private void btn_Print_Click(object sender, EventArgs e)
            {

            Form_PrintReport _Form_PrintReport = new Form_PrintReport();
            _Form_PrintReport.compannys = compannys;
            _Form_PrintReport.IsAvailable = checkBoxAvailable.Checked;
            _Form_PrintReport.Show();


            /*if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] avalues = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerReport.RunWorkerAsync(avalues);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Searching Values for Report";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }*/
            }

        private void backgroundWorkerReport_DoWork(object sender, DoWorkEventArgs e)
            {

            try
                {
                string[] avalues = (string[])e.Argument;

                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                List<Payment__c> ListPaymnt = new List<Payment__c>();
                    {

                    }
                    ListPaymnt = sf_interface.Payment_for_Print(Convert.ToDateTime(avalues[0]), avalues[1], "");
                    if (ListPaymnt != null && ListPaymnt.Count > 0)
                    //  List<string> policy_List= new List<string>();
                        {
                        var policy_List = (from z in ListPaymnt
                                           where z.Policy_Number__c != null && z.Policy_Number__c != ""
                                           select z.Policy_Number__c).Distinct();

                        List<OpportunityLineItem> ListOppLine = sf_interface.GetOpportunityLineBy_Policy_NumberforPrint(policy_List.ToList(), "", avalues[1]);

                        CreateShit _CreateShit = new CreateShit();
                        _CreateShit.FillMatrix(ListPaymnt, ListOppLine);
                        // _CreateShit.createExcel(true);
                        }
                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }

            }

        private void backgroundWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            }

        private void backgroundWorker_Wellness_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {

                string[] avalues = (string[])e.Argument;
                string agent_filter = "All";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }
                    }


                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                if (calendar_SF == null || calendar_SF.Count == 0)
                    {
                    calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues[1]);
                    }

                var this_calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();
                var old_calendar = calendar_SF.Where(c => c.Pay_Period__c + 2 == this_calendar.Pay_Period__c).FirstOrDefault();
                csvcontr = new CSV_Controller();

                List<WPA> _wellness_list = csvcontr.GetWPA("spiewak", "Insurance954", Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "102415");

                List<Import_HII_CSV> _list_Wellness = new List<Import_HII_CSV>();

                foreach (var Item in _wellness_list)
                    {

                    Import_HII_CSV _wellness = new Import_HII_CSV();
                    _wellness.Member_ID = Item.Id;


                    _wellness.Agent_id = "";//Item.Agent;
                    _wellness.Agent_Fullname = "";


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _wellness.Payroll_Type = "Commission";

                    _wellness.Verify = false;
                    _wellness.IsUpdate = false;
                    _wellness.IsManual = false;
                    _wellness.from_commission = false;
                    _wellness.Product_Name = "WPA";

                    _wellness.Application_Date = Convert.ToDateTime(Item.Created.Split('a')[0]).ToString("MM/dd/yyyy");
                    _wellness.First_Name = Item.Name;
                    _wellness.Last_Name = "";
                    _wellness.Phone = Item.Phone;

                    if (Item.Status == "Cancel")
                        {
                        _wellness.Termination_Date = Item.Next_Step.Split('a')[0];
                        _wellness.Payment_Status = _wellness.Payroll_Type = "Chargeback";
                        }
                    else
                        {
                        _wellness.Payment_Status = _wellness.Payroll_Type = "Commission";
                        }

                    _wellness.Agent__ProductProfile = avalues[1];

                    var _val = _list_Wellness.Where(c => c.First_Name == _wellness.First_Name && c.Phone == _wellness.Phone && c.Payroll_Type == _wellness.Payroll_Type && c.Member_ID == _wellness.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _list_Wellness.Add(_wellness);
                        }

                    }
                HII_CSV_list = HII_CSV_list.Concat(_list_Wellness).ToList();
                #region UNO DOS
                if (avalues[1] == "UNO" || avalues[1] == "DOS" || avalues[1] == "TRES")
                    {
                    _wellness_list = new List<WPA>();
                    csvcontr = new CSV_Controller();

                    _wellness_list = csvcontr.GetWPA("spiewak2", "Insurance954", Convert.ToDateTime(old_calendar.Start_Date__c), Convert.ToDateTime(this_calendar.Close_Date__c), "109553");

                    // _list_Wellness = new List<Import_HII_CSV>();

                    foreach (var Item in _wellness_list)
                        {

                        Import_HII_CSV _wellness = new Import_HII_CSV();
                        _wellness.Member_ID = Item.Id;

                        _wellness.Agent_id = "";//Item.Agent;
                        _wellness.Agent_Fullname = "";


                        // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                        _wellness.Payroll_Type = "Commission";

                        _wellness.Verify = false;
                        _wellness.IsUpdate = false;
                        _wellness.IsManual = false;
                        _wellness.from_commission = false;
                        _wellness.Product_Name = "WPA";

                        _wellness.Application_Date = Convert.ToDateTime(Item.Created.Split('a')[0]).ToString("MM/dd/yyyy");
                        _wellness.First_Name = Item.Name;
                        _wellness.Last_Name = "";
                        _wellness.Phone = Item.Phone;

                        if (Item.Status == "Cancel")
                            {
                            _wellness.Termination_Date = Item.Next_Step.Split('a')[0];
                            _wellness.Payment_Status = _wellness.Payroll_Type = "Chargeback";
                            }
                        else
                            {
                            _wellness.Payment_Status = _wellness.Payroll_Type = "Commission";
                            }

                        _wellness.Agent__ProductProfile = avalues[1];

                        var _val = _list_Wellness.Where(c => c.First_Name == _wellness.First_Name && c.Phone == _wellness.Phone && c.Payroll_Type == _wellness.Payroll_Type && c.Member_ID == _wellness.Member_ID).FirstOrDefault();
                        if (_val == null)
                            {
                            _list_Wellness.Add(_wellness);
                            }

                        }


                    HII_CSV_list = HII_CSV_list.Concat(_list_Wellness).ToList();

                    }
                #endregion

                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();

                csvcontr.HII_CSV_list = _list_Wellness;
                csvcontr.Calendar_Selected = this_calendar;
                csvcontr.matchSalerforce(Verify_HII_List, avalues[1]);


                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.MisMatch_HII_List).ToList();

                DeppMatch(avalues, adate, agent_filter);
                Verify_HII_List = Verify_HII_List.Where(c => c.Payroll_Type != "").ToList();
                }
            catch (Exception ex)
                {

                }
            }

        private void button7_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorker_Wellness.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Getting data from WPA";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void backgroundWorker_Wellness_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            List<Import_HII_CSV> list_SaveResult = (List<Import_HII_CSV>)e.Result;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;
                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;

                labelTotal.Text = Verify_HII_List.Count.ToString();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                rb_currentPeriod.Checked = true;
                }
            if (run_all)
                {
                System.Media.SystemSounds.Exclamation.Play();
                run_all = false;
                }

            }


        private void button8_Click(object sender, EventArgs e)
            {

            FormAudit_Sales _FormAudit_Sales = new FormAudit_Sales();
            _FormAudit_Sales.ShowDialog();

            if (!_FormAudit_Sales.isclose)
                {


                sf_interface = new Salerforce_Interface();
                Object[] _adata = new Object[] { _FormAudit_Sales.From, _FormAudit_Sales.To, comboBox_Agent.Text };
                from_Audit = _FormAudit_Sales.From;
                To_audit = _FormAudit_Sales.To;
                backgroundWorker_AUDIT.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with Saleforce";
                Proggre.Show();

                }




            }

        private void backgroundWorker_AUDIT_DoWork(object sender, DoWorkEventArgs e)
            {

            Object[] avalues = (Object[])e.Argument;
            string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
            string agent_filter = "All";
            if (avalues.Count() == 3 && agent_filter != "All")
                {
                Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == (string)avalues[2]).FirstOrDefault();
                if (_Agent__c != null)
                    {
                    agent_filter = _Agent__c.Id;
                    }
                }



            List<OpportunityLineItem> oppline = sf_interface.AUDIT_Commissios_for_Pay((DateTime)avalues[0], (DateTime)avalues[1], agent_filter);
            //List<String> opp_NoCommsionable = (from x in oppline
            //                                   where x.Commission_Payable__c == null || x.Commission_Payable__c == 0
            //                                   select x.Policy_Number__c).ToList();


            Verify_HII_List = new List<Import_HII_CSV>();
            MisMatch_HII_List = new List<Import_HII_CSV>();
            CSV_Controller csvcontr;

            // ListAdditionals = new List<OpportunityLineItem>();
            for (int i = 0; i < oppline.Count; i++)
                {


                Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();
                _Import_HII_CSV.Sf_MemberID = oppline[i].Policy_Number__c;

                _Import_HII_CSV.Agent_id = oppline[i].Opportunity.Agent_ID_Lookup__c;
                _Import_HII_CSV.Agent_Fullname = oppline[i].Opportunity.Agent_ID_Lookup__r.First_Name__c + ' ' + oppline[i].Opportunity.Agent_ID_Lookup__r.Last_Name__c;


                _Import_HII_CSV.Agent_Status = oppline[i].Opportunity.Agent_ID_Lookup__r.Status__c;
                _Import_HII_CSV.Sf_MemberID_Bkup = oppline[i].Policy_Number__c;
                _Import_HII_CSV.Opportunity_id = oppline[i].OpportunityId;
                _Import_HII_CSV.OpportunityLine_id = oppline[i].Id;
                _Import_HII_CSV.Payroll_date = Convert.ToDateTime(avalues[0]);
                _Import_HII_CSV.Payroll_Type = "Commission";

                // _Import_HII_CSV.Agent_Commision = oppline[i].Commission_Payable__c == null ? "0" : oppline[i].Commission_Payable__c.ToString();
                _Import_HII_CSV.Agent_Commision = oppline[i].Commission_Payable_Number__c == null ? "0" : oppline[i].Commission_Payable_Number__c.ToString();


                _Import_HII_CSV.Verify = false;
                _Import_HII_CSV.IsUpdate = false;
                _Import_HII_CSV.IsManual = false;
                _Import_HII_CSV.from_commission = true;
                _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                _Import_HII_CSV.Opportunity_name = oppline[i].Opportunity.Name;
                _Import_HII_CSV.Application_Date = Convert.ToDateTime(oppline[i].Enrollment_Date__c).ToString("MM/dd/yyyy");
                _Import_HII_CSV.First_Name = oppline[i].Opportunity.Account.FirstName + ' ' + oppline[i].Opportunity.Account.LastName;
                _Import_HII_CSV.Last_Name = oppline[i].Opportunity.Account.LastName;
                _Import_HII_CSV.Phone = CSV_Controller.FormatPhone(oppline[i].Opportunity.Account.Phone);
                _Import_HII_CSV.Product_Name = oppline[i].Product2.Name;
                _Import_HII_CSV.Agent__ProductProfile = oppline[i].Opportunity.Agent_ID_Lookup__r.product_profile__c;
                _Import_HII_CSV.Oppline_Status__c = oppline[i].Status__c;


                try
                    {
                    _Import_HII_CSV.Ending_Date = Convert.ToDateTime(oppline[i].Opportunity.Agent_ID_Lookup__r.Ending_Date__c);
                    }
                catch (Exception ex)
                    {

                    }

                Verify_HII_List.Add(_Import_HII_CSV);







                }



            }


        private void backgroundWorker_AUDIT_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();

                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();

                //   rb_currentPeriod.Checked = true;
                }

            backgroundWorker_HII_AUDIT.RunWorkerAsync();
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Text_value = "Conecting with Carrer";
            Proggre.Show();
            }

        private void backgroundWorker_HII_AUDIT_DoWork(object sender, DoWorkEventArgs e)
            {

            #region HII
            try
                {


                csvcontr = new CSV_Controller();
                string userName = "mspiewak2014";// System.Configuration.ConfigurationSettings.AppSettings["username"];
                string password = "zhVwIU52b2EZCx!a";// System.Configuration.ConfigurationSettings.AppSettings["pwd_agent"];
                byte[] bytes = csvcontr.Getwebfile(from_Audit, To_audit, userName, password);
                csvcontr.readFile(bytes);

                csvcontr.HII_CSV_list = csvcontr.HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();
                //Import_HII_CSV temp = csvcontr.HII_CSV_list.Where(c => c.Member_ID == "DTX0018600").FirstOrDefault();

                }
            catch (Exception ex)
                {
                throw (ex);
                }

            try
                {

                MisMatch_HII_List = MisMatch_HII_List.Concat(csvcontr.HII_CSV_list).ToList();

                //   DeppMatch(avalues, adate, agent_filter);

                foreach (var item in Verify_HII_List)
                    {

                    var _val_veri = csvcontr.HII_CSV_list.Where(c => c.Member_ID.ToUpper() == item.Sf_MemberID.ToUpper()).FirstOrDefault();
                    if (_val_veri != null)
                        {
                        item.Verify = true;
                        MisMatch_HII_List.RemoveAll(c => c.Member_ID.ToUpper() == item.Sf_MemberID.ToUpper());
                        }


                    }

                }
            catch (Exception ex)
                {
                throw (ex);
                }
            #endregion
            int cant = 0;
            #region Unified
            try
                {



                csvcontr = new CSV_Controller();

                List<Enrollment123> _enrollment_list = csvcontr.Getwebfileenrollment123(from_Audit, To_audit, "HBCInsure954", "Insurance954");


                List<Import_HII_CSV> _listUnified = new List<Import_HII_CSV>();

                foreach (var Item in _enrollment_list)
                    {

                    Import_HII_CSV _enrollmentMed = new Import_HII_CSV();
                    _enrollmentMed.Member_ID = Item.Member_ID;

                    _enrollmentMed.Agent_id = "";//Item.Agent;
                    _enrollmentMed.Agent_Fullname = "";

                    _enrollmentMed.Payroll_Type = "Commission";

                    _enrollmentMed.Verify = false;
                    _enrollmentMed.IsUpdate = false;
                    _enrollmentMed.IsManual = false;
                    _enrollmentMed.from_commission = false;
                    _enrollmentMed.Product_Name = "Unifield";

                    _enrollmentMed.Application_Date = Convert.ToDateTime(Item.Member_Created_Date).ToString("MM/dd/yyyy");
                    _enrollmentMed.First_Name = Item.First_Name + ' ' + Item.Last_Name;
                    _enrollmentMed.Last_Name = Item.Last_Name;
                    _enrollmentMed.Phone = Item.Phone_1;

                    if (Item.Hold_Reason.Contains("Cancellation"))
                        {
                        _enrollmentMed.Termination_Date = Item.Hold_Date;
                        _enrollmentMed.Payroll_Type = "Cancellation";
                        }

                    if (_enrollmentMed.Termination_Date != "")
                        {
                        _enrollmentMed.Payment_Status = _enrollmentMed.Payroll_Type = "Chargeback";
                        }

                    _enrollmentMed.Payment_Status = _enrollmentMed.Payroll_Type;


                    var _val = _listUnified.Where(c => c.First_Name == _enrollmentMed.First_Name && c.Phone == _enrollmentMed.Phone && c.Payroll_Type == _enrollmentMed.Payroll_Type && c.Member_ID == _enrollmentMed.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _listUnified.Add(_enrollmentMed);
                        }

                    var _val_veri = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper() == _enrollmentMed.Member_ID.ToUpper()).FirstOrDefault();
                    if (_val_veri == null)
                        {
                        MisMatch_HII_List.Add(_enrollmentMed);
                        }
                    else
                        {
                        _val_veri.Verify = true;
                        }

                    }


                }
            catch (Exception ex)
                {

                }
            #endregion


            cant = 0;

            #region Getmed
            try
                {


                csvcontr = new CSV_Controller();


                List<GetMed_file> _GetMed_list = csvcontr.GetMed(from_Audit, To_audit, "mspiewak", "6178");
                List<GetMed_file> Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();
                _GetMed_list = _GetMed_list.Concat(csvcontr.GetMed(from_Audit, To_audit, "BSWAIN", "3550")).ToList();
                Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();
                _GetMed_list = _GetMed_list.Concat(csvcontr.GetMed(from_Audit, To_audit, "SGRANT", "3546")).ToList();
                Listt = _GetMed_list.Where(c => c.EffectiveDate == null || Regex.IsMatch(c.EffectiveDate, "[a-zA-Z]")).ToList();

                //  var test = _GetMed_list.Where(c => c.EffectiveDate == null || c.EffectiveDate == "");



                //_GetMed_list = _GetMed_list.Where(c => Convert.ToDateTime(c.EffectiveDate) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.EffectiveDate) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();

                List<Import_HII_CSV> _listGetMed = new List<Import_HII_CSV>();

                foreach (var Item in _GetMed_list)
                    {
                    cant++;
                    Import_HII_CSV _GetMed = new Import_HII_CSV();
                    _GetMed.Member_ID = Item.Member_ID;
                    if (Item.Member_ID == "PGA9003159")
                        {
                        var jj = "kk";
                        }
                    _GetMed.Agent_id = "";//Item.Agent;
                    _GetMed.Agent_Fullname = "";


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _GetMed.Payroll_Type = Item.StatusName;
                    if (Item.StatusName == "Active")
                        {
                        _GetMed.Payroll_Type = "Commission";

                        }
                    if (Item.StatusName == "Pending Term")
                        {
                        _GetMed.Payroll_Type = "Terminated";
                        }

                    _GetMed.Verify = false;
                    _GetMed.IsUpdate = false;
                    _GetMed.IsManual = false;
                    _GetMed.from_commission = false;
                    _GetMed.Product_Name = "GetMed";

                    _GetMed.Application_Date = Convert.ToDateTime(Item.EffectiveDate).ToString("MM/dd/yyyy");
                    _GetMed.First_Name = Item.FirstName + ' ' + Item.LastName;
                    _GetMed.Last_Name = Item.LastName;
                    _GetMed.Phone = Item.Telephone;
                    _GetMed.Termination_Date = Item.TerminateDate;

                    if (Item.TerminateDate != "")
                        {
                        _GetMed.Payroll_Type = "Chargeback";
                        }
                    _GetMed.Payment_Status = _GetMed.Payroll_Type;


                    var _val = _listGetMed.Where(c => c.First_Name == _GetMed.First_Name && c.Phone == _GetMed.Phone && c.Payroll_Type == _GetMed.Payroll_Type && c.Member_ID == _GetMed.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _listGetMed.Add(_GetMed);
                        }
                    var _val_veri = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper() == _GetMed.Member_ID.ToUpper()).FirstOrDefault();
                    if (_val_veri == null)
                        {
                        MisMatch_HII_List.Add(_GetMed);
                        }
                    else
                        {
                        _val_veri.Verify = true;
                        }

                    }
                // _listGetMed = _listGetMed.Where(c => Convert.ToDateTime(c.Application_Date) >= Convert.ToDateTime(old_calendar.Start_Date__c) && Convert.ToDateTime(c.Application_Date) <= Convert.ToDateTime(this_calendar.Close_Date__c)).ToList();
                HII_CSV_list = HII_CSV_list.Concat(_listGetMed).ToList();






                // e.Result = _listFidelity;
                }
            catch (Exception ex)
                {
                e.Result = new List<Import_HII_CSV>();
                }


            #endregion
            cant = 0;

            #region Wellness
            try
                {

                csvcontr = new CSV_Controller();

                List<WPA> _wellness_list = csvcontr.GetWPA("spiewak", "Insurance954", from_Audit, To_audit, "102415");

                List<Import_HII_CSV> _list_Wellness = new List<Import_HII_CSV>();

                foreach (var Item in _wellness_list)
                    {

                    Import_HII_CSV _wellness = new Import_HII_CSV();
                    _wellness.Member_ID = Item.Id;

                    _wellness.Agent_id = "";//Item.Agent;
                    _wellness.Agent_Fullname = "";


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _wellness.Payroll_Type = "Commission";

                    _wellness.Verify = false;
                    _wellness.IsUpdate = false;
                    _wellness.IsManual = false;
                    _wellness.from_commission = false;
                    _wellness.Product_Name = "WPA";

                    _wellness.Application_Date = Convert.ToDateTime(Item.Created.Split('a')[0]).ToString("MM/dd/yyyy");
                    _wellness.First_Name = Item.Name;
                    _wellness.Last_Name = "";
                    _wellness.Phone = Item.Phone;

                    if (Item.Status.Contains("Cancel"))
                        {
                        _wellness.Termination_Date = Item.Next_Step.Split('a')[0];
                        _wellness.Payment_Status = _wellness.Payroll_Type = "Chargeback";
                        }
                    else
                        {
                        _wellness.Payment_Status = _wellness.Payroll_Type = "Commission";
                        }



                    var _val = _list_Wellness.Where(c => c.First_Name == _wellness.First_Name && c.Phone == _wellness.Phone && c.Payroll_Type == _wellness.Payroll_Type && c.Member_ID == _wellness.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {
                        _list_Wellness.Add(_wellness);
                        }

                    var _val_veri = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper() == _wellness.Member_ID.ToUpper()).FirstOrDefault();
                    if (_val_veri == null)
                        {
                        MisMatch_HII_List.Add(_wellness);
                        }
                    else
                        {
                        _val_veri.Verify = true;
                        }


                    }

                HII_CSV_list = HII_CSV_list.Concat(_list_Wellness).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();







                }
            catch (Exception ex)
                {

                }


            #endregion
            cant = 0;
            #region Fidelity
            try
                {


                csvcontr = new CSV_Controller();

                List<Fidelity_file> _Fidelity_list = csvcontr.Fidelity("16633882", "Insurance6", from_Audit);
                _Fidelity_list = _Fidelity_list.Concat(csvcontr.Fidelity("17083164", "Insurance2", from_Audit)).ToList();
                _Fidelity_list = _Fidelity_list.Concat(csvcontr.Fidelity("17034306", "1984Matt", from_Audit)).ToList();



                List<Import_HII_CSV> _listFidelity = new List<Import_HII_CSV>();

                foreach (var Item in _Fidelity_list)
                    {

                    Import_HII_CSV _fidelity = new Import_HII_CSV();
                    _fidelity.Member_ID = Item.PolicyNo;

                    _fidelity.Agent_id = Item.Agent;
                    _fidelity.Agent_Fullname = Item.WritingAgentName;


                    // _fidelity.Payroll_date = Convert.ToDateTime(avalues[0]);
                    _fidelity.Payroll_Type = Item.ReportingCategory;

                    _fidelity.Verify = false;
                    _fidelity.IsUpdate = false;
                    _fidelity.IsManual = false;
                    _fidelity.from_commission = true;
                    _fidelity.Product_Name = "Fidelity";

                    _fidelity.Application_Date = Convert.ToDateTime(Item.EntryDate).ToString("MM/dd/yyyy");
                    _fidelity.First_Name = Item.Insured.Split(',')[1] + ' ' + Item.Insured.Split(',')[0];
                    _fidelity.Last_Name = Item.Insured.Split(',')[0];
                    _fidelity.Phone = "";





                    if (_fidelity.Payroll_Type.ToUpper() == "CANCELLATION" || _fidelity.Payroll_Type.ToUpper() == "TERMINATED" || _fidelity.Payroll_Type.ToUpper() == "CHARGEBACK")
                        {
                        _fidelity.Termination_Date = _fidelity.Application_Date;//  fichero no trae termination date

                        }
                    _fidelity.Payment_Status = _fidelity.Payroll_Type;

                    if (_fidelity.Termination_Date != "")
                        {
                        _fidelity.Payment_Status = _fidelity.Payroll_Type = "Chargeback";
                        }

                    if (Convert.ToDateTime(Item.EntryDate).Year < DateTime.Today.Year)
                        {
                        var vv = "";
                        }

                    var _val = _listFidelity.Where(c => c.First_Name == _fidelity.First_Name && c.Phone == _fidelity.Phone && c.Payroll_Type == _fidelity.Payroll_Type && c.Member_ID == _fidelity.Member_ID).FirstOrDefault();
                    if (_val == null)
                        {

                        _listFidelity.Add(_fidelity);
                        }

                    var _val_veri = Verify_HII_List.Where(c => c.Sf_MemberID.ToUpper() == _fidelity.Member_ID.ToUpper()).FirstOrDefault();
                    if (_val_veri == null)
                        {
                        MisMatch_HII_List.Add(_fidelity);
                        }
                    else
                        {
                        _val_veri.Verify = true;
                        }
                    }

                HII_CSV_list = HII_CSV_list.Concat(_listFidelity).ToList();
                HII_CSV_list = HII_CSV_list.GroupBy(test => test.Member_ID).Select(grp => grp.First()).ToList();








                }
            catch (Exception ex)
                {
                e.Result = new List<Import_HII_CSV>();
                }





            #endregion



            }

        private void backgroundWorker_HII_AUDIT_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            Lb_MistMach.Text = "0";
            labelTotal.Text = "0";
            if (Verify_HII_List != null)
                {
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList(); ;
                labelTotal.Text = Verify_HII_List.Count.ToString();
                dataGridView_HII.Refresh();
                btn_MatchAgain.Enabled = true;
                }

            if (MisMatch_HII_List != null)
                {
                dataGridMismatch.DataSource = null;
                dataGridMismatch.Refresh();
                dataGridMismatch.AutoGenerateColumns = false;

                dataGridMismatch.DataSource = MisMatch_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                Lb_MistMach.Text = MisMatch_HII_List.Count.ToString();
                dataGridMismatch.Refresh();

                // rb_currentPeriod.Checked = true;
                }
            }

        private void tabPage1_Click(object sender, EventArgs e)
            {

            }

        private void btn_RunAll_Click(object sender, EventArgs e)
            {

            run_all = true;
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] _adata = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorker_HII.RunWorkerAsync(_adata);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Conecting with HII";
                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }
        private void moveNextPeriodToolStripMenuItem_Click(object sender, EventArgs e)
            {
            process_forward();
            /*
            List<string> list_policy = new List<string>();
            bool forward_flag = false;
            foreach (DataGridViewRow row in dataGridView_HII.Rows)
                {


                DataGridViewCheckBoxCell cell = row.Cells["Column_Forward"] as DataGridViewCheckBoxCell;
                string policy = row.Cells["ColumnSf_MemberID"].Value.ToString();
                //Compare to the true value because Value isn't boolean
                if (Convert.ToBoolean(cell.Value.ToString()))
                    {
                    policy = row.Cells["ColumnSf_MemberID"].Value.ToString();
                    list_policy.Add(policy);
                    forward_flag = true;
                    }

                }

            if (list_policy.Count == 0)
                {
                DataGridViewRow row_linked = this.dataGridView_HII.SelectedRows[0];
                string Sf_MemberID = row_linked.Cells["ColumnSf_MemberID"].Value.ToString();
                string Agent_id = row_linked.Cells["Agent_id_c"].Value.ToString();
                string Trans_Type = row_linked.Cells["ColumnTrans_Type"].Value.ToString();
                string Agent_Commision = row_linked.Cells["ColumnAgent_Commision"].Value.ToString();
                if (Verify_HII_List.Where(c => c.Sf_MemberID == Sf_MemberID && c.Verify == false).Count() == 1)
                    {
                    list_policy.Add(Sf_MemberID);
                    forward_flag = true;
                    }
                }



            if (forward_flag)
                {
                sf_interface = new Salerforce_Interface();
                DialogResult dr = MessageBox.Show("Do you want to move this these sale (" + sf_interface.formatString(list_policy) + ")   to next period?", "Move Sale", MessageBoxButtons.YesNo);
                switch (dr)
                    {
                    case DialogResult.Yes:
                            {

                            backgroundWorkerForwardnextPeriod.RunWorkerAsync(new dynamic[] { list_policy, cb_Profile.Text });
                            this.Enabled = false;
                            Proggre = new FormProgress();
                            Proggre.TopLevel = true;
                            Proggre.TopMost = true;
                            Proggre.Text_value = "Moving to next period";
                            Proggre.Show();

                            break;
                            }
                    case DialogResult.No:
                            {

                            break;
                            }
                    }

                }


            */

            }

        private void backgroundWorkerForwardnextPeriod_DoWork(object sender, DoWorkEventArgs e)
            {

            dynamic[] avalues = (dynamic[])e.Argument;

            List<Payment__c> entry = sf_interface.getpaymentbyPolicyNumber(avalues[0]);
            List<Import_HII_CSV> item1 = Verify_HII_List.Where(c => avalues[0].Contains(c.Sf_MemberID) && c.Verify == false && c.Payroll_Type == "Commission").ToList();
            Import_HII_CSV item = new Import_HII_CSV();
            if (item1.Count > 0)
                {
                item = item1[0];
                }

            if (entry.Count > 0)
                {
                Payment__c _payment__c = entry.Where(c => c.Payment_Type__c == "Commission" && c.Close_Period__c == false).FirstOrDefault();
                DateTime _pay_day = Convert.ToDateTime(_payment__c.Payment_Date__c);
                if (_payment__c != null)
                    {
                    _pay_day = getnextperioddate(avalues[1], _pay_day);
                    _payment__c.Payment_Date__c = _pay_day;
                    _payment__c.Payment_Date__cSpecified = true;

                    List<SaveResult> resultUpdate = new List<SaveResult>();
                    List<Payment__c> _paylist = new List<Payment__c>();
                    _paylist.Add(_payment__c);
                    resultUpdate = sf_interface.UpdatePayment(_paylist);
                    string cond = "";
                    List<object> array = new List<object>();
                    if (resultUpdate[0].errors != null)
                        {
                        for (int k = 0; k < resultUpdate[0].errors.Count(); k++)
                            {
                            if (k == 0)
                                {
                                string[] _message = resultUpdate[0].errors[0].message.Split(':');
                                cond = resultUpdate[0].errors[0].fields[0] + ".Equals(@0)";
                                array.Add(_message[2].Trim());
                                }
                            else
                                {
                                string[] _message = resultUpdate[0].errors[0].message.Split(':');
                                cond = cond + " and " + resultUpdate[0].errors[0].fields[0] + ".Equals(@" + k.ToString() + ")";
                                array.Add(_message[2]);
                                }

                            }
                        }
                    if (array.Count == 0)
                        {
                        Verify_HII_List.Remove(item);
                        }
                    else
                        {
                        e.Result = array[0];
                        }

                    }

                }
            else
                {

                Payment__c newPaymentEntry = new Payment__c();

                if (item != null)
                    {

                    newPaymentEntry.Agent__c = item.Agent_id;
                    newPaymentEntry.OpportunityLineItem_id__c = item.OpportunityLine_id;
                    newPaymentEntry.OpportunityLineItem_id__c = item.OpportunityLine_id;
                    newPaymentEntry.Payment_Date__c = item.Payroll_date.ToUniversalTime();
                    newPaymentEntry.Payment_Date__cSpecified = true;
                    newPaymentEntry.Payment_Type__c = item.Payroll_Type;
                    newPaymentEntry.Payment_Value__c = Convert.ToDouble(item.Agent_Commision);
                    newPaymentEntry.Payment_Value__cSpecified = true;
                    newPaymentEntry.Policy_Number__c = item.Sf_MemberID;
                    newPaymentEntry.Verify__c = item.Verify;
                    newPaymentEntry.Verify__cSpecified = true;

                    DateTime _pay_day = getnextperioddate(avalues[1], item.Payroll_date);
                    newPaymentEntry.Payment_Date__c = _pay_day;
                    newPaymentEntry.Payment_Date__cSpecified = true;


                    List<SaveResult> result = new List<SaveResult>();
                    List<Payment__c> newPaymentEntryList = new List<Payment__c>();
                    newPaymentEntryList.Add(newPaymentEntry);
                    if (newPaymentEntryList.Count > 0)
                        {
                        result = sf_interface.insertPayment(newPaymentEntryList);

                        }
                    List<object> array = new List<object>();
                    if (result[0].errors != null)
                        {
                        string error_message = "";
                        string cond = "";

                        for (int k = 0; k < result[0].errors.Count(); k++)
                            {
                            if (k == 0)
                                {
                                string[] _message = result[0].errors[0].message.Split(':');
                                cond = result[0].errors[0].fields[0] + ".Equals(@0)";
                                array.Add(_message[2].Trim());
                                }
                            else
                                {
                                string[] _message = result[0].errors[0].message.Split(':');
                                cond = cond + " and " + result[0].errors[0].fields[0] + ".Equals(@" + k.ToString() + ")";
                                array.Add(_message[2]);
                                }

                            }
                        }

                    if (array.Count == 0)
                        {
                        Verify_HII_List.Remove(item);
                        }
                    else
                        {
                        e.Result = array[0];
                        }



                    }


                }


            }

        private DateTime getnextperioddate(string avalues, DateTime _pay_day)
            {
            if (calendar_SF == null || calendar_SF.Count == 0)
                {
                calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(avalues);
                }

            // Payment_Period__c calendar = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == Convert.ToDateTime(avalues[0])).FirstOrDefault();

            Payment_Period__c _Payment_Period__c = calendar_SF.Where(c => Convert.ToDateTime(c.Payment_Date__c) == _pay_day).FirstOrDefault();

            if (_Payment_Period__c != null)
                {
                double _period = Convert.ToDouble(_Payment_Period__c.Pay_Period__c + 1);
                _Payment_Period__c = calendar_SF.Where(c => c.Pay_Period__c.Equals(_period)).FirstOrDefault();
                if (_Payment_Period__c != null)
                    {
                    return Convert.ToDateTime(_Payment_Period__c.Payment_Date__c);
                    }

                }

            return _pay_day;
            }

        private void backgroundWorkerForwardnextPeriod_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;

            string message = (string)e.Result;
            if (message == null || message == "")
                {

                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.AutoGenerateColumns = false;


                dataGridView_HII.DataSource = Verify_HII_List.OrderBy(x => x.First_Name).ThenByDescending(x => x.Product_Name).ToList();
                labelTotal.Text = Verify_HII_List.Count.ToString();
                dataGridView_HII.Refresh();
                MessageBox.Show("The sales was successfully move to next period");

                }
            else
                {
                MessageBox.Show(message);
                }



            }

        private void checkBoxAvailable_CheckedChanged(object sender, EventArgs e)
            {
            cb_payment_period.Items.Clear();
            comboBox_Agent.Items.Clear();
            cb_Profile.SelectedIndex = -1;
            }

        private void button9_Click(object sender, EventArgs e)
            {

            /*WPA_XML _WPA_XML = new WPA_XML()
            {
                CORPID = "102415",
                PASSWORD = "spiewak",
                USERNAME = "Insurance954"
            };
            WPA_API.postXMLData(@"https://www.enrollment123.com/api/user.getall/", _WPA_XML);*/
            }

        private void button9_Click_1(object sender, EventArgs e)
            {
            ReCalculatePolicyForm _ReCalculatePolicyForm = new ReCalculatePolicyForm();
            _ReCalculatePolicyForm.HII_CSV_list = Verify_HII_List.Where(c => !c.Verify).ToList();
            _ReCalculatePolicyForm.Show();
            if (_ReCalculatePolicyForm.Isupdate)
                {
                List<string> result_list = new List<string>();
                searchError(result_list, _ReCalculatePolicyForm.UpdateResult);
                fillError(result_list);
                }
            }
        private void dataGridView_HII_OnCellValueChanged(object sender, DataGridViewCellEventArgs e)
            {
            if (e.ColumnIndex == Column_Forward.Index && e.RowIndex != -1)
                {

                DataGridViewRow row_linked = this.dataGridView_HII.SelectedRows[0];
                if (row_linked.Cells["ColumnVerify"].Value.ToString().ToUpper() == "FALSE")
                    {
                    string Sf_MemberID = row_linked.Cells["ColumnSf_MemberID"].Value.ToString();
                    string forward = row_linked.Cells["Column_Forward"].Value.ToString();
                    if (forward.ToUpper() == "TRUE")
                        {
                        forward_List.Add(Verify_HII_List.Where(c => c.Sf_MemberID == Sf_MemberID).FirstOrDefault());
                        }

                    else
                        {
                        var element = forward_List.Where(c => c.Sf_MemberID == Sf_MemberID).FirstOrDefault();
                        if (element != null)
                            {
                            forward_List.Remove(element);
                            }

                        }

                    }
                else
                    {
                    row_linked.Cells["Column_Forward"].Value = false;
                    }


                }
            }

        private void dataGridView_HII_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
            {
            if (e.ColumnIndex == Column_Forward.Index && e.RowIndex != -1)
                {
                dataGridView_HII.EndEdit();
                }
            }

        public List<string> process_forward()
            {
            List<string> result_list = new List<string>();
            try
                {
                if (forward_List.Count > 0)
                    {

                    var PolicyNumberList = (from z in forward_List
                                            where z.Sf_MemberID != null && z.Sf_MemberID != ""
                                            select z.Sf_MemberID).Distinct().ToList();

                    List<Payment__c> entry = sf_interface.getpaymentbyPolicyNumber((List<string>)PolicyNumberList);



                    if (entry.Count > 0)
                        {
                        List<Payment__c> _payment__list = entry.Where(c => c.Payment_Type__c == "Commission" && c.Close_Period__c == false).ToList();
                        DateTime _pay_day = Convert.ToDateTime(cb_payment_period.Text);
                        List<Payment__c> _paylist = new List<Payment__c>();
                        _pay_day = getnextperioddate(cb_Profile.Text, _pay_day);
                        foreach (var _payment__c in _payment__list)
                            {
                            if (_payment__c != null)
                                {

                                _payment__c.Payment_Date__c = _pay_day;
                                _payment__c.Payment_Date__cSpecified = true;
                                _paylist.Add(_payment__c);
                                }
                            }
                        if (_paylist.Count > 0)
                            {
                            List<SaveResult> resultUpdate = new List<SaveResult>();
                            resultUpdate = sf_interface.UpdatePayment(_paylist);
                            string cond = "";

                            searchError(result_list, resultUpdate);
                            if (result_list.Count == 0)
                                {
                                foreach (var item in _paylist)
                                    {
                                    var element = Verify_HII_List.Where(c => c.Sf_MemberID == item.Policy_Number__c).FirstOrDefault();
                                    if (element != null)
                                        {
                                        Verify_HII_List.Remove(element);
                                        forward_List.Remove(element);
                                        }

                                    }
                                }


                            }

                        }

                    if (forward_List.Count > 0)
                        {
                        List<Payment__c> _paylist = new List<Payment__c>();

                        List<Payment__c> newPaymentEntryList = new List<Payment__c>();
                        foreach (var item in forward_List)
                            {
                            if (item != null)
                                {
                                Payment__c newPaymentEntry = new Payment__c();
                                newPaymentEntry.Agent__c = item.Agent_id;
                                newPaymentEntry.OpportunityLineItem_id__c = item.OpportunityLine_id;
                                newPaymentEntry.OpportunityLineItem_id__c = item.OpportunityLine_id;
                                // newPaymentEntry.OpportunityLineItem_id__cSpecified = true;
                                newPaymentEntry.Payment_Date__c = item.Payroll_date.ToUniversalTime();
                                newPaymentEntry.Payment_Date__cSpecified = true;
                                newPaymentEntry.Payment_Type__c = item.Payroll_Type;
                                newPaymentEntry.Payment_Value__c = Convert.ToDouble(item.Agent_Commision);
                                newPaymentEntry.Payment_Value__cSpecified = true;
                                newPaymentEntry.Policy_Number__c = item.Sf_MemberID;
                                newPaymentEntry.Verify__c = item.Verify;
                                newPaymentEntry.Verify__cSpecified = true;

                                DateTime _pay_day = getnextperioddate(cb_Profile.Text, item.Payroll_date);
                                newPaymentEntry.Payment_Date__c = _pay_day;
                                newPaymentEntry.Payment_Date__cSpecified = true;
                                newPaymentEntryList.Add(newPaymentEntry);
                                }
                            }

                        List<SaveResult> result = new List<SaveResult>();

                        if (newPaymentEntryList.Count > 0)
                            {
                            result = sf_interface.insertPayment(newPaymentEntryList);
                            }

                        searchError(result_list, result);

                        if (result_list.Count == 0)
                            {
                            foreach (var item in newPaymentEntryList)
                                {
                                var element = Verify_HII_List.Where(c => c.Sf_MemberID == item.Policy_Number__c).FirstOrDefault();
                                if (element != null)
                                    {
                                    Verify_HII_List.Remove(element);
                                    forward_List.Remove(element);
                                    }

                                }
                            }

                        }
                    dataGridView_HII.DataSource = Verify_HII_List;
                    dataGridView_HII.Refresh();
                    labelTotal.Text = Verify_HII_List.Where(c => !c.Verify).ToList().Count.ToString();
                    forward_List = new List<Import_HII_CSV>();



                    if (result_list.Count > 0)
                        {

                        }
                    Proggre.Close();
                    this.Enabled = true;

                    fillError(result_list);
                    }
                }
            catch (Exception ex)
                {

                MessageBox.Show(ex.Message); ;
                }


            return result_list;
            }

        private void fillError(List<string> result_list)
            {
            if (result_list == null || result_list.Count == 0)
                {
                MessageBox.Show("The data was successfully save");
                }
            else
                {

                MessageBox.Show(result_list.Count.ToString() + " error Found");
                tabPage4.Text = "Error (" + result_list.Count.ToString() + ")";
                dataGridViewError.AutoGenerateColumns = false;
                dataGridViewError.DataSource = null;
                dataGridViewError.Refresh();
                List<Payment__c> errors = new List<Payment__c>();
                foreach (var item in result_list)
                    {
                    if (item != null && item != "")
                        {
                        Payment__c error = new Payment__c()
                        {
                            Description__c = item
                        };
                        }
                    }
                dataGridViewError.DataSource = errors;//.Where(c => c.Description__c.Trim() != "").ToList();
                dataGridViewError.Refresh();
                LabelErrorTotal.Text = errors.Count.ToString();
                }
            }

        private static void searchError(List<string> result_list, List<SaveResult> result)
            {
            foreach (var item in result)
                {
                if (item.errors != null)
                    {
                    string error_message = "";
                    string cond = "";

                    for (int k = 0; k < item.errors.Count(); k++)
                        {
                        if (k == 0)
                            {
                            string[] _message = item.errors[0].message.Split(':');
                            cond = item.errors[0].fields[0] + ".Equals(@0)";
                            result_list.Add(_message[2].Trim());
                            }
                        else
                            {
                            string[] _message = item.errors[0].message.Split(':');
                            cond = cond + " and " + item.errors[0].fields[0] + ".Equals(@" + k.ToString() + ")";
                            result_list.Add(_message[2]);
                            }

                        }
                    }

                }
            }

        private void moveAllUnveryfyToolStripMenuItem_Click(object sender, EventArgs e)
            {
            forward_List = Verify_HII_List.Where(c => !c.Verify).ToList();
            if (forward_List.Count > 0)
                {
                process_forward();
                }

            }
        }
    }
