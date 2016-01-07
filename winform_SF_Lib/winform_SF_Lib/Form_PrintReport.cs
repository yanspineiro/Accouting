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
using winform_SF_Lib.com.salesforce.na5;

namespace winform_SF_Lib
    {
    public partial class Form_PrintReport : Form
        {
        Salerforce_Interface sf_interface;
        FormProgress Proggre;
        List<Agent__c> ListAgent;
        public bool IsAvailable { set; get; }
        public string compannys { set; get; }
        List<Payment_Period__c> calendar_SF;
        public Form_PrintReport()
            {
            InitializeComponent();
            }

        private void Form_PrintReport_Load(object sender, EventArgs e)
            {
            string[] companny_array = compannys.Split(',');

            if (companny_array.Count() > 0)
                {
                for (int i = 0; i < companny_array.Count(); i++)
                    cb_Profile.Items.Add(companny_array[i]);
                }
            }

        private void btn_close_Click(object sender, EventArgs e)
            {
            this.Close();
            }

        private void btn_Print_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] avalues = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerReport.RunWorkerAsync(avalues);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Searching Values for Report";
                Account a = new Account();

                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void backgroundWorkerReport_DoWork(object sender, DoWorkEventArgs e)
            {
            try
                {
                string[] avalues = (string[])e.Argument;

                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                List<Payment__c> ListPaymnt = new List<Payment__c>();


                string agent_filter = "";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }

                    }

                sf_interface = new Salerforce_Interface();
                ListPaymnt = sf_interface.Payment_for_Print(Convert.ToDateTime(avalues[0]), avalues[1], agent_filter);
                if (ListPaymnt != null && ListPaymnt.Count > 0)
                //  List<string> policy_List= new List<string>();
                    {
                    var policy_List = (from z in ListPaymnt
                                       where z.OpportunityLineItem_id__c != null && z.OpportunityLineItem_id__c != ""
                                       select z.OpportunityLineItem_id__c).Distinct();



                    List<OpportunityLineItem> ListOppLine = sf_interface.GetOpportunityLineBy_OppLineId_forPrint(policy_List.ToList(), agent_filter, avalues[1]);

                    CreateShit _CreateShit = new CreateShit();
                    _CreateShit.FillMatrix(ListPaymnt, ListOppLine);
                    _CreateShit.createExcel(true, (string)avalues[0], avalues[1]);
                    }
                }
            catch (Exception ex)
                {
                e.Result = ex.Message; 
                }
            }


        private void backgroundWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            if (e.Result != null)
                {
                string _value = (string)e.Result + ", Some data have been changed.";
                MessageBox.Show(_value);
                }

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
                sf_interface = new Salerforce_Interface();
                calendar_SF = sf_interface.GetPaymentCalendarbyCompanny(companny);

                ListAgent = sf_interface.getAgentbyProfile(companny,IsAvailable);

               

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
           

            if (calendar_SF.Count > 0)
                {

                for (int i = 0; i < calendar_SF.Count; i++)
                    cb_payment_period.Items.Add(calendar_SF[i]);
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

            }

        private void button1_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] avalues = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerGetALL.RunWorkerAsync(avalues);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Searching Values for Report";
                Account a = new Account();

                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void backgroundWorkerGetALL_DoWork(object sender, DoWorkEventArgs e)
            {

            try
                {
                string[] avalues = (string[])e.Argument;

                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                List<Payment__c> ListPaymnt = new List<Payment__c>();


                string agent_filter = "";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }

                    }

                sf_interface = new Salerforce_Interface();
                ListPaymnt = sf_interface.Payment_for_PrintGetALLfromdate(Convert.ToDateTime(avalues[0]), avalues[1], agent_filter);
                if (ListPaymnt != null && ListPaymnt.Count > 0)
                //  List<string> policy_List= new List<string>();
                    {
                    var policy_List = (from z in ListPaymnt
                                       where z.OpportunityLineItem_id__c != null && z.OpportunityLineItem_id__c != ""
                                       select z.OpportunityLineItem_id__c).Distinct();



                    List<OpportunityLineItem> ListOppLine = sf_interface.GetOpportunityLineBy_OppLineId_forPrint(policy_List.ToList(), agent_filter, avalues[1]);

                    CreateShit _CreateShit = new CreateShit();
                    _CreateShit.FillMatrix(ListPaymnt, ListOppLine);
                    _CreateShit.createExcel(true, "", avalues[1]);
                    }
                }
            catch (Exception ex)
                {
                e.Result = ex.Message;
                }

            }

        private void backgroundWorkerGetALL_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            if (e.Result != null)
                {
                string _value = (string)e.Result + ", Some data have been changed.";
                MessageBox.Show(_value);
                }
            }

        private void backgroundWorkerSalesboard_DoWork(object sender, DoWorkEventArgs e)
            {
            List<string> disable_agent = new List<string>();
            try
                {
                string[] avalues = (string[])e.Argument;

                string adate = Convert.ToDateTime(avalues[0]).ToString("yyyy-MM-dd");
                List<SalesBoard__c> ListSalesboard = new List<SalesBoard__c>();


                string agent_filter = "";
                if (avalues.Count() == 3)
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Username__c == avalues[2]).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        agent_filter = _Agent__c.Id;
                        }

                    }

               Payment_Period__c _Payment_Period= calendar_SF.Where(c=> Convert.ToDateTime( c.Payment_Date__c) == Convert.ToDateTime( avalues[0])).FirstOrDefault();
               if (_Payment_Period!= null)
                   {
                   sf_interface = new Salerforce_Interface();
                   ListSalesboard = sf_interface.get_salesborad(avalues[1], Convert.ToDateTime(_Payment_Period.Start_Date__c), Convert.ToDateTime(_Payment_Period.Close_Date__c), agent_filter);
                   List<Payment__c> ListPaymnt=  sf_interface.Payment_for_Print_agregation(Convert.ToDateTime(avalues[0]), avalues[1], agent_filter);

                   List<string> agent_id_list = (from x in ListSalesboard
                                            select x.Agent__c).ToList();
                   List<Salesboard_Vs_Payment>  list_pay_vs_board= new List<Salesboard_Vs_Payment>();
                   double total_witten = 0;
                   double terminated = 0;
                   double chargeback = 0;
                   double board = 0;

                   foreach (var agent_id in agent_id_list)
                       {
                       Salesboard_Vs_Payment item = new Salesboard_Vs_Payment();

                       Agent__c agent = ListAgent.Where(c => c.Id == agent_id).FirstOrDefault();
                       if (agent!= null)
                           {
                           item.Fronter = agent.Username__c;
                           item.Id = agent.Id;
                           item.Enable =  agent.Force_com_Enabled__c.ToString();
                           }
                       else
                           {
                           item.Fronter = agent_id;
                           disable_agent.Add(agent_id);
                           item.Id = agent_id;
                           }
                    
                       List<Payment__c> Paymnt_agent = ListPaymnt.Where(c => c.Agent__c == agent_id).ToList();

                       for (int i = 0; i < Paymnt_agent.Count; i++)
                           {
                           if (Paymnt_agent[i].Payment_Type__c == "Commission")
                               {
                               item.TotalWritten = Convert.ToDouble( Paymnt_agent[i].Payment_Value__c);
                               total_witten += item.TotalWritten;
                               }
                           if (Paymnt_agent[i].Payment_Type__c == "Terminated")
                               {
                               item.Terminated = Convert.ToDouble(Paymnt_agent[i].Payment_Value__c);
                               terminated += item.Terminated;
                               }
                           if (Paymnt_agent[i].Payment_Type__c == "Chargeback")
                               {
                               item.Chargeback = Convert.ToDouble(Paymnt_agent[i].Payment_Value__c);
                               chargeback += item.Chargeback;
                               }
                           
                           }

                       SalesBoard__c _salesBoard = ListSalesboard.Where(c => c.Agent__c == agent_id).FirstOrDefault();
                       if (_salesBoard!= null)
                           {
                           item.Board =  Convert.ToDouble(_salesBoard.Amount__c);
                           board += item.Board;
                           }

                       list_pay_vs_board.Add(item);
                       }

                   if (disable_agent.Count > 0)
                       {
                       List<Agent__c> ListAgent_temp = sf_interface.getAgentbyid(disable_agent);
                       if (ListAgent_temp.Count>0)
                           {
                           for (int k = 0; k < disable_agent.Count; k++)
                               {
                               Agent__c agent = ListAgent_temp.Where(c => c.Id == disable_agent[k]).FirstOrDefault();
                                   if (agent!= null)
                                       {
                                       Salesboard_Vs_Payment pay_vs_board = list_pay_vs_board.Where(c => c.Fronter == agent.Id).FirstOrDefault();
                                       if (pay_vs_board!= null )
                                           {
                                           pay_vs_board.Fronter = agent.Username__c;
                                           pay_vs_board.Id = agent.Id;
                                           pay_vs_board.Enable = agent.Force_com_Enabled__c.ToString();
                                           }
                                      
                                       }
                               }
                           }
                       }
                     list_pay_vs_board= list_pay_vs_board.OrderBy(c => c.Fronter).ToList();
                     list_pay_vs_board.Insert(0, new Salesboard_Vs_Payment() { Fronter = "Totals", TotalWritten = total_witten, Terminated = terminated, Chargeback = chargeback, Board = board });
                     CreateShit _CreateShit = new CreateShit();
                    var app =  _CreateShit.createExcel_salesboard(true, avalues[0], avalues[1], list_pay_vs_board);
                    /*string XMLdata= "";
                    app.WorkbookActivate.
                    app.Workbooks.Close();*/
                   }
              
               
                }
            catch (Exception ex)
                {
                e.Result = ex.Message;
                }
            }

        private void button2_Click(object sender, EventArgs e)
            {
            if (cb_Profile.Text != "" && cb_payment_period.Text != "")
                {
                sf_interface = new Salerforce_Interface();
                string[] avalues = new string[] { cb_payment_period.Text, cb_Profile.Text, comboBox_Agent.Text };
                backgroundWorkerSalesboard.RunWorkerAsync(avalues);
                this.Enabled = false;
                Proggre = new FormProgress();
                Proggre.TopLevel = true;
                Proggre.TopMost = true;
                Proggre.Text_value = "Searching Values for Report";
                Account a = new Account();

                Proggre.Show();
                }
            else
                {
                MessageBox.Show("Profile or Payroll date are empty");
                }
            }

        private void backgroundWorkerSalesboard_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            Proggre.Close();
            this.Enabled = true;
            
            }


        }
    }
