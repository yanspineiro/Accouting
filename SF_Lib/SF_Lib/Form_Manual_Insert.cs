

using SF_Lib.com.salesforce.na5;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SF_Lib
    {
    public partial class Form_Manual_Insert : Form
        {

        public List<Agent__c> ListAgent { get; set; }
        public List<Payment_Period__c> calendar_SF { get; set; }
        public List<Import_HII_CSV> HII_CSV_list { get; set; }
        private List<Import_HII_CSV> ManualHII_CSV_list { get; set; }
        int _index_selected;
        public string Agent_id_Principal { get; set; }
        public string Policy_number_Principal { get; set; }
        public string Amount { get; set; }

        public Form_Manual_Insert()
            {
            InitializeComponent();
            }

        private void cellClicOrCellSelect(DataGridViewCellEventArgs e)
            {
            if (dataGridView_HII.DataSource != null && e.RowIndex != -1 && ManualHII_CSV_list.Count > 0 && (e.RowIndex != dataGridView_HII.Rows.Count - 1 || (e.RowIndex == 0 && dataGridView_HII.Rows.Count == 1)))
                {
                _index_selected = e.RowIndex;
                cb_agent.SelectedItem = ListAgent.Where(c => c.Id.Equals(ManualHII_CSV_list[_index_selected].Agent_id)).FirstOrDefault();
                cb_commissiontype.Text = ManualHII_CSV_list[_index_selected].Payroll_Type;

                cb_PayrollDate.SelectedItem = calendar_SF.Where(c => c.Payment_Date__c == ManualHII_CSV_list[_index_selected].Payroll_date).FirstOrDefault();

                tb_policynumber.Text = ManualHII_CSV_list[_index_selected].Sf_MemberID;
                tb_Amount.Text = ManualHII_CSV_list[_index_selected].Agent_Commision;



                }
            else
                {
                _index_selected = -1;
                }

            }

        private void Form_Manual_Insert_Load(object sender, EventArgs e)
            {
            for (int i = 0; i < ListAgent.Count; i++)
                {
                cb_agent.Items.Add(ListAgent[i]);
                }

            for (int i = 0; i < calendar_SF.Count; i++)
                {
                cb_PayrollDate.Items.Add(calendar_SF[i]);
                }

            ManualHII_CSV_list = HII_CSV_list.Where(c => c.IsManual == true).ToList();

            dataGridView_HII.AutoGenerateColumns = false;

            if (ManualHII_CSV_list != null && ManualHII_CSV_list.Count > 0)
                {

                dataGridView_HII.DataSource = ManualHII_CSV_list;
                }
            clearitem();
            if ((Agent_id_Principal != null && Agent_id_Principal != "") || (Policy_number_Principal != null && Policy_number_Principal != ""))
                {
                chb_new.Checked = true;

                if (Agent_id_Principal != null && Agent_id_Principal != "")
                    {
                    Agent__c _Agent__c = ListAgent.Where(c => c.Id == Agent_id_Principal).FirstOrDefault();
                    if (_Agent__c != null)
                        {
                        cb_agent.SelectedItem = _Agent__c;
                        }
                    }
                if (Policy_number_Principal != null && Policy_number_Principal != "")
                    {
                    tb_policynumber.Text = Policy_number_Principal;
                    tb_policynumber.Focus();
                    }

                if (Amount != null && Amount != "")
                    {
                    tb_Amount.Text = Amount;
                    
                    }

                }

            }

        private void button2_Click(object sender, EventArgs e)
            {
            updateList();
            this.Close();
            }



        private void btn_delete_Click(object sender, EventArgs e)
            {
            if (_index_selected != -1)
                {
                ManualHII_CSV_list.RemoveAt(_index_selected);
                dataGridView_HII.DataSource = null;
                dataGridView_HII.Refresh();
                dataGridView_HII.DataSource = ManualHII_CSV_list;
                dataGridView_HII.Refresh();
                _index_selected = -1;
                }

            }

        private void dataGridView_HII_CellClick(object sender, DataGridViewCellEventArgs e)
            {
            cellClicOrCellSelect(e);
            }

        private void dataGridView_HII_CellEnter(object sender, DataGridViewCellEventArgs e)
            {
            cellClicOrCellSelect(e);
            }

        public void clearitem()
            {
            tb_policynumber.Text = "";
            cb_agent.SelectedIndex = -1;

            cb_commissiontype.SelectedIndex = -1;
            cb_PayrollDate.SelectedIndex = -1;
            tb_Amount.Text = "0";

            _index_selected = -1;
            }


        private void btn_Save_Click(object sender, EventArgs e)
            {

            if (cb_commissiontype.SelectedIndex != -1 && tb_policynumber.Text.Trim() != "" && cb_agent.SelectedIndex != -1 && cb_PayrollDate.SelectedIndex != -1 && (tb_Amount.Text.Trim() != "0" || tb_Amount.Text != ""))
                {
                if (chb_new.Checked)
                    {

                    Import_HII_CSV _Import_HII_CSV = new Import_HII_CSV();

                    _Import_HII_CSV.Agent_id = ((Agent__c)cb_agent.SelectedItem).Id;
                    _Import_HII_CSV.Agent_Fullname = ((Agent__c)cb_agent.SelectedItem).First_Name__c + ' ' + ((Agent__c)cb_agent.SelectedItem).Last_Name__c;


                    _Import_HII_CSV.Payroll_Type = cb_commissiontype.Text;

                    _Import_HII_CSV.Payroll_date = Convert.ToDateTime(cb_PayrollDate.Text);

                    _Import_HII_CSV.Member_ID = _Import_HII_CSV.Sf_MemberID_Bkup = _Import_HII_CSV.Sf_MemberID = tb_policynumber.Text;
                    _Import_HII_CSV.Agent_Commision = tb_Amount.Text;
                    if (_Import_HII_CSV.Payroll_Type == "Chargeback"   ||  _Import_HII_CSV.Payroll_Type =="Terminated")
                        {
                        double val = Convert.ToDouble(tb_Amount.Text);
                        if (val > 0)
                            {
                            _Import_HII_CSV.Agent_Commision = (val * -1).ToString("N");
                            }
                        _Import_HII_CSV.Termination_Date = DateTime.Today.ToShortDateString();
                        }

                    _Import_HII_CSV.IsUpdate = false;

                    _Import_HII_CSV.IsManual = true;
                    _Import_HII_CSV.Verify = true;

                    ManualHII_CSV_list.Add(_Import_HII_CSV);
                    dataGridView_HII.DataSource = null;
                    dataGridView_HII.Refresh();
                    dataGridView_HII.DataSource = ManualHII_CSV_list;
                    dataGridView_HII.Refresh();
                    clearitem();
                    chb_new.Checked = false;

                    }
                else
                    {
                    if (_index_selected != -1)
                        {

                        ManualHII_CSV_list[_index_selected].Agent_id = ((Agent__c)cb_agent.SelectedItem).Id;
                        ManualHII_CSV_list[_index_selected].Agent_Fullname = ((Agent__c)cb_agent.SelectedItem).First_Name__c + ' ' + ((Agent__c)cb_agent.SelectedItem).Last_Name__c;


                        ManualHII_CSV_list[_index_selected].Payroll_Type = cb_commissiontype.Text;

                        ManualHII_CSV_list[_index_selected].Payroll_date = Convert.ToDateTime(cb_PayrollDate.Text);

                        ManualHII_CSV_list[_index_selected].Member_ID = ManualHII_CSV_list[_index_selected].Sf_MemberID_Bkup = ManualHII_CSV_list[_index_selected].Sf_MemberID = tb_policynumber.Text;
                        ManualHII_CSV_list[_index_selected].Agent_Commision = tb_Amount.Text;

                        clearitem();
                        }

                    }

                }
            }

        private void Form_Manual_Insert_FormClosed(object sender, FormClosedEventArgs e)
            {
            updateList();
            }


        public void updateList()
            {

            // var templist = HII_CSV_list.FindAll(c => c.IsManual == true).ToList();
            HII_CSV_list.RemoveAll(c => c.IsManual == true);
            HII_CSV_list = HII_CSV_list.Concat(ManualHII_CSV_list).ToList();

            }

        private void chb_new_CheckedChanged(object sender, EventArgs e)
            {
            clearitem();
            }

        private void tb_Amount_KeyPress(object sender, KeyPressEventArgs e)
            {
            if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar)) && (e.KeyChar != '.') && (e.KeyChar != '-'))
                e.Handled = true;

            // only allow one decimal point
            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
                e.Handled = true;

            // only allow minus sign at the beginning
            if (e.KeyChar == '-' && (sender as TextBox).Text.Length > 0)
                e.Handled = true;
            }





        private void validateTextDouble(object sender, EventArgs e)
            {
            Exception X = new Exception();

            TextBox T = (TextBox)sender;

            try
                {
                if (T.Text != "-")
                    {
                    double x = double.Parse(T.Text);

                    if (T.Text.Contains(','))
                        throw X;
                    }
                }
            catch (Exception)
                {
                try
                    {
                    int CursorIndex = T.SelectionStart - 1;
                    T.Text = T.Text.Remove(CursorIndex, 1);

                    //Align Cursor to same index
                    T.SelectionStart = CursorIndex;
                    T.SelectionLength = 0;
                    }
                catch (Exception) { }
                }
            } 
        
        
        }






    }


