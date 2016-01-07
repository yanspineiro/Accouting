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

namespace winform_SF_Lib
    {
    public partial class FormMatch : Form
        {
        public List<Import_HII_CSV> Saleforce_list { set; get; }
        public List<Import_HII_CSV> CarrierList { set; get; }
        public List<Import_HII_CSV> SF_HII_list { set; get; }
        public List<Linked_class> Linked_List { set; get; }
        public FormMatch()
            {
            InitializeComponent();
            SF_HII_list = new List<Import_HII_CSV>();
            Linked_List = new List<Linked_class>();
            }
        FormProgress Proggre;
        private void button1_Click(object sender, EventArgs e)
            {
            if (dataGridView_SF.SelectedRows.Count != 0 && dataGridView_Carrier.SelectedRows.Count != 0)
                {
                DataGridViewRow row_sf = this.dataGridView_SF.SelectedRows[0];
                DataGridViewRow row_Carrier = this.dataGridView_Carrier.SelectedRows[0];
                string Sf_MemberID = row_sf.Cells["ColumnSf_MemberID"].Value.ToString();
                string Member_ID = row_Carrier.Cells["Column_Member_ID"].Value.ToString();
                if (Member_ID != null && Sf_MemberID != null)
                    {
                    Import_HII_CSV temp_sf = Saleforce_list.Where(c => c.Sf_MemberID == Sf_MemberID).FirstOrDefault();
                    SF_HII_list.Add(temp_sf);
                    Saleforce_list.Remove(temp_sf);
                    dataGridView_SF.DataSource = null;
                    dataGridView_SF.Refresh();
                    dataGridView_SF.DataSource = Saleforce_list;
                    dataGridView_SF.Refresh();

                    Import_HII_CSV temp_carrier = CarrierList.Where(c => c.Member_ID == Member_ID).FirstOrDefault();
                    SF_HII_list.Add(temp_carrier);
                    CarrierList.Remove(temp_carrier);

                    dataGridView_Carrier.DataSource = null;
                    dataGridView_Carrier.Refresh();
                    dataGridView_Carrier.DataSource = CarrierList;
                    dataGridView_Carrier.Refresh();
                    Linked_class _Linked_class = new Linked_class();
                    _Linked_class.Sf_MemberID = temp_sf.Sf_MemberID;
                    _Linked_class.SF_First_Name = temp_sf.First_Name;
                    _Linked_class.SF_phone = temp_sf.Phone;
                    _Linked_class.SF_Application_Date = temp_sf.Application_Date;

                    _Linked_class.Carrier_FirstName = temp_carrier.First_Name;
                    _Linked_class.Carrier_Phone = temp_carrier.Phone;
                    _Linked_class.Carrier_Application_Date = temp_carrier.Application_Date;
                    _Linked_class.Carrier_Policynumber = temp_carrier.Member_ID;
                    Linked_List.Add(_Linked_class);
                    dataGridViewLinked.DataSource = null;
                    dataGridViewLinked.Refresh();
                    dataGridViewLinked.DataSource = Linked_List;
                    dataGridViewLinked.Refresh();


                    }
                else
                    {
                    MessageBox.Show("Must Select one valid row in each grid for Link");
                    }
                }
            else
                {
                MessageBox.Show("Must Select one valid row in each grid for Link");
                }
            }

        private void FormMatch_Load(object sender, EventArgs e)
            {

            dataGridView_Carrier.AutoGenerateColumns = false;
            dataGridView_Carrier.DataSource = CarrierList;
            dataGridView_SF.AutoGenerateColumns = false;
            dataGridView_SF.DataSource = Saleforce_list;
            dataGridViewLinked.AutoGenerateColumns = false;

            }

        private void textBox_SF_KeyPress(object sender, KeyPressEventArgs e)
            {


            }

        private void textBox_SF_TextChanged(object sender, EventArgs e)
            {
            if (textBox_SF.Text != "")
                {
                dataGridView_SF.DataSource = null;
                dataGridView_SF.Refresh();
                var temp = Saleforce_list.Where(c => c.Sf_MemberID.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_SF.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_SF.Text.ToUpper())).ToList();
                dataGridView_SF.DataSource = temp;
                dataGridView_SF.Refresh();
                }
            else
                {
                dataGridView_SF.DataSource = null;
                dataGridView_SF.Refresh();
                dataGridView_SF.DataSource = Saleforce_list;

                dataGridView_SF.Refresh();
                }
            }

        private void textBox_Carrier_TextChanged(object sender, EventArgs e)
            {
            if (textBox_Carrier.Text != "")
                {
                dataGridView_Carrier.DataSource = null;
                dataGridView_Carrier.Refresh();
                var temp = CarrierList.Where(c => c.Member_ID.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.First_Name.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Application_Date.ToUpper().Contains(textBox_Carrier.Text.ToUpper()) || c.Phone.ToUpper().Contains(textBox_Carrier.Text.ToUpper())).ToList();
                dataGridView_Carrier.DataSource = temp;
                dataGridView_Carrier.Refresh();
                }
            else
                {
                dataGridView_Carrier.DataSource = null;
                dataGridView_Carrier.Refresh();
                dataGridView_Carrier.DataSource = CarrierList;

                dataGridView_Carrier.Refresh();
                }
            }

        private void unbindToolStripMenuItem_Click(object sender, EventArgs e)
            {

            if (dataGridViewLinked.SelectedRows.Count != 0)
                {
                DataGridViewRow row_linked = this.dataGridViewLinked.SelectedRows[0];
                string Sf_MemberID = row_linked.Cells["dataGridViewTextBoxColumn4"].Value.ToString();
                string Carrier_Policynumber = row_linked.Cells["ColumnMember_ID"].Value.ToString();

                var temp_sf = SF_HII_list.Where(c => c.Sf_MemberID == Sf_MemberID).FirstOrDefault();
                Saleforce_list.Add(temp_sf);
                var temp_Carrier = SF_HII_list.Where(c => c.Member_ID == Carrier_Policynumber).FirstOrDefault();
                CarrierList.Add(temp_Carrier);

                var _linked = Linked_List.Where(c => c.Sf_MemberID == Sf_MemberID && c.Carrier_Policynumber == Carrier_Policynumber).FirstOrDefault();
                Linked_List.Remove(_linked);

                dataGridViewLinked.DataSource = null;
                dataGridView_Carrier.DataSource = null;
                dataGridView_SF.DataSource = null;

                dataGridViewLinked.Refresh();
                dataGridView_Carrier.Refresh();
                dataGridView_SF.Refresh();

                dataGridViewLinked.DataSource = Linked_List;
                dataGridView_Carrier.DataSource = CarrierList;
                dataGridView_SF.DataSource = Saleforce_list;

                dataGridViewLinked.Refresh();
                dataGridView_Carrier.Refresh();
                dataGridView_SF.Refresh();
                }
            }

        private void button2_Click(object sender, EventArgs e)
            {
            this.Close();
            }

        private void btn_ExportReady_Click(object sender, EventArgs e)
            {
            backgroundWorkerFillExcel.RunWorkerAsync(dataGridViewLinked);
            this.Enabled = false;
            Proggre = new FormProgress();
            Proggre.TopLevel = true;
            Proggre.TopMost = true;
            Proggre.Show();
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


        }

    public class Linked_class
        {

        public string Sf_MemberID { get; set; }
        public string SF_First_Name { get; set; }
        public string SF_phone { get; set; }
        public string SF_Application_Date { get; set; }

        public string Carrier_FirstName { get; set; }
        public string Carrier_Phone { get; set; }
        public string Carrier_Application_Date { get; set; }
        public string Carrier_Policynumber { get; set; }
        }
    }
