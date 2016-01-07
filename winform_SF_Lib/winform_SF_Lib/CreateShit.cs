using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using winform_SF_Lib.com.salesforce.na5;

namespace winform_SF_Lib
    {
    class CreateShit
        {
        List<List<Shit>> Shit_Matrix { set; get; }




        public Microsoft.Office.Interop.Excel.Application createExcel(bool IsvisibleFile, string paydate, string _profile)
            {
            var border = System.Drawing.Color.Black.ToArgb();
            SimpleExcelInterface _simpleExcel = new SimpleExcelInterface(IsvisibleFile);
            System.Drawing.Color _white = System.Drawing.Color.White;
            System.Drawing.Color _red = System.Drawing.Color.DarkOrange;
            System.Drawing.Color _pink = System.Drawing.Color.Magenta;
            System.Drawing.Color _black = System.Drawing.Color.Black;


            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(1);

            List<string> names_list = new List<string>(); 




            Microsoft.Office.Interop.Excel.Sheets worksheet_List = workbook.Worksheets;
            if (Shit_Matrix != null && Shit_Matrix.Count() > 0 && Shit_Matrix[0].Count() > 0)
                {
                //_simpleExcel.addSubHeader(1, i + 2, _name, _simpleExcel.getCellLetter(i + 2) + "1", _simpleExcel.getCellLetter(i + 2) + "1", "");
                int count_sheet = 1;
                List<Shit> Shit_List_summary = Shit_Matrix[0];
                Shit_Matrix.Remove(Shit_List_summary);

                if (Shit_List_summary.Count > 0)
                    {
                    var worksheet_summary = (Microsoft.Office.Interop.Excel.Worksheet)worksheet_List.Add(worksheet_List[1], Type.Missing, Type.Missing, Type.Missing);
                    worksheet_summary.Name = "Earnings Summary";

                    SimpleExcelInterface.addSubHeader_static(1, 4, "HBO " + _profile + "- Summary & Details", worksheet_summary, false, 18, _white, null);
                    SimpleExcelInterface.addSubHeader_static(2, 6, "Earnings Summary", worksheet_summary, false, 18, _white, null);
                    SimpleExcelInterface.addSubHeader_static(3, 4, "Printed: " + DateTime.Today.ToString("MM.dd.yyyy"), worksheet_summary, false, 18, _white, null);
                    if (paydate != "")
                        {
                        SimpleExcelInterface.addSubHeader_static(2, 4, "PayDate: " + Convert.ToDateTime(paydate).ToString("MM.dd.yyyy"), worksheet_summary, false, 18, _white, null);
                        }

                    for (int i = 0; i < Shit_List_summary.Count; i++)
                        {
                        string format_cell_currency = null;
                        string format_cell_percent = null;
                       
                        System.Drawing.Color _color = System.Drawing.Color.White;
                        if (Shit_List_summary[i].First_Name == "Chargeback")
                            {
                            _color = _red;
                            format_cell_currency = "Currency_dollar";
                            format_cell_percent = "Percentage";
                            }
                        if (Shit_List_summary[i].First_Name == "Terminated")
                            {
                            _color = _pink;
                            format_cell_currency = "Currency_dollar";
                            format_cell_percent = "Percentage";
                            }
                        if (Shit_List_summary[i].First_Name == "Total Pay this Period")
                            {

                            format_cell_currency = "Currency_dollar";
                            format_cell_percent = "Percentage";
                            }
                        if (Shit_List_summary[i].First_Name == "Pushes" || Shit_List_summary[i].First_Name == "Prior Period Push" || Shit_List_summary[i].First_Name == "Terms")
                            {
                            _color = System.Drawing.Color.Aquamarine;
                            }
                        SimpleExcelInterface.addSubHeader_static(i + 4, 6, Shit_List_summary[i].First_Name, worksheet_summary, false, 18, _color, null);
                        SimpleExcelInterface.addSubHeader_static(i + 4, 7, Shit_List_summary[i].Last_Name, worksheet_summary, false, 18, _white, format_cell_currency);
                        SimpleExcelInterface.addSubHeader_static(i + 4, 8, Shit_List_summary[i].Effective_Date, worksheet_summary, false, 18, _white, format_cell_percent);
                        worksheet_summary.Columns.AutoFit();
                        }
                    }

                int row_increment = 2;
                Shit_Matrix = Shit_Matrix.OrderByDescending(x => x[0].Sortname).ToList();
                foreach (List<Shit> Shit_List in Shit_Matrix)
                    {
                    // Microsoft.Office.Interop.Excel.Worksheet worksheet;
                    var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)worksheet_List.Add(worksheet_List[1], Type.Missing, Type.Missing, Type.Missing);
                    worksheet.Columns.AutoFit();
                    Microsoft.Office.Interop.Excel.Range workSheet_range = null;
                    count_sheet++;
                    if (Shit_List.Count > 0)
                        {
                        Shit heather = Shit_List[0];
                        Shit_List.Remove(heather);
                        SimpleExcelInterface.addSubHeader_static(1, 3, "Profile :", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1, 4, _profile, worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1, 5, heather.Fronter, worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1, 6, heather.First_Name, worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1, 8, heather.Last_Name, worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1, 10, heather.Effective_Date, worksheet, false, 11, _white, null);
                        if (paydate != "")
                            {
                            SimpleExcelInterface.addSubHeader_static(1, 12, "PayDate: " + Convert.ToDateTime(paydate).ToString("MM.dd.yyyy"), worksheet, false, 11, _white, null);
                            }

                        if (Shit_List[0].Fronter.Length > 25)
                            {
                            string sheetname =Regex.Replace(Shit_List[0].Fronter.Substring(0, 25), "[^a-zA-Z0-9% ._]", string.Empty);
                            if (names_list.Contains(sheetname)){
                            sheetname = sheetname + DateTime.Today.ToShortTimeString();
                                }
                            worksheet.Name = sheetname;
                            names_list.Add(sheetname);
                            }
                        else
                            {
                            string sheetname =Regex.Replace(Shit_List[0].Fronter, "[^a-zA-Z0-9% ._]", string.Empty);
                            if (names_list.Contains(sheetname))
                                {
                                sheetname = sheetname + DateTime.Today.ToShortTimeString();
                                }
                            worksheet.Name = sheetname;
                            names_list.Add(sheetname);
                            }


                        //Add Subheader//


                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 1, "Counter", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 2, "Member_Id", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 3, "Group", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 4, "Fronter", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 5, "First Name", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 6, "Last Name", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 7, "Effective Date", worksheet, true, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 8, "Enrollment Date", worksheet, true, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 9, "Status", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 10, "Cancel Date", worksheet, true, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 11, "Plan", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 12, "Coverage_Type", worksheet, false, 11, _white, null);
                        SimpleExcelInterface.addSubHeader_static(1 + row_increment, 13, "Earnings", worksheet, false, 11, _white, null);

                        //****DATA/
                        List<string> chargeback_array = new List<string>();
                        List<string> terminated_array = new List<string>();
                        List<string> active_array = new List<string>();
                        string chargeback_formula = "=SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Chargeback\")";
                        string terminate_formula = "=SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Terminated\")";
                        string written_formula = "=SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Terminated\") + SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Active\")";
                        string active_formula = "=SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Active\")";
                        string percentage_formula = "=SUMIFS(M4:M" + Convert.ToInt16(Shit_List.Count - 1) + ", I4:I" + Convert.ToInt16(Shit_List.Count - 1) + ", \"Active\")";

                        int active_row = 0;
                        int charge_row = 0;
                        int written_row = 0;
                        int total_row = 0;
                        int new_total_row = 0;
                        for (int i = 0; i < Shit_List.Count; i++)
                            {

                            System.Drawing.Color _cellColor = _white;
                            System.Drawing.Color _fColor = _black;
                            if (Shit_List[i].Status == "Terminated")
                                {
                                _cellColor = _pink;
                                terminated_array.Add('M' + (i + 4).ToString());
                                }
                            if (Shit_List[i].Status == "Chargeback")
                                {
                                _cellColor = _red;
                                chargeback_array.Add('M' + (i + 4).ToString());
                                }

                            if (Shit_List[i].Status == "Active")
                                {

                                active_array.Add('M' + (i + 4).ToString());
                                }
                         
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 1, Shit_List[i].Counter, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 2, Shit_List[i].Member_Id, "", "", _cellColor, _fColor, worksheet, "", border);
                            if (Shit_List[i].Member_Id == "Chargeback")
                                {
                                charge_row = i + 4;
                                SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, (Convert.ToDouble(Shit_List[i].Group) * -1).ToString("N"), "", chargeback_formula, _white, _red, worksheet, "Currency_dollarP", border);
                                //  SimpleExcelInterface.addData_static(i + 2, 5, Shit_List[i].Fronter, "", "", _white, _red, worksheet, "");
                                SimpleExcelInterface.addData_static(i + 2 + row_increment, 2, Shit_List[i].Member_Id, "", "", _red, _fColor, worksheet, "", border);
                                }
                            else
                                {
                                if (Shit_List[i].Member_Id == "Terminated")
                                    {
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, (Convert.ToDouble(Shit_List[i].Group) * -1).ToString("N"), "", terminate_formula, _white, _black, worksheet, "Currency_dollarP", border);
                                    // SimpleExcelInterface.addData_static(i + 2, 5, Shit_List[i].Fronter, "", "", _pink,_fColor , worksheet, "");
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 2, Shit_List[i].Member_Id, "", "", _pink, _fColor, worksheet, "", border);
                                    }
                                else if (Shit_List[i].Member_Id == "Written")
                                    {
                                    written_row = i + 4;
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", written_formula, _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, Shit_List[i].Fronter, "", "", _cellColor, _fColor, worksheet, "", border);
                                    }

                                else if (Shit_List[i].Member_Id == "Active")
                                    {
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", active_formula, _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                    active_row = i + 4;
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, Shit_List[i].Fronter, "", "", _cellColor, _fColor, worksheet, "", border);
                                    }
                                else if (Shit_List[i].Member_Id == "Advances")
                                    {
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, Shit_List[i].Fronter, "", "", _cellColor, _fColor, worksheet, "", border);
                                    }
                                else if (Shit_List[i].Member_Id == "Total Pay this Period")
                                    {
                                    total_row = i + 4;
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", "=SUM(C" + active_row.ToString() + ",C" + charge_row.ToString() + ")", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, Shit_List[i].Fronter, "", "", _cellColor, _fColor, worksheet, "", border);
                                    }
                                else if (Shit_List[i].Member_Id == "Percentage")
                                    {
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", "=(C" + total_row.ToString() + ")/C" + written_row.ToString(), _cellColor, _fColor, worksheet, "Percentage",  -1);
                                    }
                                else
                                    {
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", "", _cellColor, _fColor, worksheet, "", border);
                                    SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, Shit_List[i].Fronter, "", "", _cellColor, _fColor, worksheet, "", border);
                                    }

                                }



                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 5, Shit_List[i].First_Name, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 6, Shit_List[i].Last_Name, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 7, Shit_List[i].Effective_Date, "", "", _cellColor, _fColor, worksheet, "Date", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 8, Shit_List[i].Enrollment_Date, "", "", _cellColor, _fColor, worksheet, "Date", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 9, Shit_List[i].Status, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 10, Shit_List[i].Cancel_Date, "", "", _cellColor, _fColor, worksheet, "Date", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 11, Shit_List[i].Plan, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 12, Shit_List[i].Coverage_Type, "", "", _cellColor, _fColor, worksheet, "", border);
                            SimpleExcelInterface.addData_static(i + 2 + row_increment, 13, Shit_List[i].Earnings, "", "", _cellColor, _fColor, worksheet, "Currency", border);

                           
                            if (Shit_List[i].Member_Id == "New Total:")
                                {
                                new_total_row = i + 4;
                                SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, Shit_List[i].Group, "", "=C" + total_row.ToString() + "-C" + (i+3).ToString() + "", _cellColor, _fColor, worksheet, "Currency", border);
                                }

                            if (Shit_List[i].Member_Id == "Payable To:")
                                {
                                SimpleExcelInterface.addData_static(i + 4 + row_increment, 2, "New Total:", "", "", _cellColor, _fColor, worksheet, "", border);
                                SimpleExcelInterface.addData_static(i + 4 + row_increment, 3, Shit_List[i].Group, "", "=(C" + new_total_row.ToString() + ")", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                
                                SimpleExcelInterface.addData_static(i + 5 + row_increment, 2, "New Push/Terms:", "", "", _cellColor, _fColor, worksheet, "", border);
                                SimpleExcelInterface.addData_static(i + 5 + row_increment, 3, "", "", "=(C" + (new_total_row - 1).ToString() + ")", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                
                                SimpleExcelInterface.addData_static(i + 7 + row_increment, 2, "Total Pay:","", "", _cellColor, _fColor, worksheet, "", border);
                                SimpleExcelInterface.addData_static(i + 7 + row_increment, 3, "", "", "=SUM(C" + (i + 4 + row_increment).ToString() + ",C" + (i + 5 + row_increment).ToString() + ")", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                                SimpleExcelInterface.addData_static(i + 7 + row_increment, 4, "", "", "=(C" + (i + 7 + row_increment).ToString() + ")/C" + written_row.ToString(), _cellColor, _fColor, worksheet, "Percentage", border);

                                }


                            }


                        }
                    worksheet.Columns.AutoFit();


                    }


                }


            app.Visible = true;
            return _simpleExcel.app;

            }

        public Microsoft.Office.Interop.Excel.Application createExcel_salesboard(bool IsvisibleFile, string paydate, string _profile, List<Salesboard_Vs_Payment> listboard)
            {
            var border = System.Drawing.Color.Black.ToArgb();
            SimpleExcelInterface _simpleExcel = new SimpleExcelInterface(IsvisibleFile);
            System.Drawing.Color _white = System.Drawing.Color.White;
            System.Drawing.Color _red = System.Drawing.Color.DarkOrange;
            System.Drawing.Color _pink = System.Drawing.Color.Magenta;
            System.Drawing.Color _black = System.Drawing.Color.Black;


            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(1);


            Microsoft.Office.Interop.Excel.Sheets worksheet_List = workbook.Worksheets;

            //_simpleExcel.addSubHeader(1, i + 2, _name, _simpleExcel.getCellLetter(i + 2) + "1", _simpleExcel.getCellLetter(i + 2) + "1", "");
            int count_sheet = 1;
            int row_increment = 3;

            var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)worksheet_List.Add(worksheet_List[1], Type.Missing, Type.Missing, Type.Missing);

            Microsoft.Office.Interop.Excel.Range workSheet_range = null;
            worksheet.Name = _profile + paydate.Replace('/', '.');

            SimpleExcelInterface.addSubHeader_static(1, 1, "HBO " + _profile + "-  Details Summary ", worksheet, false, 12, _white, null);
            SimpleExcelInterface.addSubHeader_static(2, 1, "Printed: " + DateTime.Today.ToString("MM.dd.yyyy"), worksheet, false, 12, _white, null);
            SimpleExcelInterface.addSubHeader_static(2, 3, "Paydate: " + Convert.ToDateTime(paydate).ToString("MM.dd.yyyy"), worksheet, false, 12, _white, null);


            //Add Subheader//


            // SimpleExcelInterface.addSubHeader_static(1 + row_increment, 1, "Enable", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 1, "Name", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 2, "Total Written", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 3, "Terminated", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 4, "Chargeback", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 5, "Total", worksheet, false, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 6, "Per board", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 7, "Diff", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 8, "", worksheet, true, 11, System.Drawing.Color.Yellow, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 9, "Prior Period Adjustment", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 10, "New Total", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 11, "New push", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 12, "New Terms", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 13, "Total Pay", worksheet, true, 11, _white, null);
            SimpleExcelInterface.addSubHeader_static(1 + row_increment, 14, "%", worksheet, true, 11, _white, null);

            //****DATA/
            System.Drawing.Color _cellColor = _white;
            System.Drawing.Color _fColor = _black;
            System.Drawing.Color _cellColor_blue = System.Drawing.Color.LightBlue;
            for (int i = 0; i < listboard.Count; i++)
                {


                if (i == 0)
                    {
                    _cellColor = _cellColor_blue;
                    }
                else
                    {
                    _cellColor = _white;
                    }
                // SimpleExcelInterface.addData_static(i + 2 + row_increment, 1, listboard[i].Enable, "", "", _cellColor, _fColor, worksheet, "");
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 1, listboard[i].Fronter, "", "", _cellColor, _fColor, worksheet, "",border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 2, listboard[i].TotalWritten.ToString("N"), "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 3, listboard[i].Terminated.ToString("N"), "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 4, listboard[i].Chargeback.ToString("N"), "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 5, listboard[i].Total.ToString("N"), "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 6, listboard[i].Board.ToString("N"), "", "", _cellColor, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 7, listboard[i].Diff.ToString("N"), "", "", _cellColor_blue, _fColor, worksheet, "Currency_dollar", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 8, "", "", "", System.Drawing.Color.Yellow, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 9, "0", "", "", _cellColor, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 10, "0", "", "", _cellColor, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 11, "0", "", "", _cellColor, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 12, "0", "", "", _cellColor, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 13, "0", "", "", _cellColor, _fColor, worksheet, "", border);
                SimpleExcelInterface.addData_static(i + 2 + row_increment, 14, "0", "", "", _cellColor, _fColor, worksheet, "", border);

                }


            worksheet.Columns.AutoFit();

            app.Visible = true;
            return app;// _simpleExcel.app;

            }

        public void FillMatrix(List<Payment__c> payment_list, List<OpportunityLineItem> Oppline)
            {
            Shit_Matrix = new List<List<Shit>>();
            List<string> Agent_list = new List<string>();
            Agent_list = (from x in payment_list
                          select x.Agent__c).Distinct().ToList();
            Double summary_written = 0;
            Double summary__paythisperiod = 0;
            Double summary__chargeback = 0;
            Double summary__terminated = 0;

            foreach (string agent_item in Agent_list)
                {
                List<Shit> Shit_list = new List<Shit>();
                List<Shit> Shit_list_terminated = new List<Shit>();
                List<Shit> Shit_list_Cancel = new List<Shit>();
                List<Payment__c> payment_temp = payment_list.Where(c => c.Agent__c == agent_item).ToList();

                Shit header = new Shit();
                if (payment_temp.Count > 0)
                    {
                    header.Sortname = payment_temp[0].Agent__r.First_Name__c + " " + payment_temp[0].Agent__r.Last_Name__c;
                    header.Member_Id = "Profile :";
                    header.Group = payment_temp[0].Agent__r.product_profile__c;
                    header.Fronter = "Shift:";
                    header.First_Name = payment_temp[0].Agent__r.Time_Shift__c;
                    header.Last_Name = "Created:" + Convert.ToDateTime(payment_temp[0].CreatedDate).ToString("MM.dd.yyyy");
                    header.Effective_Date = "Printed:" + DateTime.Today.ToString("MM.dd.yyyy");

                    }

                Double _written = 0;
                Double _active = 0;
                Double _chargeback = 0;
                Double _terminated = 0;
                int _count = 1;
                #region Agent
                for (int i = 0; i < payment_temp.Count; i++)
                    {
                    Shit _Shit = new Shit();
                    /* if (payment_temp[i].Policy_Number__c == "DTX0251900")
                         {
                         var x="";
                         }*/
                    _Shit.Member_Id = payment_temp[i].Policy_Number__c;
                    _Shit.Fronter = payment_temp[i].Agent__r.First_Name__c + " " + payment_temp[i].Agent__r.Last_Name__c;
                    _Shit.Sortname = payment_temp[i].Agent__r.First_Name__c + " " + payment_temp[i].Agent__r.Last_Name__c;
                    _Shit.Status = payment_temp[i].Payment_Type__c == "Commission" ? "Active" : payment_temp[i].Payment_Type__c;
                    _Shit.Earnings = payment_temp[i].Payment_Value__c.ToString();//.ToString();
                    _Shit.Cancel_Date = payment_temp[i].Cancel_Date__c == null ? string.Empty : Convert.ToDateTime(payment_temp[i].Cancel_Date__c).ToString("MM/dd/yyyy");
                    _Shit.PaymentCompany = payment_temp[i].Agent__r.Payment_company__c;
                    OpportunityLineItem item_Line = new OpportunityLineItem();
                    try 
                        {
                         item_Line = Oppline.Where(c => c.Policy_Number__c == payment_temp[i].Policy_Number__c).FirstOrDefault();
                        }
                    catch (Exception ex)
                        {
                        string message = ex.Message;
                        throw new Exception(message + "Police :" + payment_temp[i].Policy_Number__c);
                        }
                    
                    if (item_Line == null)
                        {
                        item_Line = Oppline.Where(c => c.Id == payment_temp[i].OpportunityLineItem_id__c).FirstOrDefault();
                        }


                    if (item_Line != null)
                        {
                        _Shit.Group = item_Line.Product2.Sales_Board_Category__c;
                        _Shit.Plan = item_Line.Product2.Sales_Board_Category__c == "Accident Insurance" ? "Accident Insurance" : item_Line.Product2.Plan_Type__c;

                        string plan_value = _Shit.Plan;








                        switch (plan_value.ToUpper().Trim())
                            {
                            case "GUARANTEED ISSUE":
                                    {
                                    _Shit.SortHealth = 1;
                                    break;
                                    }
                            case "SHORT TERM MEDICAL":
                                    {
                                    _Shit.SortHealth = 2;
                                    break;
                                    }
                            case "SIMPLEFIED ISSUE":
                                    {
                                    _Shit.SortHealth = 3;
                                    break;
                                    }
                            case "DENTAL & RX":
                                    {
                                    _Shit.SortHealth = 4;
                                    break;
                                    }
                            case "ACCIDENTAL INSURANCE":
                                    {
                                    _Shit.SortHealth = 5;
                                    break;
                                    }
                            default:
                                    {
                                    _Shit.SortHealth = 99;
                                    break;
                                    }
                            }

                        if (item_Line.Product2.eApp_Display_Name__c.Contains(':'))
                            {
                            string[] eappdisplay_name = item_Line.Product2.eApp_Display_Name__c.Split(':');
                            _Shit.Coverage_Type = eappdisplay_name[1];
                            }
                        else
                            {
                            if (item_Line.Product2.eApp_Display_Name__c.Contains("("))
                                {
                                string[] eappdisplay_name = item_Line.Product2.eApp_Display_Name__c.Split('(');
                                _Shit.Coverage_Type = eappdisplay_name[1].Replace(")", string.Empty);
                                }
                            else
                                {
                                _Shit.Coverage_Type = " Individual"; //item_Line.Product2.eApp_Display_Name__c;
                                }

                            }

                        _Shit.Enrollment_Date = Convert.ToDateTime(item_Line.Enrollment_Date__c).ToString("MM/dd/yyyy");
                        _Shit.Effective_Date = Convert.ToDateTime(item_Line.Effective_Date__c).ToString("MM/dd/yyyy");

                        _Shit.First_Name = item_Line.Opportunity.Account.FirstName;
                        _Shit.Last_Name = item_Line.Opportunity.Account.LastName;

                        if (payment_temp[i].Payment_Type__c != "Chargeback" && item_Line.Product2.Current_Running_Totals_Category__c != null && (item_Line.Product2.Current_Running_Totals_Category__c.Contains("1st Med STM") || item_Line.Product2.Current_Running_Totals_Category__c.Contains("HealtheMed STM") || item_Line.Product2.Current_Running_Totals_Category__c.Contains("Principle Advantage")))
                            {
                            //('1st Med STM','HealtheMed STM') or GI_Product__r.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )
                            //  _Shit.Counter = _count.ToString();
                            //    _count++;
                            }
                        }

                    if (payment_temp[i].Payment_Type__c != "Terminated" && payment_temp[i].Payment_Type__c != "Chargeback")
                        {
                        _written += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list.Add(_Shit);
                        }
                    if (payment_temp[i].Payment_Type__c == "Terminated")
                        {
                        _terminated += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list_terminated.Add(_Shit);

                        }
                    if (payment_temp[i].Payment_Type__c == "Chargeback")
                        {
                        _chargeback += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list_Cancel.Add(_Shit);
                        }




                    }
                #endregion
                if (Shit_list_terminated.Count > 0 && Shit_list.Count > 0)
                    {
                    foreach (Shit item in Shit_list_terminated)
                        {
                        Shit ter_shit = Shit_list.Where(c => c.Member_Id == item.Member_Id).FirstOrDefault();
                        if (ter_shit != null)
                            {
                            ter_shit.Status = item.Status;
                            ter_shit.Cancel_Date = item.Cancel_Date;
                            }
                        }
                    }


                Shit_list = Shit_list.OrderBy(c => c.Effective_Date).ThenByDescending(x => x.First_Name).ThenByDescending(l => l.Last_Name).ThenBy(l => l.SortHealth).ToList();
                int _no = 0;
                string memberid_exception="";
                try
                    {
                    foreach (Shit item in Shit_list)
                        {
                        memberid_exception= item.Member_Id;
                        if (item.Group != null && item.Group == "Principle Advantage" || item.Group == "Standard Life" || item.Group == "GetMed360 Freedom" || item.Group == "HealtheMed STM" || item.Group == "1st Med STM" || item.Group.Contains("Unified"))
                            {
                            _no++;
                            item.Counter = _no.ToString();

                            }
                        }
                    }
                catch(Exception ex) 
                    {
                    string message = ex.Message;
                    throw new Exception(message + " Salesboard Category missing in Police :" + memberid_exception);
                    }
               

                Shit_list_Cancel = Shit_list_Cancel.OrderBy(c => c.Enrollment_Date).ThenByDescending(x => x.First_Name).ThenByDescending(l => l.Last_Name).ThenByDescending(l => l.SortHealth).ToList();
                Shit_list = Shit_list.Concat(Shit_list_Cancel).ToList();

                Shit_list.Add(new Shit());
                Shit_list.Add(new Shit() { Member_Id = payment_temp[0].Agent__r.First_Name__c + " " + payment_temp[0].Agent__r.Last_Name__c, Group = (Convert.ToDateTime(payment_temp[0].Payment_Date__c)).ToString("MM.dd.yyyy") });

                Shit_list.Add(new Shit() { Member_Id = "Written", Group = (_written).ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Active", Group = (_written + _terminated).ToString("N") });

                Shit_list.Add(new Shit() { Member_Id = "Terminated", Group = _terminated.ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Chargeback", Group = _chargeback.ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Advances", Group = "" });
                Shit_list.Add(new Shit() { Member_Id = "Total Pay this Period", Group = ((_written + _terminated) + _chargeback).ToString("N") });

                Shit_list.Add(new Shit());
                Shit_list.Add(new Shit() { Member_Id = "Percentage", Group = (((_written + _terminated) + _chargeback) * 100 / _written).ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "", Group = "" });
                Shit_list.Add(new Shit() { Member_Id = "Push/Terms from " });
                Shit_list.Add(new Shit() { Member_Id = "New Total:" });

                Shit_list.Add(new Shit() { Member_Id = "Payable To:", Group = payment_temp[0].Agent__r.Payment_company__c });

                summary__chargeback += _chargeback;
                summary__terminated += _terminated;
                summary_written += _written;
                summary__paythisperiod += ((_written + _terminated) + _chargeback);
                Shit_list.Insert(0, header);

                Shit_Matrix.Add(Shit_list);

                }
            List<Shit> Shit_list_summary = new List<Shit>();

            Shit_list_summary.Add(new Shit() { First_Name = "", Last_Name = "Dollar Ammount", Effective_Date = "Percentage" });
            Shit_list_summary.Add(new Shit() { First_Name = "Written", Last_Name = (summary_written).ToString("N"), Effective_Date = "100.00" });
            Shit_list_summary.Add(new Shit() { First_Name = "Chargeback", Last_Name = summary__chargeback.ToString("N"), Effective_Date = (100.00 * summary__chargeback / summary_written).ToString("N") });
            Shit_list_summary.Add(new Shit() { First_Name = "Terminated", Last_Name = summary__terminated.ToString("N"), Effective_Date = (100.00 * summary__terminated / summary_written).ToString("N") });
            Shit_list_summary.Add(new Shit() { First_Name = "Pushes", Last_Name = "", Effective_Date = "" });
            Shit_list_summary.Add(new Shit() { First_Name = "Prior Period Push", Last_Name = "", Effective_Date = "" });
            Shit_list_summary.Add(new Shit() { First_Name = "Terms", Last_Name = "", Effective_Date = "" });
            Shit_list_summary.Add(new Shit() { First_Name = "Total Pay this Period", Last_Name = summary__paythisperiod.ToString("N"), Effective_Date = (100.00 * summary__paythisperiod / summary_written).ToString("N") });

            Shit_Matrix = Shit_Matrix.OrderByDescending(x => x[0].Fronter).ToList();
            Shit_Matrix.Insert(0, Shit_list_summary);
            }

        public void FillMatrix1(List<Payment__c> payment_list, List<OpportunityLineItem> Oppline)
            {
            Shit_Matrix = new List<List<Shit>>();
            List<string> Agent_list = new List<string>();
            Agent_list = (from x in payment_list
                          select x.Agent__c).Distinct().ToList();
            List<Shit> Shit_list_Cancel = new List<Shit>();
            foreach (string agent_item in Agent_list)
                {
                List<Shit> Shit_list = new List<Shit>();
                List<Shit> Shit_list_terminated = new List<Shit>();

                List<Payment__c> payment_temp = payment_list.Where(c => c.Agent__c == agent_item).ToList();

                Double _written = 0;
                Double _active = 0;
                Double _chargeback = 0;
                Double _terminated = 0;
                int _count = 1;
                #region Agent
                for (int i = 0; i < payment_temp.Count; i++)
                    {
                    Shit _Shit = new Shit();
                    /* if (payment_temp[i].Policy_Number__c == "DTX0251900")
                         {
                         var x="";
                         }*/
                    _Shit.Member_Id = payment_temp[i].Policy_Number__c;
                    _Shit.Fronter = payment_temp[i].Agent__r.First_Name__c + " " + payment_temp[i].Agent__r.Last_Name__c;
                    _Shit.Sortname = payment_temp[i].Agent__r.First_Name__c + " " + payment_temp[i].Agent__r.Last_Name__c; _Shit.Status = payment_temp[i].Payment_Type__c == "Commission" ? "Active" : payment_temp[i].Payment_Type__c;
                    _Shit.Earnings = string.Format("{0:C}", payment_temp[i].Payment_Value__c);//.ToString();
                    _Shit.Cancel_Date = payment_temp[i].Cancel_Date__c == null ? string.Empty : Convert.ToDateTime(payment_temp[i].Cancel_Date__c).ToString("MM/dd/yyyy");
                    _Shit.PaymentCompany = payment_temp[i].Agent__r.Payment_company__c;
                    OpportunityLineItem item_Line = Oppline.Where(c => c.Policy_Number__c == payment_temp[i].Policy_Number__c).FirstOrDefault();
                    if (item_Line == null)
                        {
                        item_Line = Oppline.Where(c => c.Id == payment_temp[i].OpportunityLineItem_id__c).FirstOrDefault();
                        }


                    if (item_Line != null)
                        {
                        _Shit.Group = item_Line.Product2.Sales_Board_Category__c;
                        _Shit.Plan = item_Line.Product2.Sales_Board_Category__c == "Accident Insurance" ? "Accident Insurance" : item_Line.Product2.Plan_Type__c;

                        string plan_value = _Shit.Plan;


                        switch (plan_value.ToUpper().Trim())
                            {
                            case "GUARANTEED ISSUE":
                                    {
                                    _Shit.SortHealth = 1;
                                    break;
                                    }
                            case "SHORT TERM MEDICAL":
                                    {
                                    _Shit.SortHealth = 2;
                                    break;
                                    }
                            case "SIMPLEFIED ISSUE":
                                    {
                                    _Shit.SortHealth = 3;
                                    break;
                                    }
                            case "DENTAL & RX":
                                    {
                                    _Shit.SortHealth = 4;
                                    break;
                                    }
                            case "ACCIDENTAL INSURANCE":
                                    {
                                    _Shit.SortHealth = 5;
                                    break;
                                    }
                            default:
                                    {
                                    _Shit.SortHealth = 99;
                                    break;
                                    }
                            }

                        if (item_Line.Product2.eApp_Display_Name__c.Contains(':'))
                            {
                            string[] eappdisplay_name = item_Line.Product2.eApp_Display_Name__c.Split(':');
                            _Shit.Coverage_Type = eappdisplay_name[1];
                            }
                        else
                            {
                            if (item_Line.Product2.eApp_Display_Name__c.Contains("("))
                                {
                                string[] eappdisplay_name = item_Line.Product2.eApp_Display_Name__c.Split('(');
                                _Shit.Coverage_Type = eappdisplay_name[1].Replace(")", string.Empty);
                                }
                            else
                                {
                                _Shit.Coverage_Type = " Individual"; //item_Line.Product2.eApp_Display_Name__c;
                                }

                            }

                        _Shit.Enrollment_Date = Convert.ToDateTime(item_Line.Enrollment_Date__c).ToString("MM/dd/yyyy");
                        _Shit.Effective_Date = Convert.ToDateTime(item_Line.Effective_Date__c).ToString("MM/dd/yyyy");

                        _Shit.First_Name = item_Line.Opportunity.Account.FirstName;
                        _Shit.Last_Name = item_Line.Opportunity.Account.LastName;

                        if (payment_temp[i].Payment_Type__c != "Chargeback" && item_Line.Product2.Current_Running_Totals_Category__c != null && (item_Line.Product2.Current_Running_Totals_Category__c.Contains("1st Med STM") || item_Line.Product2.Current_Running_Totals_Category__c.Contains("HealtheMed STM") || item_Line.Product2.Current_Running_Totals_Category__c.Contains("Principle Advantage")))
                            {
                            //('1st Med STM','HealtheMed STM') or GI_Product__r.CURRENT_RUNNING_TOTALS_CATEGORY__C like '%Principle Advantage%' )
                            //  _Shit.Counter = _count.ToString();
                            //    _count++;
                            }
                        }

                    if (payment_temp[i].Payment_Type__c != "Terminated" && payment_temp[i].Payment_Type__c != "Chargeback")
                        {
                        _written += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list.Add(_Shit);
                        }
                    if (payment_temp[i].Payment_Type__c == "Terminated")
                        {
                        _terminated += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list_terminated.Add(_Shit);

                        }
                    if (payment_temp[i].Payment_Type__c == "Chargeback")
                        {
                        _chargeback += Convert.ToDouble(payment_temp[i].Payment_Value__c);
                        Shit_list_Cancel.Add(_Shit);
                        }




                    }
                #endregion
                if (Shit_list_terminated.Count > 0 && Shit_list.Count > 0)
                    {
                    foreach (Shit item in Shit_list_terminated)
                        {
                        Shit ter_shit = Shit_list.Where(c => c.Member_Id == item.Member_Id).FirstOrDefault();
                        if (ter_shit != null)
                            {
                            ter_shit.Status = item.Status;
                            ter_shit.Cancel_Date = item.Cancel_Date;
                            }
                        }
                    }


                Shit_list = Shit_list.OrderBy(c => c.Effective_Date).ThenByDescending(x => x.First_Name).ThenByDescending(l => l.Last_Name).ThenBy(l => l.SortHealth).ToList();
                int _no = 0;
                foreach (Shit item in Shit_list)
                    {
                    if (item.Group == "Principle Advantage" || item.Group == "Standard Life" || item.Group == "GetMed360 Freedom" || item.Group == "HealtheMed STM" || item.Group == "1st Med STM" || item.Group.Contains("Unified"))
                        {
                        _no++;
                        item.Counter = _no.ToString();

                        }
                    }

                Shit_list_Cancel = Shit_list_Cancel.OrderBy(c => c.Effective_Date).ThenByDescending(x => x.First_Name).ThenByDescending(l => l.Last_Name).ThenByDescending(l => l.SortHealth).ToList();
                //Shit_list = Shit_list.Concat(Shit_list_Cancel).ToList();

                Shit_list.Add(new Shit());
                Shit_list.Add(new Shit() { Member_Id = payment_temp[0].Agent__r.First_Name__c + " " + payment_temp[0].Agent__r.Last_Name__c, Group = (Convert.ToDateTime(payment_temp[0].Payment_Date__c)).ToString("MM.dd.yyyy") });

                Shit_list.Add(new Shit() { Member_Id = "Written", Group = (_written).ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Active", Group = (_written + _terminated).ToString("N") });

                Shit_list.Add(new Shit() { Member_Id = "Terminated", Group = _terminated.ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Chargeback", Group = _chargeback.ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "Advances", Group = "" });
                Shit_list.Add(new Shit() { Member_Id = "Total Pay this Period", Group = ((_written + _terminated) + _chargeback).ToString("N") });

                Shit_list.Add(new Shit());
                Shit_list.Add(new Shit() { Member_Id = "Percentage", Group = (((_written + _terminated) + _chargeback) * 100 / _written).ToString("N") });
                Shit_list.Add(new Shit() { Member_Id = "", Group = "" });
                Shit_list.Add(new Shit() { Member_Id = "Push/Terms" });
                Shit_list.Add(new Shit() { Member_Id = "New Total:" });

                Shit_list.Add(new Shit() { Member_Id = "Payable To:", Group = payment_temp[0].Agent__r.Payment_company__c });

                //Shit_Matrix.Add(Shit_list);

                }

            Shit_Matrix.Add(Shit_list_Cancel);


            }

        public String formatString(List<string> NumberList)
            {
            string policyformated = "";
            for (int i = 0; i < NumberList.Count; i++)
                {
                if (i + 1 == NumberList.Count)
                    {
                    policyformated += NumberList[i];
                    }
                else
                    {
                    policyformated += NumberList[i] + ",";
                    }

                }
            return policyformated;

            }

        }


    class Shit
        {
        public string Sortname { set; get; }
        public string Counter { set; get; }
        public string Member_Id { set; get; }
        public string Group { set; get; }
        public string Fronter { set; get; }
        public string First_Name { set; get; }
        public string Last_Name { set; get; }
        public string Effective_Date { set; get; }
        public string Enrollment_Date { set; get; }
        public string Status { set; get; }
        public string Cancel_Date { set; get; }
        public string Plan { set; get; }
        public string Coverage_Type { set; get; }
        public string Earnings { set; get; }
        public int SortHealth { set; get; }
        public string PaymentCompany { set; get; }
        }
    class Salesboard_Vs_Payment
        {
        public string Enable { set; get; }
        public string Id { set; get; }
        public string Fronter { set; get; }
        public double TotalWritten { set; get; }
        public double Terminated { set; get; }
        public double Chargeback { set; get; }

        public double Total
            {

            get
                {
                return TotalWritten + Terminated + Chargeback;
                }
            }
        public double Board { set; get; }
        public double Diff
            {
            get
                {
                return Board - TotalWritten;
                }
            }
        }

    }
