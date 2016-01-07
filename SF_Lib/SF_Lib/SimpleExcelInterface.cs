using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace winform_SF_Lib
    {
    class SimpleExcelInterface
        {

        public Microsoft.Office.Interop.Excel.Application app = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        private Microsoft.Office.Interop.Excel.Range workSheet_range = null;
        private bool IsvisibleFile;
        public SimpleExcelInterface(bool _IsvisibleFile)
            {
            IsvisibleFile = _IsvisibleFile;
            }

        public bool createDoc(string _sheetName)
            {
            bool _succes = true;
            try
                {
                app = new Microsoft.Office.Interop.Excel.Application();

                app.Visible = IsvisibleFile;




                workbook = app.Workbooks.Add(1);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                if (_sheetName != "")
                    {
                    worksheet.Name = _sheetName;
                    }
                }
            catch
                {
                _succes = false;
                }

            return _succes;
            }

        public bool openDoc(string _path)
            {
            bool _succes = true;
            try
                {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Open(_path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                }
            catch
                {
                _succes = false;
                }
            return _succes;
            }

        public void showDocument()
            {
            app.Visible = true;
            }
        public void closeDocument()
            {
            app.Visible = false;
            }
        public void createHeaders(int row, int col, string htext, string cell1, string cell2, int mergeColumns, string b, bool font, int fontsize, int cellsize, string fcolor)
            {
            worksheet.Cells[row, col] = htext;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Merge(mergeColumns);
            switch (b)
                {
                case "YELLOW":
                    workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    workSheet_range.Interior.Color = System.Drawing.Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    workSheet_range.Interior.Color = System.Drawing.Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    workSheet_range.Interior.Color = System.Drawing.Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    workSheet_range.Interior.Color = System.Drawing.Color.PeachPuff.ToArgb();
                    break;
                default:
                    workSheet_range.Interior.Color = System.Drawing.Color.White.ToArgb();
                    break;
                }

            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Font.Bold = font;
            workSheet_range.Font.Size = fontsize;
            workSheet_range.ColumnWidth = cellsize;
            if (fcolor.Equals(""))
                {
                workSheet_range.Font.Color = System.Drawing.Color.White.ToArgb();
                }
            else
                {
                workSheet_range.Font.Color = System.Drawing.Color.Black.ToArgb();
                }
            }

        public void addData(int row, int col, string data, string cell1, string cell2, string format, string formula, System.Drawing.Color _color, System.Drawing.Color _fcolor)
            {

            worksheet.Cells[row, col] = data;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = _color;
            workSheet_range.NumberFormat = format;

            workSheet_range.Font.Color = _fcolor;

            if (formula != "")
                {
                workSheet_range.Formula = formula;
                }
            }

        public static void addData_static (int row, int col, string data,  string format, string formula, System.Drawing.Color _color, System.Drawing.Color _fcolor, Microsoft.Office.Interop.Excel.Worksheet worksheet, string  _format, int border)
            {

            worksheet.Cells[row, col] = data;
            Microsoft.Office.Interop.Excel.Range workSheet_range = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col], (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col]);
            if (border!= -1)
                {
                workSheet_range.Borders.Color = border;
                }
           
           
            workSheet_range.Interior.Color = _color;
            workSheet_range.NumberFormat = format;
       
            workSheet_range.Font.Color = _fcolor;

            if (formula != "")
                {
                workSheet_range.Formula = formula;
                }
            if (_format== "Date")
                {
                workSheet_range.EntireColumn.NumberFormat = "MM/DD/YYYY";
                }
            if (_format == "Currency")
                {
                workSheet_range.Cells.NumberFormat = "$##,###.00";
                }
            if (_format == "Currency_dollar")
                {

                workSheet_range.Cells.NumberFormat = "$##,###.00";
                }
            if (_format == "Currency_dollarP")
                {

                workSheet_range.Cells.NumberFormat = "($##,###.00)";
                }
            if (_format == "Percentage")
                {

                workSheet_range.Cells.NumberFormat = "##.00%";
                }
            }
            

        public void addData(int row, int col, string data, string cell1, string cell2, string format, string formula)
            {

            worksheet.Cells[row, col] = data;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();

            workSheet_range.NumberFormat = format;

            if (formula != "")
                {
                workSheet_range.Formula = formula;
                }
            }



        public void addSubHeader(int row, int col, string data, string cell1, string cell2, string format)
            {
            worksheet.Cells[row, col] = data;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Font.Bold = true;
            workSheet_range.NumberFormat = format;
            }

        public static void addSubHeader_static(int row, int col, string data,  Microsoft.Office.Interop.Excel.Worksheet worksheet, bool adate, int font_size,System.Drawing.Color _color_interior,string _format)
            {
            worksheet.Cells[row, col] = data;
            Microsoft.Office.Interop.Excel.Range workSheet_range = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col], (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col]);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            if (_color_interior!= null)
                {
                workSheet_range.Interior.Color = _color_interior;
                }

            if (_format != null && _format!= "")
                {
                if (_format == "Date")
                    {
                    workSheet_range.EntireColumn.NumberFormat = "MM/DD/YYYY";
                    }
                if (_format == "Currency")
                    {
                    workSheet_range.Cells.NumberFormat = "##,###.00";
                    }
                if (_format == "Currency_dollar")
                    {

                    workSheet_range.Cells.NumberFormat = "$##,###.00";
                    }
                if (_format == "Currency_dollarP")
                    {

                    workSheet_range.Cells.NumberFormat = "($##,###.00)";
                    }
                if (_format == "Percentage")
                    {

                    workSheet_range.NumberFormat = "";
                    }
                }
            else
                {
                workSheet_range.NumberFormat = "";
                }
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = font_size;
            
            workSheet_range.EntireColumn.AutoFit();
            if (adate)
                {
                workSheet_range.EntireColumn.NumberFormat = "MM/DD/YYYY";
                }
         
            }
        public string getData(int row, int col)
            {
            string _value = "";

            try
                {

                _value = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col]).Value2.ToString();
                }
            catch
                {
                _value = "NULL";
                }

            return _value;
            }


        public void closeDoc()
            {
            app.Quit();

            System.Diagnostics.Process[] _process = System.Diagnostics.Process.GetProcesses();

            for (int i = 0; i < _process.Length; i++)
                {
                if (_process[i].ProcessName == "EXCEL")
                    {
                    _process[i].CloseMainWindow();
                    _process[i].Close();
                    _process[i].Dispose();
                    }
                }

            }

        public string getCellLetter(int _col)
            {
            string _letter = "";

            if (_col == 1)
                {
                _letter = "A";
                }

            if (_col == 2)
                {
                _letter = "B";
                }

            if (_col == 3)
                {
                _letter = "C";
                }

            if (_col == 4)
                {
                _letter = "D";
                }

            if (_col == 5)
                {
                _letter = "E";
                }

            if (_col == 6)
                {
                _letter = "F";
                }

            if (_col == 7)
                {
                _letter = "G";
                }

            if (_col == 8)
                {
                _letter = "H";
                }

            if (_col == 9)
                {
                _letter = "I";
                }

            if (_col == 10)
                {
                _letter = "J";
                }

            if (_col == 11)
                {
                _letter = "K";
                }

            if (_col == 12)
                {
                _letter = "L";
                }

            if (_col == 13)
                {
                _letter = "M";
                }

            if (_col == 14)
                {
                _letter = "N";
                }

            if (_col == 15)
                {
                _letter = "O";
                }

            if (_col == 16)
                {
                _letter = "P";
                }

            if (_col == 17)
                {
                _letter = "Q";
                }

            if (_col == 18)
                {
                _letter = "R";
                }

            if (_col == 19)
                {
                _letter = "S";
                }

            if (_col == 20)
                {
                _letter = "T";
                }

            if (_col == 21)
                {
                _letter = "U";
                }

            if (_col == 22)
                {
                _letter = "V";
                }

            if (_col == 23)
                {
                _letter = "W";
                }

            if (_col == 24)
                {
                _letter = "X";
                }

            if (_col == 25)
                {
                _letter = "Y";
                }

            if (_col == 26)
                {
                _letter = "Z";
                }

            if (_col == 27)
                {
                _letter = "AA";
                }

            if (_col == 28)
                {
                _letter = "AB";
                }

            if (_col == 29)
                {
                _letter = "AC";
                }

            if (_col == 29)
                {
                _letter = "AD";
                }

            if (_col == 30)
                {
                _letter = "AE";
                }

            return _letter;
            }

        public void generateExcelFromDataGridView(DataGridView _dataGridView, string _name)
            {
            createDoc(_name);

            for (int i = 0; i < _dataGridView.Columns.Count; i++)
                {
                addSubHeader(1, i + 1, _dataGridView.Columns[i].HeaderText.ToUpper(), getCellLetter(i + 1) + "1", getCellLetter(i + 1) + "1", "");
                }

            for (int r = 0; r <= _dataGridView.Rows.Count - 1; r++)
                {
                for (int c = 0; c < _dataGridView.Columns.Count; c++)
                    {
                    string value = _dataGridView[c, r].Value == null ? string.Empty : _dataGridView[c, r].Value.ToString();
                    addData(r + 2, c + 1, value, getCellLetter(c + 1) + (r + 2).ToString(), getCellLetter(c + 1) + (r + 2).ToString(), "", "");
                    }
                }

            showDocument();

            }


        }
    }
