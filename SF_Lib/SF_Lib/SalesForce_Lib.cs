using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SF_Lib
{
    [Guid("192757DE-1CCC-4164-AFF6-1FD1A5EDA119")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface ISalesForce_Lib
    {
        [DispId(1)]
        void showtext();
    }

    [Guid("86290531-9089-46C3-8BD5-CE6FDA68CD6E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class SalesForce_Lib : ISalesForce_Lib
    {
        int showOnMonitor;

        bool _displaywellcome = true;

        public int ShowOnMonitor
        {
            get { return showOnMonitor; }
            set { showOnMonitor = value; }
        }

        public SalesForce_Lib()
        {

        }
        public bool Displaywellcome
        {
            get { return _displaywellcome; }
            set { _displaywellcome = value; }
        }
        public void showtext()
        {
            try
            {
                bool isMyFormOpen = false;
                foreach (Form f in Application.OpenForms)
                {
                    if (f is Principal)
                    {
                        isMyFormOpen = true;
                        DisplayTextInfo(f);
                    }
                }
                if (!isMyFormOpen)
                {
                    Principal f = new Principal();

                    DisplayTextInfo(f);

                    //logic for Display in second screeen.
                    Screen[] sc;
                    sc = Screen.AllScreens;
                    if (ShowOnMonitor >= sc.Length)
                    {
                        ShowOnMonitor = 0;
                    }

                    f.StartPosition = FormStartPosition.Manual;
                    f.Location = sc[ShowOnMonitor].Bounds.Location;
                    f.WindowState = FormWindowState.Normal;
                    f.WindowState = FormWindowState.Maximized;
                    f.ShowInTaskbar = false;
                    // f.ClientSize = new System.Drawing.Size(0, 0);
                    f.Size = Screen.PrimaryScreen.WorkingArea.Size;
                    f.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + Convert.ToString(ex.Message));
            }
        }

        private void DisplayTextInfo(Form f)
        {
            //Logic for Display Panel Middle of the Form.
            TabControl Displaypanel = (TabControl)f.Controls.Find("tabControl_Matched", true).FirstOrDefault();
            if (Displaypanel != null)
            {
                Displaypanel.Location = new System.Drawing.Point(f.ClientSize.Width / 2 - Displaypanel.Size.Width / 2, f.ClientSize.Height / 2 - Displaypanel.Size.Height / 2);
                Displaypanel.Anchor = AnchorStyles.None;

            }

          /*  TextBox txtwelcome = (TextBox)f.Controls.Find("txtwelcome", true).FirstOrDefault();

            TextBox txtline1Text = (TextBox)f.Controls.Find("txtline1Text", true).FirstOrDefault();

            TextBox txtline2Text = (TextBox)f.Controls.Find("txtline2Text", true).FirstOrDefault();

            TextBox txtline1Amount = (TextBox)f.Controls.Find("txtline1Amount", true).FirstOrDefault();

            TextBox txtline2Amount = (TextBox)f.Controls.Find("txtline2Amount", true).FirstOrDefault();

            // if (string.IsNullOrEmpty(Line1_Text)  && string.IsNullOrEmpty(Line2_Text) && string.IsNullOrEmpty(Line1_Amount) && string.IsNullOrEmpty(Line2_Amount))
            if (Displaywellcome)
            {
                txtwelcome.Visible = true;
                txtline1Text.Visible = false;
                txtline2Text.Visible = false;
                txtline1Amount.Visible = false;
                txtline2Amount.Visible = false;
            }
            else
            {


                txtwelcome.Visible = false;
                txtline1Text.Visible = true;
                txtline2Text.Visible = true;
                txtline1Amount.Visible = true;
                txtline2Amount.Visible = true;
            }*/
        }


    }
}
