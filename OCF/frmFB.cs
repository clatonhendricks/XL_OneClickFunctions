using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook =   Microsoft.Office.Interop.Outlook;


namespace XLOCF
{
    public partial class frmFB : Form
    {
        public frmFB()
        {
            InitializeComponent();
        }

        
        private bool check()
        {
            if ((txtBody.Text == "") || (txtSub.Text == ""))
            {
                MessageBox.Show("Please fill all text boxes");
                return false;
            } return true;
                
        }

        private void frmFB_Load(object sender, EventArgs e)

        {
            System.Reflection.Assembly thisAssembly = this.GetType().Assembly;

            lblVer.Text = "Ver: " +  thisAssembly.GetName().Version.ToString();
            //////////////////
            if (this.Text == "Tell A Friend")
            {
                label1.Text = "TO";
                label1.Location = new Point(90, 9);
                
                txtTO.Visible = true;
                lblVer.Visible = true;
                cboSub.Visible = false;
                
                txtSub.Text = "Try out One Click Functions";
                string Body = "Hi there," + System.Environment.NewLine +
                System.Environment.NewLine +
                "I've been using this Excel Addin 'One Click Functions' which has some built in functions where I don't have to write them myself, this addin writes it for me. " + System.Environment.NewLine +
                "I think you should try this. " + System.Environment.NewLine +
                "Install it from https://github.com/clatonhendricks/XL_OneClickFunctions" + System.Environment.NewLine +
                System.Environment.NewLine + System.Environment.NewLine +
                "Regards,";

                txtBody.Text = txtBody.Text + Body; 
            }
            else
            {
                label1.Text = "Bug or Feedback";
                label1.Location = new Point(5, 9);
                txtTO.Visible = false;
                lblVer.Visible = false;
                cboSub.Visible = true;
                txtSub.Text = string.Empty;
                txtBody.Text = string.Empty;
            }

        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if (this.Text == "Feedback")
            {
                Feedback();
            }
            else Frnd();
        }

        private void Frnd()
        {
            
            try
            {
                if (check())
                {

                    Microsoft.Office.Interop.Outlook.Application opp = new Microsoft.Office.Interop.Outlook.Application();
                    Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)opp.CreateItem(0);
                    
                    mail.To = txtTO.Text;

                    mail.Subject = txtSub.Text;
                    mail.Body = txtBody.Text;

                    mail.Send();
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void Feedback()
        {
            try
            {
                if (check())
                {
                    Microsoft.Office.Interop.Outlook.Application opp = new Microsoft.Office.Interop.Outlook.Application();
                    Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)opp.CreateItem(0);

                    mail.To = "clatonhendricks@msn.com";


                    mail.Subject = this.cboSub.Text + ": " + txtSub.Text;
                    mail.Body = this.txtBody.Text + System.Environment.NewLine + System.Environment.NewLine + "Sent using " + lblVer.Text;

                    mail.Send();
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
    }
}
