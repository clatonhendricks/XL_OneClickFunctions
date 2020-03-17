using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLOCF
{
    public partial class frmLookup : Form
    {
        public frmLookup()
        {
            InitializeComponent();
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);
            this.txtTableArray.Text = address;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Vlookup();

        }

        private void frmLookup_Load(object sender, EventArgs e)
        {
            //btnOk.Focus();
        }

        private void Vlookup()
        {
            //string celltxt = "";
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string rel = string.Empty;
            double Num;
            bool isNum = double.TryParse(txtL_Value.Text, out Num);
            try
            {
                if (isNum)
                {
                    if (rdoTrue.Checked == true)
                    { rel = "=VLookup(" + txtL_Value.Text + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "TRUE)"; }
                    else { rel = "=VLookup(" + txtL_Value.Text + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "FALSE)"; }
                }
                else
                {
                    string S_Quotes, E_Quotes;
                    S_Quotes = @"""";
                    E_Quotes = @"""";

                    if (rdoTrue.Checked == true)
                    { rel = "=VLookup(" +  S_Quotes + txtL_Value.Text + E_Quotes + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "TRUE)"; }
                    else { rel = "=VLookup(" + S_Quotes + txtL_Value.Text + E_Quotes + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "FALSE)"; }


                }

                //celltxt = "";



                MessageBox.Show(rel);
                //if (rdoTrue.Checked == true)
                //{ rel = "=VLookup(" + txtL_Value.Text + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "TRUE)"; }
                //else { rel = "=VLookup(" + txtL_Value.Text + "," + txtTableArray.Text + "," + txtCol_Index.Text + "," + "FALSE)"; }
                
                
                if (selection.Cells.Rows.Count > selection.Cells.Columns.Count)
                {
                    //row
                    int R_Count = selection.Cells.Rows.Count + 1;
                    selection.set_Item(R_Count, 1, rel);
                }
                else
                {
                    int R_Count = selection.Cells.Columns.Count + 1;
                    selection.set_Item(1, R_Count, rel);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //MessageBox.Show("NULL VALUE");
                return;

            }

            finally
            {
                this.Close();
            }
            }

    }

}
