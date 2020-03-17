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
    public partial class frmCriteria : Form
    {
        public frmCriteria()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            func_CountIf();

        }

        private void func_CountIf()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            try
            {
                string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);
                //check if the data entered is string or number
                double Num;
                string rel = string.Empty;
                string txdata = txtCri.Text.Trim();
                bool isNum = double.TryParse(txdata, out Num);

                if (isNum)
                {
                    rel = "=Countif(" + address.ToString() + "," + txtCri.Text.ToString() + ")";
                }
                else
                {
                    string S_Quotes, E_Quotes;
                    S_Quotes = @"""";
                    E_Quotes = @""")";


                    rel = "=Countif(" + address.ToString() + "," + S_Quotes + txtCri.Text.ToString() + E_Quotes;
                }

                // Formula to be set based on the greater count on rows or columns
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
                MessageBox.Show(ex.Message, "String Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                this.Close();
            }
        }

        private void txtCri_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtCri_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void txtCri_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                func_CountIf();
            }
        }

       

        

        
    }
}
