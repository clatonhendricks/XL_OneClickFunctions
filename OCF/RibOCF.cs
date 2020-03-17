using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using Microsoft.Office.Core;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RibOCF();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace XLOCF
{
    [ComVisible(true)]
    public class RibOCF : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibOCF()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("XLOCF.RibOCF.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void SimpleFunctions(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case Constants.RibbonButtonID.CAPS:
                    Caps(control);
                    break;
                case Constants.RibbonButtonID.LOWER:
                    Lower(control);
                    break;
                case Constants.RibbonButtonID.PROPER:
                    Proper(control);
                    break;
                case Constants.RibbonButtonID.TRIM:
                    Trim(control);
                    break;
                case Constants.RibbonButtonID.SUBTRACT:
                    Subtract(control);
                    break;
                case Constants.RibbonButtonID.MULTIPLY:
                    Multiply(control);
                    break;
                case Constants.RibbonButtonID.PRODUCT:
                    Product(control);
                    break;
                case Constants.RibbonButtonID.DIVIDE:
                    Divide(control);
                    break;
                case Constants.RibbonButtonID.COUNTIF:
                    countif();
                    break;
                case Constants.RibbonButtonID.COUNTA:
                    CountA(control);
                    break;
                case Constants.RibbonButtonID.COUNTB:
                    CountB(control);
                    break;
                case Constants.RibbonButtonID.NOW:
                    Now(control);
                    break;
                case Constants.RibbonButtonID.DateDiff_D:
                    Subtract(control);
                    break;
                case Constants.RibbonButtonID.DateDiff_M:
                    DateDiff_M(control);
                    break;
                case Constants.RibbonButtonID.DateDiff_Y:
                    DateDiff_Y(control);
                    break;
                case Constants.RibbonButtonID.VLOOKUP:
                    VLookup();
                    break;
                case Constants.RibbonButtonID.HLOOKUP:
                    HLookup();
                    break;
                case Constants.RibbonButtonID.FB:
                    Feedback();
                    break;
                case Constants.RibbonButtonID.Frnd:
                    Frnd();
                    break;
                case Constants.RibbonButtonID.UPDATE:
                    Update();
                    break;
            }
        }

        private void Frnd()
        {
            DialogResult diag = MessageBox.Show("Do you have Outlook Configured on this PC?", "One Click Functions", MessageBoxButtons.YesNo);
            if (diag == DialogResult.Yes)
            {
                Form FB = new frmFB();
                FB.Text = "Tell A Friend";
                FB.Show();
            }
            else { MessageBox.Show("Please refer your friends to go to https://github.com/clatonhendricks/XL_OneClickFunctions"); }
        }

        //private void Note()
        //{
        //    Form Note = new frmNote();
        //    Note.Show();
        //}

        private void Update()
        {
            //if (AutomaticUpdater.Updater.IsUpdateDownloaded)
            //{
            //    AutomaticUpdater.Apply();
            //}
        }

        private void Feedback()
        {
            DialogResult diag = MessageBox.Show("Do you have Outlook Configured on this PC?", "One Click Functions", MessageBoxButtons.YesNo);
            if (diag == DialogResult.Yes)
            {
                Form FB = new frmFB();
                FB.Text = "Feedback";
                FB.Show();
            }
            else { MessageBox.Show("Please send an email to clatonh@microsoft.com"); }
        }

        private void VLookup()
        {
            //frmLookup frm = new frmLookup();
            if (CheckMoreRn())
            {
                Form frm = new frmLookup();
                frm.Show();
                frm.Text = "VLOOKUP";
            }
            else MessageBox.Show("Please select a range and try again", "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void HLookup()
        {
            if (CheckMoreRn())
            {
                Form frm = new frmLookup();
                frm.Show();
                frm.Text = "HLOOKUP";
            }
            else MessageBox.Show("Please select a range and try again", "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void countif()
        {
            frmCriteria frm = new frmCriteria();
            frm.Show();

        }

        //code for the Image
        public System.Drawing.Bitmap OnLoadImage(string imageName)
        {
            switch (imageName)
            {
                case "Minus":
                    return Properties.Resources.Minus;
                case "DivideX":
                    return Properties.Resources.Divide;
                case "Multiply":
                    return Properties.Resources.Multiply;
                case "Product":
                    return Properties.Resources.Product;
                case "Cif":
                    return Properties.Resources.Cif;
                case "C_A":
                    return Properties.Resources.C_A;
                case "C_B":
                    return Properties.Resources.C_B;
                case "Now":
                    return Properties.Resources.Now;
                case "Days":
                    return Properties.Resources.Days;
                case "Mth":
                    return Properties.Resources.Mth;
                case "Year":
                    return Properties.Resources.Year;

                default: return null;
            }
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "Subtract":
                    return Properties.Resources.Minus;
                case "Divide":
                    return Properties.Resources.Divide;
                case "Multiply":
                    return Properties.Resources.Multiply;
                case "Product":
                    return Properties.Resources.Product;
                case "Countif":
                    return Properties.Resources.Cif;
                case "CountA":
                    return Properties.Resources.C_A;
                case "CountB":
                    return Properties.Resources.C_B;
                case "Now":
                    return Properties.Resources.Now;
                case "DateD_Days":
                    return Properties.Resources.Days;
                case "DateD_Month":
                    return Properties.Resources.Mth;
                case "DateD_Year":
                    return Properties.Resources.Year;

            }
            return null;
        }


        private static void Caps(Office.IRibbonControl control)
        {
            string celltxt = "";
           
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            foreach (object cell in selection.Cells)
            {
                celltxt = "";
                try
                {
                    celltxt = ((Excel.Range)cell).Value2.ToString();
                    ((Excel.Range)cell).Value2 = celltxt.ToUpper();

                }

                catch
                {

                    //MessageBox.Show("NULL VALUE");
                    return;

                }

                // }

            }
        }

        private static void Subtract(Office.IRibbonControl control)
        {
            string rel = string.Empty;
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            try
            {
                if (Check2Rn())
                {

                    string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);
                    //MessageBox.Show(selection.Cells.Columns.Column.ToString());
                    //MessageBox.Show(selection.Cells.Columns.Row.ToString());

                    //MessageBox.Show(selection.Next.ToString());


                    if (selection.Cells.Rows.Count == selection.Cells.Columns.Count)
                    { rel = "=" + address.Replace(",", "-"); }
                    else { rel = "=" + address.Replace(":", "-"); }



                    ////** Working code
                    ////Formula to be set based on the greater count on rows or columns
                    if (selection.Cells.Rows.Count > selection.Cells.Columns.Count)
                    {
                        //row

                        selection.set_Item(3, 1, rel);

                    }
                    else selection.set_Item(1, 3, rel);


                }
                else
                {
                    MessageBox.Show("You have either selected more than 2 cells or less. Please select only two cells as Range.", "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private static void Multiply(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string rel = string.Empty;
            if (selection.Cells.Count >= 3)
            {
                MessageBox.Show("Please select only 2 cells to Multiply", "String Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                try
                {
                    string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

                    if (selection.Cells.Rows.Count == selection.Cells.Columns.Count)
                    { rel = "=" + address.Replace(",", "*"); }
                    else { rel = "=" + address.Replace(":", "*"); }


                    // Formula to be set based on the greater count on rows or columns
                    if (selection.Cells.Rows.Count > selection.Cells.Columns.Count)
                    {
                        //row
                        selection.set_Item(3, 1, rel);
                    }
                    else selection.set_Item(1, 3, rel);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private static void Divide(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string rel = string.Empty;

            if (selection.Cells.Count >= 3)
            {
                MessageBox.Show("Please select only 2 cells to Divide", "String Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                try
                {
                    string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

                    if (selection.Cells.Rows.Count == selection.Cells.Columns.Count)
                    { rel = "=" + address.Replace(",", "/"); }
                    else { rel = "=" + address.Replace(":", "/"); }


                    // Formula to be set based on the greater count on rows or columns
                    if (selection.Cells.Rows.Count > selection.Cells.Columns.Count)
                    {
                        //row
                        selection.set_Item(3, 1, rel);
                    }
                    else selection.set_Item(1, 3, rel);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private static void CountA(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            //if (selection.Cells.Count >= 3)
            //{
            //    MessageBox.Show("Please select only 2 cells to Subtract", "String Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}
            //else
            //{
            try
            {
                string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

                //string rel = "=" + address.Replace(":", "/");
                string rel = "=CountA(" + address.ToString() + ")";
                //MessageBox.Show(address);
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
                MessageBox.Show(ex.Message);
            }
            // }
        }

        private static void CountB(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            try
            {
                string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

                //string rel = "=" + address.Replace(":", "/");
                string rel = "=CountBlank(" + address.ToString() + ")";
                //MessageBox.Show(address);
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
                MessageBox.Show(ex.Message);
            }
            // }
        }

        private static void Product(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            try
            {
                string address = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

                //string rel = "=" + address.Replace(":", "/");
                string rel = "=Product(" + address.ToString() + ")";
                //MessageBox.Show(address);
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
                MessageBox.Show(ex.Message);
            }
            // }
        }

        private static void Now(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            foreach (object cell in selection.Cells)
            {

                try
                {

                    ((Excel.Range)cell).Formula = "=Now()";

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;

                }

            }
        }

        private static void DateDiff_M(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;


            try
            {

                if (Check2Rn())
                {
                    string R1 = "=Month(" + GetLeftRn() + ")";
                    string R2 = "Month(" + GetRightRn() + ")";



                    string rel = R1 + "-" + R2;

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

                else
                { MessageBox.Show("You have either selected more than 2 cells or less. Please select only Two cells as Range.", "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

                return;

            }

        }

        private static void DateDiff_Y(Office.IRibbonControl control)
        {

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;


            try
            {
                if (Check2Rn())
                {

                    string R1 = "=Year(" + GetLeftRn() + ")";
                    string R2 = "Year(" + GetRightRn() + ")";



                    string rel = R1 + "-" + R2;

                    // Formula to be set based on the greater count on rows or columns
                    if (selection.Cells.Rows.Count > selection.Cells.Columns.Count)
                    {
                        //row
                        int R_Count = selection.Cells.Rows.Count + 1;
                        Debug.WriteLine(R_Count);
                        selection.set_Item(R_Count, 1, rel);
                    }
                    else
                    {
                        int R_Count = selection.Cells.Columns.Count + 1;
                        selection.set_Item(1, R_Count, rel);
                    }
                }
                else
                {
                    MessageBox.Show("You have either selected more than 2 cells or less. Please select only Two cells as Range.", "One Click Functions", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

                return;

            }

        }

        private static bool Check2Rn()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            if (selection.Cells.Count == 2)
            {
                return true;
            }
            else return false;

        }

        private static bool CheckMoreRn()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            if (selection.Cells.Count >= 2)
            {
                return true;
            }
            else return false;

        }

        public static string GetLeftRn()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string address = selection.Cells.Columns.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

            int div = address.Length / 2;

            string Result = Left(address, div);

            return Result;

        }

        private static string GetRightRn()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            string address = selection.Cells.Columns.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, false);

            int div = address.Length / 2;

            String Result = Right(address, div);
            return Result;

        }


        public static string Left(string param, int length)
        {
            //we start at 0 since we want to get the characters starting from the
            //left and with the specified lenght and assign it to a variable
            string result = param.Substring(0, length);

            //return the result of the operation
            return result;
        }

        public static string Right(string param, int length)
        {
            //start at the index based on the lenght of the sting minus
            //the specified lenght and assign it a variable
            string result = param.Substring(param.Length - length, length);
            //return the result of the operation
            return result;
        }

        private static void Lower(Office.IRibbonControl control)
        {
            string celltxt = "";
            //MessageBox.Show("Claton you are the dude........");
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            foreach (object cell in selection.Cells)
            {
                celltxt = "";
                try
                {

                    celltxt = ((Excel.Range)cell).Value2.ToString();
                    ((Excel.Range)cell).Value2 = celltxt.ToLower();

                }

                catch
                {

                    //MessageBox.Show("NULL VALUE");
                    return;

                }
            }
        }

        private static void Proper(Office.IRibbonControl control)
        {
            string celltxt = "";
            //MessageBox.Show("Claton you are the dude........");
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            foreach (object cell in selection.Cells)
            {
                celltxt = "";
                try
                {

                    celltxt = ((Excel.Range)cell).Value2.ToString();

                    ((Excel.Range)cell).Value2 = ProperCase(celltxt);




                }

                catch
                {

                    //MessageBox.Show("NULL VALUE");
                    return;

                }

            }
        }

        private static void Trim(Office.IRibbonControl control)
        {
            string celltxt = "";
            //MessageBox.Show("Claton you are the dude........");
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            foreach (object cell in selection.Cells)
            {
                celltxt = "";
                try
                {

                    celltxt = ((Excel.Range)cell).Value2.ToString();
                    ((Excel.Range)cell).Value2 = celltxt.Trim();

                }

                catch
                {

                    //MessageBox.Show("NULL VALUE");
                    return;
                }

            }
        }

        public static string ProperCase(string stringInput)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            bool fEmptyBefore = true;
            foreach (char ch in stringInput)
            {
                char chThis = ch;
                if (Char.IsWhiteSpace(chThis))
                    fEmptyBefore = true;
                else
                {
                    if (Char.IsLetter(chThis) && fEmptyBefore)
                        chThis = Char.ToUpper(chThis);
                    else
                        chThis = Char.ToLower(chThis);
                    fEmptyBefore = false;
                }
                sb.Append(chThis);
            }
            return sb.ToString();
        }

        #endregion


        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
