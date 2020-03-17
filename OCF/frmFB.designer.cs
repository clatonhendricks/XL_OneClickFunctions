namespace XLOCF
{
    partial class frmFB
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSend = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cboSub = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSub = new System.Windows.Forms.TextBox();
            this.lblVer = new System.Windows.Forms.Label();
            this.txtTO = new System.Windows.Forms.TextBox();
            this.lblWar = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(143, 525);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(109, 31);
            this.btnSend.TabIndex = 0;
            this.btnSend.Text = "Send Email";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(90, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "TO:";
            // 
            // cboSub
            // 
            this.cboSub.FormattingEnabled = true;
            this.cboSub.Items.AddRange(new object[] {
            "Bug",
            "Feedback"});
            this.cboSub.Location = new System.Drawing.Point(142, 6);
            this.cboSub.Name = "cboSub";
            this.cboSub.Size = new System.Drawing.Size(137, 24);
            this.cboSub.TabIndex = 1;
            this.cboSub.Text = "Bug";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(54, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 17);
            this.label2.TabIndex = 3;
            this.label2.Tag = ":";
            this.label2.Text = "Comments";
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(142, 97);
            this.txtBody.Multiline = true;
            this.txtBody.Name = "txtBody";
            this.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBody.Size = new System.Drawing.Size(517, 422);
            this.txtBody.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(69, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 17);
            this.label3.TabIndex = 5;
            this.label3.Text = "Subject:";
            // 
            // txtSub
            // 
            this.txtSub.Location = new System.Drawing.Point(142, 46);
            this.txtSub.Name = "txtSub";
            this.txtSub.Size = new System.Drawing.Size(516, 22);
            this.txtSub.TabIndex = 2;
            // 
            // lblVer
            // 
            this.lblVer.AutoSize = true;
            this.lblVer.Location = new System.Drawing.Point(592, 551);
            this.lblVer.Name = "lblVer";
            this.lblVer.Size = new System.Drawing.Size(0, 17);
            this.lblVer.TabIndex = 6;
            // 
            // txtTO
            // 
            this.txtTO.Location = new System.Drawing.Point(142, 9);
            this.txtTO.Name = "txtTO";
            this.txtTO.Size = new System.Drawing.Size(516, 22);
            this.txtTO.TabIndex = 7;
            this.txtTO.Visible = false;
            // 
            // lblWar
            // 
            this.lblWar.AutoSize = true;
            this.lblWar.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWar.ForeColor = System.Drawing.Color.Red;
            this.lblWar.Location = new System.Drawing.Point(66, 574);
            this.lblWar.Name = "lblWar";
            this.lblWar.Size = new System.Drawing.Size(606, 17);
            this.lblWar.TabIndex = 8;
            this.lblWar.Text = "Please only refer internal employees as this Project is still not released extern" +
                "ally.";
            // 
            // frmFB
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(739, 600);
            this.Controls.Add(this.lblWar);
            this.Controls.Add(this.lblVer);
            this.Controls.Add(this.txtSub);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cboSub);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.txtTO);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmFB";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Bug";
            this.Load += new System.EventHandler(this.frmFB_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboSub;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSub;
        private System.Windows.Forms.Label lblVer;
        private System.Windows.Forms.TextBox txtTO;
        private System.Windows.Forms.Label lblWar;
    }
}