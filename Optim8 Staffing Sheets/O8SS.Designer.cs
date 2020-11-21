namespace Optim8_Staffing_Sheets
{
    partial class O8SS
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
            this.rdSSbtn = new System.Windows.Forms.Button();
            this.psSSbtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rdSSbtn
            // 
            this.rdSSbtn.Location = new System.Drawing.Point(49, 90);
            this.rdSSbtn.Name = "rdSSbtn";
            this.rdSSbtn.Size = new System.Drawing.Size(311, 232);
            this.rdSSbtn.TabIndex = 0;
            this.rdSSbtn.Text = "Rides Staffing Sheet";
            this.rdSSbtn.UseVisualStyleBackColor = true;
            this.rdSSbtn.Click += new System.EventHandler(this.rdSSbtn_Click);
            // 
            // psSSbtn
            // 
            this.psSSbtn.Location = new System.Drawing.Point(409, 90);
            this.psSSbtn.Name = "psSSbtn";
            this.psSSbtn.Size = new System.Drawing.Size(311, 232);
            this.psSSbtn.TabIndex = 1;
            this.psSSbtn.Text = "Park Services Staffing Sheet";
            this.psSSbtn.UseVisualStyleBackColor = true;
            this.psSSbtn.Click += new System.EventHandler(this.psSSbtn_Click);
            // 
            // O8SS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.psSSbtn);
            this.Controls.Add(this.rdSSbtn);
            this.Name = "O8SS";
            this.Text = "O8SS";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button rdSSbtn;
        private System.Windows.Forms.Button psSSbtn;
    }
}