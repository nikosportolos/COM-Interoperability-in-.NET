namespace InterprocessCommunication
{
    partial class frmInterComm
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
            this.btSend2VB = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btSend2VB
            // 
            this.btSend2VB.Location = new System.Drawing.Point(70, 108);
            this.btSend2VB.Name = "btSend2VB";
            this.btSend2VB.Size = new System.Drawing.Size(144, 47);
            this.btSend2VB.TabIndex = 0;
            this.btSend2VB.Text = "Send message to VB6";
            this.btSend2VB.UseVisualStyleBackColor = true;
            this.btSend2VB.Click += new System.EventHandler(this.btSend2VB_Click);
            // 
            // frmInterComm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btSend2VB);
            this.Name = "frmInterComm";
            this.Text = "Interprocess Communication Tutorial";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btSend2VB;
    }
}

