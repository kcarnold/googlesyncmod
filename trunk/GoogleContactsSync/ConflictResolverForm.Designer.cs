namespace WebGear.GoogleContactsSync
{
    partial class ConflictResolverForm
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
            this.messageLabel = new System.Windows.Forms.Label();
            this.keepOutlook = new System.Windows.Forms.Button();
            this.keepGoogle = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // messageLabel
            // 
            this.messageLabel.Location = new System.Drawing.Point(14, 9);
            this.messageLabel.Name = "messageLabel";
            this.messageLabel.Size = new System.Drawing.Size(386, 62);
            this.messageLabel.TabIndex = 0;
            this.messageLabel.Text = "message";
            // 
            // keepOutlook
            // 
            this.keepOutlook.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.keepOutlook.Location = new System.Drawing.Point(14, 74);
            this.keepOutlook.Name = "keepOutlook";
            this.keepOutlook.Size = new System.Drawing.Size(140, 23);
            this.keepOutlook.TabIndex = 1;
            this.keepOutlook.Text = "Keep Outlook Entry";
            this.keepOutlook.UseVisualStyleBackColor = true;
            // 
            // keepGoogle
            // 
            this.keepGoogle.DialogResult = System.Windows.Forms.DialogResult.No;
            this.keepGoogle.Location = new System.Drawing.Point(161, 74);
            this.keepGoogle.Name = "keepGoogle";
            this.keepGoogle.Size = new System.Drawing.Size(146, 23);
            this.keepGoogle.TabIndex = 2;
            this.keepGoogle.Text = "Keep Google Entry";
            this.keepGoogle.UseVisualStyleBackColor = true;
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(314, 74);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(87, 23);
            this.cancel.TabIndex = 3;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            // 
            // ConflictResolverForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(414, 109);
            this.ControlBox = false;
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.keepGoogle);
            this.Controls.Add(this.keepOutlook);
            this.Controls.Add(this.messageLabel);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConflictResolverForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Confict detected!";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button keepOutlook;
        private System.Windows.Forms.Button keepGoogle;
        private System.Windows.Forms.Button cancel;
        public System.Windows.Forms.Label messageLabel;
    }
}