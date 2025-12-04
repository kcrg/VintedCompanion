namespace VintedCompanion
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnProcess = new Button();
            statusLabel = new Label();
            radioButton1 = new RadioButton();
            radioButton2 = new RadioButton();
            SuspendLayout();
            // 
            // btnProcess
            // 
            btnProcess.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnProcess.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnProcess.Location = new Point(33, 33);
            btnProcess.Margin = new Padding(26);
            btnProcess.Name = "btnProcess";
            btnProcess.Size = new Size(884, 189);
            btnProcess.TabIndex = 0;
            btnProcess.Text = "Wybierz plik *.html z Vinted";
            btnProcess.UseVisualStyleBackColor = true;
            btnProcess.Click += btnProcess_Click;
            // 
            // statusLabel
            // 
            statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            statusLabel.AutoSize = true;
            statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 0);
            statusLabel.Location = new Point(11, 282);
            statusLabel.Margin = new Padding(2, 0, 2, 0);
            statusLabel.MaximumSize = new Size(690, 0);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(321, 23);
            statusLabel.TabIndex = 2;
            statusLabel.Text = "Status: Wybrano generowanie pliku Excel";
            // 
            // radioButton1
            // 
            radioButton1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            radioButton1.AutoSize = true;
            radioButton1.Checked = true;
            radioButton1.Location = new Point(722, 282);
            radioButton1.Margin = new Padding(2);
            radioButton1.Name = "radioButton1";
            radioButton1.Size = new Size(111, 24);
            radioButton1.TabIndex = 3;
            radioButton1.TabStop = true;
            radioButton1.Text = "Excel (*.xslx)";
            radioButton1.UseVisualStyleBackColor = true;
            radioButton1.CheckedChanged += radioButton1_CheckedChanged;
            // 
            // radioButton2
            // 
            radioButton2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            radioButton2.AutoSize = true;
            radioButton2.Location = new Point(837, 282);
            radioButton2.Margin = new Padding(2);
            radioButton2.Name = "radioButton2";
            radioButton2.Size = new Size(102, 24);
            radioButton2.TabIndex = 4;
            radioButton2.Text = "PDF (*.pdf)";
            radioButton2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(950, 317);
            Controls.Add(radioButton2);
            Controls.Add(radioButton1);
            Controls.Add(statusLabel);
            Controls.Add(btnProcess);
            Margin = new Padding(2);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Vinted Generator";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnProcess;
        private Label statusLabel;
        private RadioButton radioButton1;
        private RadioButton radioButton2;
    }
}
