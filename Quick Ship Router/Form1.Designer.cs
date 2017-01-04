namespace Quick_Ship_Router
{
    partial class Form1
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
            this.btnPrint = new System.Windows.Forms.Button();
            this.listView = new System.Windows.Forms.ListView();
            this.loadingLabel = new System.Windows.Forms.Label();
            this.showToday = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnCreateSpecificOrder = new System.Windows.Forms.Button();
            this.specificOrder = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnInvertCustomers = new System.Windows.Forms.Button();
            this.combineOrders = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.productLineList = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.customerList = new System.Windows.Forms.CheckedListBox();
            this.btnCreateTravelers = new System.Windows.Forms.Button();
            this.login = new System.Windows.Forms.Button();
            this.btnPrintSummary = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnCreatedPrinted = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnPrint.Location = new System.Drawing.Point(12, 12);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "Print Routers";
            this.btnPrint.UseVisualStyleBackColor = false;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // listView
            // 
            this.listView.BackColor = System.Drawing.Color.WhiteSmoke;
            this.listView.CheckBoxes = true;
            this.listView.Location = new System.Drawing.Point(12, 41);
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(1560, 374);
            this.listView.TabIndex = 2;
            this.listView.UseCompatibleStateImageBehavior = false;
            // 
            // loadingLabel
            // 
            this.loadingLabel.AutoSize = true;
            this.loadingLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadingLabel.Location = new System.Drawing.Point(219, 2);
            this.loadingLabel.Name = "loadingLabel";
            this.loadingLabel.Size = new System.Drawing.Size(0, 37);
            this.loadingLabel.TabIndex = 3;
            // 
            // showToday
            // 
            this.showToday.AutoSize = true;
            this.showToday.Location = new System.Drawing.Point(6, 20);
            this.showToday.Name = "showToday";
            this.showToday.Size = new System.Drawing.Size(131, 17);
            this.showToday.TabIndex = 5;
            this.showToday.Text = "Only orders from today";
            this.showToday.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.LightGray;
            this.groupBox1.Controls.Add(this.btnCreatedPrinted);
            this.groupBox1.Controls.Add(this.btnCreateSpecificOrder);
            this.groupBox1.Controls.Add(this.specificOrder);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnInvertCustomers);
            this.groupBox1.Controls.Add(this.combineOrders);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.productLineList);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.customerList);
            this.groupBox1.Controls.Add(this.btnCreateTravelers);
            this.groupBox1.Controls.Add(this.showToday);
            this.groupBox1.Location = new System.Drawing.Point(12, 458);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(603, 391);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Refined Search";
            // 
            // btnCreateSpecificOrder
            // 
            this.btnCreateSpecificOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnCreateSpecificOrder.Location = new System.Drawing.Point(336, 73);
            this.btnCreateSpecificOrder.Name = "btnCreateSpecificOrder";
            this.btnCreateSpecificOrder.Size = new System.Drawing.Size(132, 31);
            this.btnCreateSpecificOrder.TabIndex = 15;
            this.btnCreateSpecificOrder.Text = "Add traveler";
            this.btnCreateSpecificOrder.UseVisualStyleBackColor = false;
            this.btnCreateSpecificOrder.Click += new System.EventHandler(this.btnCreateSpecificOrder_Click);
            // 
            // specificOrder
            // 
            this.specificOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.specificOrder.Location = new System.Drawing.Point(336, 36);
            this.specificOrder.Name = "specificOrder";
            this.specificOrder.Size = new System.Drawing.Size(132, 31);
            this.specificOrder.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(333, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Specific order:";
            // 
            // btnInvertCustomers
            // 
            this.btnInvertCustomers.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnInvertCustomers.Location = new System.Drawing.Point(6, 311);
            this.btnInvertCustomers.Name = "btnInvertCustomers";
            this.btnInvertCustomers.Size = new System.Drawing.Size(156, 37);
            this.btnInvertCustomers.TabIndex = 12;
            this.btnInvertCustomers.Text = "Create travelers\r\nfrom unselected";
            this.btnInvertCustomers.UseVisualStyleBackColor = false;
            this.btnInvertCustomers.Click += new System.EventHandler(this.btnInvertCustomers_Click);
            // 
            // combineOrders
            // 
            this.combineOrders.AutoSize = true;
            this.combineOrders.Checked = true;
            this.combineOrders.CheckState = System.Windows.Forms.CheckState.Checked;
            this.combineOrders.Location = new System.Drawing.Point(168, 19);
            this.combineOrders.Name = "combineOrders";
            this.combineOrders.Size = new System.Drawing.Size(99, 17);
            this.combineOrders.TabIndex = 11;
            this.combineOrders.Text = "Combine orders";
            this.combineOrders.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(165, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(163, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Product Lines (currently disabled)";
            // 
            // productLineList
            // 
            this.productLineList.BackColor = System.Drawing.Color.Wheat;
            this.productLineList.CheckOnClick = true;
            this.productLineList.Enabled = false;
            this.productLineList.FormattingEnabled = true;
            this.productLineList.Location = new System.Drawing.Point(168, 78);
            this.productLineList.Name = "productLineList";
            this.productLineList.Size = new System.Drawing.Size(156, 184);
            this.productLineList.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Customers";
            // 
            // customerList
            // 
            this.customerList.BackColor = System.Drawing.Color.Wheat;
            this.customerList.FormattingEnabled = true;
            this.customerList.Location = new System.Drawing.Point(6, 78);
            this.customerList.Name = "customerList";
            this.customerList.Size = new System.Drawing.Size(156, 184);
            this.customerList.TabIndex = 7;
            // 
            // btnCreateTravelers
            // 
            this.btnCreateTravelers.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnCreateTravelers.Location = new System.Drawing.Point(6, 268);
            this.btnCreateTravelers.Name = "btnCreateTravelers";
            this.btnCreateTravelers.Size = new System.Drawing.Size(156, 37);
            this.btnCreateTravelers.TabIndex = 6;
            this.btnCreateTravelers.Text = "Create travelers\r\nfrom selected";
            this.btnCreateTravelers.UseVisualStyleBackColor = false;
            this.btnCreateTravelers.Click += new System.EventHandler(this.btnCreateTravelers_Click);
            // 
            // login
            // 
            this.login.Location = new System.Drawing.Point(1468, 15);
            this.login.Name = "login";
            this.login.Size = new System.Drawing.Size(104, 23);
            this.login.TabIndex = 7;
            this.login.Text = "Login to MAS";
            this.login.UseVisualStyleBackColor = true;
            this.login.Click += new System.EventHandler(this.login_Click);
            // 
            // btnPrintSummary
            // 
            this.btnPrintSummary.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnPrintSummary.Location = new System.Drawing.Point(93, 12);
            this.btnPrintSummary.Name = "btnPrintSummary";
            this.btnPrintSummary.Size = new System.Drawing.Size(92, 23);
            this.btnPrintSummary.TabIndex = 8;
            this.btnPrintSummary.Text = "Print Summary";
            this.btnPrintSummary.UseVisualStyleBackColor = false;
            this.btnPrintSummary.Click += new System.EventHandler(this.btnPrintSummary_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(805, 421);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(204, 39);
            this.label3.TabIndex = 9;
            this.label3.Text = "Currently only supports tables\r\n\r\nStay tuned to find out what happens next!";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnClear
            // 
            this.btnClear.BackColor = System.Drawing.Color.Maroon;
            this.btnClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.btnClear.Location = new System.Drawing.Point(12, 421);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(84, 31);
            this.btnClear.TabIndex = 16;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = false;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnCreatedPrinted
            // 
            this.btnCreatedPrinted.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnCreatedPrinted.Location = new System.Drawing.Point(168, 268);
            this.btnCreatedPrinted.Name = "btnCreatedPrinted";
            this.btnCreatedPrinted.Size = new System.Drawing.Size(156, 37);
            this.btnCreatedPrinted.TabIndex = 16;
            this.btnCreatedPrinted.Text = "Only printed travelers";
            this.btnCreatedPrinted.UseVisualStyleBackColor = false;
            this.btnCreatedPrinted.Click += new System.EventHandler(this.btnCreatedPrinted_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(1584, 861);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnPrintSummary);
            this.Controls.Add(this.login);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.loadingLabel);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.btnPrint);
            this.Name = "Form1";
            this.Text = "Quick Ship Traveler    v1.1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Shown += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.Label loadingLabel;
        private System.Windows.Forms.CheckBox showToday;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckedListBox customerList;
        private System.Windows.Forms.Button btnCreateTravelers;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckedListBox productLineList;
        private System.Windows.Forms.Button login;
        private System.Windows.Forms.Button btnPrintSummary;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox combineOrders;
        private System.Windows.Forms.Button btnInvertCustomers;
        private System.Windows.Forms.Button btnCreateSpecificOrder;
        private System.Windows.Forms.TextBox specificOrder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnCreatedPrinted;
    }
}

