﻿namespace Quick_Ship_Router
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
            this.showToday = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.deleteTravelers = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.toTraveler = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.fromTraveler = new System.Windows.Forms.TextBox();
            this.btnLoadEATS = new System.Windows.Forms.Button();
            this.clearBefore = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCreatedPrinted = new System.Windows.Forms.Button();
            this.login = new System.Windows.Forms.Button();
            this.btn = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.btnCreateSpecificOrder = new System.Windows.Forms.Button();
            this.specificOrder = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.combineOrders = new System.Windows.Forms.CheckBox();
            this.customerList = new System.Windows.Forms.CheckedListBox();
            this.btnInvertCustomers = new System.Windows.Forms.Button();
            this.btnCreateTravelers = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tableListView = new System.Windows.Forms.ListView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chairListView = new System.Windows.Forms.ListView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.miscListView = new System.Windows.Forms.ListView();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.infoLabel = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chckChairs = new System.Windows.Forms.RadioButton();
            this.chckTables = new System.Windows.Forms.RadioButton();
            this.btnPrintLabels = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnCreateAll = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // showToday
            // 
            this.showToday.AutoSize = true;
            this.showToday.Location = new System.Drawing.Point(6, 17);
            this.showToday.Name = "showToday";
            this.showToday.Size = new System.Drawing.Size(187, 24);
            this.showToday.TabIndex = 5;
            this.showToday.Text = "Only orders from today";
            this.showToday.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.BackColor = System.Drawing.Color.LightGray;
            this.groupBox1.Controls.Add(this.deleteTravelers);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.toTraveler);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.fromTraveler);
            this.groupBox1.Controls.Add(this.btnLoadEATS);
            this.groupBox1.Controls.Add(this.clearBefore);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnCreatedPrinted);
            this.groupBox1.Controls.Add(this.login);
            this.groupBox1.Controls.Add(this.btn);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.btnCreateSpecificOrder);
            this.groupBox1.Controls.Add(this.specificOrder);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.combineOrders);
            this.groupBox1.Controls.Add(this.customerList);
            this.groupBox1.Controls.Add(this.showToday);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(1272, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 737);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Refined Search";
            // 
            // deleteTravelers
            // 
            this.deleteTravelers.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.deleteTravelers.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.deleteTravelers.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.deleteTravelers.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.deleteTravelers.Location = new System.Drawing.Point(104, 551);
            this.deleteTravelers.Name = "deleteTravelers";
            this.deleteTravelers.Size = new System.Drawing.Size(95, 34);
            this.deleteTravelers.TabIndex = 29;
            this.deleteTravelers.Text = "Delete";
            this.deleteTravelers.UseVisualStyleBackColor = false;
            this.deleteTravelers.Click += new System.EventHandler(this.deleteTravelers_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(101, 591);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 13);
            this.label3.TabIndex = 28;
            this.label3.Text = "To";
            // 
            // toTraveler
            // 
            this.toTraveler.Location = new System.Drawing.Point(122, 591);
            this.toTraveler.Name = "toTraveler";
            this.toTraveler.Size = new System.Drawing.Size(68, 26);
            this.toTraveler.TabIndex = 27;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(0, 591);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 26;
            this.label2.Text = "From";
            // 
            // fromTraveler
            // 
            this.fromTraveler.Location = new System.Drawing.Point(33, 591);
            this.fromTraveler.Name = "fromTraveler";
            this.fromTraveler.Size = new System.Drawing.Size(68, 26);
            this.fromTraveler.TabIndex = 25;
            this.fromTraveler.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // btnLoadEATS
            // 
            this.btnLoadEATS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoadEATS.BackColor = System.Drawing.Color.MediumPurple;
            this.btnLoadEATS.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnLoadEATS.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadEATS.Location = new System.Drawing.Point(6, 623);
            this.btnLoadEATS.Name = "btnLoadEATS";
            this.btnLoadEATS.Size = new System.Drawing.Size(188, 34);
            this.btnLoadEATS.TabIndex = 24;
            this.btnLoadEATS.Text = "Load EATS travelers";
            this.btnLoadEATS.UseVisualStyleBackColor = false;
            this.btnLoadEATS.Click += new System.EventHandler(this.btnLoadEATS_Click);
            // 
            // clearBefore
            // 
            this.clearBefore.AutoSize = true;
            this.clearBefore.Checked = true;
            this.clearBefore.CheckState = System.Windows.Forms.CheckState.Checked;
            this.clearBefore.Location = new System.Drawing.Point(6, 64);
            this.clearBefore.Name = "clearBefore";
            this.clearBefore.Size = new System.Drawing.Size(174, 24);
            this.clearBefore.TabIndex = 23;
            this.clearBefore.Text = "Clear before new set";
            this.clearBefore.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.MediumTurquoise;
            this.button4.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.button4.ForeColor = System.Drawing.Color.Black;
            this.button4.Location = new System.Drawing.Point(6, 331);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(188, 31);
            this.button4.TabIndex = 12;
            this.button4.Text = "Everything else";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.btnInvertCustomers_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.Teal;
            this.button3.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.Black;
            this.button3.Location = new System.Drawing.Point(6, 294);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(188, 31);
            this.button3.TabIndex = 6;
            this.button3.Text = "Amazon / WF";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.btnCreateTravelers_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(30, 708);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 26);
            this.label1.TabIndex = 22;
            this.label1.Text = "Powered by C#\r\nDeveloped by Gage Coates";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCreatedPrinted
            // 
            this.btnCreatedPrinted.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreatedPrinted.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnCreatedPrinted.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnCreatedPrinted.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreatedPrinted.Location = new System.Drawing.Point(6, 551);
            this.btnCreatedPrinted.Name = "btnCreatedPrinted";
            this.btnCreatedPrinted.Size = new System.Drawing.Size(95, 34);
            this.btnCreatedPrinted.TabIndex = 16;
            this.btnCreatedPrinted.Text = "Load Printed Travelers";
            this.btnCreatedPrinted.UseVisualStyleBackColor = false;
            this.btnCreatedPrinted.Click += new System.EventHandler(this.btnCreatedPrinted_Click);
            // 
            // login
            // 
            this.login.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.login.BackColor = System.Drawing.Color.DimGray;
            this.login.ForeColor = System.Drawing.Color.White;
            this.login.Location = new System.Drawing.Point(6, 671);
            this.login.Name = "login";
            this.login.Size = new System.Drawing.Size(188, 34);
            this.login.TabIndex = 7;
            this.login.Text = "Login to MAS";
            this.login.UseVisualStyleBackColor = false;
            this.login.Click += new System.EventHandler(this.login_Click);
            // 
            // btn
            // 
            this.btn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btn.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn.ForeColor = System.Drawing.Color.Black;
            this.btn.Location = new System.Drawing.Point(6, 511);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(188, 34);
            this.btn.TabIndex = 21;
            this.btn.Text = "Clear";
            this.btn.UseVisualStyleBackColor = false;
            this.btn.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.button5.Location = new System.Drawing.Point(6, 471);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(188, 34);
            this.button5.TabIndex = 16;
            this.button5.Text = "Been Printed?";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCreateSpecificOrder
            // 
            this.btnCreateSpecificOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnCreateSpecificOrder.Location = new System.Drawing.Point(6, 434);
            this.btnCreateSpecificOrder.Name = "btnCreateSpecificOrder";
            this.btnCreateSpecificOrder.Size = new System.Drawing.Size(188, 31);
            this.btnCreateSpecificOrder.TabIndex = 15;
            this.btnCreateSpecificOrder.Text = "Add Order / Traveler";
            this.btnCreateSpecificOrder.UseVisualStyleBackColor = false;
            this.btnCreateSpecificOrder.Click += new System.EventHandler(this.btnCreateSpecificOrder_Click);
            // 
            // specificOrder
            // 
            this.specificOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.specificOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.specificOrder.Location = new System.Drawing.Point(6, 397);
            this.specificOrder.Name = "specificOrder";
            this.specificOrder.Size = new System.Drawing.Size(188, 26);
            this.specificOrder.TabIndex = 14;
            this.specificOrder.Text = "Order # / Traveler #";
            this.specificOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 374);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(110, 20);
            this.label4.TabIndex = 13;
            this.label4.Text = "Specific order:";
            // 
            // combineOrders
            // 
            this.combineOrders.AutoSize = true;
            this.combineOrders.Checked = true;
            this.combineOrders.CheckState = System.Windows.Forms.CheckState.Checked;
            this.combineOrders.Location = new System.Drawing.Point(6, 41);
            this.combineOrders.Name = "combineOrders";
            this.combineOrders.Size = new System.Drawing.Size(140, 24);
            this.combineOrders.TabIndex = 11;
            this.combineOrders.Text = "Combine orders";
            this.combineOrders.UseVisualStyleBackColor = true;
            // 
            // customerList
            // 
            this.customerList.BackColor = System.Drawing.Color.Wheat;
            this.customerList.FormattingEnabled = true;
            this.customerList.Location = new System.Drawing.Point(6, 93);
            this.customerList.Name = "customerList";
            this.customerList.Size = new System.Drawing.Size(188, 193);
            this.customerList.TabIndex = 7;
            // 
            // btnInvertCustomers
            // 
            this.btnInvertCustomers.BackColor = System.Drawing.Color.Blue;
            this.btnInvertCustomers.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInvertCustomers.ForeColor = System.Drawing.Color.White;
            this.btnInvertCustomers.Location = new System.Drawing.Point(168, 15);
            this.btnInvertCustomers.Name = "btnInvertCustomers";
            this.btnInvertCustomers.Size = new System.Drawing.Size(156, 50);
            this.btnInvertCustomers.TabIndex = 12;
            this.btnInvertCustomers.Text = "Everything else";
            this.btnInvertCustomers.UseVisualStyleBackColor = false;
            this.btnInvertCustomers.Click += new System.EventHandler(this.btnInvertCustomers_Click);
            // 
            // btnCreateTravelers
            // 
            this.btnCreateTravelers.BackColor = System.Drawing.Color.Green;
            this.btnCreateTravelers.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateTravelers.ForeColor = System.Drawing.Color.White;
            this.btnCreateTravelers.Location = new System.Drawing.Point(6, 15);
            this.btnCreateTravelers.Name = "btnCreateTravelers";
            this.btnCreateTravelers.Size = new System.Drawing.Size(156, 50);
            this.btnCreateTravelers.TabIndex = 6;
            this.btnCreateTravelers.UseVisualStyleBackColor = false;
            this.btnCreateTravelers.Click += new System.EventHandler(this.btnCreateTravelers_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(12, 106);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1254, 643);
            this.tabControl1.TabIndex = 17;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.checkBox1);
            this.tabPage1.Controls.Add(this.tableListView);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1246, 610);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Tables";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tableListView
            // 
            this.tableListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableListView.BackColor = System.Drawing.Color.Beige;
            this.tableListView.CheckBoxes = true;
            this.tableListView.FullRowSelect = true;
            this.tableListView.GridLines = true;
            this.tableListView.Location = new System.Drawing.Point(6, 36);
            this.tableListView.Name = "tableListView";
            this.tableListView.Size = new System.Drawing.Size(1234, 568);
            this.tableListView.TabIndex = 0;
            this.tableListView.UseCompatibleStateImageBehavior = false;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.chairListView);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1246, 610);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Chairs";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // chairListView
            // 
            this.chairListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chairListView.BackColor = System.Drawing.Color.Beige;
            this.chairListView.CheckBoxes = true;
            this.chairListView.ForeColor = System.Drawing.Color.Black;
            this.chairListView.FullRowSelect = true;
            this.chairListView.GridLines = true;
            this.chairListView.Location = new System.Drawing.Point(6, 6);
            this.chairListView.Name = "chairListView";
            this.chairListView.Size = new System.Drawing.Size(1234, 598);
            this.chairListView.TabIndex = 0;
            this.chairListView.UseCompatibleStateImageBehavior = false;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.miscListView);
            this.tabPage3.Location = new System.Drawing.Point(4, 29);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1246, 610);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Misc";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // miscListView
            // 
            this.miscListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.miscListView.BackColor = System.Drawing.Color.Beige;
            this.miscListView.CheckBoxes = true;
            this.miscListView.ForeColor = System.Drawing.Color.Black;
            this.miscListView.FullRowSelect = true;
            this.miscListView.GridLines = true;
            this.miscListView.Location = new System.Drawing.Point(6, 6);
            this.miscListView.Name = "miscListView";
            this.miscListView.Size = new System.Drawing.Size(1237, 598);
            this.miscListView.TabIndex = 1;
            this.miscListView.UseCompatibleStateImageBehavior = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnCreateTravelers);
            this.groupBox3.Controls.Add(this.btnInvertCustomers);
            this.groupBox3.Location = new System.Drawing.Point(12, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(331, 72);
            this.groupBox3.TabIndex = 19;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Generate";
            // 
            // infoLabel
            // 
            this.infoLabel.AutoSize = true;
            this.infoLabel.BackColor = System.Drawing.Color.Transparent;
            this.infoLabel.Font = new System.Drawing.Font("Monotype Corsiva", 27.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infoLabel.ForeColor = System.Drawing.Color.Cyan;
            this.infoLabel.Location = new System.Drawing.Point(758, 39);
            this.infoLabel.Name = "infoLabel";
            this.infoLabel.Size = new System.Drawing.Size(153, 45);
            this.infoLabel.TabIndex = 3;
            this.infoLabel.Text = "Loading...";
            // 
            // groupBox4
            // 
            this.groupBox4.BackColor = System.Drawing.Color.LightGray;
            this.groupBox4.Controls.Add(this.chckChairs);
            this.groupBox4.Controls.Add(this.chckTables);
            this.groupBox4.Controls.Add(this.btnPrintLabels);
            this.groupBox4.Controls.Add(this.button1);
            this.groupBox4.Controls.Add(this.button2);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.ForeColor = System.Drawing.Color.Black;
            this.groupBox4.Location = new System.Drawing.Point(349, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(403, 72);
            this.groupBox4.TabIndex = 18;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Print";
            // 
            // chckChairs
            // 
            this.chckChairs.AutoSize = true;
            this.chckChairs.Location = new System.Drawing.Point(330, 41);
            this.chckChairs.Name = "chckChairs";
            this.chckChairs.Size = new System.Drawing.Size(64, 20);
            this.chckChairs.TabIndex = 27;
            this.chckChairs.Text = "Chairs";
            this.chckChairs.UseVisualStyleBackColor = true;
            // 
            // chckTables
            // 
            this.chckTables.AutoSize = true;
            this.chckTables.Checked = true;
            this.chckTables.Location = new System.Drawing.Point(330, 19);
            this.chckTables.Name = "chckTables";
            this.chckTables.Size = new System.Drawing.Size(69, 20);
            this.chckTables.TabIndex = 26;
            this.chckTables.TabStop = true;
            this.chckTables.Text = "Tables";
            this.chckTables.UseVisualStyleBackColor = true;
            this.chckTables.CheckedChanged += new System.EventHandler(this.chckTables_CheckedChanged);
            // 
            // btnPrintLabels
            // 
            this.btnPrintLabels.BackColor = System.Drawing.Color.MediumSeaGreen;
            this.btnPrintLabels.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnPrintLabels.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrintLabels.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrintLabels.ForeColor = System.Drawing.Color.Black;
            this.btnPrintLabels.Location = new System.Drawing.Point(222, 16);
            this.btnPrintLabels.Name = "btnPrintLabels";
            this.btnPrintLabels.Size = new System.Drawing.Size(102, 49);
            this.btnPrintLabels.TabIndex = 22;
            this.btnPrintLabels.Text = "Labels";
            this.btnPrintLabels.UseVisualStyleBackColor = false;
            this.btnPrintLabels.Click += new System.EventHandler(this.btnPrintLabels_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Aquamarine;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Location = new System.Drawing.Point(6, 15);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 49);
            this.button1.TabIndex = 0;
            this.button1.Text = "Travelers";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.MediumSpringGreen;
            this.button2.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.Black;
            this.button2.Location = new System.Drawing.Point(114, 15);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(102, 49);
            this.button2.TabIndex = 8;
            this.button2.Text = "Summary";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.btnPrintSummary_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.BackColor = System.Drawing.Color.LightGray;
            this.groupBox5.Controls.Add(this.btnCreateAll);
            this.groupBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox5.ForeColor = System.Drawing.Color.Black;
            this.groupBox5.Location = new System.Drawing.Point(12, 12);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(331, 72);
            this.groupBox5.TabIndex = 19;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Generate";
            // 
            // btnCreateAll
            // 
            this.btnCreateAll.BackColor = System.Drawing.Color.MediumTurquoise;
            this.btnCreateAll.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnCreateAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateAll.Font = new System.Drawing.Font("Impact", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateAll.ForeColor = System.Drawing.Color.Black;
            this.btnCreateAll.Location = new System.Drawing.Point(10, 17);
            this.btnCreateAll.Name = "btnCreateAll";
            this.btnCreateAll.Size = new System.Drawing.Size(314, 49);
            this.btnCreateAll.TabIndex = 23;
            this.btnCreateAll.Text = "Everything";
            this.btnCreateAll.UseVisualStyleBackColor = false;
            this.btnCreateAll.Click += new System.EventHandler(this.btnCreateAll_Click);
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.BackColor = System.Drawing.Color.DimGray;
            this.progressBar.ForeColor = System.Drawing.Color.Teal;
            this.progressBar.Location = new System.Drawing.Point(12, 90);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(1254, 10);
            this.progressBar.Step = 1;
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 20;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "travelers.json";
            this.openFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog_FileOk);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(6, 6);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(48, 24);
            this.checkBox1.TabIndex = 1;
            this.checkBox1.Text = "All";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DimGray;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1484, 761);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.infoLabel);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Quick Ship Traveler    v2.0";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckBox showToday;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckedListBox customerList;
        private System.Windows.Forms.Button btnCreateTravelers;
        private System.Windows.Forms.Button login;
        private System.Windows.Forms.CheckBox combineOrders;
        private System.Windows.Forms.Button btnInvertCustomers;
        private System.Windows.Forms.Button btnCreateSpecificOrder;
        private System.Windows.Forms.TextBox specificOrder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnCreatedPrinted;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ListView tableListView;
        private System.Windows.Forms.ListView chairListView;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label infoLabel;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btn;
        private System.Windows.Forms.Button btnPrintLabels;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCreateAll;
        private System.Windows.Forms.CheckBox clearBefore;
        private System.Windows.Forms.RadioButton chckChairs;
        private System.Windows.Forms.RadioButton chckTables;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.ListView miscListView;
        private System.Windows.Forms.Button btnLoadEATS;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox toTraveler;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox fromTraveler;
        private System.Windows.Forms.Button deleteTravelers;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}

