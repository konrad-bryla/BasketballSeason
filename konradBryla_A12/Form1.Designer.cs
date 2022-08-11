namespace konradBryla_A12
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
            this.components = new System.ComponentModel.Container();
            this._2018_SeasonDataSet1 = new konradBryla_A12._2018_SeasonDataSet();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.season_2018TableAdapter = new konradBryla_A12._2018_SeasonDataSetTableAdapters.Season_2018TableAdapter();
            this.doExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this._2018_SeasonDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // _2018_SeasonDataSet1
            // 
            this._2018_SeasonDataSet1.DataSetName = "_2018_SeasonDataSet";
            this._2018_SeasonDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataMember = "Season 2018";
            this.bindingSource1.DataSource = this._2018_SeasonDataSet1;
            // 
            // season_2018TableAdapter
            // 
            this.season_2018TableAdapter.ClearBeforeFill = true;
            // 
            // doExcel
            // 
            this.doExcel.Location = new System.Drawing.Point(76, 118);
            this.doExcel.Name = "doExcel";
            this.doExcel.Size = new System.Drawing.Size(168, 89);
            this.doExcel.TabIndex = 0;
            this.doExcel.Text = "Open Excel";
            this.doExcel.UseVisualStyleBackColor = true;
            this.doExcel.Click += new System.EventHandler(this.doExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Start Date";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "End Date";
            // 
            // dtpStartDate
            // 
            this.dtpStartDate.Location = new System.Drawing.Point(104, 31);
            this.dtpStartDate.MaxDate = new System.DateTime(2018, 12, 31, 0, 0, 0, 0);
            this.dtpStartDate.MinDate = new System.DateTime(2018, 1, 1, 0, 0, 0, 0);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(200, 20);
            this.dtpStartDate.TabIndex = 3;
            this.dtpStartDate.Value = new System.DateTime(2018, 1, 1, 0, 0, 0, 0);
            // 
            // dtpEndDate
            // 
            this.dtpEndDate.Location = new System.Drawing.Point(104, 69);
            this.dtpEndDate.MaxDate = new System.DateTime(2018, 12, 31, 0, 0, 0, 0);
            this.dtpEndDate.MinDate = new System.DateTime(2018, 1, 1, 0, 0, 0, 0);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(200, 20);
            this.dtpEndDate.TabIndex = 4;
            this.dtpEndDate.Value = new System.DateTime(2018, 12, 31, 0, 0, 0, 0);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(315, 231);
            this.Controls.Add(this.dtpEndDate);
            this.Controls.Add(this.dtpStartDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.doExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this._2018_SeasonDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private _2018_SeasonDataSet _2018_SeasonDataSet1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private _2018_SeasonDataSetTableAdapters.Season_2018TableAdapter season_2018TableAdapter;
        private System.Windows.Forms.Button doExcel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpStartDate;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
    }
}

