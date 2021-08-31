
namespace TimeTableGeneratorDbce
{
    partial class main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(main));
            this.mainHeading = new System.Windows.Forms.Label();
            this.deptHeading = new System.Windows.Forms.Label();
            this.course = new System.Windows.Forms.Button();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.viewTables = new System.Windows.Forms.Button();
            this.generateTables = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.teacher = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainHeading
            // 
            this.mainHeading.AutoSize = true;
            this.mainHeading.BackColor = System.Drawing.Color.Transparent;
            this.mainHeading.Font = new System.Drawing.Font("Impact", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mainHeading.ForeColor = System.Drawing.Color.MidnightBlue;
            this.mainHeading.Location = new System.Drawing.Point(167, 33);
            this.mainHeading.Name = "mainHeading";
            this.mainHeading.Size = new System.Drawing.Size(434, 45);
            this.mainHeading.TabIndex = 0;
            this.mainHeading.Text = "DBCE TIME TABLE GENERATOR";
            // 
            // deptHeading
            // 
            this.deptHeading.AutoSize = true;
            this.deptHeading.BackColor = System.Drawing.Color.RoyalBlue;
            this.deptHeading.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.deptHeading.Font = new System.Drawing.Font("Corbel", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.deptHeading.ForeColor = System.Drawing.Color.White;
            this.deptHeading.Location = new System.Drawing.Point(203, 95);
            this.deptHeading.Name = "deptHeading";
            this.deptHeading.Size = new System.Drawing.Size(340, 28);
            this.deptHeading.TabIndex = 1;
            this.deptHeading.Text = "Department Of Computer Engineering";
            this.deptHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // course
            // 
            this.course.BackColor = System.Drawing.Color.DarkBlue;
            this.course.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.course.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.course.ForeColor = System.Drawing.Color.White;
            this.course.Location = new System.Drawing.Point(103, 197);
            this.course.Name = "course";
            this.course.Size = new System.Drawing.Size(289, 61);
            this.course.TabIndex = 2;
            this.course.Text = "COURSES";
            this.course.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.course.UseVisualStyleBackColor = false;
            this.course.Click += new System.EventHandler(this.course_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.White;
            this.pictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(38, 100);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(46, 42);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 5;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(116, 208);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(46, 41);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // viewTables
            // 
            this.viewTables.BackColor = System.Drawing.Color.MidnightBlue;
            this.viewTables.Font = new System.Drawing.Font("Arial Rounded MT Bold", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.viewTables.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.viewTables.Location = new System.Drawing.Point(341, 24);
            this.viewTables.Name = "viewTables";
            this.viewTables.Size = new System.Drawing.Size(202, 128);
            this.viewTables.TabIndex = 6;
            this.viewTables.Text = "VIEW TABLES";
            this.viewTables.UseVisualStyleBackColor = false;
            this.viewTables.Click += new System.EventHandler(this.viewTables_Click);
            // 
            // generateTables
            // 
            this.generateTables.BackColor = System.Drawing.Color.LightSeaGreen;
            this.generateTables.Font = new System.Drawing.Font("Consolas", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateTables.ForeColor = System.Drawing.Color.White;
            this.generateTables.Location = new System.Drawing.Point(25, 186);
            this.generateTables.Name = "generateTables";
            this.generateTables.Size = new System.Drawing.Size(539, 41);
            this.generateTables.TabIndex = 7;
            this.generateTables.Text = "GENERATE TIME TABLE";
            this.generateTables.UseVisualStyleBackColor = false;
            this.generateTables.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Gainsboro;
            this.groupBox1.Controls.Add(this.generateTables);
            this.groupBox1.Controls.Add(this.viewTables);
            this.groupBox1.Controls.Add(this.pictureBox2);
            this.groupBox1.Controls.Add(this.teacher);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.groupBox1.Location = new System.Drawing.Point(78, 173);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(585, 255);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "OPERATIONS";
            // 
            // teacher
            // 
            this.teacher.BackColor = System.Drawing.Color.RoyalBlue;
            this.teacher.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.teacher.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.teacher.ForeColor = System.Drawing.Color.White;
            this.teacher.Location = new System.Drawing.Point(25, 91);
            this.teacher.Name = "teacher";
            this.teacher.Size = new System.Drawing.Size(289, 61);
            this.teacher.TabIndex = 4;
            this.teacher.Text = "TEACHERS";
            this.teacher.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.teacher.UseVisualStyleBackColor = false;
            this.teacher.Click += new System.EventHandler(this.teacher_Click);
            // 
            // main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(762, 462);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.deptHeading);
            this.Controls.Add(this.mainHeading);
            this.Controls.Add(this.course);
            this.Controls.Add(this.groupBox1);
            this.ForeColor = System.Drawing.SystemColors.Control;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "main";
            this.Text = "DBCE TIME TABLE MANAGEMENT";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label mainHeading;
        private System.Windows.Forms.Label deptHeading;
        private System.Windows.Forms.Button course;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button viewTables;
        private System.Windows.Forms.Button generateTables;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button teacher;
    }
}

