namespace Lab_1.Table_Editor
{
    partial class TableEditor
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.TabledataGridView = new System.Windows.Forms.DataGridView();
            this.DeleteColumnbutton = new System.Windows.Forms.Button();
            this.AddColumnbutton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.DeleteRowbutton = new System.Windows.Forms.Button();
            this.AddRowbutton = new System.Windows.Forms.Button();
            this.ParsertextBox = new System.Windows.Forms.TextBox();
            this.CurentElLabel = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.зберегтиToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.зберегтиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.відкритиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.таблицяToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.додатиСтовпчикToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.видалитиСтовпчикToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.додатиРядокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.видалитиРядокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.допомогаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.проПроектToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.умоваToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.TabledataGridView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabledataGridView
            // 
            this.TabledataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TabledataGridView.Location = new System.Drawing.Point(12, 138);
            this.TabledataGridView.Name = "TabledataGridView";
            this.TabledataGridView.Size = new System.Drawing.Size(2164, 1055);
            this.TabledataGridView.TabIndex = 0;
            this.TabledataGridView.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.TabledataGridView_CellEnter);
            this.TabledataGridView.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.TabledataGridView_CellLeave);
            // 
            // DeleteColumnbutton
            // 
            this.DeleteColumnbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DeleteColumnbutton.Location = new System.Drawing.Point(600, 81);
            this.DeleteColumnbutton.Name = "DeleteColumnbutton";
            this.DeleteColumnbutton.Size = new System.Drawing.Size(49, 51);
            this.DeleteColumnbutton.TabIndex = 1;
            this.DeleteColumnbutton.Text = "-";
            this.DeleteColumnbutton.UseVisualStyleBackColor = true;
            this.DeleteColumnbutton.Click += new System.EventHandler(this.DeleteColumn);
            // 
            // AddColumnbutton
            // 
            this.AddColumnbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.AddColumnbutton.Location = new System.Drawing.Point(655, 81);
            this.AddColumnbutton.Name = "AddColumnbutton";
            this.AddColumnbutton.Size = new System.Drawing.Size(49, 51);
            this.AddColumnbutton.TabIndex = 2;
            this.AddColumnbutton.Text = "+";
            this.AddColumnbutton.UseVisualStyleBackColor = true;
            this.AddColumnbutton.Click += new System.EventHandler(this.AddColumn);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(595, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "Стовпчики";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(758, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 25);
            this.label2.TabIndex = 4;
            this.label2.Text = "Рядки";
            // 
            // DeleteRowbutton
            // 
            this.DeleteRowbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DeleteRowbutton.Location = new System.Drawing.Point(763, 81);
            this.DeleteRowbutton.Name = "DeleteRowbutton";
            this.DeleteRowbutton.Size = new System.Drawing.Size(49, 51);
            this.DeleteRowbutton.TabIndex = 5;
            this.DeleteRowbutton.Text = "-";
            this.DeleteRowbutton.UseVisualStyleBackColor = true;
            this.DeleteRowbutton.Click += new System.EventHandler(this.DeleteRow);
            // 
            // AddRowbutton
            // 
            this.AddRowbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.AddRowbutton.Location = new System.Drawing.Point(818, 81);
            this.AddRowbutton.Name = "AddRowbutton";
            this.AddRowbutton.Size = new System.Drawing.Size(49, 51);
            this.AddRowbutton.TabIndex = 6;
            this.AddRowbutton.Text = "+";
            this.AddRowbutton.UseVisualStyleBackColor = true;
            this.AddRowbutton.Click += new System.EventHandler(this.AddRow);
            // 
            // ParsertextBox
            // 
            this.ParsertextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ParsertextBox.Location = new System.Drawing.Point(233, 98);
            this.ParsertextBox.Name = "ParsertextBox";
            this.ParsertextBox.Size = new System.Drawing.Size(361, 26);
            this.ParsertextBox.TabIndex = 7;
            this.ParsertextBox.TextChanged += new System.EventHandler(this.ParsertextBox_TextChanged);
            this.ParsertextBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.ParsertextBox_KeyUp);
            // 
            // CurentElLabel
            // 
            this.CurentElLabel.AutoSize = true;
            this.CurentElLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CurentElLabel.Location = new System.Drawing.Point(176, 101);
            this.CurentElLabel.Name = "CurentElLabel";
            this.CurentElLabel.Size = new System.Drawing.Size(39, 20);
            this.CurentElLabel.TabIndex = 8;
            this.CurentElLabel.Text = "------";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.таблицяToolStripMenuItem,
            this.допомогаToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(2188, 24);
            this.menuStrip1.TabIndex = 9;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.зберегтиToolStripMenuItem1,
            this.зберегтиToolStripMenuItem,
            this.відкритиToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // зберегтиToolStripMenuItem1
            // 
            this.зберегтиToolStripMenuItem1.Name = "зберегтиToolStripMenuItem1";
            this.зберегтиToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
            this.зберегтиToolStripMenuItem1.Text = "Зберегти";
            this.зберегтиToolStripMenuItem1.Click += new System.EventHandler(this.зберегтиToolStripMenuItem1_Click);
            // 
            // зберегтиToolStripMenuItem
            // 
            this.зберегтиToolStripMenuItem.Name = "зберегтиToolStripMenuItem";
            this.зберегтиToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.зберегтиToolStripMenuItem.Text = "Зберегти як";
            this.зберегтиToolStripMenuItem.Click += new System.EventHandler(this.зберегтиToolStripMenuItem_Click);
            // 
            // відкритиToolStripMenuItem
            // 
            this.відкритиToolStripMenuItem.Name = "відкритиToolStripMenuItem";
            this.відкритиToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.відкритиToolStripMenuItem.Text = "Відкрити";
            this.відкритиToolStripMenuItem.Click += new System.EventHandler(this.відкритиToolStripMenuItem_Click);
            // 
            // таблицяToolStripMenuItem
            // 
            this.таблицяToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.додатиСтовпчикToolStripMenuItem,
            this.видалитиСтовпчикToolStripMenuItem,
            this.додатиРядокToolStripMenuItem,
            this.видалитиРядокToolStripMenuItem});
            this.таблицяToolStripMenuItem.Name = "таблицяToolStripMenuItem";
            this.таблицяToolStripMenuItem.Size = new System.Drawing.Size(66, 20);
            this.таблицяToolStripMenuItem.Text = "Таблиця";
            // 
            // додатиСтовпчикToolStripMenuItem
            // 
            this.додатиСтовпчикToolStripMenuItem.Name = "додатиСтовпчикToolStripMenuItem";
            this.додатиСтовпчикToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.додатиСтовпчикToolStripMenuItem.Text = "Додати стовпчик";
            this.додатиСтовпчикToolStripMenuItem.Click += new System.EventHandler(this.AddColumn);
            // 
            // видалитиСтовпчикToolStripMenuItem
            // 
            this.видалитиСтовпчикToolStripMenuItem.Name = "видалитиСтовпчикToolStripMenuItem";
            this.видалитиСтовпчикToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.видалитиСтовпчикToolStripMenuItem.Text = "Видалити стовпчик";
            this.видалитиСтовпчикToolStripMenuItem.Click += new System.EventHandler(this.DeleteColumn);
            // 
            // додатиРядокToolStripMenuItem
            // 
            this.додатиРядокToolStripMenuItem.Name = "додатиРядокToolStripMenuItem";
            this.додатиРядокToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.додатиРядокToolStripMenuItem.Text = "Додати рядок";
            this.додатиРядокToolStripMenuItem.Click += new System.EventHandler(this.AddRow);
            // 
            // видалитиРядокToolStripMenuItem
            // 
            this.видалитиРядокToolStripMenuItem.Name = "видалитиРядокToolStripMenuItem";
            this.видалитиРядокToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.видалитиРядокToolStripMenuItem.Text = "Видалити рядок";
            this.видалитиРядокToolStripMenuItem.Click += new System.EventHandler(this.DeleteRow);
            // 
            // допомогаToolStripMenuItem
            // 
            this.допомогаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.проПроектToolStripMenuItem,
            this.умоваToolStripMenuItem});
            this.допомогаToolStripMenuItem.Name = "допомогаToolStripMenuItem";
            this.допомогаToolStripMenuItem.Size = new System.Drawing.Size(75, 20);
            this.допомогаToolStripMenuItem.Text = "Допомога";
            // 
            // проПроектToolStripMenuItem
            // 
            this.проПроектToolStripMenuItem.Name = "проПроектToolStripMenuItem";
            this.проПроектToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.проПроектToolStripMenuItem.Text = "Про проект";
            this.проПроектToolStripMenuItem.Click += new System.EventHandler(this.проПроектToolStripMenuItem_Click);
            // 
            // умоваToolStripMenuItem
            // 
            this.умоваToolStripMenuItem.Name = "умоваToolStripMenuItem";
            this.умоваToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.умоваToolStripMenuItem.Text = "Умова";
            this.умоваToolStripMenuItem.Click += new System.EventHandler(this.умоваToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // TableEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2188, 1205);
            this.Controls.Add(this.CurentElLabel);
            this.Controls.Add(this.ParsertextBox);
            this.Controls.Add(this.AddRowbutton);
            this.Controls.Add(this.DeleteRowbutton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AddColumnbutton);
            this.Controls.Add(this.DeleteColumnbutton);
            this.Controls.Add(this.TabledataGridView);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "TableEditor";
            this.Text = "TableEditor";
            ((System.ComponentModel.ISupportInitialize)(this.TabledataGridView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView TabledataGridView;
        private System.Windows.Forms.Button DeleteColumnbutton;
        private System.Windows.Forms.Button AddColumnbutton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button DeleteRowbutton;
        private System.Windows.Forms.Button AddRowbutton;
        private System.Windows.Forms.TextBox ParsertextBox;
        private System.Windows.Forms.Label CurentElLabel;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem зберегтиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem відкритиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem таблицяToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem допомогаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem додатиСтовпчикToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem видалитиСтовпчикToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem додатиРядокToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem видалитиРядокToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem проПроектToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem умоваToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem зберегтиToolStripMenuItem1;
    }
}

