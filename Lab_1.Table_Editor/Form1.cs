using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace Lab_1.Table_Editor
{
    public partial class TableEditor : Form
    {
        int columns = 10;
        int rows = 10;
        Table Main_table;
        bool Is_Formula_Saved;
        string path;
        //string[] Alpeabet = new string[]( "A", "B", "C", "D", "E",
        //    "F", "G", "H", "I", "J", "K", "L", "M", "N",
        //    "O", "P", "Q", "R", "S", "T", "U", "V", "W",
        //    "X", "Y", "Z" );
        public TableEditor()
        {
            InitializeComponent();
            Main_table = new Table(rows, columns);
            FillDataGrid();
        }

        public TableEditor(string Path, int Rows, int Columns)
        {
            InitializeComponent();
            this.path = Path;
            this.rows = Rows;
            this.columns = Columns;
            Main_table = new Table(Path, rows, columns);
            FillDataGrid();
        }

        private void FillDataGrid()
        {
            TabledataGridView.RowHeadersWidth = 60;
            TabledataGridView.ColumnCount = columns;
            TabledataGridView.RowCount = rows;
            for (int i = 0; i < columns; i++)
            {
                if (i < 26)
                    TabledataGridView.Columns[i].Name = Convert.ToChar(65 + i).ToString();
                else
                {
                    TabledataGridView.Columns[i].Name = "";
                    //string name = "";
                    int n = 26;
                    for (; ((i / n) >= 26); n *= 26) ;
                    for (; n != 1; n /= 26)
                        TabledataGridView.Columns[i].Name += Convert.ToChar(64 + ((i / n) % 26)).ToString();
                    TabledataGridView.Columns[i].Name += Convert.ToChar(65 + (i % 26)).ToString();
                }
            }
            for (int i = 0; i < rows; i++)
                TabledataGridView.Rows[i].HeaderCell.Value = (i + 1).ToString();
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    if (!(Main_table.Cells[i, j].IsNULL()))
                        TabledataGridView.Rows[i].Cells[j].Value = Main_table.Cells[i, j].value;
                    else TabledataGridView.Rows[i].Cells[j].Value = "";
        }

        private void AddColumn(object sender, EventArgs e)
        {
            TabledataGridView.Columns.Add(new DataGridViewColumn(TabledataGridView.Rows[0].Cells[0]));
            columns++;
            Main_table = new Table(rows, columns);
            FillDataGrid();
        }

        private void AddRow(object sender, EventArgs e)
        {
            TabledataGridView.Rows.Add(new DataGridViewColumn());
            rows++;
            Main_table = new Table(rows, columns);
            FillDataGrid();
        }

        private void DeleteRow(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Видалити рядок?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (res == DialogResult.Yes)
            {
                if (rows == 1)
                {
                    MessageBox.Show("У таблиці один рядок", "Неможливо виконати операцію", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                rows--;
                TabledataGridView.Rows.RemoveAt(rows - 1);
                Main_table.rows--;
                //Main_table.SaveTable(path);
                Main_table = new Table(rows, columns);
                FillDataGrid();
            }
        }

        private void DeleteColumn(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Видалити стовпчик?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (res == DialogResult.Yes)
            {
                if (columns == 1)
                {
                    MessageBox.Show("У таблиці один стовпчик", "Неможливо виконати операцію", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                columns--;
                TabledataGridView.Columns.RemoveAt(columns - 1);
                Main_table.rows--;
                //Main_table.SaveTable(path);
                Main_table = new Table(rows, columns);
                FillDataGrid();
            }
        }

        private void TabledataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurentElLabel.Text = TabledataGridView.Columns[e.ColumnIndex].Name + 
                TabledataGridView.Rows[e.RowIndex].HeaderCell.Value;
            ParsertextBox.Text = Main_table.Cells[e.RowIndex, e.ColumnIndex].expression;
        }
        
        private void TabledataGridView_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void ParsertextBox_TextChanged(object sender, EventArgs e)
        {
            Is_Formula_Saved = false;
        }

        private void ParsertextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SaveExp_from_Textbox();
            }
        }
        void SaveExp_from_Textbox ()
        {
            int save_c = TabledataGridView.SelectedCells[0].ColumnIndex;
            int save_r = TabledataGridView.SelectedCells[0].RowIndex;
            Main_table.Cells[save_r, save_c].expression = ParsertextBox.Text;
            Main_table.EvaluateCells();
            if (Main_table.IsError)
            {
                MessageBox.Show(Main_table.ParsMess);
                Is_Formula_Saved = false;
            }
            else
            {
                FillDataGrid();
                Is_Formula_Saved = true;
            }
        }
        private void SaveTable(object sender, EventArgs e)
        {
            Main_table.SaveTable(path);
        }

        private void проПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(@"Лабораторна робота №1:
Розробка програми для роботи з електронними таблицями
Виконав студент групи К-26 Назарчук Сергій");
        }

        private void умоваToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(@"Варіант 19: 
Операції:
1) +, -, (унарні операції);
2) ^ (піднесення у ступінь);
3) max(x,y), min(x,y);
4) mmax(x1,x2,...,xN), mmin(x1,x2,...,xN) (N>=1);");
        }

        private void зберегтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Path = new SaveFileDialog();
            Path.InitialDirectory = "f:\\";
            Path.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
            Path.ShowDialog();
            Path.DefaultExt = ".xml";
            this.path = Path.FileName;
            Main_table.SaveTable(Path.FileName);
        }

        private void відкритиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            openFileDialog1.InitialDirectory = "f:\\";
            openFileDialog1.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
                TableEditor table = new TableEditor(filePath, Main_table.rows, Main_table.columns);
                table.Show();
            }
        }

        private void зберегтиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Main_table.SaveTable(path);
        }
    }
    public class Table
    {
        public bool IsError;
        public string ParsMess;
        public Cell[,] Cells = { };
        private Parser parser;
        public int rows;
        public int columns;
        public Table(string Path, int Rows, int Columns)
        {
            Cells = new Cell[Rows, Columns];
            for (int i = 0; i < Rows; i++)
                for (int j = 0; j < Columns; j++)
                    Cells[i, j] = new Cell();
            this.rows = Rows;
            this.columns = Columns;
            ReadTable(Path);
            parser = new Parser(Cells, Rows, Columns);
            EvaluateCells();
        }
        public Table (int Rows, int Columns)
        {
            Cells = new Cell[Rows, Columns];
            for (int i = 0; i < Rows; i++)
                for (int j = 0; j < Columns; j++)
                    Cells[i, j] = new Cell();
            this.rows = Rows;
            this.columns = Columns;
            parser = new Parser(Cells, Rows, Columns);
        }
        public void EvaluateCells()
        {
            IsError = false;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    if (Cells[i, j].expression != "")
                    {
                        Cells[i, j].value = parser.Eval(Cells[i, j].expression);
                        Parser.varsInFormula.Clear();
                        if (parser.Error)
                        {
                            IsError = true;
                            ParsMess = parser.message;
                            return;
                        }
                    }
        }
        public void SaveTable(string path)
        {

            XDocument xdoc = new XDocument();
            XElement table = new XElement("table");
            XAttribute colAttr = new XAttribute("columns", columns.ToString());
            XAttribute rowAttr = new XAttribute("rows", rows.ToString());
            table.Add(colAttr);
            table.Add(rowAttr);
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    if (!Cells[i, j].IsNULL())
                    {
                        XElement cell = new XElement("cell", Cells[i, j].expression);
                        XAttribute XIndex = new XAttribute("X", (i + 1).ToString());
                        XAttribute YIndex = new XAttribute("Y", (j + 1).ToString());
                        cell.Add(XIndex);
                        cell.Add(YIndex);
                        table.Add(cell);
                    }
            xdoc.Add(table);
            xdoc.Save(path);
        }
        public void ReadTable(string path)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            int X;
            int Y;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    Cells[i, j] = new Cell();
            foreach (XmlElement table in xmlDoc.DocumentElement.ChildNodes)
            {
                foreach (XmlElement cellinfo in xmlDoc.DocumentElement.ChildNodes)
                {
                    Int32.TryParse(cellinfo.GetAttribute("X"), out X);
                    Int32.TryParse(cellinfo.GetAttribute("Y"), out Y);
                    Cells[X - 1, Y - 1] = new Cell(cellinfo.InnerText);
                }
            }
        }
    }
    public class Cell
    {
        public string name;
        public string expression;
        public double value;
        public Cell (string expr)
        {
            expression = expr;
            value = 0;
        }
        public Cell()
        {
            expression = "";
            value = 0;
        }
        public bool IsNULL()
        {
            if (expression == null)
                return true;
            return false;
        }
    }
}
