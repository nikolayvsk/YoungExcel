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

namespace MyExcel1
{
    public partial class Form1 : Form
    {
        int currRow, currCol;
        int rows = 5;
        int cols = 5;
        public static Dictionary<string, Cell> dictionary = new Dictionary<string, Cell>();
        public Form1()
        {
            InitializeComponent();
            createSpreadsheet(rows, cols);
        }

        private void createSpreadsheet(int rows, int cols) // Створення початкової таблиці
        {
            for (int i = 0; i < cols; i++)
            {
                DataGridViewColumn column = new DataGridViewColumn();
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                column.CellTemplate = cell;
                string name = setColNum(i);
                column.HeaderText = name;
                column.Name = name;
                dgv.Columns.Add(column);
            }

            DataGridViewRow r = new DataGridViewRow();
            for (int i = 0; i < rows; i++)
            {
                dgv.Rows.Add();
            }
            setRowNum(dgv);

            Cell _cell;
            for (int j = 0; j < rows; j++)
            {
                for (int i = 0; i < cols; i++)
                {
                    string cellName = setColNum(i) + (j + 1).ToString();
                    _cell = new Cell();
                    _cell.Val = "0";
                    _cell.Exp = "";
                    dictionary.Add(cellName, _cell);
                }
            }
        }

        public string setColNum(int num) // виставлення значень A B C D
        {
            const int alphabet = 26;
            string name = "";
            int n = num;
            if (num < alphabet)
            {
                char c = (char)(num + 65);
                return (name + c);
            }

            for (int j = num; j >= alphabet; j = j / alphabet)
            {
                int mod = num % alphabet;
                int div = num / alphabet - 1;

                name += (char)(div + 65);

                if (div < alphabet)
                {
                    name += (char)(mod + 65);
                }
            }

            return name;
        }

        public void setRowNum(DataGridView dgv1) // виставлення 1, 2, 3, 4 і так далі
        {
            foreach (DataGridViewRow row in dgv1.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);
            }
        }

        private void addRowButton_Click(object sender, EventArgs e) // кнопочка додавання рядку
        {
            addRow(dgv);
            refreshCell();
        }

        public void addRow(DataGridView dgv1) // додавання рядку
        {
            DataGridViewRow row = new DataGridViewRow();
            dgv1.Rows.Add(row);
            Cell cell;
            for (int i = 0; i < dgv1.ColumnCount; i++)
            {
                string cellName = dgv1.Columns[i].HeaderText + (dgv1.RowCount).ToString();
                cell = new Cell();
                cell.Val = "0";
                cell.Exp = "";
                cell.Name = cellName;
                dictionary.Add(cellName, cell);
            }
            setRowNum(dgv1);
        }

        private void deleteRowButton_Click(object sender, EventArgs e) // кнопочка видалення рядку
        {
            if (dgv.Rows.Count == 1)
                MessageBox.Show("Попередження: кількість рядків вже 1, куди 0...");
            else
            {
                deleteRow(dgv);
                refreshCell();
            }
        }

        public void deleteRow(DataGridView dgv1) // видалення рядку
        {
            dgv1.Rows.RemoveAt(dgv1.Rows.Count - 1);
            for (int i = 0; i < dgv1.ColumnCount; i++)
            {
                string deletedCell;
                deletedCell = dgv1.Columns[i].HeaderText + (dgv1.Rows.Count + 1).ToString();
                dictionary.Remove(deletedCell);
            }
        }

        private void addColumnButton_Click(object sender, EventArgs e) // кнопочка додавання стовпчика
        {
            addColumn(dgv);
            refreshCell();
        }

        public void addColumn(DataGridView dgv1) // додавання стовпчика
        {
            DataGridViewColumn column = new DataGridViewColumn();
            DataGridViewCell cell = new DataGridViewTextBoxCell();
            column.CellTemplate = cell;
            string name = setColNum(dgv1.Columns.Count);
            column.Name = name;
            column.HeaderText = name;
            dgv1.Columns.Add(column);
            Cell new_cell;
            for (int i = 0; i < dgv1.RowCount; i++)
            {
                string cellName = name + (i + 1).ToString();
                new_cell = new Cell();
                new_cell.Val = "0";
                new_cell.Exp = "";
                new_cell.Name = cellName;
                dictionary.Add(cellName, new_cell);
            }
        }
        private void deleteColumnButton_Click(object sender, EventArgs e) // кнопочка видалення стовпчика
        {
            if (dgv.Columns.Count == 1)
                MessageBox.Show("Попередження: кількість стовпчиків вже 1, куди 0...");
            else
            {
                deleteCol(dgv);
                refreshCell();
            }
        }

        public void deleteCol(DataGridView dgv1) // видалення стовпчика
        {
            string colName = dgv1.Columns[dgv1.Columns.Count - 1].HeaderText;
            for (int i = 0; i < dgv1.RowCount; i++)
            {
                string deletedCell;
                deletedCell = colName + (i + 1).ToString();
                dictionary.Remove(deletedCell);
            }
            dgv1.Columns.RemoveAt(dgv1.Columns.Count - 1);
        }

        private void refreshCell() // оновлювання данних ячейки в табличці
        {
            string errCell = "";
            for (int i = 0; i < dgv.RowCount; i++)
            {
                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    string cellName = dgv.Columns[j].HeaderText + (i + 1).ToString();
                    if (!dictionary[cellName].WasContained)
                    {
                        dgv[j, i].Value = "";
                        calculateCell(cellName, ref errCell);
                    }
                    if (dictionary[cellName].Exp != "")
                    {
                        if (!(dictionary[cellName].Val == "Error" || dictionary[cellName].Val == "ErrorCycle" || dictionary[cellName].Val == "ErrorDivZero"))
                        {
                            dgv[j, i].Value = dictionary[cellName].Val;
                        }
                    }
                    else
                    {
                        dgv[j, i].Value = "";
                    }
                }
            }
            foreach (var item in dictionary)
            {
                dictionary[item.Key].WasContained = false;
            }
            if (errCell != "")
            {
                MessageBox.Show("Помилка в комірці:\n" + errCell);
            }
        }

        public void calculateCell(string cellName, ref string errCel) // Обрахунок ячеєчки
        {
            Calculator calculator = new Calculator(dictionary);
            dictionary[cellName].WasContained = true;
            foreach (string el in dictionary[cellName].CellReference)
            {
                if (!dictionary.ContainsKey(el))
                {
                    dictionary[cellName].Val = "Error";
                    errCel += "[" + cellName + "]";
                    dictionary[cellName].CellReference.Remove(el);
                    return;
                }
                else if (!dictionary[el].WasContained)
                {
                    calculateCell(el, ref errCel);
                }
            }

            string expr = dictionary[cellName].Exp;
            if (expr != "")
            {
                var res = calculator.Evaluate(expr, cellName, ref errCel);
                if (dictionary[cellName].Val != "Error" && dictionary[cellName].Val != "ErrorCycle" && dictionary[cellName].Val != "ErrorDivZero")
                    dictionary[cellName].Val = res.ToString();
            }
        }

        private void buttonEnter_Click(object sender, EventArgs e) // занесення значень в ячейку через кнопку "Enter"
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
            string str = textBox1.Text;
            dictionary[cellName].Exp = str;
            if (str == "")
            {
                dictionary[cellName].Val = "0";
            }
            dictionary[cellName].CellReference = new List<string>();
            refreshCell();
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
            textBox1.Text = dictionary[cellName].Exp;
        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            if (dgv[currCol, currRow].Value != null)
            {
                string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
                textBox1.Text = dictionary[cellName].Exp;
            }
            else
            {
                textBox1.Text = "";
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e) // кнопочка збереження файлу
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog1.OpenFile();
                StreamWriter sw = new StreamWriter(fs);
                save(sw, dgv);
                sw.Close();
                fs.Close();
            }
            refreshCell();
        }

        public void save(StreamWriter sw, DataGridView dgv1) // збереження файлу
        {
            int row = dgv1.RowCount;
            int col = dgv1.ColumnCount;
            sw.WriteLine(row);
            sw.WriteLine(col);
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    string cellName = setColNum(j) + (i + 1).ToString();
                    sw.WriteLine(cellName);
                    sw.WriteLine(dictionary[cellName].Exp);
                    int count = dictionary[cellName].CellReference.Count;
                    sw.WriteLine(count);
                    if (count != 0)
                    {
                        foreach (string el in dictionary[cellName].CellReference)
                        {
                            sw.WriteLine(el);
                        }
                    }
                }
            }
        }

        private void openAsToolStripMenuItem_Click(object sender, EventArgs e) // кнопочка відкриття файлу
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            dgv.Rows.Clear();
            dgv.Columns.Clear();
            dictionary.Clear();
            int rows, cols;
            try
            {
                string str = sr.ReadLine();
                rows = int.Parse(str);
                str = sr.ReadLine();
                cols = int.Parse(str);
                createSpreadsheet(rows, cols);
                open(sr, rows, cols);
                refreshCell();
            }
            catch
            {
                MessageBox.Show("Помилка при відкритті файлу!");
                dgv.Rows.Clear();
                dgv.Columns.Clear();
                dictionary.Clear();
            }
        }

        public void open(StreamReader sr, int row, int col) // відкриття файлу
        {
            string cellName;
            string expr;
            int count;
            string reference;
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    cellName = sr.ReadLine();
                    expr = sr.ReadLine();
                    dictionary[cellName].Name = cellName;
                    dictionary[cellName].Exp = expr;
                    count = int.Parse(sr.ReadLine());
                    for (int k = 0; k < count; k++)
                    {
                        reference = sr.ReadLine();
                        dictionary[cellName].CellReference.Add(reference);
                    }
                }
            }
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e) // Інформація, що взагалі файл вміє
        {
            MessageBox.Show("Привіт! Програма виконує такі операції та функції:\n" +
                            "1) +, -, *, / (бінарні операції);\n" +
                            "2) mod, div;\n" +
                            "3) +, - (унарні операції);\n" +
                            "4) ^ (піднесеня у степінь).\n" +
                            "Функціонал програми:\n"+
                            "Додавання/видалення рядку, стовпчика;\n"+
                            "Обрахунок значень в ячейці та ячейок між собою;\n" +
                            "Також програма вміє зберігати та відкривати файли.\n" +
                            "Щоб не виникало непорозумінь, програма рахує тільки цілі числа.\n" +
                            "Щасти!");
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Підтверджуєте вихід з програми?", "Значочок", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question))
            {
                //Application.Exit();
            }
            else { e.Cancel = true; }
        }
    }
}
