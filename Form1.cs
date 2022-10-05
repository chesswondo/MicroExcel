using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace MicroExcel
{
    public partial class MicroExcel : Form
    {
        public const int _maxCols = 16;
        public const int _maxRows = 32;
        private const int _rowHeaderWidth = 70;  // Ширина заголовка рядків

        private bool _formulaView = false;       // Показувати формули чи значення

        static public MEParser parser = new MEParser();   // Доступний всім

        public MicroExcel()
        {
            InitializeComponent();
            InitializeDataGridView();
            InitializeAllCells();
        }

        private void InitializeDataGridView()
        {
            dataGridView.AllowUserToAddRows = false;
            dataGridView.ColumnCount = _maxCols;
            dataGridView.RowCount = _maxRows;

            FillHeaders();

            dataGridView.AutoResizeRows();
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView.RowHeadersWidth = _rowHeaderWidth;
        }

        // Sets up the data grid view's headers with corresponding indices.
        private void FillHeaders()
        {
            foreach (DataGridViewColumn col in dataGridView.Columns)
            {
                col.HeaderText = "C" + (col.Index + 1);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.HeaderCell.Value = "R" + (row.Index + 1);
            }
        }
        // Ініціалізуємо одну комірку за її координатами
        private void InitializeSingleCell(DataGridViewRow row, DataGridViewCell cell)
        {
            // Компонуємо її ім'я
            string cellName = "R" + (row.Index + 1).ToString() + "C" + (cell.ColumnIndex + 1).ToString();
            // Створюємо комірку; cell — батьківський об'єкт
            cell.Tag = new Cell(cell, cellName, "");
            cell.Value = "";
            // cell.Value = cellName;  // Для відладки - ім'я комірки
        }
        // Ініціалізуємо всі комірки таблиці
        private void InitializeAllCells()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // можна змінити шрифт, але поки не треба :)
                // row.DefaultCellStyle.Font = new Font(_defaultFontName, _defaultFontSize, GraphicsUnit.Point);
                foreach (DataGridViewCell cell in row.Cells)
                {
                    InitializeSingleCell(row, cell);
                }
            }
        }

        // Оновлення комірок
        // Оновлення всіх комірок
        private void UpdateCellValues()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    UpdateSingleCellValue(dgvCell);
                }
            }
        }
        // Оновлення всіх комірок за виключенням одної
        private void UpdateCellValues(DataGridViewCell invoker)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    if (invoker != dgvCell)
                    {
                        UpdateSingleCellValue(dgvCell);
                    }
                }
            }
        }

        // Оновлення значення комірки 
        private void UpdateSingleCellValue(DataGridViewCell dgvCell)
        {
            Cell cell = (Cell)dgvCell.Tag;

            if (!_formulaView)
            {
                // Виводити значення комірки
                /* doit
                if (cell.Formula.Equals("") || Regex.IsMatch(cell.Formula,@"^\d+$"))
                {
                    dgvCell.Value = cell.Value;
                }
                else
                {
                    dgvCell.Value = cell.Evaluate();
                }
                */
                dgvCell.Value = cell.Value;
            }
            else
            {
                dgvCell.Value = cell.Formula;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            InitializeAllCells();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void QuitMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Quit?", "Exit", MessageBoxButtons.YesNo);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                Close();
            }
        }

        private void formulaView_CheckedChanged(object sender, EventArgs e)
        {
            _formulaView = !formulaView.Checked;
            miValues.Checked = !_formulaView;
            miFormulas.Checked = _formulaView;
            //if (formulaView.Checked)
            //    MessageBox.Show("Checked");
            //else
            //    MessageBox.Show("UnChecked");
            UpdateCellValues();

        }

         private void miValues_Click(object sender, EventArgs e)
        {
            formulaView.Checked = true;
            miValues.Checked = true;
            miFormulas.Checked = false;
        }

        private void miFormulas_Click(object sender, EventArgs e)
        {
            formulaView.Checked = false;
            miValues.Checked = false;
            miFormulas.Checked = true;

        }

        private void miNew_Click(object sender, EventArgs e)
        {
            InitializeAllCells();
        }
        // Натискання на комірку
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 || e.ColumnIndex == -1)
            {
                return;
            }

            Cell cell = (Cell)dataGridView[e.ColumnIndex, e.RowIndex].Tag;
            DataGridViewCell dgvCell = cell.Parent;

            if (!dgvCell.ReadOnly)
            {
                dataGridView.BeginEdit(true);
            }
        }
        // Редагування комірки
        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Знайшли об'єкт, що відповідає комірці
            Cell cell = (Cell)dataGridView[e.ColumnIndex, e.RowIndex].Tag;

            //doit CellManager.Instance.CurrentCell = cell;

            // Значення комірки для редагування береться як формула з об'єкту
            DataGridViewCell dgvCell = cell.Parent;
            dgvCell.Value = cell.Formula;

        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // По закінченню редагування оновлюємо наш об'єкт комірки
            Cell cell = (Cell)dataGridView[e.ColumnIndex, e.RowIndex].Tag;
            DataGridViewCell dgvCell = cell.Parent;

            if (dgvCell.Value == null)
            {
                cell.Formula = "";
                cell.Value = "";
                dgvCell.Value = "";
            }

            // doit
            //ClearRemovedReferences(cell);
            //ResolveCellFormula(cell, dgvCell);
            // Вместо doit пока такое...
            cell.Formula = dgvCell.Value.ToString();
            if (cell.Formula == null || cell.Formula.Length == 0)
                cell.Value = "";
            else
            {
                try
                {
                    if (!checkParens(cell.Formula))
                        throw new ArgumentException("Проблема з дужками у виразі");
                    cell.Value = parser.Evaluate(cell.Formula, this).ToString();
                }
                catch (ArgumentException ee)
                {
                    MessageBox.Show(ee.Message, "Помилка", MessageBoxButtons.OK);
                    cell.Value = "";
                    cell.Formula = "";
                }
                catch (Exception ee)
                {
                    MessageBox.Show("Невірна формула", "Помилка", MessageBoxButtons.OK);
                    cell.Value = "";
                    cell.Formula = "";
                }
            }
            UpdateSingleCellValue(dgvCell);
        }

        private bool checkParens(string f)
        {
            int stk = 0;
            foreach(var c in f)
            {
                if (c == '(') stk++; 
                else if (c == ')') stk--;
                if (stk < 0) return false;
            }
            return stk == 0;
        }

        private void Save()
        {
            Stream mystream;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((mystream = saveFileDialog.OpenFile()) != null)
                {
                    StreamWriter sw = new StreamWriter(mystream);
                    sw.WriteLine(dataGridView.RowCount.ToString() + " " + dataGridView.ColumnCount.ToString());

                    for (int i=0; i < dataGridView.RowCount; i++)
                    {
                        for (int j = 0; j < dataGridView.ColumnCount; j++)
                        {
                            Cell cell = (Cell)dataGridView.Rows[i].Cells[j].Tag;
                            sw.WriteLine(i.ToString()+" " + j.ToString());
                            sw.WriteLine(cell.Formula);
                            sw.WriteLine(cell.Value);
                        }
                    }
                    sw.Close();
                    mystream.Close();
                }

            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Open()
        {
            Stream mystream;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((mystream = openFileDialog.OpenFile()) != null)
                {
                    StreamReader sr = new StreamReader(mystream);
                    var Size = sr.ReadLine().Split(' ');
                    dataGridView.AllowUserToAddRows = false;
                    dataGridView.ColumnCount = int.Parse(Size[1]);
                    dataGridView.RowCount = int.Parse(Size[0]);

                    FillHeaders();

                    dataGridView.AutoResizeRows();
                    dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    dataGridView.RowHeadersWidth = _rowHeaderWidth;

                    while(sr.Peek() != -1)
                    {
                        var coord = sr.ReadLine().Split(' ');
                        DataGridViewRow row = dataGridView.Rows[int.Parse(coord[0])];
                        DataGridViewCell cell = row.Cells[int.Parse(coord[1])];
                        string cellName = "R" + (row.Index + 1).ToString() + "C" + (cell.ColumnIndex + 1).ToString();
                        Cell ns = new Cell(cell, cellName, sr.ReadLine());
                        ns.Value = sr.ReadLine();
                        cell.Tag = ns;
                    }
                    UpdateCellValues();
                    sr.Close();
                    mystream.Close();
                }

            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            Open();
        }

        private void miSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void miOpen_Click(object sender, EventArgs e)
        {
            Open();
        }

        private void miAbout_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.Show();
        }

        public Cell getCell(int row, int col)
        {
            //Cell cell = (Cell)dataGridView[col, row].Tag;
            //string g = "GetCell [" + row + "," + col + "] = " + cell.Value;
            //MessageBox.Show(g, "", MessageBoxButtons.OK);
            return (Cell)dataGridView[col, row].Tag;
        }
    }
}