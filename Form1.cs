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
        private bool _saved = true;

        static public MEParser parser = new MEParser();   // Доступний всім

       public int RowCount { get { return dataGridView.RowCount; }  }
        public int ColCount { get { return dataGridView.ColumnCount; } }

        public MicroExcel()
        {
            InitializeComponent();
            InitializeDataGridView(_maxRows,_maxCols);
            InitializeAllCells();
        }

        private void InitializeDataGridView(int rows, int cols)
        {
            dataGridView.AllowUserToAddRows = true;
            dataGridView.AllowUserToDeleteRows = true;
            dataGridView.ColumnCount = cols;
            dataGridView.RowCount = rows;

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
            Cell ns = new Cell(cell, cellName, row.Index + 1, cell.ColumnIndex + 1, "");
            cell.Tag = ns;
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

        private void ReinitializeAllCells()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null)
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
        private void resetCalculated()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    Cell cl = (Cell)dgvCell.Tag;
                    cl.Calculed = false;
                }
            }
        }

        private void reCalcAll()
        {

            for (bool ok = true; ; )
            {
                ok = true;
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell dgvCell in row.Cells)
                    {
                        Cell c = (Cell)dgvCell.Tag;
                        if (c.Formula != "")
                        {
                            resetCalculated();
                            c.Calculed = true;
                            try
                            {
                                c.Value = parser.Evaluate(c.Formula, this).ToString();
                            }
                            catch (Exception ee)
                            {
                                MessageBox.Show("Невірна формула: " + ee.Message, "Помилка", MessageBoxButtons.OK);
                                c.Value = "";
                                c.Formula = "";
                                ok = false;
                            }
                        }
                    }
                }
                if (ok) break;
            }
        }

        // Оновлення значення комірки 
        private void UpdateSingleCellValue(DataGridViewCell dgvCell)
        {
            Cell cell = (Cell)dgvCell.Tag;

            if (!_formulaView)
            {
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
                Close();
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
                resetCalculated();
                cell.Calculed = true;
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
                    MessageBox.Show("Невірна формула: " + ee.Message, "Помилка", MessageBoxButtons.OK);
                    cell.Value = "";
                    cell.Formula = "";
                }
            }
            reCalcAll();
            UpdateCellValues();
            //UpdateSingleCellValue(dgvCell);
            _saved = false;
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
                    _saved = true;
                }

            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Open()
        {
            if (!_saved)
            {
                DialogResult result = MessageBox.Show("Є незбережені зміни! Підтверждуєте загрузку?", "Незбережені зміни!", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.No) return;
            }
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
                        Cell ns = new Cell(cell, cellName, row.Index + 1, cell.ColumnIndex + 1, sr.ReadLine());
                        ns.Value = sr.ReadLine();
                        cell.Tag = ns;
                    }
                    UpdateCellValues();
                    sr.Close();
                    mystream.Close();
                    _saved = true;
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
            form.ShowDialog();
        }

        public Cell getCell(int row, int col)
        {
            //Cell cell = (Cell)dataGridView[col, row].Tag;
            //string g = "GetCell [" + row + "," + col + "] = " + cell.Value;
            //MessageBox.Show(g, "", MessageBoxButtons.OK);
            return (Cell)dataGridView[col, row].Tag;
        }

        private void ClosingForm(object sender, FormClosingEventArgs e)
        {
            if (!_saved)
            {
                DialogResult result = MessageBox.Show("Є незбережені зміни! Підтверждуєте завершення роботи?", "Незбережені зміни!", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.No) e.Cancel = true;
            }
            else
            {
                DialogResult result = MessageBox.Show("Вийти з програми?", "Завершення роботи", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.No)
                    e.Cancel = true;
            }
        }

        private void DelRow(object sender, EventArgs e)
        {
            if (RowCount <= 2) { MessageBox.Show("Подальше видалення неможливе"); return; }
            int ROWS = RowCount - 1;
            dataGridView.Rows.RemoveAt(ROWS-1);
            dataGridView.Rows[ROWS-1].HeaderCell.Value = "R" + ROWS.ToString();
            dataGridView.Refresh();
            reCalcAll();
        }

        private void DelCol(object sender, EventArgs e)
        {
            if (ColCount <= 2) { MessageBox.Show("Подальше видалення неможливе"); return; }
            dataGridView.Columns.RemoveAt(ColCount-1);
            dataGridView.Refresh();
            reCalcAll();

        }

        private void AddRow(object sender, EventArgs e)
        {
            dataGridView.Rows.Add();
            ReinitializeAllCells();
            FillHeaders();
            dataGridView.Refresh();
        }

        private void AddCol(object sender, EventArgs e)
        {
            DataGridViewColumn col = (DataGridViewColumn)dataGridView.Columns[0].Clone();
            dataGridView.Columns.Add(col);
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                DataGridViewCell cell = row.Cells[ColCount - 1];
                cell.Tag = null;
            }

            ReinitializeAllCells();
            FillHeaders();
            dataGridView.Refresh();
        }
    }
}