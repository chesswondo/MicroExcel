using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MicroExcel
{
    public partial class MicroExcel : Form
    {
        private const int _maxCols = 16;
        private const int _maxRows = 32;
        private const int _rowHeaderWidth = 70;  // Ширина заголовка рядків

        private bool _formulaView = false;       // Показувати формули чи значення


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
            cell.Value = cell.Formula.Length.ToString();

            UpdateSingleCellValue(dgvCell);
        }
    }
}
