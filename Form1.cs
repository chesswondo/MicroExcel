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
            cell.Tag = new Cell(cell, cellName, "0");
            cell.Value = "0";
            // cell.Value = cellName;  // Для відладки - ім'я комірки
        }
        // Ініціалізуємо всі комірки таблиці
        private void InitializeAllCells()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // row.DefaultCellStyle.Font = new Font(_defaultFontName, _defaultFontSize, GraphicsUnit.Point);
                foreach (DataGridViewCell cell in row.Cells)
                {
                    InitializeSingleCell(row, cell);
                }
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

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
    }
}
