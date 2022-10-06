using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MicroExcel
{
    public class Cell
    {
        private DataGridViewCell _parent;
        private string _value;           // Є значення чи немає
        private string _formula;         // Формула в комірці
        private int row_, col_;          // Координати комірки
        public int Row { get { return row_; } }
        public int Col { get { return col_; } }

        // Геттери та сеттери
        public bool Calculed { get; set; }          // При оновленні таблиці
        public string Value { get { return _value; } set { _value = value; /* .Equals("0") ? "0" : "1"; */ } }
        public string Formula { get { return _formula; } set { _formula = value; } }
        public DataGridViewCell Parent { get { return _parent; } }
        public Cell(DataGridViewCell parent, string name, int row, int col, string formula, string value = "")
        {
            row_ = row;
            col_ = col;
            _parent = parent;
            _formula = formula;
            _value = value;
        }
        public Cell()  // Конструктор за замовчуванням (такої комірки немає)
        {
            row_ = col_ = 0;
        }

        private void Calculate()
        {
        }


    }
}
