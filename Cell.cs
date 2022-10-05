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
        private string _error;
        private string _name;
        private int row_, col_;          // Координати комірки
        public bool calcflag;
        public int Row { get { return row_; } }
        public int Col { get { return col_; } }

        // Геттери та сеттери
        public bool Cacluled { get { return calcflag; } set { calcflag = value; } }          // При оновленні таблиці
        public string Value { get { return _value; } set { _value = value; /* .Equals("0") ? "0" : "1"; */ } }
        public string Formula { get { return _formula; } set { _formula = value; } }
        public DataGridViewCell Parent { get { return _parent; } }
        public string Name { get { return _name; } }
        public string Error { get { return _error; } set { _error = value; } }
        public List<Cell> CellReferences { get; set; }  // Список комірок, від яких залежить значення даної комірки
        public Cell(DataGridViewCell parent, string name, int row, int col, string formula, string value = "")
        {
            row_ = row;
            col_ = col;
            _parent = parent;
            _name = name;
            _formula = formula;
            _value = value;
            _error = "";
            CellReferences = new List<Cell>();
            CellReferences.Add(new Cell());    // Добавити пусту (відсутню в таблиці) комірку
        }
        public Cell()  // Конструктор за замовчуванням (такої комірки немає)
        {
            _name = "";
            row_ = col_ = 0;
        }

        private void Calculate()
        {
        }


    }
}
