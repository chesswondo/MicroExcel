using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MicroExcel
{
    class Cell
    {
        private DataGridViewCell _parent;
        private string _value;           // Є значення чи немає
        private string _formula;         // Формула в комірці
        private string _error;
        private string _name;

        // Геттери та сеттери
        public string Value { get { return _value; } set { _value = value; /* .Equals("0") ? "0" : "1"; */ } }
        public string Formula { get { return _formula; } set { _formula = value; } }
        public DataGridViewCell Parent { get { return _parent; } }
        public string Name { get { return _name; } }
        public string Error { get { return _error; } set { _error = value; } }
        public List<Cell> CellReferences { get; set; }  // Список комірок, від яких залежить значення даної комірки
        public Cell(DataGridViewCell parent, string name, string formula)
        {
            _parent = parent;
            _name = name;
            _formula = formula;
            _value = "";
            _error = "";
            CellReferences = new List<Cell>();
            CellReferences.Add(new Cell());    // Добавити пусту (відсутню в таблиці) комірку
        }
        public Cell()  // Конструктор за замовчуванням (такої комірки немає)
        {
            _name = "";
        }


    }
}
