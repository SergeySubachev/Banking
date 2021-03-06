using System.Collections.Generic;

namespace Banking
{
    class Person
    {
        public string Id { private set; get; }
        public double InitBalance { private set; get; }
        public double Balance { set; get; }
        public int SheetNumber { private set; get; }
        public int Row { private set; get; }
        public List<(int month, double value)> Costs { private set; get; } = new List<(int, double)>();

        public Person(string id, double initBalance, int sheetNumber, int row)
        {
            Id = id;
            InitBalance = Balance = initBalance;
            SheetNumber = sheetNumber;
            Row = row;
        }
    }
}
