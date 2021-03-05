using System.Collections.Generic;

namespace Banking
{
    class Person
    {
        public string Id { private set; get; }
        public double InitBalance { private set; get; }
        public int Row { private set; get; }
        public List<(int month, double sum)> Costs { private set; get; } = new List<(int, double)>();

        public Person(string id, double initBalance, int row)
        {
            Id = id;
            InitBalance = initBalance;
            Row = row;
        }
    }
}
