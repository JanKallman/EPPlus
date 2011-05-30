using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    public class TestDTO
    {
        public string NameVar;

        public int Id { get; set; }
        public string Name { get; set; }
        public TestDTO dto { get; set; }
        public DateTime Date { get; set; }
        public bool Boolean { get; set; }

        public string GetNameID()
        {
            return Id + "," + Name;
        }
    }
}
