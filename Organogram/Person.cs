using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Organogram
{
    public class Person
    {
        public int id;
        public int parentID;
        public string name;
        public string surname;
        public string company;
        public string city;
        public string position;
        public double firstNumber;
        public double secondNumber;
        public double thirdNumber;
        public Person()
        {
            id = 0;
            parentID = 0;
            name = "";
            surname = "";
            company = "";
            city = "";
            position = "";
            firstNumber = 0;
            secondNumber = 0;
            thirdNumber = 0;
        }
    }
}
