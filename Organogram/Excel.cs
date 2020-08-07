using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace Organogram
{
    public class Excel
    {
        public Application excelApp;
        public Workbook excelBook;
        public _Worksheet excelSheet;
        public Range excelRange;
        public string path;

        public int rowCount;
        public int colCount;

        public Person[] people;

        /// <summary>
        /// Gets access to .csv file, reads amount of rows, cols etc. Writes all data to "people" array.
        /// </summary>
        /// <param name="p">Path to .csv file</param>
        public Excel(string p)
        {
            excelApp = new Application();
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            path = p;
            excelBook = excelApp.Workbooks.Open(path);
            excelSheet = excelBook.Sheets[1];
            excelRange = excelSheet.UsedRange;

            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;

            people = ParseFields(people);
        }

        // Parsing all data.
        public Person[] ParseFields(Person[] people)
        {
            Regex CSVParser = new Regex(";");
            String[] Fields;
            people = new Person[rowCount + 1];
            for (int i = 1; i < rowCount + 1; i++)
            {
                people[i] = new Person();

                Fields = CSVParser.Split(excelRange.Cells[i, 1].Value2);
                Fields = Fields.ToArray();
                people[i].id = Int32.Parse(Fields.ElementAt(0));
                people[i].parentID = Int32.Parse(Fields.ElementAt(1));
                people[i].name = Fields.ElementAt(2);
                people[i].surname = Fields.ElementAt(3);
                people[i].company = Fields.ElementAt(4);
                people[i].city = Fields.ElementAt(5);
                people[i].position = Fields.ElementAt(6);
                people[i].firstNumber = Int32.Parse(Fields.ElementAt(7));
                people[i].secondNumber = Int32.Parse(Fields.ElementAt(8));
                people[i].thirdNumber = Int32.Parse(Fields.ElementAt(9));
            }
            return people;
        }
        

        /// <summary>
        /// Gets all the children.
        /// </summary>
        /// <param name="id">Used to get children of a node with this id. 0 returns whole tree</param>
        /// <param name="generation">Used to keep track how deep we are. 0 at the begining</param>
        public void GetChildren(int id, int generation)
        {
            Person[] children = new Person[rowCount + 1];
            int howMany = 0;

            // Searching for children.
            for (int i = 1; i < rowCount + 1; i++)
            {
                if (people[i].parentID == id)
                {
                    children[howMany] = people[i];
                    howMany++;
                }
            }

            // Sorting children by ID ASC.
            children = children.Where(x => x != null).ToArray();
            if (children.Length > 0)
            {
                children = children.OrderBy(x => x.id).ToArray();

                for (int i = 0; i < children.Length; i++)
                {
                    if (children[i].parentID != 0)
                    {
                        string arrow = String.Concat(Enumerable.Repeat("--", generation));
                        Console.Write(arrow + ">");
                    }
                    Console.WriteLine(children[i].name + " " + children[i].surname + ", " + children[i].company + ", " + children[i].position);
                    GetChildren(children[i].id, generation+1);
                }
            }
        }

    }
}
