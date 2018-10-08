using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace movieData
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook movieWorkbook = xlApp.Workbooks.Open(@"C:\Users\Class2018\source\repos\movieData\movieData\Movie Data.xlsx");
            Excel._Worksheet xlWorksheet = movieWorkbook.Sheets[1];

            Excel.Range titleCol = xlWorksheet.Columns[1];
            Excel.Range yearCol = xlWorksheet.Columns[2];
            Excel.Range genreCol = xlWorksheet.Columns[3];
            Excel.Range dirCol = xlWorksheet.Columns[4];
            Excel.Range runCol = xlWorksheet.Columns[5];
            Excel.Range grossCol = xlWorksheet.Columns[6];
            Excel.Range ratingCol = xlWorksheet.Columns[7];

            System.Array listVals = (System.Array)titleCol.Cells.Value;
            string[] allTitles = listVals.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals1 = (System.Array)yearCol.Cells.Value;
            string[] allYears = listVals1.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals2 = (System.Array)genreCol.Cells.Value;
            string[] allGenre = listVals2.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals3 = (System.Array)dirCol.Cells.Value;
            string[] allDir = listVals3.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals4 = (System.Array)runCol.Cells.Value;
            string[] allRun = listVals4.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals5 = (System.Array)grossCol.Cells.Value;
            string[] allGross = listVals5.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals6 = (System.Array)ratingCol.Cells.Value;
            string[] allRating = listVals6.OfType<object>().Select(o => o.ToString()).ToArray();

            Stats s = new Stats();

            Dictionary<string, int> DictYears = s.AddYears(allYears);
            Dictionary<string, int> DictGenres = s.AddGenres(allGenre);
            Dictionary<string, int> DictDirs = s.AddDirectors(allDir);
            Dictionary<string, int> DictRating = s.AddGenres(allRating);

            s.SetGrossStats(allGross, allTitles);
            s.SetRunStats(allRun, allTitles);

            //write to excel sheets

            Excel._Worksheet xlWorksheet2 = movieWorkbook.Sheets[2];
            xlWorksheet2.Cells[1, 1] = "Directors";
            xlWorksheet2.Cells[1, 2] = "Count";

            int a = 1;
            foreach (var d in DictDirs)
            {
                a++;
                xlWorksheet2.Cells[a, 1] = d.Key;
                xlWorksheet2.Cells[a, 2] = d.Value;
            }

            Excel._Worksheet xlWorksheet3 = movieWorkbook.Sheets[3];
            xlWorksheet3.Cells[1, 1] = "Years";
            xlWorksheet3.Cells[1, 2] = "Count";

            int b = 1;
            foreach (var y in DictYears)
            {
                b++;
                xlWorksheet3.Cells[b, 1] = y.Key;
                xlWorksheet3.Cells[b, 2] = y.Value;
            }

            Excel._Worksheet xlWorksheet4 = movieWorkbook.Sheets[4];
            xlWorksheet4.Cells[1, 1] = "Genre";
            xlWorksheet4.Cells[1, 2] = "Count";

            int c = 1;
            foreach (var g in DictGenres)
            {
                c++;
                xlWorksheet4.Cells[c, 1] = g.Key;
                xlWorksheet4.Cells[c, 2] = g.Value;
            }

            Excel._Worksheet xlWorksheet5 = movieWorkbook.Sheets[5];
            xlWorksheet5.Cells[1, 1] = "Rating";
            xlWorksheet5.Cells[1, 2] = "Count";

            int e = 1;
            foreach (var r in DictRating)
            {
                e++;
                xlWorksheet5.Cells[e, 1] = r.Key;
                xlWorksheet5.Cells[e, 2] = r.Value;
            }

            Console.WriteLine("Done!");

            Marshal.ReleaseComObject(yearCol);
            Marshal.ReleaseComObject(genreCol);
            Marshal.ReleaseComObject(dirCol);
            Marshal.ReleaseComObject(runCol);
            Marshal.ReleaseComObject(grossCol);
            Marshal.ReleaseComObject(ratingCol);
            Marshal.ReleaseComObject(xlWorksheet);

            movieWorkbook.Close();
            Marshal.ReleaseComObject(movieWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
