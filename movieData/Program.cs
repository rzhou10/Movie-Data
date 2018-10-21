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
            Excel.Range franchiseCol = xlWorksheet.Columns[8];
            Excel.Range studioCol = xlWorksheet.Columns[9];

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

            System.Array listVals7 = (System.Array)franchiseCol.Cells.Value;
            string[] allFran = listVals7.OfType<object>().Select(o => o.ToString()).ToArray();

            System.Array listVals8 = (System.Array)studioCol.Cells.Value;
            string[] allStudio = listVals8.OfType<object>().Select(o => o.ToString()).ToArray();

            Stats s = new Stats();

            Dictionary<string, int> DictYears = s.AddYears(allYears);
            Dictionary<string, int> DictGenres = s.AddGenres(allGenre);
            Dictionary<string, int> DictDirs = s.AddDirectors(allDir);
            Dictionary<string, int> DictRating = s.AddGenres(allRating);
            Dictionary<string, int> DictFran = s.AddGenres(allFran);
            Dictionary<string, int> DictStudio = s.AddGenres(allStudio);

            s.SetGrossStats(allGross, allTitles);
            s.SetRunStats(allRun, allTitles);

            Console.WriteLine("Time Average: {0} min", s.GetRunAvg());
            Console.WriteLine("Shortest: {0} at {1} min", s.GetShortestFilm(), s.GetMinRun());
            Console.WriteLine("Longest: {0} at {1} min", s.GetLongestFilm(), s.GetMaxRun());

            Console.WriteLine("Gross Average: ${0}", s.GetGrossAvg());
            Console.WriteLine("Lowest: {0} at ${1}", s.GetLowestFilm(), s.GetMinGross());
            Console.WriteLine("Highest: {0} at ${1}", s.GetHighestFilm(), s.GetMaxGross());

            //write to excel sheets
            Excel._Worksheet xlWorksheet2 = movieWorkbook.Sheets[2];
            xlWorksheet2.Cells[1, 1] = "Directors";
            xlWorksheet2.Cells[1, 2] = "Count";

            int a = 1;
            foreach (var directors in DictDirs)
            {
                a++;
                xlWorksheet2.Cells[a, 1] = directors.Key;
                xlWorksheet2.Cells[a, 2] = directors.Value;
            }

            Excel._Worksheet xlWorksheet3 = movieWorkbook.Sheets[3];
            xlWorksheet3.Cells[1, 1] = "Years";
            xlWorksheet3.Cells[1, 2] = "Count";

            int b = 1;
            foreach (var years in DictYears)
            {
                b++;
                xlWorksheet3.Cells[b, 1] = years.Key;
                xlWorksheet3.Cells[b, 2] = years.Value;
            }

            Excel._Worksheet xlWorksheet4 = movieWorkbook.Sheets[4];
            xlWorksheet4.Cells[1, 1] = "Genre";
            xlWorksheet4.Cells[1, 2] = "Count";

            int c = 1;
            foreach (var genres in DictGenres)
            {
                c++;
                xlWorksheet4.Cells[c, 1] = genres.Key;
                xlWorksheet4.Cells[c, 2] = genres.Value;
            }

            Excel._Worksheet xlWorksheet5 = movieWorkbook.Sheets[5];
            xlWorksheet5.Cells[1, 1] = "Rating";
            xlWorksheet5.Cells[1, 2] = "Count";

            int d = 1;
            foreach (var ratings in DictRating)
            {
                d++;
                xlWorksheet5.Cells[d, 1] = ratings.Key;
                xlWorksheet5.Cells[d, 2] = ratings.Value;
            }

            Excel._Worksheet xlWorksheet6 = movieWorkbook.Sheets[6];
            xlWorksheet6.Cells[1, 1] = "Franchise";
            xlWorksheet6.Cells[1, 2] = "Count";

            int e = 1;
            foreach (var franchises in DictFran)
            {
                e++;
                xlWorksheet6.Cells[e, 1] = franchises.Key;
                xlWorksheet6.Cells[e, 2] = franchises.Value;
            }

            Excel._Worksheet xlWorksheet7 = movieWorkbook.Sheets[7];
            xlWorksheet7.Cells[1, 1] = "Studio";
            xlWorksheet7.Cells[1, 2] = "Count";

            int x = 1;
            foreach (var studio in DictStudio)
            {
                x++;
                xlWorksheet7.Cells[x, 1] = studio.Key;
                xlWorksheet7.Cells[x, 2] = studio.Value;
            }

            Console.WriteLine("Done!");

            Marshal.ReleaseComObject(yearCol);
            Marshal.ReleaseComObject(genreCol);
            Marshal.ReleaseComObject(dirCol);
            Marshal.ReleaseComObject(runCol);
            Marshal.ReleaseComObject(grossCol);
            Marshal.ReleaseComObject(ratingCol);
            Marshal.ReleaseComObject(franchiseCol);
            Marshal.ReleaseComObject(studioCol);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorksheet2);
            Marshal.ReleaseComObject(xlWorksheet3);
            Marshal.ReleaseComObject(xlWorksheet4);
            Marshal.ReleaseComObject(xlWorksheet5);
            Marshal.ReleaseComObject(xlWorksheet6);
            Marshal.ReleaseComObject(xlWorksheet7);

            movieWorkbook.Save();

            movieWorkbook.Close(0);
            Marshal.ReleaseComObject(movieWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}