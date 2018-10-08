using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Numerics;

namespace movieData
{
    class Stats
    {
        private BigInteger GrossAvg { get; set; }
        private int MaxGross { get; set; }
        private string HighestFilm { get; set; }
        private string LowestFilm { get; set; }
        private int MinGross { get; set; }
        private double RunAvg { get; set; }
        private string LongestFilm { get; set; }
        private string ShortestFilm { get; set; }
        private int MaxRun { get; set; }
        private int MinRun { get; set; }

        public Dictionary<string, int> AddDirectors(string[] arr)
        {
            Dictionary<string, int> directors = new Dictionary<string, int>();

            for (int i = 1; i < arr.Length; i++)
            {
                if (directors.ContainsKey(arr[i]))
                {
                    directors[arr[i]] = directors[arr[i]] += 1;
                }
                else
                {
                    directors.Add(arr[i], 1);
                }
            }

            return directors;
        }

        public Dictionary<string, int> AddGenres(string[] arr)
        {
            Dictionary<string, int> genre = new Dictionary<string, int>();

            for (int i = 1; i < arr.Length; i++)
            {
                if (genre.ContainsKey(arr[i]))
                {
                    genre[arr[i]] = genre[arr[i]] += 1;
                }
                else
                {
                    genre.Add(arr[i], 1);
                }
            }

            return genre;
        }

        public Dictionary<string, int> AddYears(string[] arr)
        {
            Dictionary<string, int> years = new Dictionary<string, int>();

            for (int i = 1; i < arr.Length; i++)
            {
                if (years.ContainsKey(arr[i]))
                {
                    years[arr[i]] = years[arr[i]] += 1;
                }
                else
                {
                    years.Add(arr[i], 1);
                }
            }

            return years;
        }

        public Dictionary<string, int> AddRatings(string[] arr)
        {
            Dictionary<string, int> ratings = new Dictionary<string, int>();

            for (int i = 1; i < arr.Length; i++)
            {
                if (ratings.ContainsKey(arr[i]))
                {
                    ratings[arr[i]] = ratings[arr[i]] += 1;
                }
                else
                {
                    ratings.Add(arr[i], 1);
                }
            }

            return ratings;
        }

        public void SetGrossStats(string[] arr, string[] arr1)
        {
            BigInteger sum = 0;
            int NON_NA = 0;
            MinGross = Int32.MaxValue;
            for (int i = 1; i < arr.Length; i++)
            {
                if (arr[i] != "N/A")
                {
                    int ValToString = System.Convert.ToInt32(arr[i]);
                    sum += ValToString;
                    NON_NA++;

                    if (ValToString > MaxGross)
                    {
                        MaxGross = ValToString;
                        HighestFilm = arr1[i];
                        
                    }
                    else if (ValToString < MinGross)
                    {
                        MinGross = ValToString;
                        LowestFilm = arr1[i];
                    }
                }
            }

            GrossAvg = sum / NON_NA;
        }

        public BigInteger GetGrossAvg()
        {
            return GrossAvg;
        }

        public int GetMaxGross()
        {
            return MaxGross;
        }

        public int GetMinGross()
        {
            return MinGross;
        }

        public string GetHighestFilm()
        {
            return HighestFilm;
        }

        public string GetLowestFilm()
        {
            return LowestFilm;
        }

        public void SetRunStats(string[] arr, string[] arr1)
        {
            int sum = 0;
            MinRun = Int32.MaxValue;
            for (int i = 1; i < arr.Length; i++)
            {
                int ValToString = System.Convert.ToInt32(arr[i]);
                sum += ValToString;

                if (ValToString > MaxRun)
                {
                    MaxRun = ValToString;
                    LongestFilm = arr1[i];
                }
                else if (ValToString < MinRun)
                {
                    MinRun = ValToString;
                    ShortestFilm = arr1[i];
                }
            }

            RunAvg = sum / arr.Length;
        }

        public double GetRunAvg()
        {
            return RunAvg;
        }

        public int GetMaxRun()
        {
            return MaxRun;
        }

        public int GetMinRun()
        {
            return MinRun;
        }

        public string GetLongestFilm()
        {
            return LongestFilm;
        }

        public string GetShortestFilm()
        {
            return ShortestFilm;
        }
    }
}