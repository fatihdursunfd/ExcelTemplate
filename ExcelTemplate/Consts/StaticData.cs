namespace ExcelTemplate.Consts
{
    public static class StaticData
    {
        public static readonly string[] ColumnNames = new string[]
        {
            "Name",
            "Surname",
            "Age",
            "Country",
            "City",
            "Phone Number",
            "Email",
            "Point"
        };

        public static readonly Dictionary<string, List<string>> CountriesWithCities = new()
        {
            { "Turkey", new List<string> { "Trabzon", "Istanbul", "Ankara", "Izmir", "Bursa", "Antalya", "Adana", "Konya" } },
            { "USA", new List<string> { "New York", "Los Angeles", "Chicago", "Houston", "Phoenix", "Philadelphia", "San Antonio" } },
            { "Germany", new List<string> { "Berlin", "Munich", "Frankfurt", "Hamburg", "Cologne", "Stuttgart", "Düsseldorf" } },
            { "France", new List<string> { "Paris", "Lyon", "Marseille", "Nice", "Toulouse", "Nice", "Nantes" } },
            { "Italy", new List<string> { "Rome", "Milan", "Naples", "Turin", "Palermo", "Genoa", "Bologna" } },
            { "Spain", new List<string> { "Madrid", "Barcelona", "Valencia", "Seville", "Zaragoza", "Malaga", "Murcia" } },
            { "Canada", new List<string> { "Toronto", "Vancouver", "Montreal", "Calgary", "Ottawa", "Edmonton", "Quebec City" } },
            { "Australia", new List<string> { "Sydney", "Melbourne", "Brisbane", "Perth", "Adelaide", "Gold Coast", "Canberra" } }
        };

    }
}