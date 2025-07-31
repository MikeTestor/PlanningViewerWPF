using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

public class WeeklyEntry
{
    public string ResourceName { get; set; }
    public string CustomerProject { get; set; }
    public int WeekNumber { get; set; }
    public int Year { get; set; }
    public double AllocatedHours { get; set; }
    public double AvailableHours { get; set; }
}

class Program
{
    static void Main(string[] args)
    {
        string path = "report28-07-small.xml";

        XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
        XDocument doc = XDocument.Load(path);

        var table = doc.Descendants(ss + "Worksheet")
                       .FirstOrDefault(w => (string)w.Attribute(ss + "Name") == "NavigateDailyPlanning")
                       ?.Element(ss + "Table");

        var rows = table?.Elements(ss + "Row").ToList();
        if (rows == null || rows.Count == 0)
        {
            Console.WriteLine("No rows found.");
            return;
        }

        var weeklyEntries = new List<WeeklyEntry>();
        string currentResource = null;
        List<(int Week, int Year)> weekMap = new();

        // Week titles are on row 7 (index 6)
        var headerRow = rows[6];
        var weekTitles = headerRow.Elements(ss + "Cell").Skip(2).ToList(); // skip "Resource" and "Customer"
        foreach (var titleCell in weekTitles)
        {
            var data = titleCell.Element(ss + "Data")?.Value;
            if (!string.IsNullOrEmpty(data) && data.StartsWith("Week"))
            {
                var parts = data.Split(' ');
                if (int.TryParse(parts[1], out int week) && int.TryParse(parts[2], out int year))
                {
                    for (int i = 0; i < 3; i++) // 3 columns per week
                        weekMap.Add((week, year));
                }
            }
        }

        // Data starts from row 9 (index 8)
        for (int i = 8; i < rows.Count; i++)
        {
            var row = rows[i];
            var cells = row.Elements(ss + "Cell").ToList();

            if (cells.Count == 0)
                continue;

            var firstCell = cells[0].Element(ss + "Data")?.Value;
            if (!string.IsNullOrWhiteSpace(firstCell) && !firstCell.StartsWith("Cus"))
            {
                currentResource = firstCell.Trim();
                continue;
            }

            if (currentResource == null || cells.Count < 3)
                continue;

            var customerProject = cells[1].Element(ss + "Data")?.Value ?? "(unknown)";
            var cellIndex = 2;

            for (int w = 0; w < weekMap.Count / 3; w++)
            {
                var (week, year) = weekMap[w * 3];

                double allocated = GetCellNumber(cells, cellIndex++);
                double available = GetCellNumber(cells, cellIndex++);
                cellIndex++; // skip utilization

                weeklyEntries.Add(new WeeklyEntry
                {
                    ResourceName = currentResource,
                    CustomerProject = customerProject,
                    WeekNumber = week,
                    Year = year,
                    AllocatedHours = allocated,
                    AvailableHours = available
                });
            }
        }

        // Display result

        foreach(var entry in weekTitles)
        {
            var data = entry.Element(ss + "Data")?.Value;
            if (!string.IsNullOrEmpty(data))
            {
                Console.WriteLine(data);
            }
        }

        foreach (var entry in weeklyEntries)
        {
            Console.WriteLine($"{entry.ResourceName} | {entry.CustomerProject} | Week {entry.WeekNumber}/{entry.Year} | Allocated: {entry.AllocatedHours}, Available: {entry.AvailableHours}");
        }
    }

    static double GetCellNumber(List<XElement> cells, int index)
    {
        if (index >= cells.Count) return 0;
        var data = cells[index].Element("{urn:schemas-microsoft-com:office:spreadsheet}Data");
        if (data != null && double.TryParse(data.Value, out double val))
            return val;
        return 0;
    }
}