using System;
using IronXL;

namespace ExcelExportApp
{
    class Program
    {
        static void Main(string[] args)
        {
            License.LicenseKey = "IRONSUITE.ZADORT.KKSZKI.HU.26764-DEC9CFA078-BDGVTOX-4S5W2SYZ734V-62QGSBQ5THZG-6BDGF57AKERJ-TSTE7QBH7XX7-RWNJ33MXTHMJ-H32SW3PB4RO6-22Q4SD-TWTCQIBBCJ2NUA-DEPLOYMENT.TRIAL-NM5PSK.TRIAL.EXPIRES.13.OCT.2024";

            var workBook = WorkBook.Create();
            var workSheet = workBook.DefaultWorkSheet;

            Console.WriteLine("Adja meg a sorok számát:");
            if (!int.TryParse(Console.ReadLine(), out int numberOfRows) || numberOfRows <= 0)
            {
                Console.WriteLine("Érvénytelen sor szám.");
                return;
            }

            workSheet["A1"].Value = "Név";
            workSheet["B1"].Value = "Kor";

            for (int i = 0; i < numberOfRows; i++)
            {
                Console.WriteLine($"Adja meg az {i + 1}. sor nevét:");
                var name = Console.ReadLine();

                Console.WriteLine($"Adja meg az {i + 1}. sor korát:");
                if (!int.TryParse(Console.ReadLine(), out int age) || age < 0)
                {
                    Console.WriteLine("Érvénytelen kor.");
                    return;
                }

                workSheet[$"A{i + 2}"].Value = name;
                workSheet[$"B{i + 2}"].Value = age;
            }


            var filePath = "exportált_adatok.xlsx";
            workBook.SaveAs(filePath);

            Console.WriteLine($"Az adatok sikeresen mentve: {filePath}");
        }
    }
}
