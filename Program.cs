using IronXL;
using System;
using System.Linq;

License.LicenseKey = "IRONSUITE.ZADORT.KKSZKI.HU.26764-DEC9CFA078-BDGVTOX-4S5W2SYZ734V-62QGSBQ5THZG-6BDGF57AKERJ-TSTE7QBH7XX7-RWNJ33MXTHMJ-H32SW3PB4RO6-22Q4SD-TWTCQIBBCJ2NUA-DEPLOYMENT.TRIAL-NM5PSK.TRIAL.EXPIRES.13.OCT.2024";
WorkBook workBook = WorkBook.Load("mintainput.xlsx");
WorkSheet workSheet = workBook.WorkSheets[0];
WorkSheet firstSheet = workBook.DefaultWorkSheet;

int cellValue = workSheet["A2"].IntValue;

foreach (var cell in workSheet["A1:B2"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}
