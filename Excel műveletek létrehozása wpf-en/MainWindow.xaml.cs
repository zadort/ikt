using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using IronXL;
using IronXL.Styles;

namespace WpfApp2
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			//Licenszkulcs az IronXL Package használatához
			License.LicenseKey = "IRONSUITE.ZADORT.KKSZKI.HU.26764-DEC9CFA078-BDGVTOX-4S5W2SYZ734V-62QGSBQ5THZG-6BDGF57AKERJ-TSTE7QBH7XX7-RWNJ33MXTHMJ-H32SW3PB4RO6-22Q4SD-TWTCQIBBCJ2NUA-DEPLOYMENT.TRIAL-NM5PSK.TRIAL.EXPIRES.13.OCT.2024";
		}

		private void Letrehoz_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
			var workSheet = workBook.CreateWorkSheet("example_sheet");
			workSheet["A1"].Value = "A";
			workSheet["B1"].Value = "programot";
			workSheet["C1"].Value = "Zádor";
			workSheet["D1"].Value = "Tamás";
			workSheet["E1"].Value = "készítette!";
			workSheet["A2"].Value = "20";
			workSheet["A3"].Value = "40";
			workSheet["A4"].Value = "60";
			workSheet["A5"].Value = "80";
			workSheet["A6"].Value = "100";

			//Mentés
			workBook.SaveAs("feladat.xlsx");
			MessageBox.Show("Az Excel dokumentum sikeresen létrehozva!");
		}

		private void Betolt_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			//Szűrő kódsor ami kiválasztja az üres cellákat
			var cellak = workSheet.Where(cell => !string.IsNullOrEmpty(cell.Text));

			//A1-es cellától a B2-ig végigfut majd hozzáad egy sor szöveget
			//ami tartalmazza a cella nevét és értékét
			foreach (var cell in cellak)
			{
				ExcelLista.Items.Add($"A(z) {cell.AddressString} cella értéke: {cell.Text}");
			}
		}

		private void BetuSzin_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.Font.Color = "#FFFFFF";

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void BetuSzinHatter_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.BackgroundColor = "#000000";

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void BetuMeret_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.Font.Height = 9;

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void BetuStilus_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.Font.Name = "Times New Roman";

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void Felkover_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.Font.Bold = true;

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void Dolt_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1:E1"].Style.Font.Italic = true;

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}
		private void CellaTorol_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1"].ClearContents();

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void CellaRendez_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			var column = workSheet.GetColumn(0);
			column.SortAscending();

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

		private void CellaMasol_Click(object sender, RoutedEventArgs e)
		{
			WorkBook workBook = WorkBook.Load("feladat.xlsx");
			WorkSheet workSheet = workBook.WorkSheets.First();

			workSheet["A1"].Copy(workBook.WorkSheets.First(), "A2");

			MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			workBook.SaveAs("feladat.xlsx");
		}

        private void Osszead_Click(object sender, RoutedEventArgs e)
        {
			try {
				WorkBook workBook = WorkBook.Load("feladat.xlsx");
				WorkSheet workSheet = workBook.WorkSheets.First();

				var range = workSheet["A2:A6"];
				decimal sum = range.Sum();

				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
				workBook.SaveAs("feladat.xlsx");
			}
			catch(Exception ex) { 
				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			}
		}

        private void Kivon_Click(object sender, RoutedEventArgs e)
        {
			try {
				WorkBook workBook = WorkBook.Load("feladat.xlsx");
				WorkSheet workSheet = workBook.WorkSheets.First();

				var range = workSheet["A2:A6"];

				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
				workBook.SaveAs("feladat.xlsx");
			}
			catch(Exception ex) { 
				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			}
		}

        private void Szoroz_Click(object sender, RoutedEventArgs e)
        {
			try
			{
				WorkBook workBook = WorkBook.Load("feladat.xlsx");
				WorkSheet workSheet = workBook.WorkSheets.First();

				var range = workSheet["A2:A6"];

				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
				workBook.SaveAs("feladat.xlsx");
			}
			catch (Exception ex) {
				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			}
		}

        private void Oszt_Click(object sender, RoutedEventArgs e)
        {
			try
			{
				WorkBook workBook = WorkBook.Load("feladat.xlsx");
				WorkSheet workSheet = workBook.WorkSheets.First();

				var range = workSheet["A2:A6"];

				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
				workBook.SaveAs("feladat.xlsx");
			}
			catch(Exception ex) {
				MessageBox.Show("Az Excel dokumentum sikeresen módosítva!");
			}
		}
	}
}