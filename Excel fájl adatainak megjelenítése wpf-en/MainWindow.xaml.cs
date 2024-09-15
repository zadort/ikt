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
using System.Linq;

namespace WpfApp1
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
            IronXL.License.LicenseKey = "IRONSUITE.ZADORT.KKSZKI.HU.26764-DEC9CFA078-BDGVTOX-4S5W2SYZ734V-62QGSBQ5THZG-6BDGF57AKERJ-TSTE7QBH7XX7-RWNJ33MXTHMJ-H32SW3PB4RO6-22Q4SD-TWTCQIBBCJ2NUA-DEPLOYMENT.TRIAL-NM5PSK.TRIAL.EXPIRES.13.OCT.2024";

            //Excel fájl betöltése és első munkalap kiválasztása a munkafüzetből
            WorkBook workBook = WorkBook.Load("mintainput.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();

            //Szűrő kódsor ami kiválasztja az üres cellákat
            var cellak = workSheet.Where(cell => !string.IsNullOrEmpty(cell.Text));

            //A1-es cellától a B2-ig végigfut majd hozzáad egy sor szöveget
            //ami tartalmazza a cella nevét és értékét
            foreach (var cell in cellak)
            {   
                ExcelLista.Items.Add($"A(z) {cell.AddressString} cella értéke: {cell.Text}");
            }

            //LINQ kifejezés ami kiszűri a számként értelmezhető cellákat
            //Majd ezek összegzése
            decimal sum = cellak.Where(c => decimal.TryParse(c.Text, out _)).Sum(c => c.DecimalValue);
            decimal max = cellak.Where(c => decimal.TryParse(c.Text, out _)).Max(c => c.DecimalValue);

            //Értékek szöveggé alakítása
            SumTextBlock.Text = sum.ToString();
            MaxTextBlock.Text = max.ToString();
        }
    }
}