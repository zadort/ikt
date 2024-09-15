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
            License.LicenseKey = "IRONSUITE.ZADORT.KKSZKI.HU.26764-DEC9CFA078-BDGVTOX-4S5W2SYZ734V-62QGSBQ5THZG-6BDGF57AKERJ-TSTE7QBH7XX7-RWNJ33MXTHMJ-H32SW3PB4RO6-22Q4SD-TWTCQIBBCJ2NUA-DEPLOYMENT.TRIAL-NM5PSK.TRIAL.EXPIRES.13.OCT.2024";

            //Excel fájl betöltése és első munkalap kiválasztása a munkafüzetből
            WorkBook workBook = WorkBook.Load("mintainput.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();

            //A1-es cellától a B2-ig végigfut majd hozzáad egy sor szöveget
            //ami tartalmazza a cella nevét és értékét
            foreach (var cell in workSheet["A1:B2"])
            {   
                excelLista.Items.Add($"A(z) {cell.AddressString} cella értéke: {cell.Text}");
            }

            //Sum() metódus a cella összértékeinek meghatározásához
            decimal sum = workSheet["A1:B2"].Sum();
            decimal max = workSheet["A1:B2"].Max(c => c.DecimalValue);

            //Értékek szöveggé alakítása
            sumTextBlock.Text = sum.ToString();
            maxTextBlock.Text = max.ToString();
        }
    }
}
