using System.Data;
using System.Data.OleDb;
using System.Windows;
using DevExpress.Xpf.PivotGrid;

namespace HowToBindToMDB {

    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) {
            // Create a connection object.
            OleDbConnection connection =
                new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=NWIND.MDB");
            // Create a data adapter.
            OleDbDataAdapter adapter =
                new OleDbDataAdapter("SELECT * FROM SalesPerson", connection);

            // Create and fill a dataset.
            DataSet sourceDataSet = new DataSet();
            adapter.Fill(sourceDataSet, "SalesPerson");

            // Assign the data source to the PivotGrid control.
            pivotGridControl1.DataSource = sourceDataSet.Tables["SalesPerson"];

            pivotGridControl1.BeginUpdate();

            // Create a row pivot grid field bound to the Country datasource field.
            PivotGridField fieldCountry = new PivotGridField("Country", FieldArea.RowArea);

            // Create a row pivot grid field bound to the Sales Person datasource field.
            PivotGridField fieldCustomer = new PivotGridField("Sales Person", FieldArea.RowArea);
            fieldCustomer.Caption = "Customer";

            // Create a column pivot grid field bound to the OrderDate datasource field.
            PivotGridField fieldYear = new PivotGridField("OrderDate", FieldArea.ColumnArea);
            fieldYear.Caption = "Year";
            // Group field values by years.
            fieldYear.GroupInterval = FieldGroupInterval.DateYear;

            // Create a column pivot grid field bound to the CategoryName datasource field.
            PivotGridField fieldCategoryName = new PivotGridField("CategoryName", FieldArea.ColumnArea);
            fieldCategoryName.Caption = "Product Category";

            // Create a filter pivot grid field bound to the ProductName datasource field.
            PivotGridField fieldProductName = new PivotGridField("ProductName", FieldArea.FilterArea);
            fieldProductName.Caption = "Product Name";

            // Create a data pivot grid field bound to the 'Extended Price' datasource field.
            PivotGridField fieldExtendedPrice = new PivotGridField("Extended Price", FieldArea.DataArea);

            // Specify the formatting setting to format summary values as integer currency amount.
            fieldExtendedPrice.CellFormat = "c0";

            // Add the fields to the control's field collection.         
            pivotGridControl1.Fields.AddRange(fieldCountry, fieldCustomer,
                fieldCategoryName, fieldProductName, fieldYear, fieldExtendedPrice);

            // Arrange the row fields within the Row Header Area.
            fieldCountry.AreaIndex = 0;
            fieldCustomer.AreaIndex = 1;

            // Arrange the column fields within the Column Header Area.
            fieldCategoryName.AreaIndex = 0;
            fieldYear.AreaIndex = 1;

            pivotGridControl1.EndUpdate();
        }
    }
}
