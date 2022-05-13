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
            pivotGridControl1.DataProcessingEngine = DataProcessingEngine.Optimized;

            pivotGridControl1.BeginUpdate();

            // Create a row pivot grid field bound to the Country data source column.
            PivotGridField fieldCountry = new PivotGridField();
            fieldCountry.Caption = "Country";
            fieldCountry.Area = FieldArea.RowArea;

            DataSourceColumnBinding countryBinding = new DataSourceColumnBinding("Country");
            fieldCountry.DataBinding = countryBinding;

            // Create a row pivot grid field bound to the Sales Person data source column.
            PivotGridField fieldCustomer = new PivotGridField();
            fieldCustomer.Caption = "Customer";
            fieldCustomer.Area = FieldArea.RowArea;

            DataSourceColumnBinding customerBinding = new DataSourceColumnBinding("Sales Person");
            fieldCustomer.DataBinding = customerBinding;

            // Create a column pivot grid field bound to the OrderDate data source column.
            PivotGridField fieldYear = new PivotGridField();
            fieldYear.Caption = "Year";
            fieldYear.Area = FieldArea.ColumnArea;

            DataSourceColumnBinding fieldOrderDate1Binding = new DataSourceColumnBinding("OrderDate");
            fieldOrderDate1Binding.GroupInterval = FieldGroupInterval.DateYear;
            fieldYear.DataBinding = fieldOrderDate1Binding;

            // Create a column pivot grid field bound to the CategoryName data source column.
            PivotGridField fieldCategoryName = new PivotGridField();
            fieldCategoryName.Caption = "Product Category";
            fieldCategoryName.Area = FieldArea.ColumnArea;

            DataSourceColumnBinding categoryNameBinding = new DataSourceColumnBinding("CategoryName");
            fieldCategoryName.DataBinding = categoryNameBinding;

            // Create a filter pivot grid field bound to the ProductName data source column.
            PivotGridField fieldProductName = new PivotGridField();
            fieldProductName.Caption = "Product Name";
            fieldProductName.Area = FieldArea.FilterArea;

            DataSourceColumnBinding productNameBinding = new DataSourceColumnBinding("ProductName");
            fieldProductName.DataBinding = productNameBinding;

            // Create a data pivot grid field bound to the 'Extended Price' data source column.
            PivotGridField fieldExtendedPrice = new PivotGridField();
            fieldExtendedPrice.Area = FieldArea.DataArea;

            DataSourceColumnBinding extendedPriceBinding = new DataSourceColumnBinding("Extended Price");
            fieldExtendedPrice.DataBinding = extendedPriceBinding;

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
