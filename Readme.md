<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128578388/21.1.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E2118)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/HowToBindToMDB/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/HowToBindToMDB/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/HowToBindToMDB/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/HowToBindToMDB/MainWindow.xaml.vb))
<!-- default file list end -->
# How to: Bind a PivotGrid to an MS Access Database Programmatically


<p>The following example demonstrates how to programmatically bind the <strong>PivotGridControl</strong> to a "SalesPerson" view in the nwind.mdb database, which is shipped with the installation. The control will be used to analyze sales per country, customers, product categories and years.</p>


<h3>Description</h3>

<p>First, data is obtained from an MDB database via the <strong>OleDbConnection</strong>, <strong>OleDbDataAdapter</strong> and <strong>DataSet</strong> components. Secondly, the <strong>PivotGridControl</strong> is bound to a table in the dataset via the <strong>PivotGridControl.DataSource</strong> property. Lastly, pivot grid fields are created that will represent datasource fields. They are positioned within appropriate areas to analyze the data in the way you want.</p><p>Note that if you want to see an example of how to add pivot grid fields in XAML, please refer to the <a data-ticket="E2121">How to: Bind a PivotGrid to an MS Access Database</a> tutorial.</p>

<br/>


