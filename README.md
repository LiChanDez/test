# test

<Window x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="MainWindow" Height="450" Width="800">
    <Grid Background="Bisque">

        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="150"/>
            <RowDefinition Height="60"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0"
                Orientation="Horizontal"
                Margin="10">

            <TextBox x:Name="SearchBox"
                 Width="300"
                 Height="30"
                 Margin="5"
                 TextChanged="SearchBox_TextChanged"/>
        </StackPanel>

        <DataGrid x:Name="OrdersGrid"
              Grid.Row="1"
              Margin="10"
              AutoGenerateColumns="True"
              Background="White"
              SelectionChanged="OrdersGrid_SelectionChanged"/>

        <StackPanel Grid.Row="2" Margin="10">

            <StackPanel Orientation="Horizontal">
                <TextBox x:Name="CarIdBox" Width="120" Margin="5"/>
                <TextBox x:Name="ServiceIdBox" Width="120" Margin="5"/>
                <TextBox x:Name="StatusBox" Width="120" Margin="5"/>
            </StackPanel>

            <StackPanel Orientation="Horizontal">
                <TextBox x:Name="PriceBox" Width="120" Margin="5"/>
                <DatePicker x:Name="DateBox" Width="150" Margin="5"/>
            </StackPanel>

        </StackPanel>

        <StackPanel Grid.Row="3"
                Orientation="Horizontal"
                HorizontalAlignment="Left"
                Margin="10">

            <Button Content="Добавить"
        Width="120"
        Margin="5"
        Background="#4CAF50"
        Foreground="White"
        Click="Add_Click"/>

            <Button Content="Редактировать"
        Width="120"
        Margin="5"
        Background="#2196F3"
        Foreground="White"
        Click="Edit_Click"/>

            <Button Content="Удалить"
        Width="120"
        Margin="5"
        Background="#F44336"
        Foreground="White"
        Click="Delete_Click"/>

            <Button Content="Экспорт в Excel"
                Width="150"
                Margin="5"
                Click="ExportCsv_Click"/>

        </StackPanel>

    </Grid>
</Window>







----------------------------------------------------











using Npgsql;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{

    public partial class MainWindow : Window
    {
        string connection =
            "Host=localhost;Port=5432;Username=postgres;Password=123;Database=postgres";
        DataTable ordersTable = new DataTable();
        public MainWindow()
        {
            InitializeComponent();
            LoadOrders();
        }

        private void LoadOrders()
        {
            using var con = new NpgsqlConnection(connection);
            con.Open();

            var query = @"
    SELECT order_id, car_id, service_id, status, total_price, date::timestamp as date
    FROM ""Orders""";

            var adapter = new NpgsqlDataAdapter(query, con);

            ordersTable.Clear();
            adapter.Fill(ordersTable);

            OrdersGrid.ItemsSource = ordersTable.DefaultView;
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            using var con = new NpgsqlConnection(connection);
            con.Open();

            var query = @"
    INSERT INTO ""Orders""
    (car_id, service_id, status, total_price, date)
    VALUES (@car, @service, @status, @price, @date)";

            using var cmd = new NpgsqlCommand(query, con);

            cmd.Parameters.AddWithValue("car", int.Parse(CarIdBox.Text));
            cmd.Parameters.AddWithValue("service", int.Parse(ServiceIdBox.Text));
            cmd.Parameters.AddWithValue("status", StatusBox.Text);
            cmd.Parameters.AddWithValue("price", decimal.Parse(PriceBox.Text));
            cmd.Parameters.AddWithValue("date", DateBox.SelectedDate.Value);

            cmd.ExecuteNonQuery();
            LoadOrders();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (OrdersGrid.SelectedItem == null) return;

            var row = (DataRowView)OrdersGrid.SelectedItem;

            using var con = new NpgsqlConnection(connection);
            con.Open();

            var query = @"
    UPDATE ""Orders""
    SET status = @status,
        total_price = @price
    WHERE order_id = @id";

            using var cmd = new NpgsqlCommand(query, con);

            cmd.Parameters.AddWithValue("id", Convert.ToInt32(row["order_id"]));
            cmd.Parameters.AddWithValue("status", StatusBox.Text);
            cmd.Parameters.AddWithValue("price", decimal.Parse(PriceBox.Text));

            cmd.ExecuteNonQuery();
            LoadOrders();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (OrdersGrid.SelectedItem == null) return;

            var row = (DataRowView)OrdersGrid.SelectedItem;

            using var con = new NpgsqlConnection(connection);
            con.Open();

            var query = @"DELETE FROM ""Orders"" WHERE order_id = @id";

            using var cmd = new NpgsqlCommand(query, con);

            cmd.Parameters.AddWithValue("id", row["order_id"]);

            cmd.ExecuteNonQuery();

            LoadOrders();
        }
        private void OrdersGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OrdersGrid.SelectedItem is not DataRowView row)
                return;

            StatusBox.Text = row["status"].ToString();
            PriceBox.Text = row["total_price"].ToString();
            DateBox.SelectedDate = (DateTime)row["date"];
        }
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ordersTable == null || ordersTable.Columns.Count == 0)
                return;

            string filter = SearchBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(filter))
            {
                ordersTable.DefaultView.RowFilter = "";
                return;
            }

            ordersTable.DefaultView.RowFilter =
                $"Convert(order_id, 'System.String') LIKE '%{filter}%' OR " +
                $"status LIKE '%{filter}%' OR " +
                $"Convert(total_price, 'System.String') LIKE '%{filter}%' OR " +
                $"Convert(car_id, 'System.String') LIKE '%{filter}%' OR " +
                $"Convert(service_id, 'System.String') LIKE '%{filter}%'";
        }
        private void ExportCsv_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                              + @"\Orders.csv";

                using (var writer = new StreamWriter(path))
                {
                    for (int i = 0; i < ordersTable.Columns.Count; i++)
                    {
                        writer.Write(ordersTable.Columns[i]);
                        if (i < ordersTable.Columns.Count - 1)
                            writer.Write(";");
                    }
                    writer.WriteLine();

                    foreach (DataRow row in ordersTable.Rows)
                    {
                        for (int i = 0; i < ordersTable.Columns.Count; i++)
                        {
                            writer.Write(row[i]);
                            if (i < ordersTable.Columns.Count - 1)
                                writer.Write(";");
                        }
                        writer.WriteLine();
                    }
                }

                MessageBox.Show("CSV создан: " + path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
