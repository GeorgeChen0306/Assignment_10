using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace Assignment_10
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /**
          * Event handler that will ask the user if they want to close the application when exit is selected
          */
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            // Message box that prompt user to select yes or no to close the application
            MessageBoxResult answer = MessageBox.Show("Do you really want to exit?", "Exit Application",
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (answer == MessageBoxResult.Yes) Application.Current.Shutdown(); // if yes, close the application

        }

        // Global array to store data information
        readonly string[] lowest_name = new string[6];
        readonly string[] highest_name = new string[6];
        readonly string[] all = new string[6];

        readonly int[] lowest_unit_value = new int[3];
        readonly int[] highest_unit_value = new int[3];

        readonly float[] lowest_revenue = new float[3];
        readonly float[] highest_revenue = new float[3];

        // Global variable to store the excel file
        string inputfile = "";

        /**
         * Event handler that open a window dialog when the user select the open submenu
         * It will open to the user Documents directory and filter only excel files
         * When the file is selected, it will process all 18 possible data combinations and store
         * them in the global array
         */
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog myOpenDialog = new OpenFileDialog();

            //Open the dialog window that filters only excel files in the Documents directory.
            myOpenDialog.Title = "Select an Excel File to Process";
            myOpenDialog.Filter = "Excel files(*.xlsx)|*.xlsx";
            myOpenDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (myOpenDialog.ShowDialog() == true)
            {
                inputfile = myOpenDialog.FileName; // store the file path to the global variable
                Output.Content = "Select Data to Continue"; // Prompt user to select different data combinations
            }
            else
            {
                return; // Exit the function if no file is provided
            }

            // variables for using Excel
            Excel.Application myApp;
            Excel.Workbook myBook;
            Excel.Worksheet mySheet;
            Excel.Range myRange;

            // connect to the Excel file data
            myApp = new Excel.Application();
            myApp.Visible = false;
            myBook = myApp.Workbooks.Open(inputfile);
            mySheet = myBook.Sheets[1];
            myRange = mySheet.UsedRange;

            // Section 1
            // Highest/Lowest -> Items -> Units sold

            // create a dictionary
            Dictionary<string, int> item_units = new Dictionary<string, int>();

            string next_item; // item on the next row, such as pencil or desk
            int next_units; // how many items were sold

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_item = (string)(mySheet.Cells[r, 4] as Excel.Range).Value;
                next_units = (int)(mySheet.Cells[r, 5] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!item_units.ContainsKey(next_item))
                {
                    item_units.Add(next_item, next_units);
                }
                // if item already in list, then add the new values
                else
                {
                    item_units[next_item] += next_units; // add current item value to the existing item value     
                }
            }

            //variables to store the lowest and the highest number of units sold
            int min_item_unit = int.MaxValue; //initialize with the largest int 
            int max_item_unit = int.MinValue; //initialize with the smallest int

            // variables to store the item name with the least and the most units
            string min_item_unit_name = "";
            string max_item_unit_name = "";

            // variable to store all units sold for each item
            string all_item_units_sold = "Units sold for all Items:\n-------------------------\n";

            //Loop through all dictionary key and check for the smallest value and also the largest value
            foreach (KeyValuePair<string, int> item in item_units)
            {
                if (item.Value < min_item_unit)
                {
                    min_item_unit_name = item.Key; // assign the item name to the current min units sold
                    min_item_unit = item.Value; // assign the value to the current min units sold
                }
                if (item.Value > max_item_unit)
                {
                    max_item_unit_name = item.Key; //assign the item name to the current max units sold
                    max_item_unit = item.Value; // assign the value to the current max units sold
                }
                // concatenate the total number of units sold for each item
                all_item_units_sold += $"{item.Key} - {item.Value} units\n";
            }

            // Store the lowest/highest/all names and unit values.
            lowest_name[0] = min_item_unit_name;
            lowest_unit_value[0] = min_item_unit;
            highest_name[0] = max_item_unit_name;
            highest_unit_value[0] = max_item_unit;
            all[0] = all_item_units_sold;


            // Section 2
            // Highest/Lowest -> Item -> Revenue

            // create a dictionary
            Dictionary<string, float> item_revenue = new Dictionary<string, float>();
            float next_revenue; // variable to store the revenue in each row

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_item = (string)(mySheet.Cells[r, 4] as Excel.Range).Value;
                next_revenue = (float)(mySheet.Cells[r, 7] as Excel.Range).Value;


                // is that a new item for the dictionary?
                if (!item_revenue.ContainsKey(next_item))
                {
                    item_revenue.Add(next_item, next_revenue);
                }
                // if item already in list, then add the new values
                else
                {
                    item_revenue[next_item] += next_revenue; // add current item value to the existing item value     
                }
            }

            // variables to store the lowest and the highest revenue
            float min_item_revenue = int.MaxValue;
            float max_item_revenue = int.MinValue;

            // variables to store the lowest and the highest item revenue name
            string min_item_revenue_name = "";
            string max_item_revenue_name = "";

            // variable to store all items total revenue
            string all_item_revenue = "Revenue for all Items:\n------------------------\n";

            // loop through each key in the dictionary and check for the smallest and the largest value
            foreach (KeyValuePair<string, float> item in item_revenue)
            {
                if (item.Value < min_item_revenue)
                {
                    min_item_revenue_name = item.Key; // assign the current min revenue name
                    min_item_revenue = item.Value; // assign the current value of the min revenue
                }
                if (item.Value > max_item_revenue)
                {
                    max_item_revenue_name = item.Key; // assign the current max revenue name
                    max_item_revenue = item.Value; // assign the current max value of revenue
                }
                // concatenate each item name and the revenue
                all_item_revenue += $"{item.Key} - ${item.Value}\n";
            }

            // Store the lowest/highest/all names and revenue values.
            lowest_name[1] = min_item_revenue_name;
            highest_name[1] = max_item_revenue_name;
            lowest_revenue[0] = min_item_revenue;
            highest_revenue[0] = max_item_revenue;
            all[1] = all_item_revenue;


            // Section 3
            // Highest/Lowest -> Sales Rep -> Units sold

            //create a dictionary
            Dictionary<string, int> sales_rep_units_sold = new Dictionary<string, int>();
            string next_rep; // variable to store the cell data on the Sales Rep column

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_rep = (string)(mySheet.Cells[r, 3] as Excel.Range).Value;
                next_units = (int)(mySheet.Cells[r, 5] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!sales_rep_units_sold.ContainsKey(next_rep))
                {
                    sales_rep_units_sold.Add(next_rep, next_units);
                }
                // if item already in list, then add the new values
                else
                {
                    sales_rep_units_sold[next_rep] += next_units; // add current item value to the existing item value     
                }
            }
            // variables to store the smallest and largest value
            int min_rep_unit = int.MaxValue; //initialize with largest int 
            int max_rep_unit = int.MinValue; //initialize with smallest int

            // variables to store the sales rep name
            string min_rep_unit_name = "";
            string max_rep_unit_name = "";

            // variable to store the number of units sold by each Sales Rep 
            string all_rep_units_sold = "Units sold by each Sales Rep:\n------------------------------\n";

            // loop through each key in the dictionary and assign the smallest and the largest value
            foreach (KeyValuePair<string, int> item in sales_rep_units_sold)
            {
                if (item.Value < min_rep_unit)
                {
                    min_rep_unit_name = item.Key; // assign the current Sales rep name that sold the least
                    min_rep_unit = item.Value; // assign the current smallest units sold 
                }
                if (item.Value > max_rep_unit)
                {
                    max_rep_unit_name = item.Key; // assign the current Sales rep name that sold the most
                    max_rep_unit = item.Value; // assign the current largest units sold
                }
                // concatenate the number of units sold by each Sales Rep
                all_rep_units_sold += $"{item.Key} - {item.Value} units\n";
            }

            // Store the lowest/highest/all names and unit values.
            lowest_name[2] = min_rep_unit_name;
            highest_name[2] = max_rep_unit_name;
            lowest_unit_value[1] = min_rep_unit;
            highest_unit_value[1] = max_rep_unit;
            all[2] = all_rep_units_sold;


            // Section 4
            // Highest/Lowest -> Sales Rep -> Revenue

            // create a dictionary
            Dictionary<string, float> rep_revenue = new Dictionary<string, float>();

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_rep = (string)(mySheet.Cells[r, 3] as Excel.Range).Value;
                next_revenue = (float)(mySheet.Cells[r, 7] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!rep_revenue.ContainsKey(next_rep))
                {
                    rep_revenue.Add(next_rep, next_revenue);
                }
                // if item already in list, then add the new values
                else
                {
                    rep_revenue[next_rep] += next_revenue; // add current item value to the existing item value     
                }
            }

            // variables to store the lowest and the highest revenue
            float min_rep_revenue = int.MaxValue;
            float max_rep_revenue = int.MinValue;

            // variables to store the lowest and the highest revenue rep name
            string min_rep_revenue_name = "";
            string max_rep_revenue_name = "";

            // variable to store all Sales Rep revenue
            string all_rep_revenue = "Revenue by each Sales Rep:\n---------------------------\n";

            // loop through each key in the dictionary and assign the smallest and the largest value
            foreach (KeyValuePair<string, float> item in rep_revenue)
            {
                if (item.Value < min_rep_revenue)
                {
                    min_rep_revenue_name = item.Key; // assign the current name of the smallest revenue
                    min_rep_revenue = item.Value; // assign the current value of the smallest revenue
                }
                if (item.Value > max_rep_revenue)
                {
                    max_rep_revenue_name = item.Key; // assign the current name of the largest revenue
                    max_rep_revenue = item.Value; // assign the current value of the largest revenue
                }
                // concatenate each Sales Rep name and their revenue 
                all_rep_revenue += $"{item.Key} - ${item.Value}\n";
            }

            // Store the lowest/highest/all names and revenue values.
            lowest_name[3] = min_rep_revenue_name;
            highest_name[3] = max_rep_revenue_name;
            lowest_revenue[1] = min_rep_revenue;
            highest_revenue[1] = max_rep_revenue;
            all[3] = all_rep_revenue;


            // Section 5
            // Highest/Lowest -> Region -> Units sold

            // create a dictionary
            Dictionary<string, int> region_units_sold = new Dictionary<string, int>();
            string next_region;

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_region = (string)(mySheet.Cells[r, 2] as Excel.Range).Value;
                next_units = (int)(mySheet.Cells[r, 5] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!region_units_sold.ContainsKey(next_region))
                {
                    region_units_sold.Add(next_region, next_units);
                }
                // if item already in list, then add the new values
                else
                {
                    region_units_sold[next_region] += next_units; // add current item value to the existing item value     
                }
            }

            // variables to store the smallest and largest units sold value
            int min_region_unit = int.MaxValue; //initialize with largest int 
            int max_region_unit = int.MinValue; //initialize with smallest int

            // variables to store the smallest and largest units sold by their name
            string min_region_unit_name = "";
            string max_region_unit_name = "";

            // variable to store all units sold by the region
            string all_region_units_sold = "Units sold by each Region:\n-----------------------------\n";

            // loop through each key in the dictionary and assign the smallest and largest value
            foreach (KeyValuePair<string, int> item in region_units_sold)
            {
                if (item.Value < min_region_unit)
                {
                    min_region_unit_name = item.Key; // assign the current smallest region units sold name
                    min_region_unit = item.Value; // assign the current smallest region units sold value
                }
                if (item.Value > max_item_unit)
                {
                    max_region_unit_name = item.Key; // assign the current largest region units sold name
                    max_region_unit = item.Value; // assign the current largest region units sold value
                }
                // concatenate each units sold by the region
                all_region_units_sold += $"{item.Key} - {item.Value} units\n";
            }

            // Store the lowest/highest/all names and unit values.
            lowest_name[4] = min_region_unit_name;
            highest_name[4] = max_region_unit_name;
            lowest_unit_value[2] = min_region_unit;
            highest_unit_value[2] = max_region_unit;
            all[4] = all_region_units_sold;


            // Section 6
            // Highest/Lowest -> Region -> Revenue

            // create a dictionary
            Dictionary<string, float> region_revenue = new Dictionary<string, float>();

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_region = (string)(mySheet.Cells[r, 2] as Excel.Range).Value;
                next_revenue = (float)(mySheet.Cells[r, 7] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!region_revenue.ContainsKey(next_region))
                {
                    region_revenue.Add(next_region, next_revenue);
                }
                // if item already in list, then add the new values
                else
                {
                    region_revenue[next_region] += next_revenue; // add current item value to the existing item value     
                }
            }

            // variables to store the lowest and the highest revenue
            float min_region_revenue = int.MaxValue;
            float max_region_revenue = int.MinValue;

            // variables to store the lowest and the highest revenue rep name
            string min_region_revenue_name = "";
            string max_region_revenue_name = "";

            // variable to store all region revenue
            string all_region_revenue = "Revenue by each Region:\n--------------------------\n";

            // loop through each key in the dictionary and assign the smallest and largest revenue
            foreach (KeyValuePair<string, float> item in region_revenue)
            {
                if (item.Value < min_region_revenue)
                {
                    min_region_revenue_name = item.Key; // assign the curr smallest region revenue name
                    min_region_revenue = item.Value; // assign the curr smallest region revenue value
                }
                if (item.Value > max_region_revenue)
                {
                    max_region_revenue_name = item.Key; // assign the curr largest region revenue name
                    max_region_revenue = item.Value; // assign the curr largest region revenue value
                }
                // concatenate each key and their revenue value 
                all_region_revenue += $"{item.Key} - ${item.Value}\n";
            }

            // Store the lowest/highest/all names and revenue values.
            lowest_name[5] = min_region_revenue_name;
            highest_name[5] = max_region_revenue_name;
            lowest_revenue[2] = min_region_revenue;
            highest_revenue[2] = max_region_revenue;
            all[5] = all_region_revenue;

            myBook.Close(); // Close the file
            myApp.Quit(); // Close excel
        }

        /**
         * Event handler that runs when users click the Run Reports menu
         * It will check if the file exists and warn them there is no data to process
         * Output results based on which radio buttons are selected. 
         */
        private void Run_Click(object sender, RoutedEventArgs e)
        {
            // check does file exist
            if (!System.IO.File.Exists(inputfile))
            {
                Output.Content = "No data to process";
                return;
            }

            // Check for which radio buttons are selected and post the result accordingly

            // Highest is selected
            if (Max.IsChecked == true)
            {
                // Item is selected
                if (Item.IsChecked == true)
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        // Highest -> Items -> Units sold
                        Output.Content = $"Most popular item = {highest_name[0]} ({highest_unit_value[0]} units)";
                        return;
                    }
                    //Revenue is selected
                    else
                    {
                        //Highest value -> Item -> Revenue
                        Output.Content = $"Item with the highest revenue = {highest_name[1]} (${highest_revenue[0]})"; 
                    }
                }
                // Sales is selected
                else if (Sales.IsChecked == true)
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        // Highest value -> Sales -> Units sold
                        Output.Content = $"Sales Rep with the most units sold = {highest_name[2]} ({highest_unit_value[1]} units)";
                    }
                    // Revenue is selected
                    else
                    {
                        // Highest value -> Sales -> Revenue
                        Output.Content = $"Sales Rep with the highest revenue = {highest_name[3]} (${highest_revenue[1]})";
                    }
                }
                // Region is selected
                else
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        //Highest value -> Region -> Units sold
                        Output.Content = $"Region with the most units sold = {highest_name[4]} ({highest_unit_value[2]} units)";
                    }
                    // Revenue is selected
                    else
                    {
                        //Highest value -> Region -> Revenue
                        Output.Content = $"Region with the highest revenue = {highest_name[5]} (${highest_revenue[2]})";
                    }
                }
            }

            // Lowest is selected
            if (Min.IsChecked == true)
            {
                // Item is selected
                if (Item.IsChecked == true)
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        Output.Content = $"Least popular item = {lowest_name[0]} ({lowest_unit_value[0]} units)";
                    }
                    // Revenue is selected
                    else
                    {
                        // Lowest value -> Item -> Revenue
                        Output.Content = $"Item with the lowest revenue = {lowest_name[1]} (${lowest_revenue[0]})";
                    }
                }
                // Sales rep is selected
                else if (Sales.IsChecked == true)
                {
                    // Unit solds is selected
                    if (Sold.IsChecked == true)
                    {
                        // Lowest value -> Sales -> Units sold
                        Output.Content = $"Sales Rep with the least units sold = {lowest_name[2]} ({lowest_unit_value[1]} units)";
                    }
                    // Revenue is selected
                    else
                    {
                        // Lowest value -> Sales -> Revenue
                        Output.Content = $"Sales Rep with the lowest revenue = {lowest_name[3]} (${lowest_revenue[1]})";
                    }
                }
                // Region is selected
                else
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        // Lowest value -> Region -> Units sold
                        Output.Content = $"Region with the least units sold = {lowest_name[4]} ({lowest_unit_value[2]} units)";
                    }
                    // Revenue is selected
                    else
                    {
                        // Lowest value -> Region -> Revenue
                        Output.Content = $"Region with the lowest revenue = {lowest_name[5]} (${lowest_revenue[2]})";
                    }
                }
            }

            // All is selected
            if (All.IsChecked == true)
            {
                // Items is selected
                if (Item.IsChecked == true)
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        // All values -> Item -> Units sold
                        Output.Content = all[0];
                    }
                    // Revenue is selected
                    else
                    {
                        // All values -> Item -> Revenue
                        Output.Content = all[1];
                    }
                }
                // Sales Rep is selected
                else if (Sales.IsChecked == true)
                {
                    // Units sold is selected
                    if (Sold.IsChecked == true)
                    {
                        // All values -> Sales rep -> Units sold
                        Output.Content = all[2];
                    }
                    // Revenue is selected
                    else
                    {
                        // All values -> Sales rep -> Revenue
                        Output.Content = all[3];
                    }
                }
                // Region is selected
                else
                {
                    // Unit solds is selected
                    if (Sold.IsChecked == true)
                    {
                        // All values -> Region -> Units sold
                        Output.Content = all[4];
                    }
                    // Revenue is selected
                    else
                    {
                        // All values -> Region -> Revenue
                        Output.Content = all[5];
                    }
                }
            }
        }
    }
}
