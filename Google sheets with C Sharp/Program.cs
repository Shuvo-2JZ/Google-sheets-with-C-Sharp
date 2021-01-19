using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;

namespace Google_sheets_with_C_Sharp
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        static readonly string ApplicationName = "Legislators";

        static readonly string SpreadsheetId = "1Donbqk8p0LEmPbHY_nRVxTMdiK_x1fy10sT-4x-K9cc";

        static readonly string sheet = "Mar-2016";

        static SheetsService service;

        static void Main(string[] args)
        {
            GoogleCredential credential;

            using(var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            ReadEntries();


            Console.ReadLine();
        }

        static void ReadEntries()
        {
            var range = $"{sheet}!A3:G46";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;

            while(true)
            {
                Console.WriteLine("1. Read data from google sheet and store it in the database.");
                Console.WriteLine("2. Show the data from database.");
                Console.WriteLine("3. Remove all data from the database.");
                Console.WriteLine("4. Exit.");

                Console.WriteLine();

                Console.WriteLine("Enter Choice: ");

                int choice = Convert.ToInt32(Console.ReadLine());

                Console.WriteLine();

                if(choice == 1)
                {
                    SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + "C:\\Users\\shuvo\\source\\repos\\Google sheets with C Sharp\\Google sheets with C Sharp\\Test.mdf" + ";Integrated Security=True");

                    cn.Open();

                    int numberOfRows = 0;

                    if (values != null && values.Count > 0)
                    {
                        foreach (var row in values)
                        {
                            SqlCommand mycommand = new SqlCommand("INSERT INTO Mar_2016 VALUES(@Date, @Status, @Tahsinur_Refat_Emon, @Niger, @Nishu, @Asifuzzaman, @Sanjoy)", cn);

                            mycommand.Parameters.AddWithValue("@Date", row[0]);
                            mycommand.Parameters.AddWithValue("@Status", row[1]);
                            mycommand.Parameters.AddWithValue("@Tahsinur_Refat_Emon", row[2]);
                            mycommand.Parameters.AddWithValue("@Niger", row[3]);
                            mycommand.Parameters.AddWithValue("@Nishu", row[4]);
                            mycommand.Parameters.AddWithValue("@Asifuzzaman", row[5]);

                            if(row.Count < 7)
                            {
                                mycommand.Parameters.AddWithValue("@Sanjoy", "NULL");
                            }
                            else
                            {
                                mycommand.Parameters.AddWithValue("@Sanjoy", row[6]);
                            }

                            numberOfRows = mycommand.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        Console.WriteLine("No data found.");
                    }

                    if (numberOfRows > 0)
                    {
                        Console.WriteLine("Insertion successful.");
                    }

                    cn.Close();

                    Console.WriteLine();
                    Console.WriteLine();
                }

                if(choice == 2)
                {
                    SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + "C:\\Users\\shuvo\\source\\repos\\Google sheets with C Sharp\\Google sheets with C Sharp\\Test.mdf" + ";Integrated Security=True");

                    cn.Open();


                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT * FROM Mar_2016", cn);

                    DataTable dataTable = new DataTable();

                    sqlDataAdapter.Fill(dataTable);

                    Console.WriteLine();

                    Console.WriteLine("\t\t\t\tToggle Fullscreen For Better Experience........");

                    Console.WriteLine();
                    Console.WriteLine();
                    Console.WriteLine();

                    Console.WriteLine("\t\t\t\tNBR DC/DR Attendance Register, March 2016");

                    Console.WriteLine();

                    Console.WriteLine("Status" + "\t\t" + "Tahsinur Refat Emon" + "\t" + "Niger" + "\t\t" + "Nishu" + "\t\t" + "Asifuzzaman" + "\t\t" + "Sanjoy" + "\t\t" + "Date");

                    Console.WriteLine();

                    foreach (DataRow row in dataTable.Rows)
                    {
                        Console.Write(row["Status"].ToString() + "\t\t");
                        Console.Write(row["Tahsinur_Refat_Emon"].ToString() + "\t\t\t");
                        Console.Write(row["Niger"].ToString() + "\t\t");
                        Console.Write(row["Nishu"].ToString() + "\t\t");
                        Console.Write(row["Asifuzzaman"].ToString() + "\t\t\t");
                        Console.Write(row["Sanjoy"].ToString() + "\t\t");
                        Console.Write(row["Date"].ToString());

                        Console.WriteLine();
                    }                 

                    Console.WriteLine();
                    Console.WriteLine();

                    cn.Close();
                }

                if(choice == 3)
                {
                    SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + "C:\\Users\\shuvo\\source\\repos\\Google sheets with C Sharp\\Google sheets with C Sharp\\Test.mdf" + ";Integrated Security=True");

                    cn.Open();

                    SqlCommand mycommand = new SqlCommand("DELETE FROM Mar_2016", cn);
                    mycommand.ExecuteNonQuery();

                    Console.WriteLine("Deletion successful.");

                    Console.WriteLine();
                    Console.WriteLine();

                    cn.Close();
                }

                if(choice == 4)
                {
                    break;
                }

                if(choice > 4 || choice < 1)
                {
                    Console.WriteLine();

                    Console.WriteLine("Wrong Input !!!");

                    Console.WriteLine();
                    Console.WriteLine();
                }
            }

            Console.WriteLine("Programe ended.");
        }
    }
}