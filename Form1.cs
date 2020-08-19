using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParseEcxellTarifsToSQl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private int daysOnTheMonth = 31;
        private string[] filenames = new string[] { };
        private string date = string.Empty;
        private List<List<string>> listTarifsCk1 = new List<List<string>>();
        private List<List<string>> listTarifsCk2 = new List<List<string>>();
        private List<List<string>> listTarifsCk3 = new List<List<string>>(); 
        private List<List<string>> listTarifsCk4 = new List<List<string>>();
        private List<List<string>> listTarifsCk5 = new List<List<string>>();
        private List<List<string>> listTarifsCk6 = new List<List<string>>();
        private List<List<string>> listOtherTarifs = new List<List<string>>();

        private List<List<float>> listFloatTarifsCk1 = new List<List<float>>();
        private List<List<float>> listFloatTarifsCk2 = new List<List<float>>();
        private List<List<float>> listFloatTarifsCk3 = new List<List<float>>();
        private List<List<float>> listFloatTarifsCk4 = new List<List<float>>();
        private List<List<float>> listFloatTarifsCk5 = new List<List<float>>();
        private List<List<float>> listFloatTarifsCk6 = new List<List<float>>();
        private List<List<float>> listFloatOtherTarifs = new List<List<float>>();

        #region ButtonForms
        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = true;
            openFileDialog1.InitialDirectory = @"D:\RenatWork\TarifsForCalculator\2020";
            openFileDialog1.ShowDialog();
            filenames = openFileDialog1.FileNames;
            if (filenames.Length > 0)
            {
                btnOkFile.Image = Properties.Resources.tick;
                btnStart.Enabled = true;
            }
        }
        private void Calendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            int month = e.Start.Month;
            int year = e.Start.Year;
            daysOnTheMonth = e.Start.AddMonths(1).AddDays(-e.Start.Day).Day;
            date = month.ToString() + year.ToString();
            btnOkDate.Image = Properties.Resources.tick;
        }
        #endregion

        #region FormLoad
        private void Form1_Load(object sender, EventArgs e)
        {
            btnOkFile.Image = Properties.Resources.cancel;
            btnOkDate.Image = Properties.Resources.cancel;
            btnOkWrite.Image = Properties.Resources.cancel;
            LabFinish.Visible = false;
            btnStart.Enabled = false;
            string[] rowsProcess1 = { "Чтение 1 ЦК", "Чтение 2 ЦК", "Чтение 3 ЦК", "Чтение 4 ЦК", "Чтение 5 ЦК", "Чтение 6 ЦК" };
            for (int i = 0; i < rowsProcess1.Length; i++)
            {
                tableProcess1.Rows.Add();
                tableProcess1.Rows[i].Cells["ColProcess"].Value = rowsProcess1[i];
                tableProcess1.Rows[i].Cells["ColExe"].Value = Properties.Resources.cross;
            }
            string[] rowsProcess2 = { "Чтение 3 ЦК", "Чтение 4 ЦК", "Чтение 5 ЦК", "Чтение 6 ЦК" };
            for (int i = 0; i < rowsProcess2.Length; i++)
            {
                tableProcess2.Rows.Add();
                tableProcess2.Rows[i].Cells["ColProcess2"].Value = rowsProcess2[i];
                tableProcess2.Rows[i].Cells["ColExe2"].Value = Properties.Resources.cross;
            }
            string[] rowsProcess3 = { "Чтение 3 ЦК", "Чтение 4 ЦК", "Чтение 5 ЦК", "Чтение 6 ЦК" };
            for (int i = 0; i < rowsProcess3.Length; i++)
            {
                tableProcess3.Rows.Add();
                tableProcess3.Rows[i].Cells["ColProcess3"].Value = rowsProcess3[i];
                tableProcess3.Rows[i].Cells["ColExe3"].Value = Properties.Resources.cross;
            }
        }
        #endregion

        #region Support Funccions For Calc
        private List<List<string>> ArrayBuilder(Excel.Range xlRange, string idPower, string voltLvl, int rowUp, int colUp)
        {
            List<List<string>> tmpListCk3 = new List<List<string>>();
            for (int i = 1; i <= daysOnTheMonth; i++)
            {
                List<string> listRow = new List<string>();
                listRow.Add("1"); // provider
                listRow.Add(idPower); // idpower
                listRow.Add(voltLvl); // voltLevel
                listRow.Add(i.ToString()); // date
                for (int j = 1; j < 25; j++)
                {
                    listRow.Add(xlRange.Cells[i + rowUp, j + colUp].Value2.ToString()); // tarifs
                }
                tmpListCk3.Add(listRow);
            }
            return tmpListCk3;
        }
        private List<List<string>> ArrayBuilder5(Excel.Range xlRange, string idPower, string voltLvl, string planfact, int rowUp, int colUp)
        {
            List<List<string>> tmpListCk3 = new List<List<string>>();
            for (int i = 1; i <= daysOnTheMonth; i++)
            {
                List<string> listRow = new List<string>();
                listRow.Add("1"); // provider
                listRow.Add(idPower); // idpower
                listRow.Add(voltLvl); // voltLevel
                listRow.Add(planfact); // planfact
                listRow.Add(i.ToString()); // date
                for (int j = 1; j < 25; j++)
                {
                    listRow.Add(xlRange.Cells[i + rowUp, j + colUp].Value2.ToString()); // tarifs
                }
                tmpListCk3.Add(listRow);
            }
            return tmpListCk3;
        }
        private void OtherTarifsBuilder(string idPower, string ck, string contract, string power, string nwp1, string nwp2, string nwp3, string nwp4, string sumplan, string abs)
        {
            List<string> listOther = new List<string>();
            listOther.Add("1"); // provider
            listOther.Add(idPower); // idPower
            listOther.Add(ck); // ck
            listOther.Add(contract); // contract
            listOther.Add(power); // power
            listOther.Add(nwp1); // networkPower 1
            listOther.Add(nwp2); // nwp2
            listOther.Add(nwp3); // nwp3
            listOther.Add(nwp4); // nwp4
            listOther.Add(sumplan); // sumplan
            listOther.Add(abs); // abs
            listOtherTarifs.Add(listOther);
        }
        private List<List<float>> FloatParseTables(List<List<string>> listStr)
        {
            CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.CurrencyDecimalSeparator = ".";
            List<List<float>> floatTable = new List<List<float>>();
            List<float> listFloat = new List<float>();

            foreach (var list in listStr)
            {
                listFloat = new List<float>();
                foreach (var j in list)
                {
                    listFloat.Add(float.Parse(j, NumberStyles.Any, ci));
                }
                floatTable.Add(listFloat);
            }
            return floatTable;
        }
        private List<List<string>> FloatParseTables(List<List<float>> listFloat)
        {
            List<List<string>> StringTable = new List<List<string>>();
            List<string> listString = new List<string>();

            foreach (var list in listFloat)
            {
                listString = new List<string>();
                foreach (var j in list)
                {
                    listString.Add(j.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
                }
                StringTable.Add(listString);
            }
            return StringTable;
        }
        #endregion

        #region Create and Write to Sql
        private void CreateSqlTable()
        {
            string connectionString = @"Data Source = 192.168.0.101;Initial Catalog=tarifsForCk;Persist Security Info = True; User ID = matrix; Password = matrix";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command1 = new SqlCommand();
                SqlCommand command2 = new SqlCommand();
                SqlCommand command3 = new SqlCommand();
                SqlCommand command4 = new SqlCommand();
                SqlCommand command5 = new SqlCommand();
                SqlCommand command6 = new SqlCommand();
                SqlCommand command7 = new SqlCommand();

                command1.CommandText = $"CREATE TABLE C1_{date} ( [provider] [float] NOT NULL, [contract] [float] NOT NULL, [bh] [float] NULL, [ch1] [float] NULL, [ch2] [float] NULL, [hh] [float] NULL, CONSTRAINT [C1_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [contract] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                command2.CommandText = $"CREATE TABLE C2_{date} ( [provider] [float] NOT NULL, [contract] [float] NOT NULL, [zone] [float] NOT NULL, [bh] [float] NULL, [ch1] [float] NULL, [ch2] [float] NULL, [hh] [float] NULL, CONSTRAINT [C2_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [contract] ASC, [zone] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                command3.CommandText = $"CREATE TABLE C3_{date} ( [provider] [float] NOT NULL,  [idPower] [float] NOT NULL, [bh] [float] NOT NULL, [date] [nvarchar](255) NOT NULL, [h0] [float] NULL, [h1] [float] NULL, [h2] [float] NULL, [h3] [float] NULL, [h4] [float] NULL, [h5] [float] NULL, [h6] [float] NULL, [h7] [float] NULL, [h8] [float] NULL, [h9] [float] NULL, [h10] [float] NULL, [h11] [float] NULL,  [h12] [float] NULL,  [h13] [float] NULL,  [h14] [float] NULL,  [h15] [float] NULL,  [h16] [float] NULL,  [h17] [float] NULL,  [h18] [float] NULL,  [h19] [float] NULL,  [h20] [float] NULL,  [h21] [float] NULL,  [h22] [float] NULL,  [h23] [float] NULL,  CONSTRAINT [C3_{date}_pk1] PRIMARY KEY NONCLUSTERED  (  [provider] ASC,  [idPower] ASC,  [bh] ASC,  [date] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]";
                command4.CommandText = $"CREATE TABLE C4_{date}( [provider] [float] NOT NULL, [idPower] [float] NOT NULL, [bh] [float] NOT NULL, [date] [nvarchar](255) NOT NULL, [h0] [float] NULL, [h1] [float] NULL, [h2] [float] NULL, [h3] [float] NULL, [h4] [float] NULL, [h5] [float] NULL, [h6] [float] NULL, [h7] [float] NULL, [h8] [float] NULL, [h9] [float] NULL, [h10] [float] NULL, [h11] [float] NULL, [h12] [float] NULL, [h13] [float] NULL, [h14] [float] NULL, [h15] [float] NULL, [h16] [float] NULL, [h17] [float] NULL, [h18] [float] NULL, [h19] [float] NULL, [h20] [float] NULL, [h21] [float] NULL, [h22] [float] NULL, [h23] [float] NULL, CONSTRAINT [C4_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [idPower] ASC, [bh] ASC, [date] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                command5.CommandText = $"CREATE TABLE C5_{date}( [provider] [float] NOT NULL, [idPower] [float] NOT NULL, [bh] [float] NOT NULL, [planfact] [float] NOT NULL, [date] [nvarchar](255) NOT NULL, [h0] [float] NULL, [h1] [float] NULL, [h2] [float] NULL, [h3] [float] NULL, [h4] [float] NULL, [h5] [float] NULL, [h6] [float] NULL, [h7] [float] NULL, [h8] [float] NULL, [h9] [float] NULL, [h10] [float] NULL, [h11] [float] NULL, [h12] [float] NULL, [h13] [float] NULL, [h14] [float] NULL, [h15] [float] NULL, [h16] [float] NULL, [h17] [float] NULL, [h18] [float] NULL, [h19] [float] NULL, [h20] [float] NULL, [h21] [float] NULL, [h22] [float] NULL, [h23] [float] NULL, CONSTRAINT [C5_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [idPower] ASC, [bh] ASC, [planfact] ASC, [date] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]";
                command6.CommandText = $"CREATE TABLE C6_{date}( [provider] [float] NOT NULL, [idPower] [float] NOT NULL, [bh] [float] NOT NULL, [planfact] [float] NOT NULL, [date] [nvarchar](255) NOT NULL, [h0] [float] NULL, [h1] [float] NULL, [h2] [float] NULL, [h3] [float] NULL, [h4] [float] NULL, [h5] [float] NULL, [h6] [float] NULL, [h7] [float] NULL, [h8] [float] NULL, [h9] [float] NULL, [h10] [float] NULL, [h11] [float] NULL, [h12] [float] NULL, [h13] [float] NULL, [h14] [float] NULL, [h15] [float] NULL, [h16] [float] NULL, [h17] [float] NULL, [h18] [float] NULL, [h19] [float] NULL, [h20] [float] NULL, [h21] [float] NULL, [h22] [float] NULL, [h23] [float] NULL, CONSTRAINT [C6_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [idPower] ASC, [bh] ASC, [planfact] ASC, [date] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]";
                command7.CommandText = $"CREATE TABLE otherTarifs_{date}( [provider] [float] NOT NULL, [idPower] [float] NOT NULL, [ck] [float] NOT NULL, [contract] [float] NOT NULL, [power] [float] NULL, [network1] [float] NULL, [network2] [float] NULL, [network3] [float] NULL, [network4] [float] NULL, [sumplan] [float] NULL, [abs] [float] NULL, CONSTRAINT [otherTarifs_{date}_pk1] PRIMARY KEY NONCLUSTERED ( [provider] ASC, [idPower] ASC, [ck] ASC, [contract] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]";

                command1.Connection = connection;
                command2.Connection = connection;
                command3.Connection = connection;
                command4.Connection = connection;
                command5.Connection = connection;
                command6.Connection = connection;
                command7.Connection = connection;

                command1.ExecuteNonQuery();
                command2.ExecuteNonQuery();
                command3.ExecuteNonQuery();
                command4.ExecuteNonQuery();
                command5.ExecuteNonQuery();
                command6.ExecuteNonQuery();
                command7.ExecuteNonQuery();
            }
        }
        private void InsertSqlTable()
        {
            listFloatTarifsCk1 = new List<List<float>>();
            listFloatTarifsCk2 = new List<List<float>>();
            listFloatTarifsCk3 = new List<List<float>>();
            listFloatTarifsCk4 = new List<List<float>>();
            listFloatTarifsCk5 = new List<List<float>>();
            listFloatTarifsCk6 = new List<List<float>>();
            listFloatOtherTarifs = new List<List<float>>();

            listFloatTarifsCk1.AddRange(FloatParseTables(listTarifsCk1));
            listFloatTarifsCk2.AddRange(FloatParseTables(listTarifsCk2));
            listFloatTarifsCk3.AddRange(FloatParseTables(listTarifsCk3));
            listFloatTarifsCk4.AddRange(FloatParseTables(listTarifsCk4));
            listFloatTarifsCk5.AddRange(FloatParseTables(listTarifsCk5));
            listFloatTarifsCk6.AddRange(FloatParseTables(listTarifsCk6));
            listFloatOtherTarifs.AddRange(FloatParseTables(listOtherTarifs));

            listTarifsCk1 = new List<List<string>>();
            listTarifsCk2 = new List<List<string>>();
            listTarifsCk3 = new List<List<string>>();
            listTarifsCk4 = new List<List<string>>();
            listTarifsCk5 = new List<List<string>>();
            listTarifsCk6 = new List<List<string>>();
            listOtherTarifs = new List<List<string>>();

            listTarifsCk1.AddRange(FloatParseTables(listFloatTarifsCk1));
            listTarifsCk2.AddRange(FloatParseTables(listFloatTarifsCk2));
            listTarifsCk3.AddRange(FloatParseTables(listFloatTarifsCk3));
            listTarifsCk4.AddRange(FloatParseTables(listFloatTarifsCk4));
            listTarifsCk5.AddRange(FloatParseTables(listFloatTarifsCk5));
            listTarifsCk6.AddRange(FloatParseTables(listFloatTarifsCk6));
            listOtherTarifs.AddRange(FloatParseTables(listFloatOtherTarifs));

            string connectionString = @"Data Source = 192.168.0.101;Initial Catalog=tarifsForCk;Persist Security Info = True; User ID = matrix; Password = matrix";
            CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.CurrencyDecimalSeparator = ".";
            foreach (var row in listTarifsCk1)
            {
                string sqlExpression = $"INSERT INTO C1_{date} ([provider], [contract], [bh], [ch1], [ch2], [hh]) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listTarifsCk2)
            {
                string sqlExpression = $"INSERT INTO C2_{date} (provider, contract, zone, bh, ch1, ch2, hh) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listTarifsCk3)
            {
                string sqlExpression = $"INSERT INTO C3_{date} (provider ,idPower ,bh ,date ,h0 ,h1 ,h2 ,h3 ,h4 ,h5 ,h6 ,h7 ,h8 ,h9 ,h10 ,h11 ,h12 ,h13 ,h14 ,h15 ,h16 ,h17 ,h18 ,h19 ,h20 ,h21 ,h22 ,h23) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]}, {row[7]}, {row[8]}, {row[9]}, {row[10]}, {row[11]}, {row[12]}, {row[13]}, {row[14]}, {row[15]}, {row[16]}, {row[17]}, {row[18]}, {row[19]}, {row[20]}, {row[21]}, {row[22]}, {row[23]}, {row[24]}, {row[25]}, {row[26]}, {row[27]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listTarifsCk4)
            {
                string sqlExpression = $"INSERT INTO C4_{date} (provider ,idPower ,bh ,date ,h0 ,h1 ,h2 ,h3 ,h4 ,h5 ,h6 ,h7 ,h8 ,h9 ,h10 ,h11 ,h12 ,h13 ,h14 ,h15 ,h16 ,h17 ,h18 ,h19 ,h20 ,h21 ,h22 ,h23) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]}, {row[7]}, {row[8]}, {row[9]}, {row[10]}, {row[11]}, {row[12]}, {row[13]}, {row[14]}, {row[15]}, {row[16]}, {row[17]}, {row[18]}, {row[19]}, {row[20]}, {row[21]}, {row[22]}, {row[23]}, {row[24]}, {row[25]}, {row[26]}, {row[27]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listTarifsCk5)
            {
                string sqlExpression = $"INSERT INTO C5_{date} (provider ,idPower ,bh ,planfact,date ,h0 ,h1 ,h2 ,h3 ,h4 ,h5 ,h6 ,h7 ,h8 ,h9 ,h10 ,h11 ,h12 ,h13 ,h14 ,h15 ,h16 ,h17 ,h18 ,h19 ,h20 ,h21 ,h22 ,h23) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]}, {row[7]}, {row[8]}, {row[9]}, {row[10]}, {row[11]}, {row[12]}, {row[13]}, {row[14]}, {row[15]}, {row[16]}, {row[17]}, {row[18]}, {row[19]}, {row[20]}, {row[21]}, {row[22]}, {row[23]}, {row[24]}, {row[25]}, {row[26]}, {row[27]}, {row[28]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listTarifsCk6)
            {
                string sqlExpression = $"INSERT INTO C6_{date} (provider ,idPower ,bh ,planfact,date ,h0 ,h1 ,h2 ,h3 ,h4 ,h5 ,h6 ,h7 ,h8 ,h9 ,h10 ,h11 ,h12 ,h13 ,h14 ,h15 ,h16 ,h17 ,h18 ,h19 ,h20 ,h21 ,h22 ,h23) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]}, {row[7]}, {row[8]}, {row[9]}, {row[10]}, {row[11]}, {row[12]}, {row[13]}, {row[14]}, {row[15]}, {row[16]}, {row[17]}, {row[18]}, {row[19]}, {row[20]}, {row[21]}, {row[22]}, {row[23]}, {row[24]}, {row[25]}, {row[26]}, {row[27]}, {row[28]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            foreach (var row in listOtherTarifs)
            {
                string sqlExpression = $"INSERT INTO otherTarifs_{date} (provider,idPower,ck,contract,power,network1,network2,network3,network4,sumplan,abs) VALUES ({row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, {row[5]}, {row[6]}, {row[7]}, {row[8]}, {row[9]}, {row[10]})";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }

            listTarifsCk1 = new List<List<string>>();
            listTarifsCk2 = new List<List<string>>();
            listTarifsCk3 = new List<List<string>>();
            listTarifsCk4 = new List<List<string>>();
            listTarifsCk5 = new List<List<string>>();
            listTarifsCk6 = new List<List<string>>();
            listOtherTarifs = new List<List<string>>();

            LabFinish.Visible = true;
            btnOkWrite.Image = Properties.Resources.accept;
        }
        #endregion

        #region Button Start
        private void MainCalc()
        {
            if (string.IsNullOrEmpty(date))
            {
                MessageBox.Show("Упс! Выберите дату!");
                return;
            }

            for (int index = 0; index < filenames.Length; index++)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filenames[index]);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                if (filenames[index].Contains("не менее 10000кВт"))
                {
                    Excel.Application xlApp1 = new Excel.Application();
                    Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(filenames[index]);
                    Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                    Excel.Range xlRange1 = xlWorksheet1.UsedRange;

                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "16 %";
                    // Для 3 ценовой категории
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "10", "1", 10, 1)); // bh
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "32 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "10", "2", 44, 1)); // ch 1
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "48 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "10", "3", 78, 1)); // ch 2
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "64 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "10", "4", 112, 1)); // hh 
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "80 %";
                    string power = xlRange1.Cells[145, 18].Value2.ToString();
                    OtherTarifsBuilder("10", "3", "1", power, "0", "0", "0", "0", "0", "0");

                    Marshal.ReleaseComObject(xlRange1);
                    Marshal.ReleaseComObject(xlWorksheet1);
                    Excel._Worksheet xlWorksheet3 = xlWorkbook1.Sheets[5];
                    Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "90 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "10", "5", 10, 1)); // for contract = 2
                    tableProcess3.Rows[0].Cells["ColProc3"].Value = "100 %";
                    tableProcess3.Rows[0].Cells["ColExe3"].Value = Properties.Resources.accept;

                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "16 %";
                    // Для 4 ценовой категории
                    Marshal.ReleaseComObject(xlRange3);
                    Marshal.ReleaseComObject(xlWorksheet3);
                    Excel._Worksheet xlWorksheet4 = xlWorkbook1.Sheets[2];
                    Excel.Range xlRange4 = xlWorksheet4.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "10", "1", 10, 1)); // bh
                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "32 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "10", "2", 44, 1)); // ch 1
                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "48 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "10", "3", 78, 1)); // ch 2
                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "64 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "10", "4", 112, 1)); // hh 

                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "80 %";
                    string nwp1 = xlRange4.Cells[151, 7].Value2.ToString();
                    string nwp2 = xlRange4.Cells[151, 10].Value2.ToString();
                    string nwp3 = xlRange4.Cells[151, 13].Value2.ToString();
                    string nwp4 = xlRange4.Cells[151, 16].Value2.ToString();
                    OtherTarifsBuilder("10", "4", "1", power, nwp1, nwp2, nwp3, nwp4, "0", "0");
                    OtherTarifsBuilder("10", "4", "2", power, "0", "0", "0", "0", "0", "0");

                    Marshal.ReleaseComObject(xlRange4);
                    Marshal.ReleaseComObject(xlWorksheet4);
                    Excel._Worksheet xlWorksheet44 = xlWorkbook1.Sheets[6];
                    Excel.Range xlRange44 = xlWorksheet44.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange44, "10", "5", 10, 1)); // for contract = 2
                    tableProcess3.Rows[1].Cells["ColProc3"].Value = "100 %";
                    tableProcess3.Rows[1].Cells["ColExe3"].Value = Properties.Resources.accept;

                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "16 %";
                    // Для 5 ценовой категории
                    Marshal.ReleaseComObject(xlRange44);
                    Marshal.ReleaseComObject(xlWorksheet44);
                    Excel._Worksheet xlWorksheet5 = xlWorkbook1.Sheets[3];
                    Excel.Range xlRange5 = xlWorksheet5.UsedRange;

                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "1", "0", 10, 1)); // bh
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "32 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "2", "0", 44, 1)); // ch 1
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "48 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "3", "0", 78, 1)); // ch 2
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "64 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "4", "0", 112, 1)); // hh 
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "80 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "5", "1", 146, 1));  // planfact
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "90 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "10", "5", "2", 180, 1));  // planfact

                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "95 %";
                    string sumplan = xlRange5.Cells[214, 20].Value2.ToString();
                    string abs = xlRange5.Cells[215, 20].Value2.ToString();
                    OtherTarifsBuilder("10", "5", "1", power, "0", "0", "0", "0", sumplan, abs);
                    OtherTarifsBuilder("10", "5", "2", power, "0", "0", "0", "0", sumplan, abs);

                    Marshal.ReleaseComObject(xlRange5);
                    Marshal.ReleaseComObject(xlWorksheet5);
                    Excel._Worksheet xlWorksheet55 = xlWorkbook1.Sheets[7];
                    Excel.Range xlRange55 = xlWorksheet55.UsedRange;
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange55, "10", "5", "0", 10, 1));
                    tableProcess3.Rows[2].Cells["ColProc3"].Value = "100 %";
                    tableProcess3.Rows[2].Cells["ColExe3"].Value = Properties.Resources.accept;

                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "15 %";
                    // Для 6 ценовой категории
                    Marshal.ReleaseComObject(xlRange55);
                    Marshal.ReleaseComObject(xlWorksheet55);
                    Excel._Worksheet xlWorksheet6 = xlWorkbook1.Sheets[4];
                    Excel.Range xlRange6 = xlWorksheet6.UsedRange;

                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "1", "0", 10, 1)); // bh
                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "30 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "2", "0", 44, 1)); // ch 1
                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "45 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "3", "0", 78, 1)); // ch 2
                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "60 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "4", "0", 112, 1)); // hh 
                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "75 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "5", "1", 146, 1));  // planfact
                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "90 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "10", "5", "2", 180, 1));  // planfact

                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "95 %";
                    OtherTarifsBuilder("10", "6", "1", power, nwp1, nwp2, nwp3, nwp4, sumplan, abs);
                    OtherTarifsBuilder("10", "6", "2", power, "0", "0", "0", "0", sumplan, abs);

                    Marshal.ReleaseComObject(xlRange6);
                    Marshal.ReleaseComObject(xlWorksheet6);
                    Excel._Worksheet xlWorksheet66 = xlWorkbook1.Sheets[8];
                    Excel.Range xlRange66 = xlWorksheet66.UsedRange;
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange66, "10", "5", "0", 10, 1));

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlRange66);
                    Marshal.ReleaseComObject(xlWorksheet66);

                    //close and release
                    xlWorkbook1.Close();
                    Marshal.ReleaseComObject(xlWorkbook1);

                    //quit and release
                    xlApp1.Quit();
                    Marshal.ReleaseComObject(xlApp1);

                    tableProcess3.Rows[3].Cells["ColProc3"].Value = "100 %";
                    tableProcess3.Rows[3].Cells["ColExe3"].Value = Properties.Resources.accept;
                }

                if (filenames[index].Contains("от 670кВт до 10000кВт"))
                {
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "0 %";
                    Excel.Application xlApp1 = new Excel.Application();
                    Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(filenames[index]);
                    Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                    Excel.Range xlRange1 = xlWorksheet1.UsedRange;

                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "16 %";
                    // Для 3 ценовой категории
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "670", "1", 10, 1)); // bh
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "32 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "670", "2", 44, 1)); // ch 1
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "48 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "670", "3", 78, 1)); // ch 2
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "64 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange1, "670", "4", 112, 1)); // hh 
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "80 %";
                    string power = xlRange1.Cells[145, 18].Value2.ToString();
                    OtherTarifsBuilder("670", "3", "1", power, "0", "0", "0", "0", "0", "0");

                    Marshal.ReleaseComObject(xlRange1);
                    Marshal.ReleaseComObject(xlWorksheet1);
                    Excel._Worksheet xlWorksheet3 = xlWorkbook1.Sheets[5];
                    Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "90 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "670", "5", 10, 1)); // for contract = 2
                    tableProcess2.Rows[0].Cells["ColProc2"].Value = "100 %";
                    tableProcess2.Rows[0].Cells["ColExe2"].Value = Properties.Resources.accept;

                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "0 %";
                    // Для 4 ценовой категории
                    Marshal.ReleaseComObject(xlRange3);
                    Marshal.ReleaseComObject(xlWorksheet3);
                    Excel._Worksheet xlWorksheet4 = xlWorkbook1.Sheets[2];
                    Excel.Range xlRange4 = xlWorksheet4.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "670", "1", 10, 1)); // bh
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "15 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "670", "2", 44, 1)); // ch 1
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "30 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "670", "3", 78, 1)); // ch 2
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "45 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "670", "4", 112, 1)); // hh 

                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "60 %";
                    string nwp1 = xlRange4.Cells[151, 7].Value2.ToString();
                    string nwp2 = xlRange4.Cells[151, 10].Value2.ToString();
                    string nwp3 = xlRange4.Cells[151, 13].Value2.ToString();
                    string nwp4 = xlRange4.Cells[151, 16].Value2.ToString();
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "75 %";
                    OtherTarifsBuilder("670", "4", "1", power, nwp1, nwp2, nwp3, nwp4, "0", "0");
                    OtherTarifsBuilder("670", "4", "2", power, "0", "0", "0", "0", "0", "0");
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "90 %";
                    Marshal.ReleaseComObject(xlRange4);
                    Marshal.ReleaseComObject(xlWorksheet4);
                    Excel._Worksheet xlWorksheet44 = xlWorkbook1.Sheets[6];
                    Excel.Range xlRange44 = xlWorksheet44.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange44, "670", "5", 10, 1)); // for contract = 2
                    tableProcess2.Rows[1].Cells["ColProc2"].Value = "100 %";
                    tableProcess2.Rows[1].Cells["ColExe2"].Value = Properties.Resources.accept;

                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "0 %";
                    // Для 5 ценовой категории
                    Marshal.ReleaseComObject(xlRange44);
                    Marshal.ReleaseComObject(xlWorksheet44);
                    Excel._Worksheet xlWorksheet5 = xlWorkbook1.Sheets[3];
                    Excel.Range xlRange5 = xlWorksheet5.UsedRange;

                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "1", "0", 10, 1)); // bh
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "15 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "2", "0", 44, 1)); // ch 1
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "30 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "3", "0", 78, 1)); // ch 2
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "45 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "4", "0", 112, 1)); // hh 
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "60 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "5", "1", 146, 1));  // planfact
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "70 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "670", "5", "2", 180, 1));  // planfact

                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "80 %";
                    string sumplan = xlRange5.Cells[214, 20].Value2.ToString();
                    string abs = xlRange5.Cells[215, 20].Value2.ToString();
                    OtherTarifsBuilder("670", "5", "1", power, "0", "0", "0", "0", sumplan, abs);
                    OtherTarifsBuilder("670", "5", "2", power, "0", "0", "0", "0", sumplan, abs);

                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "90 %";
                    Marshal.ReleaseComObject(xlRange5);
                    Marshal.ReleaseComObject(xlWorksheet5);
                    Excel._Worksheet xlWorksheet55 = xlWorkbook1.Sheets[7];
                    Excel.Range xlRange55 = xlWorksheet55.UsedRange;
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange55, "670", "5", "0", 10, 1));
                    tableProcess2.Rows[2].Cells["ColProc2"].Value = "100 %";
                    tableProcess2.Rows[2].Cells["ColExe2"].Value = Properties.Resources.accept;

                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "0 %";
                    // Для 6 ценовой категории
                    Marshal.ReleaseComObject(xlRange55);
                    Marshal.ReleaseComObject(xlWorksheet55);
                    Excel._Worksheet xlWorksheet6 = xlWorkbook1.Sheets[4];
                    Excel.Range xlRange6 = xlWorksheet6.UsedRange;

                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "1", "0", 10, 1)); // bh
                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "15 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "2", "0", 44, 1)); // ch 1
                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "30 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "3", "0", 78, 1)); // ch 2
                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "45 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "4", "0", 112, 1)); // hh 
                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "60 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "5", "1", 146, 1));  // planfact
                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "75 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "670", "5", "2", 180, 1));  // planfact

                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "85 %";
                    OtherTarifsBuilder("670", "6", "1", power, nwp1, nwp2, nwp3, nwp4, sumplan, abs);
                    OtherTarifsBuilder("670", "6", "2", power, "0", "0", "0", "0", sumplan, abs);

                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "95 %";
                    Marshal.ReleaseComObject(xlRange6);
                    Marshal.ReleaseComObject(xlWorksheet6);
                    Excel._Worksheet xlWorksheet66 = xlWorkbook1.Sheets[8];
                    Excel.Range xlRange66 = xlWorksheet66.UsedRange;
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange66, "670", "5", "0", 10, 1));

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlRange66);
                    Marshal.ReleaseComObject(xlWorksheet66);

                    //close and release
                    xlWorkbook1.Close();
                    Marshal.ReleaseComObject(xlWorkbook1);

                    //quit and release
                    xlApp1.Quit();
                    Marshal.ReleaseComObject(xlApp1);

                    tableProcess2.Rows[3].Cells["ColProc2"].Value = "100 %";
                    tableProcess2.Rows[3].Cells["ColExe2"].Value = Properties.Resources.accept;
                }

                if (filenames[index].Contains("до 670кВт"))
                {
                    tableProcess1.Rows[0].Cells["ColProc"].Value = "0 %";
                    Excel.Application xlApp1 = new Excel.Application();
                    Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(filenames[index]);
                    Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                    Excel.Range xlRange1 = xlWorksheet1.UsedRange;

                    List<string> tmpCk1 = new List<string>();
                    tmpCk1.Add("1");
                    tmpCk1.Add("1");
                    tmpCk1.Add(xlRange1.Cells[23, 7].Value2.ToString());
                    tmpCk1.Add(xlRange1.Cells[23, 10].Value2.ToString());
                    tmpCk1.Add(xlRange1.Cells[23, 13].Value2.ToString());
                    tmpCk1.Add(xlRange1.Cells[23, 16].Value2.ToString());
                    listTarifsCk1.Add(tmpCk1);

                    tableProcess1.Rows[0].Cells["ColProc"].Value = "50 %";
                    Marshal.ReleaseComObject(xlRange1);
                    Marshal.ReleaseComObject(xlWorksheet1);
                    Excel._Worksheet xlWorksheet11 = xlWorkbook1.Sheets[7];
                    Excel.Range xlRange11 = xlWorksheet11.UsedRange;
                    tmpCk1 = new List<string>();
                    tmpCk1.Add("1");
                    tmpCk1.Add("2");
                    tmpCk1.Add(xlRange11.Cells[23, 7].Value2.ToString());
                    tmpCk1.Add(xlRange11.Cells[23, 10].Value2.ToString());
                    tmpCk1.Add(xlRange11.Cells[23, 13].Value2.ToString());
                    tmpCk1.Add(xlRange11.Cells[23, 16].Value2.ToString());
                    listTarifsCk1.Add(tmpCk1);

                    tableProcess1.Rows[0].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[0].Cells["ColExe"].Value = Properties.Resources.accept;

                    Marshal.ReleaseComObject(xlRange11);
                    Marshal.ReleaseComObject(xlWorksheet11);
                    Excel._Worksheet xlWorksheet2 = xlWorkbook1.Sheets[2];
                    Excel.Range xlRange2 = xlWorksheet2.UsedRange;

                    tableProcess1.Rows[1].Cells["ColProc"].Value = "0 %";
                    List<string> tmpCk2 = new List<string>();

                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add(xlRange2.Cells[12, 10].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[12, 12].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[12, 14].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[12, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add("2");
                    tmpCk2.Add(xlRange2.Cells[13, 10].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[13, 12].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[13, 14].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[13, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add("3");
                    tmpCk2.Add(xlRange2.Cells[14, 10].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[14, 12].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[14, 14].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[14, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add("4");
                    tmpCk2.Add(xlRange2.Cells[20, 10].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[20, 12].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[20, 14].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[20, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("1");
                    tmpCk2.Add("5");
                    tmpCk2.Add(xlRange2.Cells[21, 10].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[21, 12].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[21, 14].Value2.ToString());
                    tmpCk2.Add(xlRange2.Cells[21, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tableProcess1.Rows[1].Cells["ColProc"].Value = "50 %";

                    Marshal.ReleaseComObject(xlRange2);
                    Marshal.ReleaseComObject(xlWorksheet2);
                    Excel._Worksheet xlWorksheet22 = xlWorkbook1.Sheets[8];
                    Excel.Range xlRange22 = xlWorksheet22.UsedRange;

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("2");
                    tmpCk2.Add("1");
                    tmpCk2.Add(xlRange22.Cells[12, 10].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[12, 12].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[12, 14].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[12, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("2");
                    tmpCk2.Add("2");
                    tmpCk2.Add(xlRange22.Cells[13, 10].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[13, 12].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[13, 14].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[13, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);

                    tmpCk2 = new List<string>();
                    tmpCk2.Add("1");
                    tmpCk2.Add("2");
                    tmpCk2.Add("3");
                    tmpCk2.Add(xlRange22.Cells[14, 10].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[14, 12].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[14, 14].Value2.ToString());
                    tmpCk2.Add(xlRange22.Cells[14, 16].Value2.ToString());
                    listTarifsCk2.Add(tmpCk2);
                    tableProcess1.Rows[1].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[1].Cells["ColExe"].Value = Properties.Resources.accept;

                    Marshal.ReleaseComObject(xlRange22);
                    Marshal.ReleaseComObject(xlWorksheet22);
                    Excel._Worksheet xlWorksheet3 = xlWorkbook1.Sheets[3];
                    Excel.Range xlRange3 = xlWorksheet3.UsedRange;

                    tableProcess1.Rows[2].Cells["ColProc"].Value = "0 %";
                    // Для 3 ценовой категории
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "150", "1", 10, 1)); // bh
                    tableProcess1.Rows[2].Cells["ColProc"].Value = "15 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "150", "2", 44, 1)); // ch 1
                    tableProcess1.Rows[2].Cells["ColProc"].Value = "30 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "150", "3", 78, 1)); // ch 2
                    tableProcess1.Rows[2].Cells["ColProc"].Value = "45 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange3, "150", "4", 112, 1)); // hh 

                    string power = xlRange3.Cells[145, 18].Value2.ToString();
                    OtherTarifsBuilder("150", "3", "1", power, "0", "0", "0", "0", "0", "0");

                    tableProcess1.Rows[2].Cells["ColProc"].Value = "60 %";
                    Marshal.ReleaseComObject(xlRange3);
                    Marshal.ReleaseComObject(xlWorksheet3);
                    Excel._Worksheet xlWorksheet33 = xlWorkbook1.Sheets[9];
                    Excel.Range xlRange33 = xlWorksheet33.UsedRange;
                    tableProcess1.Rows[2].Cells["ColProc"].Value = "75 %";
                    listTarifsCk3.AddRange(ArrayBuilder(xlRange33, "150", "5", 10, 1)); // for contract = 2
                    tableProcess1.Rows[2].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[2].Cells["ColExe"].Value = Properties.Resources.accept;

                    tableProcess1.Rows[3].Cells["ColProc"].Value = "0 %";
                    // Для 4 ценовой категории
                    Marshal.ReleaseComObject(xlRange33);
                    Marshal.ReleaseComObject(xlWorksheet33);
                    Excel._Worksheet xlWorksheet4 = xlWorkbook1.Sheets[4];
                    Excel.Range xlRange4 = xlWorksheet4.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "150", "1", 10, 1)); // bh
                    tableProcess1.Rows[3].Cells["ColProc"].Value = "15 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "150", "2", 44, 1)); // ch 1
                    tableProcess1.Rows[3].Cells["ColProc"].Value = "30 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "150", "3", 78, 1)); // ch 2
                    tableProcess1.Rows[3].Cells["ColProc"].Value = "45 %";
                    listTarifsCk4.AddRange(ArrayBuilder(xlRange4, "150", "4", 112, 1)); // hh 

                    tableProcess1.Rows[3].Cells["ColProc"].Value = "60 %";
                    string nwp1 = xlRange4.Cells[151, 7].Value2.ToString();
                    string nwp2 = xlRange4.Cells[151, 10].Value2.ToString();
                    string nwp3 = xlRange4.Cells[151, 13].Value2.ToString();
                    string nwp4 = xlRange4.Cells[151, 16].Value2.ToString();
                    tableProcess1.Rows[3].Cells["ColProc"].Value = "75 %";
                    OtherTarifsBuilder("150", "4", "1", power, nwp1, nwp2, nwp3, nwp4, "0", "0");
                    OtherTarifsBuilder("150", "4", "2", power, "0", "0", "0", "0", "0", "0");
                    tableProcess1.Rows[3].Cells["ColProc"].Value = "90 %";
                    Marshal.ReleaseComObject(xlRange4);
                    Marshal.ReleaseComObject(xlWorksheet4);
                    Excel._Worksheet xlWorksheet44 = xlWorkbook1.Sheets[10];
                    Excel.Range xlRange44 = xlWorksheet44.UsedRange;

                    listTarifsCk4.AddRange(ArrayBuilder(xlRange44, "150", "5", 10, 1)); // for contract = 2

                    tableProcess1.Rows[3].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[3].Cells["ColExe"].Value = Properties.Resources.accept;

                    tableProcess1.Rows[4].Cells["ColProc"].Value = "0 %";
                    // Для 5 ценовой категории
                    Marshal.ReleaseComObject(xlRange44);
                    Marshal.ReleaseComObject(xlWorksheet44);
                    Excel._Worksheet xlWorksheet5 = xlWorkbook1.Sheets[5];
                    Excel.Range xlRange5 = xlWorksheet5.UsedRange;

                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "1", "0", 10, 1)); // bh
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "15 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "2", "0", 44, 1)); // ch 1
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "30 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "3", "0", 78, 1)); // ch 2
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "45 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "4", "0", 112, 1)); // hh 
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "60 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "5", "1", 146, 1));  // planfact
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "75 %";
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange5, "150", "5", "2", 180, 1));  // planfact

                    tableProcess1.Rows[4].Cells["ColProc"].Value = "90 %";
                    string sumplan = xlRange5.Cells[214, 20].Value2.ToString();
                    string abs = xlRange5.Cells[215, 20].Value2.ToString();
                    OtherTarifsBuilder("150", "5", "1", power, "0", "0", "0", "0", sumplan, abs);
                    OtherTarifsBuilder("150", "5", "2", power, "0", "0", "0", "0", sumplan, abs);

                    Marshal.ReleaseComObject(xlRange5);
                    Marshal.ReleaseComObject(xlWorksheet5);
                    Excel._Worksheet xlWorksheet55 = xlWorkbook1.Sheets[11];
                    Excel.Range xlRange55 = xlWorksheet55.UsedRange;
                    listTarifsCk5.AddRange(ArrayBuilder5(xlRange55, "150", "5", "0", 10, 1));
                    tableProcess1.Rows[4].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[4].Cells["ColExe"].Value = Properties.Resources.accept;

                    tableProcess1.Rows[5].Cells["ColProc"].Value = "0 %";
                    // Для 6 ценовой категории
                    Marshal.ReleaseComObject(xlRange55);
                    Marshal.ReleaseComObject(xlWorksheet55);
                    Excel._Worksheet xlWorksheet6 = xlWorkbook1.Sheets[6];
                    Excel.Range xlRange6 = xlWorksheet6.UsedRange;

                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "1", "0", 10, 1)); // bh
                    tableProcess1.Rows[5].Cells["ColProc"].Value = "15 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "2", "0", 44, 1)); // ch 1
                    tableProcess1.Rows[5].Cells["ColProc"].Value = "30 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "3", "0", 78, 1)); // ch 2
                    tableProcess1.Rows[5].Cells["ColProc"].Value = "45 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "4", "0", 112, 1)); // hh 
                    tableProcess1.Rows[5].Cells["ColProc"].Value = "60 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "5", "1", 146, 1));  // planfact
                    tableProcess1.Rows[5].Cells["ColProc"].Value = "75 %";
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange6, "150", "5", "2", 180, 1));  // planfact

                    tableProcess1.Rows[5].Cells["ColProc"].Value = "90 %";
                    OtherTarifsBuilder("150", "6", "1", power, nwp1, nwp2, nwp3, nwp4, sumplan, abs);
                    OtherTarifsBuilder("150", "6", "2", power, "0", "0", "0", "0", sumplan, abs);

                    Marshal.ReleaseComObject(xlRange6);
                    Marshal.ReleaseComObject(xlWorksheet6);
                    Excel._Worksheet xlWorksheet66 = xlWorkbook1.Sheets[12];
                    Excel.Range xlRange66 = xlWorksheet66.UsedRange;
                    listTarifsCk6.AddRange(ArrayBuilder5(xlRange66, "150", "5", "0", 10, 1));

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlRange66);
                    Marshal.ReleaseComObject(xlWorksheet66);

                    //close and release
                    xlWorkbook1.Close();
                    Marshal.ReleaseComObject(xlWorkbook1);

                    //quit and release
                    xlApp1.Quit();
                    Marshal.ReleaseComObject(xlApp1);

                    tableProcess1.Rows[5].Cells["ColProc"].Value = "100 %";
                    tableProcess1.Rows[5].Cells["ColExe"].Value = Properties.Resources.tick;
                }
            }
            try
            {
                CreateSqlTable();
            }
            catch (Exception exc)
            {
                MessageBox.Show("Упс! " + exc.Message);
            }
            try
            {
                InsertSqlTable();
            }
            catch (Exception exc)
            {
                MessageBox.Show("Упс! " + exc.Message);
            }
        }
        private void BtnStart_Click(object sender, EventArgs e)
        {
            Thread myThread = new Thread(new ThreadStart(MainCalc));
            myThread.Start();
        }
        #endregion
       
    }
}
