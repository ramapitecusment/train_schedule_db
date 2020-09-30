using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ZD
{
    public partial class Zd_Scedule : Form
    {
        public static string connection_string = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=zd.mdb;";

        string querry, querry1, finalQuerry, querryInsert;

        string standart_querry = "SELECT TRAIN.ID_TRAIN, TRAIN.NAME_TRAIN, STATION.NAME_STATION, STATION_1.NAME_STATION, TRAIN.DEP_DATE, TRAIN.ARR_DATE, " +
                    "TYPE_VAGON.TYPE_VAGON, VAGON_TRAIN.LEFT_SEATS, VAGON_TRAIN.PRICE " +
                    "FROM STATION AS STATION_1 " +
                    "INNER JOIN(TYPE_VAGON INNER JOIN ((STATION INNER JOIN TRAIN ON STATION.ID_STATION = TRAIN.FROM_STATION) " +
                    "INNER JOIN VAGON_TRAIN ON TRAIN.ID_TRAIN = VAGON_TRAIN.ID_TRAIN) ON TYPE_VAGON.ID_VAGON_TYPE = VAGON_TRAIN.ID_TYPE_VAGON) " +
                    "ON STATION_1.ID_STATION = TRAIN.TO_STATION";

        string login, from_station, to_station;

        List<string> querryList = new List<string>();
        List<string> querryList1 = new List<string>();

        List<string> idTrain = new List<string>();
        List<string> vagonType = new List<string>();

        public Zd_Scedule()
        {
            InitializeComponent();
        }

        public Zd_Scedule(string login, List<string> logins)
        {
            InitializeComponent();

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy HH:mm";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yyyy HH:mm";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd.MM.yyyy HH:mm";
            this.login = login;

        }

        private void Zd_Scedule_Shown(object sender, EventArgs e)
        {
            tabPage2.Parent = null;
            if (login == "admin")
            {
                //MessageBox.Show(login);
                tabPage2.Parent = tabControl1;
            }
            updateBase();
        }

        private void updateBase()
        {
            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();
                querry = "SELECT NAME_STATION FROM STATION";

                OleDbCommand command = new OleDbCommand(querry, conn);
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader[0]);
                    comboBox2.Items.Add(reader[0]);
                    comboBox3.Items.Add(reader[0]);
                    comboBox4.Items.Add(reader[0]);
                }
            }
        }

        private void поискПоездаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            querry = "";
            querry1 = "";
            querryList.Clear();
            querryList1.Clear();
            if (comboBox1.SelectedIndex != -1 && comboBox1.Text != "")
            {
                querry = "STATION.NAME_STATION = '" + comboBox1.Text + "'";
                querryList.Add(querry);
            }
            if (comboBox2.SelectedIndex != -1 && comboBox2.Text != "")
            {
                querry = "STATION_1.NAME_STATION = '" + comboBox2.Text + "'";
                querryList.Add(querry);
            }
            if(textBox1.Text != "")
            {
                querry = "VAGON_TRAIN.PRICE >= " + textBox1.Text;
                querryList.Add(querry);
            }
            if (textBox2.Text != "")
            {
                querry = "VAGON_TRAIN.PRICE <= " + textBox2.Text;
                querryList.Add(querry);
            }
            if(dateTimePicker1.Value != null)
            {
                querry = "TRAIN.DEP_DATE >= @DT";
                querryList.Add(querry);
            }

            foreach (object itemChecked in checkedListBox1.CheckedItems)
            {
                querry1 = "TYPE_VAGON.TYPE_VAGON = '" + itemChecked.ToString() + "'";
                querryList1.Add(querry1);
            }

            if (querry1 != "")
            {
                querry = String.Join(" OR ", querryList1);
                querryList.Add(querry);
            }

            fillDataGrid(String.Join(" AND ", querryList));
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void выходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearAll();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (checkAllFilled())
            {
                MessageBox.Show("Данные успешно добавлены в таблицу");
                fillDataBase();
                clearAll();
            }
            else
            {
                MessageBox.Show("Все поля должны быть заполнены!");
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            clearAll();
        }

        private void fillDataGrid(string querry)
        {
            finalQuerry = "";
            if(querry == "")
            {
                finalQuerry = standart_querry + ";";
            }
            else
            {
                finalQuerry = standart_querry + " WHERE " + querry + ";";
            }
            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();
                
                dataGridView1.Rows.Clear();
                idTrain.Clear();
                vagonType.Clear();

                int i = 0;
                OleDbCommand cmd = new OleDbCommand(finalQuerry, conn);

                cmd.Parameters.Add("@DT", OleDbType.Date).Value = dateTimePicker1.Value;

                OleDbDataReader reader = cmd.ExecuteReader();
                
                while (reader.Read())
                {
                    dataGridView1.Rows.Add();
                    idTrain.Add(reader[0].ToString());
                    vagonType.Add(reader[6].ToString());
                    for (int j = 1; j < reader.FieldCount; j++)
                    {
                        dataGridView1.Rows[i].Cells[j-1].Value = reader[j];
                    }
                    dataGridView1.Rows[i].Cells[reader.FieldCount - 1].Value = Convert.ToDateTime(reader["ARR_DATE"]) - Convert.ToDateTime(reader["DEP_DATE"]);
                    i++;
                }
            }
        }
        
        private string idTrainQuerry()
        {
            string idTrain = "";
            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                string sm_querry = "SELECT MAX(ID_TRAIN) FROM TRAIN;";

                OleDbCommand cmd = new OleDbCommand(sm_querry, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    idTrain = reader[0].ToString();
                }
            }
            //MessageBox.Show(idTrain);
            return idTrain;

        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            ShowTrain sh1 = new ShowTrain(idTrain[dataGridView1.CurrentRow.Index], vagonType[dataGridView1.CurrentRow.Index], dataGridView1.CurrentRow.Cells[5].Value.ToString());
            sh1.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for(int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (i == 0)
                    {
                        ExcelWorkSheet.Cells[1, j+1] = dataGridView1.Columns[j].HeaderCell.Value.ToString();
                    }
                    else
                    {
                        //MessageBox.Show(dataGridView1.Rows[i].Cells[j].Value);
                        ExcelWorkSheet.Cells[i+1, j+1] = dataGridView1.Rows[i-1].Cells[j].Value.ToString();
                    }
                }
            }

            ExcelWorkSheet.Cells[1, 1].CurrentRegion.Borders.LineStyle = XlLineStyle.xlContinuous;
            ExcelWorkSheet.Rows[1].Font.Bold = true;
            ExcelWorkSheet.Range["A:I"].EntireColumn.AutoFit();
            ExcelApp.Visible = true;
            //ExcelApp.UserControl = true;
        }

        private string makeStationQuerry(string station)
        {
            string MyStation = "";
            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                string sm_querry = "SELECT ID_STATION FROM STATION WHERE NAME_STATION = '" + station + "';";

                OleDbCommand cmd = new OleDbCommand(sm_querry, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    MyStation = reader[0].ToString();
                }
            }
            //MessageBox.Show(MyStation);
            return MyStation;
        }

        private string makeVagonQuerry(string vagon)
        {
            string MyVagon = "";

            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                string sm_querry = "SELECT ID_VAGON_TYPE FROM TYPE_VAGON WHERE TYPE_VAGON = '" + vagon + "';";

                OleDbCommand cmd = new OleDbCommand(sm_querry, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    MyVagon = reader[0].ToString();
                }
            }
            //MessageBox.Show(MyVagon);
            return MyVagon;
        }

        private void fillDataBase()
        {
            if (checkAllFilled())
            {
                using (OleDbConnection conn = new OleDbConnection(connection_string))
                {
                    conn.Open();
                    from_station = makeStationQuerry(comboBox3.Text);
                    to_station = makeStationQuerry(comboBox4.Text);
                    querryInsert = "INSERT INTO TRAIN(NAME_TRAIN, FROM_STATION, TO_STATION, DEP_DATE, ARR_DATE) " +
                        "VALUES(@train_name, @from_station, @to_station, @dep_date, @arr_date);";
                    
                    OleDbCommand cmd = new OleDbCommand(querryInsert, conn);
                    cmd.Parameters.Add("@train_name", OleDbType.VarChar);
                    cmd.Parameters["@train_name"].Value = textBox7.Text;
                    cmd.Parameters.Add("@from_station", OleDbType.Integer);
                    cmd.Parameters["@from_station"].Value = from_station;
                    cmd.Parameters.Add("@to_station", OleDbType.Integer);
                    cmd.Parameters["@to_station"].Value = to_station;
                    cmd.Parameters.Add("@dep_date", OleDbType.Date);
                    cmd.Parameters["@dep_date"].Value = dateTimePicker2.Value;
                    cmd.Parameters.Add("@arr_date", OleDbType.Date);
                    cmd.Parameters["@arr_date"].Value = dateTimePicker3.Value;

                    cmd.ExecuteNonQuery();


                }

                using (OleDbConnection conn = new OleDbConnection(connection_string))
                {
                    conn.Open();

                    string id_train = idTrainQuerry();

                    if(textBox4.Text != "")
                    {
                        querryInsert = "INSERT INTO VAGON_TRAIN(ID_TRAIN, ID_TYPE_VAGON, TOTAL_SEATS, LEFT_SEATS, PRICE)" +
                        "VALUES(@id_train, @id_type_vagon, @total_seats, @left_seats, @price);";
                        OleDbCommand cmd = new OleDbCommand(querryInsert, conn);
                        cmd.Parameters.Add("@id_train", OleDbType.Integer);
                        cmd.Parameters["@id_train"].Value = id_train;
                        cmd.Parameters.Add("@id_type_vagon", OleDbType.Integer);
                        cmd.Parameters["@id_type_vagon"].Value = makeVagonQuerry("Сидячий");
                        cmd.Parameters.Add("@total_seats", OleDbType.Integer);
                        cmd.Parameters["@total_seats"].Value = textBox4.Text;
                        cmd.Parameters.Add("@left_seats", OleDbType.Integer);
                        cmd.Parameters["@left_seats"].Value = textBox4.Text;
                        cmd.Parameters.Add("@price", OleDbType.Integer);
                        cmd.Parameters["@price"].Value = textBox3.Text;

                        cmd.ExecuteNonQuery();
                    }
                    if(textBox5.Text != "")
                    {
                        querryInsert = "INSERT INTO VAGON_TRAIN(ID_TRAIN, ID_TYPE_VAGON, TOTAL_SEATS, LEFT_SEATS, PRICE)" +
                        "VALUES(@id_train, @id_type_vagon, @total_seats, @left_seats, @price*1.5);";
                        OleDbCommand cmd = new OleDbCommand(querryInsert, conn);
                        cmd.Parameters.Add("@id_train", OleDbType.Integer);
                        cmd.Parameters["@id_train"].Value = id_train;
                        cmd.Parameters.Add("@id_type_vagon", OleDbType.Integer);
                        cmd.Parameters["@id_type_vagon"].Value = makeVagonQuerry("Плацкарт");
                        cmd.Parameters.Add("@total_seats", OleDbType.Integer);
                        cmd.Parameters["@total_seats"].Value = textBox5.Text;
                        cmd.Parameters.Add("@left_seats", OleDbType.Integer);
                        cmd.Parameters["@left_seats"].Value = textBox5.Text;
                        cmd.Parameters.Add("@price", OleDbType.Integer);
                        cmd.Parameters["@price"].Value = textBox3.Text;

                        cmd.ExecuteNonQuery();
                    }
                    if(textBox6.Text != "")
                    {
                        querryInsert = "INSERT INTO VAGON_TRAIN(ID_TRAIN, ID_TYPE_VAGON, TOTAL_SEATS, LEFT_SEATS, PRICE)" +
                        "VALUES(@id_train, @id_type_vagon, @total_seats, @left_seats, @price*2);";
                        OleDbCommand cmd = new OleDbCommand(querryInsert, conn);
                        cmd.Parameters.Add("@id_train", OleDbType.Integer);
                        cmd.Parameters["@id_train"].Value = id_train;
                        cmd.Parameters.Add("@id_type_vagon", OleDbType.Integer);
                        cmd.Parameters["@id_type_vagon"].Value = makeVagonQuerry("Люкс");
                        cmd.Parameters.Add("@total_seats", OleDbType.Integer);
                        cmd.Parameters["@total_seats"].Value = textBox6.Text;
                        cmd.Parameters.Add("@left_seats", OleDbType.Integer);
                        cmd.Parameters["@left_seats"].Value = textBox6.Text;
                        cmd.Parameters.Add("@price", OleDbType.Integer);
                        cmd.Parameters["@price"].Value = textBox3.Text;

                        cmd.ExecuteNonQuery();
                    }

                    
                }
            }
        }

        private bool checkAllFilled()
        {
            if ((comboBox3.SelectedIndex != -1 && comboBox4.SelectedIndex != -1 && textBox3.Text != "") && (textBox4.Text != "" || textBox5.Text != "" || textBox6.Text != ""))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void clearAll()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            checkedListBox1.SetItemChecked(0, false);
            checkedListBox1.SetItemChecked(1, false);
            checkedListBox1.SetItemChecked(2, false);
        }
    }
}
