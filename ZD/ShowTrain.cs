using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Data.OleDb;

namespace ZD
{
    public partial class ShowTrain : Form
    {
        private string result = "";
        private string finalQuerry, querryInsert, updateQuerry;
        private string idTrain, vagonType, seat, vagonName;

        public static string connection_string = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=zd.mdb;";

        string standart_querry = "SELECT TRAIN.ID_TRAIN, TRAIN.NAME_TRAIN, STATION.NAME_STATION, STATION_1.NAME_STATION, TRAIN.DEP_DATE, TRAIN.ARR_DATE, " +
                    "TYPE_VAGON.TYPE_VAGON, VAGON_TRAIN.LEFT_SEATS, VAGON_TRAIN.PRICE " +
                    "FROM STATION AS STATION_1 " +
                    "INNER JOIN(TYPE_VAGON INNER JOIN ((STATION INNER JOIN TRAIN ON STATION.ID_STATION = TRAIN.FROM_STATION) " +
                    "INNER JOIN VAGON_TRAIN ON TRAIN.ID_TRAIN = VAGON_TRAIN.ID_TRAIN) ON TYPE_VAGON.ID_VAGON_TYPE = VAGON_TRAIN.ID_TYPE_VAGON) " +
                    "ON STATION_1.ID_STATION = TRAIN.TO_STATION";


        public ShowTrain()
        {
            InitializeComponent();
        }

        public ShowTrain(string idTrain, string vagonType, string vagonName)
        {
            InitializeComponent();
            this.vagonType = vagonType;
            this.idTrain = idTrain;
            this.vagonName = vagonName;
            MessageBox.Show(vagonName);
        }

        private void fillLabels()
        {
            finalQuerry = standart_querry + " WHERE TRAIN.ID_TRAIN = " + idTrain + " AND TYPE_VAGON.TYPE_VAGON = '" + vagonName + "'";

            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                result = "";

                OleDbCommand cmd = new OleDbCommand(finalQuerry, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    setLabeltext(label1, reader[1].ToString());
                    setLabeltext(label2, reader[2].ToString());
                    setLabeltext(label3, reader[3].ToString());
                    setLabeltext(label4, reader[4].ToString());
                    setLabeltext(label5, reader[5].ToString());
                    setLabeltext(label6, reader[6].ToString());
                    setLabeltext(label7, reader[7].ToString());
                    setLabeltext(label8, reader[8].ToString());
                    setLabeltext(label9, (Convert.ToDateTime(reader["ARR_DATE"]) - Convert.ToDateTime(reader["DEP_DATE"])).ToString());
                }
            }
        }

        private void makeQuerry()
        {
            finalQuerry = standart_querry + " WHERE TRAIN.ID_TRAIN = " + idTrain + " AND TYPE_VAGON.TYPE_VAGON = '" + vagonName + "'";

            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                result = "";
                seat = "";

                OleDbCommand cmd = new OleDbCommand(finalQuerry, conn);
                OleDbDataReader reader = cmd.ExecuteReader();
                //int i = 0;

                while (reader.Read())
                {
                    int delRez_ost = Convert.ToInt32(reader[7]) % 4;
                    string delrez = (Convert.ToInt32(reader[7]) / 4).ToString();

                    if (delRez_ost == 0)
                    {
                        seat = delrez + "A";
                    }
                    else if (delRez_ost == 1)
                    {
                        seat = delrez + "B";
                    }
                    else if (delRez_ost == 2)
                    {
                        seat = delrez + "C";
                    }
                    else if (delRez_ost == 3)
                    {
                        seat = delrez + "D";
                    }

                    result += "\n\n\t\t" + getLabeltext(label1);
                    result += "\n\n\t\t" + getLabeltext(label2);
                    result += "\n\n\t\t" + getLabeltext(label3);
                    result += "\n\n\t\t" + getLabeltext(label4);
                    result += "\n\n\t\t" + getLabeltext(label5);
                    result += "\n\n\t\t" + getLabeltext(label6);
                    result += "\n\n\t\t" + "Место: " + seat;
                    result += "\n\n\t\t" + getLabeltext(label8);
                    result += "\n\n\t\t" + getLabeltext(label9);

                }
            }
        }

        private void setLabeltext(Label label, string text)
        {
            label.Text += " " + text;
        }

        private string getLabeltext(Label label)
        {
            return label.Text;
        }

        private void makeUpdate()
        {
            updateQuerry = standart_querry + " WHERE TRAIN.ID_TRAIN = " + idTrain;

            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();
                querryInsert = "UPDATE VAGON_TRAIN SET LEFT_SEATS = LEFT_SEATS - 1 WHERE ID_TRAIN = " + idTrain + " AND ID_TYPE_VAGON = " + makeVagonQuerry(vagonType);

                OleDbCommand cmd = new OleDbCommand(querryInsert, conn);

                cmd.ExecuteNonQuery();

            }

        }

        private void ShowTrain_Load(object sender, EventArgs e)
        {
            fillLabels();
        }

        private void printingSMTH()
        {
            // объект для печати
            PrintDocument printDocument = new PrintDocument();

            // обработчик события печати
            printDocument.PrintPage += PrintPageHandler;

            // диалог настройки печати
            PrintDialog printDialog = new PrintDialog();

            // установка объекта печати для его настройки
            printDialog.Document = printDocument;

            // если в диалоге было нажато ОК
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                printDialog.Document.Print(); // печатаем
                makeUpdate();
            }
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

        private void PrintPageHandler(object sender, PrintPageEventArgs e)
        {
            // печать строки result
            e.Graphics.DrawString(result, new Font("Times New Roman", 14), Brushes.Black, 0, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            makeQuerry();
            printingSMTH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
