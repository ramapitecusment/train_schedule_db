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

namespace ZD
{
    public partial class Form1 : Form
    {
        public static string connection_string = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=zd.mdb;";

        List<string> logons = new List<string>();
        string querry, querryInsert;
        string login, password;
        bool isLogged;

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                using (OleDbConnection conn = new OleDbConnection(connection_string))
                {
                    conn.Open();
                    querryInsert = "INSERT INTO LOGIN([LOGIN], [PASSWORD]) VALUES('" + textBox1.Text + "', '" + textBox2.Text + "');";

                    OleDbCommand cmd = new OleDbCommand(querryInsert, conn);

                    cmd.ExecuteNonQuery();
                }
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OleDbConnection conn = new OleDbConnection(connection_string))
            {
                conn.Open();

                querry = "SELECT * FROM LOGIN";

                OleDbCommand command = new OleDbCommand(querry, conn);
                OleDbDataReader reader = command.ExecuteReader();

                isLogged = false;
                while (reader.Read())
                {
                    password = reader[2].ToString();
                    login = reader[1].ToString();

                    logons.Add(login);

                    if (textBox2.Text == password && textBox1.Text == login)
                    {
                        Zd_Scedule z = new Zd_Scedule(login, logons);
                        z.Show();
                        isLogged = true;
                    }

                }
                reader.Close();
            }
            if (!isLogged)
            {
                MessageBox.Show("Логин и парль, которые вы указали, не соответствуют ни одному аккаунту.");
            }
        }
    }
}
