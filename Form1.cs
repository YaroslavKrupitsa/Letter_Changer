using Letter_Changer.PractikDataSetTableAdapters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Letter_Changer
{
    public partial class Form1 : Form
    {
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\MainDataBase.accdb;";
        private OleDbConnection myConnection;

        public Form1()
        {
            myConnection = new OleDbConnection(connectString);
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "mainDataBaseDataSet1.Decans". При необходимости она может быть перемещена или удалена.
            this.decansTableAdapter.Fill(this.mainDataBaseDataSet1.Decans);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "mainDataBaseDataSet.Organisations". При необходимости она может быть перемещена или удалена.
            this.organisationsTableAdapter1.Fill(this.mainDataBaseDataSet.Organisations);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "organisationBaseDataSet.Organisations". При необходимости она может быть перемещена или удалена.
            this.organisationsTableAdapter.Fill(this.organisationBaseDataSet.Organisations);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "practikDataSet.Таблица1". При необходимости она может быть перемещена или удалена.
            this.таблица1TableAdapter.Fill(this.practikDataSet.Таблица1);

        }
        string nameOfFile;
        private void button1_Click(object sender, EventArgs e)
        {
            string path = null;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
            }
            var helper = new WordHelper("Лист_щодо_моніторингу.docx");
            var items = new Dictionary<string, string>
            {
                {"<TYPE>", Work_type_box.Text},
                {"<ORG>", Work_name_box.Text},
                {"<NAME>", Work_director_box.Text},
                {"<FULLNAME>", Work_fullname_box.Text},
            };

            helper.Process(items, path, nameOfFile);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int selected_row = int.Parse(dataGridView1.CurrentRow.Index.ToString());

            Work_name_box.Text = dataGridView1.Rows[selected_row].Cells[1].Value.ToString();
            Work_type_box.Text = dataGridView1.Rows[selected_row].Cells[2].Value.ToString();
            Work_director_box.Text = dataGridView1.Rows[selected_row].Cells[3].Value.ToString();
            Work_fullname_box.Text = dataGridView1.Rows[selected_row].Cells[7].Value.ToString();
            nameOfFile = Work_name_box.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            myConnection.Open();
            string search = Search_name_box.Text;
            string query = "SELECT * FROM Organisations WHERE Work_Name LIKE '%"+search+"%';";

            DataTable dt = new DataTable();

            using (OleDbCommand command = new OleDbCommand(query, myConnection))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;

            myConnection.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            myConnection.Open();
            string query = "SELECT * FROM Organisations;";

            DataTable dt = new DataTable();

            using (OleDbCommand command = new OleDbCommand(query, myConnection))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(dt);
                }
            }

            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = dt;

            myConnection.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int selected_row = int.Parse(dataGridView1.CurrentRow.Index.ToString());

            Work_name_box.Text = dataGridView1.Rows[selected_row].Cells[0].Value.ToString();
            Work_type_box.Text = dataGridView1.Rows[selected_row].Cells[1].Value.ToString();
            Work_director_box.Text = dataGridView1.Rows[selected_row].Cells[2].Value.ToString();
            Work_fullname_box.Text = dataGridView1.Rows[selected_row].Cells[6].Value.ToString();
            nameOfFile = Work_name_box.Text;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int selected_row = int.Parse(dataGridView2.CurrentRow.Index.ToString());

            OrgBox.Text = dataGridView2.Rows[selected_row].Cells[0].Value.ToString();
            TypeBox.Text = dataGridView2.Rows[selected_row].Cells[1].Value.ToString();
            DirBox.Text = dataGridView2.Rows[selected_row].Cells[2].Value.ToString();
            AdressBox.Text = dataGridView2.Rows[selected_row].Cells[3].Value.ToString();
            EmailBox.Text = dataGridView2.Rows[selected_row].Cells[4].Value.ToString();
            CodeBox.Text = dataGridView2.Rows[selected_row].Cells[5].Value.ToString();
            PhoneBox.Text = dataGridView2.Rows[selected_row].Cells[7].Value.ToString();
            nameOfFile = OrgBox.Text;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            myConnection.Open();
            string search = SOrgBox.Text;
            string query = "SELECT * FROM Organisations WHERE Work_Name LIKE '%" + search + "%';";

            DataTable dt = new DataTable();

            using (OleDbCommand command = new OleDbCommand(query, myConnection))
            {
                command.Parameters.AddWithValue("@Name", SOrgBox.Text);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(dt);
                }
            }

            dataGridView2.DataSource = dt;

            myConnection.Close();
        }
        string sp = null;
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            myConnection.Open();
            string query = "SELECT * FROM Specialities WHERE Number = @Number;";

            DataTable dt = new DataTable();

            using (OleDbCommand command = new OleDbCommand(query, myConnection))
            {
                command.Parameters.AddWithValue("@Number", textBox2.Text);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(dt);
                   
                    comboBox1.DisplayMember = "Education_program";
                    comboBox1.ValueMember = "Код";
                    comboBox1.DataSource = dt;
                    if (dt.Rows.Count > 0)
                    {
                        sp = dt.Rows[0][2].ToString();
                    }
                }
            }

            myConnection.Close();
        }
        public bool CheckEnteredData()
        {
            if (textBox2.Text == "" || comboBox1.Text == "") 
            {
                MessageBox.Show("Не всі потрібні дані введено");
                return false;
            }
            if (ComKurs1.Text == "" || ComPrak1.Text == "")
            {
                MessageBox.Show("Не всі потрібні дані введено");
                return false;
            }

            return true;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string path = null;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
            }
            var helper = new WordHelper("Договір_на_практику.docx");
            var items = new Dictionary<string, string>
            {
                {"<TYPE>", TypeBox.Text},
                {"<ORG>", OrgBox.Text},
                {"<NAME>", DirBox.Text},
                {"<CODE>", CodeBox.Text},
                {"<ADRESS>", AdressBox.Text},
                {"<EMAIL>", EmailBox.Text},
                {"<NM>", textBox2.Text},
                {"<SPNM>", sp},
                {"<SP2>", comboBox1.Text},
                {"<K1>", ComKurs1.Text},
                {"<K2>", ComKurs2.Text},
                {"<K3>", ComKurs3.Text},
                {"<K4>", ComKurs4.Text},
                {"<K5>", ComKurs5.Text},
                {"<K6>", ComKurs6.Text},
                {"<TYPEPRACRIK1>", ComPrak1.Text},
                {"<TYPEPRACRIK2>", ComPrak2.Text},
                {"<TYPEPRACRIK3>", ComPrak3.Text},
                {"<TYPEPRACRIK4>", ComPrak4.Text},
                {"<TYPEPRACRIK5>", ComPrak5.Text},
                {"<TYPEPRACRIK6>", ComPrak6.Text},
                {"<S11>", Num11.Text},
                {"<S12>", "-"+Num12.Text},
                {"<S21>", Num21.Text},
                {"<S22>", "-"+Num22.Text},
                {"<S31>", Num31.Text},
                {"<S32>", "-"+Num32.Text},
                {"<S41>", Num41.Text},
                {"<S42>", "-"+Num42.Text},
                {"<S51>", Num51.Text},
                {"<S52>", "-"+Num52.Text},
                {"<S61>", Num61.Text},
                {"<S62>", "-"+Num62.Text},
                {"<PHONE>", PhoneBox.Text},
            };

            helper.Process(items, path, nameOfFile);
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox14.SelectedIndex)
            { 
                case 0: 
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
                    break;
                case 1:
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
                    ComKurs2.Visible = true;
                    ComPrak2.Visible = true;
                    Num21.Visible = true;
                    Num21.Text = "";
                    Num22.Visible = true;
                    Num22.Text = "";
                    labelNum2.Visible = true;
                    break;
                case 2:
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
                    ComKurs2.Visible = true;
                    ComPrak2.Visible = true;
                    Num21.Visible = true;
                    Num21.Text = "";
                    Num22.Visible = true;
                    Num22.Text = "";
                    labelNum2.Visible = true;
                    ComKurs3.Visible = true;
                    ComPrak3.Visible = true;
                    Num31.Visible = true;
                    Num31.Text = "";
                    Num32.Visible = true;
                    Num32.Text = "";
                    labelNum3.Visible = true;
                    break;
                case 3:
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
//----------------------------------------------------
                    ComKurs2.Visible = true;
                    ComPrak2.Visible = true;
                    Num21.Visible = true;
                    Num21.Text = "";
                    Num22.Visible = true;
                    Num21.Text = "";
                    labelNum2.Visible = true;
//----------------------------------------------------
                    ComKurs3.Visible = true;
                    ComPrak3.Visible = true;
                    Num31.Visible = true;
                    Num31.Text = "";
                    Num32.Visible = true;
                    Num32.Text = "";
                    labelNum3.Visible = true;
 //----------------------------------------------------
                    ComKurs4.Visible = true;
                    ComPrak4.Visible = true;
                    Num41.Visible = true;
                    Num41.Text = "";
                    Num42.Visible = true;
                    Num42.Text = "";
                    labelNum4.Visible = true;
                    break;
                case 4:
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
//----------------------------------------------------
                    ComKurs2.Visible = true;
                    ComPrak2.Visible = true;
                    Num21.Visible = true;
                    Num21.Text = "";
                    Num22.Visible = true;
                    Num22.Text = "";
                    labelNum2.Visible = true;
//----------------------------------------------------
                    ComKurs3.Visible = true;
                    ComPrak3.Visible = true;
                    Num31.Visible = true;
                    Num31.Text = "";
                    Num32.Visible = true;
                    Num32.Text = "";
                    labelNum3.Visible = true;
//----------------------------------------------------
                    ComKurs4.Visible = true;
                    ComPrak4.Visible = true;
                    Num41.Visible = true;
                    Num41.Text = "";
                    Num42.Visible = true;
                    Num42.Text = "";
                    labelNum4.Visible = true;
//----------------------------------------------------
                    ComKurs5.Visible = true;
                    ComPrak5.Visible = true;
                    Num51.Visible = true;
                    Num51.Text = "";
                    Num52.Visible = true;
                    Num52.Text = "";
                    labelNum5.Visible = true;
                    break;
                case 5:
                    ComKurs1.Visible = true;
                    ComPrak1.Visible = true;
                    Num11.Visible = true;
                    Num12.Visible = true;
                    labelNum1.Visible = true;
//----------------------------------------------------
                    ComKurs2.Visible = true;
                    ComPrak2.Visible = true;
                    Num21.Visible = true;
                    Num21.Text = "";
                    Num22.Visible = true;
                    Num22.Text = "";
                    labelNum2.Visible = true;
//----------------------------------------------------
                    ComKurs3.Visible = true;
                    ComPrak3.Visible = true;
                    Num31.Visible = true;
                    Num31.Text = "";
                    Num32.Visible = true;
                    Num32.Text = "";
                    labelNum3.Visible = true;
//----------------------------------------------------
                    ComKurs4.Visible = true;
                    ComPrak4.Visible = true;
                    Num41.Visible = true;
                    Num41.Text = "";
                    Num42.Visible = true;
                    Num42.Text = "";
                    labelNum4.Visible = true;
//----------------------------------------------------
                    ComKurs5.Visible = true;
                    ComPrak5.Visible = true;
                    Num51.Visible = true;
                    Num51.Text = "";
                    Num52.Visible = true;
                    Num52.Text = "";
                    labelNum6.Visible = true;
//----------------------------------------------------
                    ComKurs6.Visible = true;
                    ComPrak6.Visible = true;
                    Num61.Visible = true;
                    Num61.Text = "";
                    Num62.Visible = true;
                    Num62.Text = "";
                    labelNum6.Visible = true;
                    break;
            }
            switch (comboBox14.SelectedIndex)
            {
                case 0:
                    //----------------------------------------------------
                    ComKurs2.Visible = false;
                    ComKurs2.Text = "";
                    ComPrak2.Visible = false;
                    ComPrak2.Text = "";
                    Num21.Visible = false;
                    Num21.Text = "";
                    Num22.Visible = false;
                    Num22.Text = "";
                    labelNum2.Visible = false;
                    //----------------------------------------------------
                    ComKurs3.Visible = false;
                    ComKurs3.Text = "";
                    ComPrak3.Visible = false;
                    ComPrak3.Text = "";
                    Num31.Visible = false;
                    Num31.Text = "";
                    Num32.Visible = false;
                    Num32.Text = "";
                    labelNum3.Visible = false;
                    //----------------------------------------------------
                    ComKurs4.Visible = false;
                    ComKurs4.Text = "";
                    ComPrak4.Visible = false;
                    ComPrak4.Text = "";
                    Num41.Visible = false;
                    Num41.Text = "";
                    Num42.Visible = false;
                    Num42.Text = "";
                    labelNum4.Visible = false;
                    //----------------------------------------------------
                    ComKurs5.Visible = false;
                    ComKurs5.Text = "";
                    ComPrak5.Visible = false;
                    ComPrak5.Text = "";
                    Num51.Visible = false;
                    Num51.Text = "";
                    Num52.Visible = false;
                    Num52.Text = "";
                    labelNum6.Visible = false;
                    //----------------------------------------------------
                    ComKurs6.Visible = false;
                    ComKurs6.Text = "";
                    ComPrak6.Visible = false;
                    ComPrak6.Text = "";
                    Num61.Visible = false;
                    Num61.Text = "";
                    Num62.Visible = false;
                    Num62.Text = "";
                    labelNum6.Visible = false;
                    break;
                case 1:                    
                    ComKurs3.Visible = false;
                    ComKurs3.Text = "";
                    ComPrak3.Visible = false;
                    ComPrak3.Text = "";
                    Num31.Visible = false;
                    Num31.Text = "";
                    Num32.Visible = false;
                    Num32.Text = "";
                    labelNum3.Visible = false;
                    //----------------------------------------------------
                    ComKurs4.Visible = false;
                    ComKurs4.Text = "";
                    ComPrak4.Visible = false;
                    ComPrak4.Text = "";
                    Num41.Visible = false;
                    Num41.Text = "";
                    Num42.Visible = false;
                    Num42.Text = "";
                    labelNum4.Visible = false;
                    //----------------------------------------------------
                    ComKurs5.Visible = false;
                    ComKurs5.Text = "";
                    ComPrak5.Visible = false;
                    ComPrak5.Text = "";
                    Num51.Visible = false;
                    Num51.Text = "";
                    Num52.Visible = false;
                    Num52.Text = "";
                    labelNum5.Visible = false;
                    //----------------------------------------------------
                    ComKurs6.Visible = false;
                    ComKurs6.Text = "";
                    ComPrak6.Visible = false;
                    ComPrak6.Text = "";
                    Num61.Visible = false;
                    Num61.Text = "";
                    Num62.Visible = false;
                    Num62.Text = "";
                    labelNum6.Visible = false;
                    break;
                case 2:
                    ComKurs4.Visible = false;
                    ComKurs4.Text = "";
                    ComPrak4.Visible = false;
                    ComPrak4.Text = "";
                    Num41.Visible = false;
                    Num41.Text = "";
                    Num42.Visible = false;
                    Num42.Text = "";
                    labelNum4.Visible = false;
                    //----------------------------------------------------
                    ComKurs5.Visible = false;
                    ComKurs5.Text = "";
                    ComPrak5.Visible = false;
                    ComPrak5.Text = "";
                    Num51.Visible = false;
                    Num51.Text = "";
                    Num52.Visible = false;
                    Num52.Text = "";
                    labelNum5.Visible = false;
                    //----------------------------------------------------
                    ComKurs6.Visible = false;
                    ComKurs6.Text = "";
                    ComPrak6.Visible = false;
                    ComPrak6.Text = "";
                    Num61.Visible = false;
                    Num61.Text = "";
                    Num62.Visible = false;
                    Num62.Text = "";
                    labelNum6.Visible = false;
                    break;
                case 3:
                    ComKurs5.Visible = false;
                    ComKurs5.Text = "";
                    ComPrak5.Visible = false;
                    ComPrak5.Text = "";
                    Num51.Visible = false;
                    Num51.Text = "";
                    Num52.Visible = false;
                    Num52.Text = "";
                    labelNum5.Visible = false;
                    //----------------------------------------------------
                    ComKurs6.Visible = false;
                    ComKurs6.Text = "";
                    ComPrak6.Visible = false;
                    ComPrak6.Text = "";
                    Num61.Visible = false;
                    Num61.Text = "";
                    Num62.Visible = false;
                    Num62.Text = "";
                    labelNum6.Visible = false;
                    break;
                case 4:
                    ComKurs6.Visible = false;
                    ComKurs6.Text = "";
                    ComPrak6.Visible = false;
                    ComPrak6.Text = "";
                    Num61.Visible = false;
                    Num61.Text = "";
                    Num62.Visible = false;
                    Num62.Text = "";
                    labelNum6.Visible = false;
                    break;
                case 5:                    
                    break;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string path = null;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
            }
            var helper = new WordHelper("Лист_запрошення.docx");
            var items = new Dictionary<string, string>
            {
                {"<TYPE>", Work_type_box.Text},
                {"<ORG>", Work_name_box.Text},
                {"<NAME>", Work_director_box.Text},
                {"<FULLNAME>", Work_fullname_box.Text},
            };

            helper.Process(items, path, nameOfFile);
        }

        private void ComPrak1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int selected_row = int.Parse(dataGridView3.CurrentRow.Index.ToString());
            ZavBox.Text = dataGridView3.Rows[selected_row].Cells[0].Value.ToString();
            KafBox.Text = dataGridView3.Rows[selected_row].Cells[1].Value.ToString();
            nameOfFile = ZavBox.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string path = null;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
            }
            var helper = new WordHelper("Лист_до_завкафедри.docx");
            var items = new Dictionary<string, string>
            {
                {"<KAFEDRA>", KafBox.Text},
                {"<ZAVIDUYUCHIY>", ZavBox.Text}, 
            };

            helper.Process(items, path, nameOfFile);
        }

        private void button9_Click(object sender, EventArgs e)
        {
       
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
