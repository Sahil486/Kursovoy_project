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

namespace Букетный_мир
{
    public partial class FormAddSupplies : Form
    {
        FormDatabase main;
        public FormAddSupplies(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        public void Start()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Поставщики]", ClassCon.con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = dt.Columns[1].ColumnName;
            comboBox1.ValueMember = dt.Columns[0].ColumnName;

            OleDbDataAdapter da1 = new OleDbDataAdapter("SELECT * FROM [Товары]", ClassCon.con);
            DataTable dt1 = new DataTable();
            da1.Fill(dt1);
            comboBox2.DataSource = dt1;
            comboBox2.DisplayMember = dt1.Columns[1].ColumnName;
            comboBox2.ValueMember = dt1.Columns[0].ColumnName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int provider = Convert.ToInt32(comboBox1.SelectedValue);
                int product = Convert.ToInt32(comboBox2.SelectedValue);
                int kolvo = Convert.ToInt32(textBox1.Text);
                string date = dateTimePicker1.Text;

                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"INSERT INTO Поставки ([Код поставщика], [Код товара], Количество, Дата) VALUES ({provider}, {product}, {kolvo}, '{date}')", ClassCon.con);
                OleDbCommand com2 = new OleDbCommand($"SELECT Количество FROM Товары WHERE [Код товара] = {product}", ClassCon.con);
                OleDbCommand com3 = new OleDbCommand($"UPDATE Товары SET Количество = {Convert.ToInt32(com2.ExecuteScalar()) + kolvo} WHERE [Код товара] = {product}", ClassCon.con);
                com1.ExecuteNonQuery();
                com3.ExecuteNonQuery();
                MessageBox.Show("Успешная поставка товара");

                ClassCon.CloseCon();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            main.UpdateDataGrid();
            Hide();
        }
    }
}
