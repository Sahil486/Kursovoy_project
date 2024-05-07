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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Букетный_мир
{
    public partial class FormAddSale : Form
    {
        FormDatabase main;
        public FormAddSale(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        public void Start()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Товары]", ClassCon.con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = dt.Columns[1].ColumnName;
            comboBox1.ValueMember = dt.Columns[0].ColumnName;

            OleDbDataAdapter da1 = new OleDbDataAdapter($"SELECT * FROM [Пользователь] WHERE [Код пользователя] = {Account.codeUser}", ClassCon.con);
            DataTable dt1 = new DataTable();
            da1.Fill(dt1);
            comboBox2.DataSource = dt1;
            comboBox2.DisplayMember = dt1.Columns[1].ColumnName;
            comboBox2.ValueMember = dt1.Columns[0].ColumnName;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                int product = Convert.ToInt32(comboBox1.SelectedValue);
                int kolvo = Convert.ToInt32(textBox1.Text);
                int saler = Convert.ToInt32(comboBox2.SelectedValue);
                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"INSERT INTO Продажи ([Код товара], Количество, [Код пользователя]) VALUES ({product}, {kolvo}, {saler})", ClassCon.con);
                OleDbCommand com2 = new OleDbCommand($"SELECT Количество FROM Товары WHERE [Код товара] = {product}", ClassCon.con);
                OleDbCommand com3 = new OleDbCommand($"UPDATE Товары SET Количество = {Convert.ToInt32(com2.ExecuteScalar()) - kolvo} WHERE [Код товара] = {product}", ClassCon.con);

                com1.ExecuteNonQuery();
                com3.ExecuteNonQuery();
                MessageBox.Show("Успешная продажа товара");

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
