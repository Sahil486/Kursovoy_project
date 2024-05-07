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

namespace Букетный_мир
{
    public partial class FormChangeProducts : Form
    {
        FormDatabase main;
        public FormChangeProducts(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        public void Start()
        {
            ClassCon.OpenCon();

            OleDbCommand com1 = new OleDbCommand("SELECT [Название] FROM [Товары] WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);
            OleDbCommand com2 = new OleDbCommand("SELECT [Количество] FROM [Товары] WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);
            OleDbCommand com3 = new OleDbCommand("SELECT [Цена] FROM [Товары] WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);

            textBox1.Text = Convert.ToString(com1.ExecuteScalar());
            textBox2.Text = Convert.ToString(com2.ExecuteScalar());
            textBox3.Text = Convert.ToString(com3.ExecuteScalar());

            ClassCon.CloseCon();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                int kolvo = Convert.ToInt32(textBox2.Text);
                int cost = Convert.ToInt32(textBox3.Text);

                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"UPDATE [Товары] SET [Название] = '{name}' WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);
                OleDbCommand com2 = new OleDbCommand($"UPDATE [Товары] SET [Количество] = {kolvo} WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);
                OleDbCommand com3 = new OleDbCommand($"UPDATE [Товары] SET [Цена] = {cost} WHERE [Код товара] = " + main.kodProd + "", ClassCon.con);

                com1.ExecuteNonQuery();
                com2.ExecuteNonQuery();
                com3.ExecuteNonQuery();

                ClassCon.CloseCon();

                MessageBox.Show("Запись о товаре успешно изменена");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            main.UpdateDataGrid();
            this.Hide();
        }
    }
}
