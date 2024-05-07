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
    public partial class FormAddProducts : Form
    {
        FormDatabase main;
        public FormAddProducts(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                int kolvo = Convert.ToInt32(textBox2.Text);
                int cost = Convert.ToInt32(textBox3.Text);

                ClassCon.OpenCon();

                OleDbCommand com = new OleDbCommand($"INSERT INTO Товары (Название, Количество, Цена) VALUES ('{name}', {kolvo}, {cost})", ClassCon.con);
                com.ExecuteNonQuery();
                MessageBox.Show("Запись успешно добавлена");

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
