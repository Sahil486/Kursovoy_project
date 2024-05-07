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
    public partial class FormChangeProviders : Form
    {
        FormDatabase main;
        public FormChangeProviders(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        public void Start()
        {
            ClassCon.OpenCon();

            OleDbCommand com1 = new OleDbCommand("SELECT [Организация] FROM [Поставщики] WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);
            OleDbCommand com2 = new OleDbCommand("SELECT [Юридический адрес] FROM [Поставщики] WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);
            OleDbCommand com3 = new OleDbCommand("SELECT [ФИО представителя] FROM [Поставщики] WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);

            textBox1.Text = Convert.ToString(com1.ExecuteScalar());
            textBox2.Text = Convert.ToString(com2.ExecuteScalar());
            textBox3.Text = Convert.ToString(com3.ExecuteScalar());

            ClassCon.CloseCon();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string organization = textBox1.Text;
                string uradress = textBox2.Text;
                string FIO = textBox3.Text;

                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"UPDATE [Поставщики] SET [Организация] = '{organization}' WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);
                OleDbCommand com2 = new OleDbCommand($"UPDATE [Поставщики] SET [Юридический адрес] = '{uradress}' WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);
                OleDbCommand com3 = new OleDbCommand($"UPDATE [Поставщики] SET [ФИО представителя] = '{FIO}' WHERE [Код поставщика] = " + main.kodProv + "", ClassCon.con);

                com1.ExecuteNonQuery();
                com2.ExecuteNonQuery();
                com3.ExecuteNonQuery();

                ClassCon.CloseCon();

                MessageBox.Show("Запись о поставщике успешно изменена");
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
