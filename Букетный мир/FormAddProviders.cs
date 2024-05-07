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
    public partial class FormAddProviders : Form
    {
        FormDatabase main;
        public FormAddProviders(FormDatabase owner)
        {
            InitializeComponent();
            main = owner;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string ua = textBox2.Text;
                string fio = textBox3.Text;

                ClassCon.OpenCon();

                OleDbCommand com = new OleDbCommand($"INSERT INTO Поставщики (Организация, [Юридический адрес], [ФИО представителя]) VALUES ('{name}', '{ua}', '{fio}')", ClassCon.con);
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
