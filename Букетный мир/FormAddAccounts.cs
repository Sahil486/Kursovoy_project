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
    public partial class FormAddAccounts : Form
    {
        FormAccounts main;
        public FormAddAccounts(FormAccounts owner)
        {
            InitializeComponent();
            main = owner;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string password = textBox2.Text;
                bool admin = Convert.ToBoolean(checkBox1.Checked);

                ClassCon.OpenCon();

                OleDbCommand com = new OleDbCommand($"INSERT INTO Пользователь (Имя, Пароль, Админ) VALUES ('{name}', '{password}', {admin})", ClassCon.con);
                com.ExecuteNonQuery();
                MessageBox.Show("Пользователь успешно добавлен");

                ClassCon.CloseCon();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            main.UpdateDateGridAccounts();
            Hide();
        }

        private void FormAddAccounts_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }
    }
}
