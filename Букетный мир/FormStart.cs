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
    public partial class FormStart : Form
    {
        public FormStart()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text != null && maskedTextBox2.Text != null)
            {
                try
                {
                    ClassCon.OpenCon();

                    OleDbCommand com1 = new OleDbCommand("SELECT Пароль FROM Пользователь WHERE Пароль = ('" + maskedTextBox2.Text + "')", ClassCon.con);
                    OleDbCommand com2 = new OleDbCommand("SELECT Админ FROM Пользователь WHERE Пароль = ('" + maskedTextBox2.Text + "')", ClassCon.con);
                    OleDbCommand com3 = new OleDbCommand("SELECT Имя FROM Пользователь WHERE Имя = ('" + maskedTextBox1.Text + "')", ClassCon.con);
                    OleDbCommand com4 = new OleDbCommand("SELECT [Код пользователя] FROM Пользователь WHERE Пароль = ('" + maskedTextBox2.Text + "')", ClassCon.con);

                    string query = "SELECT COUNT(*) FROM Пользователь WHERE Имя = '" + maskedTextBox1.Text + "' AND Пароль = '" + maskedTextBox2.Text + "'";
                    OleDbCommand cmd = new OleDbCommand(query, ClassCon.con);
                    int count = Convert.ToInt32(cmd.ExecuteScalar());

                    if (count > 0 && Convert.ToString(com3.ExecuteScalar()) == maskedTextBox1.Text && Convert.ToString(com1.ExecuteScalar()) == maskedTextBox2.Text)
                    {
                        Account.isadmin = Convert.ToBoolean(com2.ExecuteScalar());
                        Account.nameUser = Convert.ToString(com3.ExecuteScalar());
                        Account.codeUser = Convert.ToInt32(com4.ExecuteScalar());
                        MessageBox.Show($"Добро пожаловать, {Account.nameUser}!");
                        Hide();

                        FormDatabase fdatabase = new FormDatabase();
                        fdatabase.Show();
                    }
                    else
                    {
                        MessageBox.Show("Неправильное имя или пароль");
                    }

                    ClassCon.CloseCon();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void FormStart_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
            Environment.Exit(0);
        }

        private void maskedTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click_1(sender, e);
            }
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            maskedTextBox2.PasswordChar = checkBox1.Checked ? '\0' : '*';
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormReset f = new FormReset();
            f.ShowDialog();
        }
    }
}
