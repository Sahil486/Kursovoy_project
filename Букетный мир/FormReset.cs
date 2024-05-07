using Microsoft.Office.Interop.Excel;
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
    public partial class FormReset : Form
    {
        public FormReset()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string kodslovo = textBox2.Text;
                string password1 = textBox3.Text;
                string password2 = textBox4.Text;

                ClassCon.OpenCon();

                if (textBox3.Text == textBox4.Text && textBox3.Text != "")
                {
                    OleDbCommand com1 = new OleDbCommand($"SELECT [Кодовое слово] FROM Пользователь WHERE Имя = '{name}'", ClassCon.con);
                    com1.ExecuteScalar();

                    string slovo = Convert.ToString(com1.ExecuteScalar());

                    if (slovo == kodslovo)
                    {
                        DialogResult result = MessageBox.Show("Вы действительно хотите изменить пароль данного пользователя?", "Изменение пароля", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            OleDbCommand com3 = new OleDbCommand($"UPDATE [Пользователь] SET [Пароль] = '{password1}' WHERE [Имя] = '{name}' AND [Кодовое слово] = '{kodslovo}'", ClassCon.con);
                            com3.ExecuteNonQuery();

                            MessageBox.Show("Пароль успешно изменен");
                            Hide();
                        }
                        if (result == DialogResult.No)
                        {
                            
                        }
                    }
                    else
                    {
                        MessageBox.Show("Неверное кодовое слово!");
                    }
                }

                ClassCon.CloseCon();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
