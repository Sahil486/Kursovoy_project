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
    public partial class FormAccounts : Form
    {
        FormAddAccounts fAddAccounts;
        //FormChangeAccounts fChangeAccounts;

        public int kodUser;

        public FormAccounts()
        {
            InitializeComponent();
        }

        private void FormAccounts_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "цветочныйМагазинDataSet.Пользователь". При необходимости она может быть перемещена или удалена.
            this.пользовательTableAdapter.Fill(this.цветочныйМагазинDataSet.Пользователь);
        }

        public void UpdateDateGridAccounts()
        {
            this.пользовательTableAdapter.Fill(this.цветочныйМагазинDataSet.Пользователь);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (fAddAccounts == null)
            {
                fAddAccounts = new FormAddAccounts(this);
            }
            fAddAccounts.ShowDialog();
        }

        private void FormAccounts_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }

        private void dataGridViewAccounts_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            kodUser = Convert.ToInt32(dataGridViewAccounts[0, e.RowIndex].Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите удалить данного пользователя?", "Удаление пользователя", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"DELETE * FROM [Пользователь] WHERE [Код пользователя] = {Convert.ToInt32(dataGridViewAccounts[0, dataGridViewAccounts.CurrentRow.Index].Value)}", ClassCon.con);
                com1.ExecuteNonQuery();

                ClassCon.CloseCon();

                MessageBox.Show("Пользователь успешно удален");
                UpdateDateGridAccounts();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text.ToLower();
            dataGridViewAccounts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewAccounts.ClearSelection();
            foreach (DataGridViewRow row in dataGridViewAccounts.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    string cellValue = cell.Value.ToString().ToLower();
                    if (cellValue.Contains(searchValue))
                    {
                        row.Selected = true;
                        break;
                    }
                }
            }
        }
    }
}
