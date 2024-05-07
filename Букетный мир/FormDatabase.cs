using System;
using System.Collections.Generic;
using System.ComponentModel;
using SD = System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Букетный_мир
{
    public partial class FormDatabase : Form
    {
        FormRef fRef;
        FormAccounts fAccounts;
        
        FormAddProducts fAddProducts;
        FormAddProviders fAddProviders;
        FormAddSale fAddSale;
        FormAddSupplies fAddSupplies;

        FormChangeProducts fChangeProducts;
        FormChangeProviders fChangeProviders;

        public int kodProd;
        public int kodProv;

        public FormDatabase()
        {
            InitializeComponent();
        }

        private void FormDatabase_Load(object sender, EventArgs e)
        {
            this.поставки_ЗапросTableAdapter.Fill(this.цветочныйМагазинDataSet1.Поставки_Запрос);
            this.поставщикиTableAdapter.Fill(this.цветочныйМагазинDataSet1.Поставщики);
            this.продажи_ЗапросTableAdapter.Fill(this.цветочныйМагазинDataSet1.Продажи_Запрос);
            this.товарыTableAdapter.Fill(this.цветочныйМагазинDataSet.Товары);

            dataGridViewProducts.Columns[0].Visible = false;
            dataGridViewProviders.Columns[0].Visible = false;
            dataGridViewSale.Columns[0].Visible = false;
            dataGridViewDelivery.Columns[0].Visible = false;

            if (Account.isadmin == true)
            {
                labelName.Text = "Администратор";          
            }
            else
            {
                labelName.Text = "Пользователь";
                пользовательToolStripMenuItem.Enabled = false;

                button3.Enabled = false;
                button2.Enabled = false;
                button1.Enabled = false;
                button6.Enabled = false;

                button10.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button7.Enabled = false;
            }
        }

        //метод обновления таблиц программы

        public void UpdateDataGrid()
        {
            this.поставки_ЗапросTableAdapter.Fill(this.цветочныйМагазинDataSet1.Поставки_Запрос);
            this.поставщикиTableAdapter.Fill(this.цветочныйМагазинDataSet1.Поставщики);
            this.продажи_ЗапросTableAdapter.Fill(this.цветочныйМагазинDataSet1.Продажи_Запрос);
            this.товарыTableAdapter.Fill(this.цветочныйМагазинDataSet.Товары);

            dataGridViewProducts.Columns[0].Visible = false;
            dataGridViewProviders.Columns[0].Visible = false;
            dataGridViewSale.Columns[0].Visible = false;
            dataGridViewDelivery.Columns[0].Visible = false;
        }

        private void разработчикToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FormRef f = new FormRef();
            f.ShowDialog();
        }

        private void FormDatabase_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Вы действительно хотите выйти из программы?", "Подтверждение выхода", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    Application.Exit();
                    Environment.Exit(0);
                }
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Пожалуйста, закройте все окна программы", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (fAddProducts == null)
            {
                fAddProducts = new FormAddProducts(this);
            }
            fAddProducts.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (fAddProviders == null)
            {
                fAddProviders = new FormAddProviders(this);
            }
            fAddProviders.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (fAddSale == null)
            {
                fAddSale = new FormAddSale(this);
            }
            fAddSale.Start();
            fAddSale.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (fAddSupplies == null)
            {
                fAddSupplies = new FormAddSupplies(this);
            }
            fAddSupplies.Start();
            fAddSupplies.ShowDialog();
        }

        //переменная для выборки из полей таблиц для изменения/удаления (выбрали - удалили)

        private void dataGridViewProducts_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            kodProd = Convert.ToInt32(dataGridViewProducts[0, e.RowIndex].Value);
        }

        private void dataGridViewProviders_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            kodProv = Convert.ToInt32(dataGridViewProviders[0, e.RowIndex].Value);
        }

        //формы открытия изменения товаров и поставщиков

        private void button2_Click(object sender, EventArgs e)
        {
            if (fChangeProducts == null)
            {
                fChangeProducts = new FormChangeProducts(this);
            }
            fChangeProducts.Start();
            fChangeProducts.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (fChangeProviders == null)
            {
                fChangeProviders = new FormChangeProviders(this);
            }
            fChangeProviders.Start();
            fChangeProviders.ShowDialog();
        }

        //открытие формы-справки о разработчике

        private void разработчикToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fRef = new FormRef();
            fRef.ShowDialog();
        }

        //удаление записей

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите удалить запись с товаром?", "Удаление товара", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"DELETE * FROM [Товары] WHERE [Код товара] = {Convert.ToInt32(dataGridViewProducts[0, dataGridViewProducts.CurrentRow.Index].Value)}", ClassCon.con);
                com1.ExecuteNonQuery();

                ClassCon.CloseCon();

                MessageBox.Show("Запись о товаре успешно удалена");
                UpdateDataGrid();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите удалить запись с поставщиком?", "Удаление поставщика", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                ClassCon.OpenCon();

                OleDbCommand com1 = new OleDbCommand($"DELETE * FROM [Поставщики] WHERE [Код поставщика] = {Convert.ToInt32(dataGridViewProviders[0, dataGridViewProviders.CurrentRow.Index].Value)}", ClassCon.con);
                com1.ExecuteNonQuery();

                ClassCon.CloseCon();

                MessageBox.Show("Запись о поставщике успешно удалена");
                UpdateDataGrid();
            }
        }

        //поиск

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text.ToLower();
            dataGridViewProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewProducts.ClearSelection();
            foreach (DataGridViewRow row in dataGridViewProducts.Rows)
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox4.Text.ToLower();
            dataGridViewSale.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewSale.ClearSelection();
            foreach (DataGridViewRow row in dataGridViewSale.Rows)
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox2.Text.ToLower();
            dataGridViewProviders.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewProviders.ClearSelection();
            foreach (DataGridViewRow row in dataGridViewProviders.Rows)
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox3.Text.ToLower();
            dataGridViewDelivery.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewDelivery.ClearSelection();
            foreach (DataGridViewRow row in dataGridViewDelivery.Rows)
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

        //вывод данных в Excel

        private void ExporttoExcel(DataTable dt, Excel.Worksheet sheet)
        {
            int _excelHeader = 1;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                _excelHeader = 1;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Columns[j].ColumnName == "Код поставщика")
                        continue;
                    if (dt.Columns[j].ColumnName == "Код поставки")
                        continue;
                    if (dt.Columns[j].ColumnName == "Код товара")
                        continue;
                    if (dt.Columns[j].ColumnName == "Код продажи")
                        continue;
                    sheet.Cells[i + 2, _excelHeader] = dt.Rows[i][j];
                    _excelHeader++;
                }
            }
            _excelHeader = 1;
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName == "Код поставщика")
                    continue;
                if (column.ColumnName == "Код поставки")
                    continue;
                if (column.ColumnName == "Код товара")
                    continue;
                if (column.ColumnName == "Код продажи")
                    continue;
                sheet.Cells[1, _excelHeader] = column.ColumnName;
                _excelHeader++;
            }
            sheet.Columns.ColumnWidth = 30;
            sheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }

        private void товарыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClassCon.OpenCon();

            string query = "SELECT * FROM [Товары]";
            OleDbDataAdapter command = new OleDbDataAdapter(query, ClassCon.con);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            command.Fill(dt);

            ClassCon.CloseCon();

            Excel.Application ExceApp = new Excel.Application();

            ExceApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)ExceApp.ActiveSheet;

            worksheet.Name = "Отчет по товарам";

            ExporttoExcel(dt, worksheet);

            ExceApp.Visible = true;
            ExceApp.ActiveWindow.Activate();
        }

        private void поставщикиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClassCon.OpenCon();

            string query = "SELECT * FROM [Поставщики]";
            OleDbDataAdapter command = new OleDbDataAdapter(query, ClassCon.con);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            command.Fill(dt);

            ClassCon.CloseCon();

            Excel.Application ExceApp = new Excel.Application();

            ExceApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)ExceApp.ActiveSheet;

            worksheet.Name = "Отчет по поставщикам";

            ExporttoExcel(dt, worksheet);

            ExceApp.Visible = true;
            ExceApp.ActiveWindow.Activate();
        }

        private void продажиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClassCon.OpenCon();

            string query = "SELECT * FROM [Продажи Запрос]";
            OleDbDataAdapter command = new OleDbDataAdapter(query, ClassCon.con);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            command.Fill(dt);

            ClassCon.CloseCon();

            Excel.Application ExceApp = new Excel.Application();

            ExceApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)ExceApp.ActiveSheet;

            worksheet.Name = "Отчет по продажам товаров";

            ExporttoExcel(dt, worksheet);

            ExceApp.Visible = true;
            ExceApp.ActiveWindow.Activate();
        }

        private void поставкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClassCon.OpenCon();

            string query = "SELECT * FROM [Поставки Запрос]";
            OleDbDataAdapter command = new OleDbDataAdapter(query, ClassCon.con);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            command.Fill(dt);

            ClassCon.CloseCon();

            Excel.Application ExceApp = new Excel.Application();

            ExceApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)ExceApp.ActiveSheet;

            worksheet.Name = "Отчет по поставкам товаров";

            ExporttoExcel(dt, worksheet);

            ExceApp.Visible = true;
            ExceApp.ActiveWindow.Activate();
        }

        private void пользовательToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fAccounts == null)
            {
                fAccounts = new FormAccounts();
            }
            fAccounts.Show();
        }

        private void выйтиИзСистемыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormStart f = new FormStart();
            f.Show();
            Hide();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
            Environment.Exit(0);
        }
    }
}
