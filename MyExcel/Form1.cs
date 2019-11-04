using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MyExcel
{
    public partial class Form1 : Form
    {
        static private int rowCount = 29;
        static private int columnCount = 15;
        private Parser p = new Parser();
        private Based26Sys name = new Based26Sys();
        string firstCell;
        public Form1()
        {
            InitializeComponent();
            InitializeDataGridView(columnCount, rowCount);
            dataGridView.CellStateChanged += DataGridViewCellStateChangedEventHandler;
            dataGridView.CellEndEdit += DataGridViewCellEndEditEventHandler;
            dataGridView.CellBeginEdit += DataGridViewCellBeginEditEventHandler;
            this.Closing += new System.ComponentModel.CancelEventHandler(this.FormClosingEventHandler);
        }
        int recurr = 0;
        private void ButtonClickEventHandler(object sender, EventArgs e)
        {
        }

        private void InitializeDataGridView(int columnCount, int rowCount)
        {
            dataGridView.ColumnCount = columnCount;
            dataGridView.RowCount = rowCount;

           

            for (int i = 0; i < columnCount; ++i)
            {
                dataGridView.Columns[i].Name = name.To26Sys(i);
            }

            for (int i = 0; i < rowCount; ++i)
            {
                dataGridView.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }

            dataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView.RowHeadersVisible = true;
            dataGridView.RowHeadersWidth = 55;
            textBox.ReadOnly = true;
            dataGridView.AllowUserToAddRows = false;
        }

        private void Refresh(string nameCell, int road)
        {
            
            if (road == 0)
            {
                firstCell = nameCell;
            }
            List<String> dependentList;
            if (p.GetDependent(nameCell) != null)
            {
                dependentList = p.GetDependent(nameCell);
            }
            else
                dependentList = null;

            int i = 0;

            string columnHeader = null;
            string rowHeader = null;
            int currentRow = -1;
            int currentColumn = -1;

            while (Char.IsLetter(nameCell[i]))
            {

                columnHeader += nameCell[i];
                i++;
            }
            while (i < nameCell.Length)
            {
                rowHeader += nameCell[i];
                i++;
            }
            currentColumn = Convert.ToInt32(name.From26Sys(columnHeader));
            currentRow = Convert.ToInt32(rowHeader) - 1;

            if (dataGridView[currentColumn, currentRow].Value != null)
            {
                p.SetCurrentName(nameCell);


                if (road == 0)
                {
                    dataGridView[currentColumn, currentRow].Tag = dataGridView[currentColumn, currentRow].Value;
                }
                textBox.Text = dataGridView[currentColumn, currentRow].Tag.ToString();

                if (road == 0)
                {

                    dataGridView[currentColumn, currentRow].Value = p.Evaluate(dataGridView[currentColumn, currentRow].Value.ToString());
                }
                else
                {
                    dataGridView[currentColumn, currentRow].Value = p.Evaluate(dataGridView[currentColumn, currentRow].Tag.ToString());
                }

                if (dependentList != null)
                {

                    for (int c = 0; c < dependentList.Count; c++)
                    {
                        recurr++;
                       // MessageBox.Show(dependentList[c]);
                        if (recurr > 2000)
                        //if (firstCell == dependentList[c])
                        {
                            MessageBox.Show("Recurr", "MiniExcel");
                            return;
                        }
                        Refresh(dependentList[c], 1);

                    }
                }
            }
            else
            {
                p.SetCurrentName(nameCell);
                p.Evaluate("");
                if (dependentList != null)
                {
                    for (int c = 0; c < dependentList.Count; c++)
                    {

                        Refresh(dependentList[c], 1);

                    }
                }
                dataGridView[currentColumn, currentRow].Tag = null;
            }
        }

        private void DataGridViewCellStateChangedEventHandler(object sender, DataGridViewCellStateChangedEventArgs e)
        {
            try
            {
                if (e.StateChanged.ToString() == "Selected")
                {
                    if (dataGridView[e.Cell.ColumnIndex, e.Cell.RowIndex].Tag != null)
                    {
                        textBox.Text = dataGridView[e.Cell.ColumnIndex, e.Cell.RowIndex].Tag.ToString();
                    }
                    else
                    {
                        textBox.Text = "";
                    }
                }
            }
            catch
            {

            }
        }

        private void DataGridViewCellEndEditEventHandler(object sender, DataGridViewCellEventArgs e)
        {
                Refresh(name.To26Sys(e.ColumnIndex) + (e.RowIndex + 1).ToString(), 0);     
        }

        private void DataGridViewCellBeginEditEventHandler(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (dataGridView[e.ColumnIndex, e.RowIndex].Tag != null)
            {
                dataGridView[e.ColumnIndex, e.RowIndex].Value = dataGridView[e.ColumnIndex, e.RowIndex].Tag;
            }
        }

        private void SaveDataGridView()
        {
           
            if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string fileName = saveFileDialog.FileName;
            FileStream aFile = new FileStream(fileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(aFile);
            aFile.Seek(0, SeekOrigin.End);
            sw.WriteLine(columnCount);
            sw.WriteLine(rowCount);
            for (int i = 0; i < columnCount; i++)
            {
                for (int j = 0; j < rowCount - 2; j++)
                    if (dataGridView[i, j].Value != null)
                    {
                        sw.WriteLine(i);
                        sw.WriteLine(j);
                        sw.WriteLine(dataGridView[i, j].Tag);
                    }
            }
            sw.Close();
            MessageBox.Show("Файл збережено", "MiniExcel");
        }

        private void SaveToolStripMenuItemClickEventHandler(object sender, EventArgs e)
        {
            SaveDataGridView();
        }

        private void OpenToolStripMenuItemClickEventHandler(object sender, EventArgs e)
        {
            OpenDataGridView();
        }

        private void HelpToolStripMenuItemClickEventHandler(object sender, EventArgs e)
        {
            using (StreamReader sr = new StreamReader("ReadMe.txt"))
            {
                MessageBox.Show(sr.ReadToEnd(), "MiniExcel");
            }
        }

        private void ClearDataGridView()
        {
            for (int m = 0; m < columnCount; m++)
            {
                for (int n = 0; n < rowCount - 2; n++)
                {
                    dataGridView[m, n].Value = null;
                    dataGridView[m, n].Tag = null;
                }
            }
        }

        private void OpenDataGridView()
        {
            DialogResult dialogResult = MessageBox.Show("Зберегти поточну таблицю?", "MiniExcel", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                SaveDataGridView();
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                return;
            }

            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
                return;
            ClearDataGridView();
            // получаем выбранный файл
            string fileName = openFileDialog.FileName;
            FileStream aFile = new FileStream(fileName, FileMode.Open);
            StreamReader sw = new StreamReader(aFile);
            string line = "";
            columnCount = Convert.ToInt32(sw.ReadLine());
            rowCount = Convert.ToInt32(sw.ReadLine());
            InitializeDataGridView(columnCount, rowCount);
            int i;
            int j;
            while (true)
            {
                line = sw.ReadLine();
                if (line != null)
                {
                    i = Convert.ToInt32(line);
                }
                else
                {
                    break;
                }
                line = sw.ReadLine();
                if (line != null)
                {
                    j = Convert.ToInt32(line);
                }
                else
                {
                    MessageBox.Show("Invalid file", "MiniExcel");
                    break;
                }
                line = sw.ReadLine();
                if (line != null)
                {
                   dataGridView[i, j].Tag = line;
                    dataGridView[i, j].Value = line;
                    Refresh(name.To26Sys(i) + (j + 1).ToString(), 1);
                }
                else
                {
                    MessageBox.Show("Invalid file", "MiniExcel");
                    break;
                }
            }
            sw.Close();
        }

        private void FormClosingEventHandler(object sender, System.ComponentModel.CancelEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Зберегти таблицю перед завершенням роботи?", "MiniExcel", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                SaveDataGridView();
                e.Cancel = true;
                return;
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            columnCount++;
            dataGridView.Columns.Add(columnCount.ToString(), name.To26Sys(columnCount - 1));
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RemoveColumn();
        }

        private void addToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            rowCount++;
            dataGridView.Rows.Add();

            dataGridView.Rows[rowCount - 2].HeaderCell.Value = (rowCount - 1).ToString();
        }

        private void removeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RemoveRow();
        }

        private void RemoveRow()
        {
            List<String> dependentList;
            for (int i = 0; i < columnCount; i++)
            {
                p.SetCurrentName(name.To26Sys(i) + (rowCount + 1).ToString());
                p.Evaluate("");
                p.removeVariable(name.To26Sys(i) + (rowCount + 1).ToString());

                if (p.GetDependent(name.To26Sys(i) + (rowCount + 1).ToString()) != null)
                {
                    dependentList = p.GetDependent(name.To26Sys(i) + (rowCount + 1).ToString());
                }
                else
                    dependentList = null;
                if (dependentList != null)
                {
                    for (int c = 0; c < dependentList.Count; c++)
                    {

                        Refresh(dependentList[c], 1);

                    }
                }
            }
            dataGridView.AllowUserToAddRows = true;
            dataGridView.Rows.RemoveAt(rowCount- 2);
            rowCount--;
            MessageBox.Show("Не забудьте перевірити зв'зки між комірками в таблиці", "MiniExcel");
            dataGridView.AllowUserToAddRows = false;
        }

        private void RemoveColumn()
        {

            List<String> dependentList;
            for (int i = 0; i < rowCount; i++)
            {
                p.SetCurrentName(name.To26Sys(columnCount - 1) + (i + 1).ToString());
                p.Evaluate("");
                p.removeVariable(name.To26Sys(columnCount - 1) + (i + 1).ToString());
                if (p.GetDependent(name.To26Sys(columnCount - 1) + (i + 1).ToString()) != null)
                {
                    dependentList = p.GetDependent(name.To26Sys(columnCount - 1) + (i + 1).ToString());
                }
                else
                    dependentList = null;
                if (dependentList != null)
                {
                    for (int c = 0; c < dependentList.Count; c++)
                    {

                        Refresh(dependentList[c], 1);

                    }
                }
            }
            dataGridView.Columns.RemoveAt(columnCount - 1);
            columnCount--;
            MessageBox.Show("Не забудьте перевірити зв'зки між комірками в таблиці", "MiniExcel");
        }
    }
}
