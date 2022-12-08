using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using NCalc;
namespace EXCEL
{

    public partial class Form1 : Form
    {
        List<List<string>> formulas;
        List<List<int>> values;
        List<List<bool>> good;
        //DataGridViewCell template;
        public Form1()
        {
            formulas = new List<List<string>>();
            values = new List<List<int>>();
            good = new List<List<bool>>();
            //template = dataGridView1.Columns[0].CellTemplate;
            InitializeComponent();
        }

        void calculate()
        {
            const int nll = -874533;
            List<Tuple<string, int>> pars = new List<Tuple<string, int>>();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (formulas[i][j] != "" && dataGridView1[j, i].Value != "CANT EVAL")
                    {
                        try
                        {
                            values[i][j] = Int32.Parse(formulas[i][j].ToString());
                            dataGridView1[j, i].Value = values[i][j].ToString();
                        }
                        catch (Exception ex) { values[i][j] = nll; }
                    }
            for (int tt = 0; tt < dataGridView1.RowCount * dataGridView1.ColumnCount; tt++)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (values[i][j] != nll && formulas[i][j] != "" && dataGridView1[j, i].Value != "CANT EVAL")
                        {
                            try
                            {
                                var s = values[i][j];
                                pars.Add(new Tuple<string, int>(HelperFuncs.cv(j, i + 1), s));
                            }
                            catch (Exception ex) { }
                        }
                bool any = false;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (formulas[i][j] != "" && values[i][j] == nll)
                        {
                            var s = formulas[i][j];
                            try
                            {
                                var exp = new Expression(s);
                                foreach (var t in pars)
                                    exp.Parameters[t.Item1] = t.Item2;
                                var ans = exp.Evaluate();
                                if (HelperFuncs.bad(formulas[i][j], ((char)('A' + j)).ToString() + ((char)('1' + i)).ToString())) { throw new Exception(); }
                                values[i][j] = Int32.Parse(ans.ToString());
                                any = true;
                                dataGridView1[j, i].Value = values[i][j].ToString();
                            }
                            catch (Exception ex)
                            {
                                dataGridView1[j, i].Value = "CANT EVAL";
                                // MessageBox.Show(String.Format("Trash at {0}, {1}", i, j));
                            }

                        }
                if (!any) break;
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        void addRow()
        {
            formulas.Add(new List<string>(Enumerable.Repeat("", dataGridView1.ColumnCount)));
            values.Add(new List<int>(Enumerable.Repeat(0, dataGridView1.ColumnCount)));
            good.Add(new List<bool>(Enumerable.Repeat(false, dataGridView1.ColumnCount)));
            dataGridView1.Rows.Add();
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = dataGridView1.Rows.Count.ToString();
        }
        void addColumn()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                formulas[i].Add("");
                good[i].Add(false);
                values[i].Add(0);
            }
            dataGridView1.Columns.Add(new DataGridViewColumn(dataGridView1.Columns[0].CellTemplate));
            dataGridView1.Columns[dataGridView1.ColumnCount - 1].HeaderCell.Value = HelperFuncs.toA(dataGridView1.ColumnCount - 1);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            addRow();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            addColumn();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < values.Count; i++)
                for (int j = 0; j < values[i].Count; j++)
                    dataGridView1[j, i].Value = values[i][j].ToString();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            int j = e.ColumnIndex;
            if (i < 0 || j < 0) return;
            textBox1.Text = formulas[i][j].ToString();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            int j = e.ColumnIndex;
            try
            {
                formulas[i][j] = dataGridView1[j, i].Value.ToString();
                good[i][j] = false;
                calculate();

            }
            catch (Exception ex)
            {

            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentCell.RowIndex;
            int j = dataGridView1.CurrentCell.ColumnIndex;
            formulas[i][j] = textBox1.Text;
            good[i][j] = false;
            calculate();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var path = string.Empty;
            if (FileFuncs.saveFile(formulas, out path))
                MessageBox.Show("Succesfully creared save file at: " + path);
            else
                MessageBox.Show("Error while saving data");
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

            var lines = FileFuncs.loadFile();
            List<List<string>> retf = new List<List<string>>();
            try
            {
                retf = HelperFuncs.get_formulas(lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while reading from file : " + ex.Message);
                return;
            }


            formulas = new List<List<string>>();
            values = new List<List<int>>();
            good = new List<List<bool>>();
            dataGridView1.Rows.Clear();
            while (dataGridView1.ColumnCount > 4) dataGridView1.Columns.RemoveAt(dataGridView1.ColumnCount - 1);


            for (int i = 0; i < retf[0].Count - 4; i++) addColumn();
            for (int i = 0; i < retf.Count; i++) addRow();

            formulas = retf;

            calculate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кнопки <<Add new row/column>> додають відповідно один рядок, одну колонку. Налаштований інтерфейс, який дозволяє зберігати таблицю до файлу з довільним розширенням і завантажувати звідти таблицю. За нього відповідають кнопки Save to file, Load from file\nТаблиця підтримує числові значення, математичні вирази та посилання на клітини. При введені некорректних значень розрахунки не проводяться.", "Програма - аналог EXCEL");
        }
    }
    static class HelperFuncs
    {
        public static String toA(int x)
        {
            return ((char)('A' + x)).ToString();

        }
        public static String cv(int x, int y)
        {
            return toA(x) + y.ToString();
        }
        public static bool bad(string x, string y)
        {
            return x.Contains(y);
        }
        public static List<List<string>> get_formulas(string line)
        {
            var lines = line.Split('\n');
            var fline = lines[0].Split(' ');
            var n = Int32.Parse(fline[0]);
            var m = Int32.Parse(fline[1]);
            List<List<string>> res = new List<List<string>>();
            for (int i = 0; i < n; i++)
            {
                res.Add(new List<string>());
                var xd = lines[1 + i].Split('\t');
                for (int j = 0; j < m; j++)
                    res[i].Add(xd[j]);
            }
            return res;
        }
    }

    static class FileFuncs
    {
        public static string loadFile()
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            var lines = string.Empty;
            using (OpenFileDialog openfile = new OpenFileDialog())
            {
                openfile.InitialDirectory = "c:\\";
                openfile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openfile.FilterIndex = 2;
                openfile.RestoreDirectory = true;

                if (openfile.ShowDialog() == DialogResult.OK)
                    lines = File.ReadAllText(openfile.FileName);
            }
            return lines;
        }
        public static bool saveFile(List<List<string>> ar, out string path)
        {
            path = string.Empty;
            var filePath = string.Empty;
            using (SaveFileDialog savefile = new SaveFileDialog())
            {
                savefile.InitialDirectory = "c:\\";
                savefile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                savefile.FilterIndex = 2;
                savefile.RestoreDirectory = true;

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    filePath = savefile.FileName;
                    var fileStream = savefile.OpenFile();
                    path = filePath;

                    using (StreamWriter s = new StreamWriter(fileStream))
                    {

                        int n = ar.Count;
                        if (n == 0)
                        {
                            MessageBox.Show("nothing to save");
                            return false;
                        }
                        int m = ar[0].Count;
                        s.WriteLine($"{n} {m}");
                        for (int i = 0; i < n; i++)
                        {
                            for (int j = 0; j < m; j++)
                                s.Write($"{ar[i][j]}\t");
                            s.WriteLine();
                        }
                    }
                }
                else return false;
            }
            return true;
        }
    }

}
