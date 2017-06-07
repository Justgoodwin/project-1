using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace diplom
{
    public partial class Form1 : Form
    {
        public bool slope_angleVisible { get; private set; }

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            comboBox1.Items.AddRange(new string[] {"Всё включено", "Выключен монитор", "Всё выключено" });
        }
        private void Take()
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                FileStream stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);

                Excel.IExcelDataReader IEDR;

                int fileformat = ofd.SafeFileName.IndexOf(".xlsx");

                if (fileformat > -1)
                {
                    //2007 format *.xlsx
                    IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else
                {
                    //97-2003 format *.xls
                    IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                }

                //Если данное значение установлено в true
                //то первая строка используется в качестве 
                //заголовков для колонок
                IEDR.IsFirstRowAsColumnNames = true;

                DataSet ds = IEDR.AsDataSet();

                //Устанавливаем в качестве источника данных dataset 
                //с указанием номера таблицы. Номер таблицы указавает 
                //на соответствующий лист в файле нумерация листов 
                //начинается с нуля.
               
                
                dataGridView1.DataSource = ds.Tables[0];
                IEDR.Close();
                
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void print()
        {
                   
            
            
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    int x = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                    int y = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);              
                    chart1.Series[0].Points.AddXY(x, y);

                }
           
        }
   


    private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Take();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            print();
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {            
            string selectedState = comboBox1.SelectedItem.ToString();            
        }
    }
}

