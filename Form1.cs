using System;
using System.Data;
using System.Windows.Forms;
using System.IO;


namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        
        public object comboBox;

        public bool Flag { get; private set; }

        private void Take()
        {
            OpenFileDialog ofd = new OpenFileDialog();            
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";


            if (ofd.ShowDialog() == DialogResult.OK)
            {
                
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ofd.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;';";
                System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(constr);
                con.Open();
                DataSet ds = new DataSet();
                DataTable schemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(select, con);
                ad.Fill(ds);
                DataTable tb = ds.Tables[0];
                con.Close();
                dataGridView1.DataSource = tb;
                con.Close();
                                
            }

            else
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }
        

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            comboBox1.Items.AddRange(new string[] { "1"+" метр", "1,5"+" метра", "2"+" метра", "2,5"+" метра", "3"+" метра" });
            comboBox2.Items.AddRange(new string[] { "Всё включено", "Выключен монитор", "Всё выключено"});            
            
        }
        

        private void button1_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                int x = Convert.ToInt32(dataGridView1.Rows[i].Cells[i].Value);
                int y = Convert.ToInt32(dataGridView1.Rows[i].Cells[i].Value);                
                chart1.Series[0].Points.AddXY(x, y);

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedState = comboBox1.SelectedItem.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            Take();            
        }
      
        public int RowCount { get; set; }
    }

    internal class dataGridView1
    {
        public static DataTable DataSource { get; internal set; }


    }

    internal class ExcelReaderFactory
    {
        internal static IExcelDataReader CreateBinaryReader(FileStream stream)
        {
            throw new NotImplementedException();
        }

        internal static IExcelDataReader CreateOpenXmlReader(FileStream stream)
        {
            throw new NotImplementedException();
        }
       
    }

}

   
    internal class openFileDialog1
    {
        internal static DialogResult ShowDialog()
        {
            throw new NotImplementedException();
        }
    }