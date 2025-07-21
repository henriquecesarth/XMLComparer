
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Globalization;
using System.IO;

namespace XMLComparer
{
    public partial class Form1 : Form
    {
        private string path = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Conectando, aguarde.";
            statusStrip1.Refresh();
            LimparTabelas();

            try
            {
                ExibirLabels(true);
                BuscarArquivos(path);

                string query = $"SELECT Modelo, Emissao AS Emissao, NF AS Numero, Serie, TotalNota AS Total " +
                               $"FROM Fiscal.Documento " +
                               $"WHERE CONVERT(DATE, Emissao) BETWEEN '{dateTimePicker1.Value:yyyy-MM-dd}' AND '{dateTimePicker2.Value:yyyy-MM-dd}' " +
                               $"AND NF != '' ORDER BY NF;";

                ConectarBanco(query, consulta: true);
                AvaliarComparacao();
            }
            catch (XmlException ex)
            {
                toolStripStatusLabel1.Text = "Falha.";
                MessageBox.Show($"Falha ao buscar XMLs{ex.Message}");
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = "Falha.";
                MessageBox.Show($"Falha ao tentar conectar{ex.Message}");
            }
        }

        private void ExibirLabels(bool visivel)
        {
            label7.Visible = visivel;
            label8.Visible = visivel;
            label14.Visible = visivel;
            label15.Visible = visivel;
            label16.Visible = visivel;
        }

        private void LimparTabelas()
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
        }

        private void AvaliarComparacao()
        {
            int countXml = dataGridView2.RowCount;
            int countBanco = dataGridView3.RowCount;

            if (countXml == 0)
            {
                MessageBox.Show("Nenhum XML encontrado");
                label12.Text = CalcularTotal(dataGridView3).ToString("N2");
                label18.Text = countBanco.ToString();
                label18.Visible = true;
                return;
            }

            if (countBanco == 0)
            {
                MessageBox.Show("Nenhum documento encontrado no banco de dados nesse período");
                label11.Text = CalcularTotal(dataGridView2).ToString("N2");
                label17.Text = countXml.ToString();
                label17.Visible = true;
                return;
            }

            decimal totalXml = CalcularTotal(dataGridView2);
            decimal totalBanco = CalcularTotal(dataGridView3);

            label11.Text = totalXml.ToString("N2");
            label12.Text = totalBanco.ToString("N2");
            label13.Text = Math.Abs(totalXml - totalBanco).ToString("F2");

            label17.Text = countXml.ToString();
            label18.Text = countBanco.ToString();
            label17.Visible = label18.Visible = true;

            DiferencaBancoXml();
        }

        private decimal CalcularTotal(DataGridView dgv)
        {
            return dgv.Rows.Cast<DataGridViewRow>()
                     .Sum(row => Convert.ToDecimal(row.Cells["TOTAL"].Value, CultureInfo.InvariantCulture));
        }

        private void DiferencaBancoXml()
        {
            int min1 = int.Parse((string)dataGridView2[2, 0].Value);
            int min2 = int.Parse(dataGridView3[2, 0].Value.ToString());
            int min = min1 > min2 ? min2 : min1;

            int max1 = int.Parse((string)dataGridView2[2, dataGridView2.Rows.Count - 1].Value);
            int max2 = int.Parse(dataGridView3[2, dataGridView3.Rows.Count - 1].Value.ToString());
            int max = max1 < max2 ? max2 : max1;

            string[] campo = new string[7];
            List<string> listaDfeXml = new List<string>();
            List<string> listaDfeBanco = new List<string>();

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                listaDfeXml.Add(int.Parse(dataGridView2[2, i].Value.ToString()).ToString());
            }

            for (int i = 0; i < dataGridView3.RowCount; i++)
            {
                listaDfeBanco.Add(int.Parse(dataGridView3[2, i].Value.ToString()).ToString());
            }

            for (int i = 0; i < max; i++)
            {
                campo[4] = "0"; campo[5] = "0";
                campo[2] = (min + i).ToString();

                if (listaDfeXml.Contains(campo[2]))
                {
                    for (int j = 0; j < dataGridView2.RowCount; j++)
                    {
                        if (int.Parse(dataGridView2[2, j].Value.ToString()).ToString() == int.Parse(campo[2]).ToString())
                        {
                            campo[0] = dataGridView2[0, j].Value.ToString();
                            campo[1] = dataGridView2[1, j].Value.ToString();
                            campo[3] = dataGridView2[3, j].Value.ToString();
                            campo[4] = dataGridView2[4, j].Value.ToString();
                            if (listaDfeBanco.Contains(campo[2]))
                            {
                                for (int k = 0; k < dataGridView3.RowCount; k++)
                                {
                                    if (int.Parse(dataGridView3[2, k].Value.ToString()).ToString() == int.Parse(campo[2]).ToString() && int.Parse(dataGridView3[3, k].Value.ToString()).ToString() == int.Parse(campo[3]).ToString())
                                    {
                                        campo[5] = dataGridView3[4, k].Value.ToString();
                                        campo[5] = double.Parse(campo[5]).ToString("f2", CultureInfo.InvariantCulture);
                                        break;
                                    }
                                }
                                campo[6] = (Math.Abs(double.Parse(campo[5], CultureInfo.InvariantCulture) - double.Parse(campo[4], CultureInfo.InvariantCulture))).ToString("f2",CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                campo[5] = "NÃO CONSTA";
                                campo[6] = campo[4];
                            }
                            dataGridView1.Rows.Add(campo);
                            break;
                        }
                    }
                }
                else if (listaDfeBanco.Contains(campo[2]))
                {
                    for (int j = 0; j < dataGridView3.RowCount; j++)
                    {
                        if (dataGridView3[2, j].Value.ToString() == campo[2])
                        {
                            campo[0] = dataGridView3[0, j].Value.ToString();
                            campo[1] = dataGridView3[1, j].Value.ToString();
                            campo[3] = dataGridView3[3, j].Value.ToString();
                            campo[5] = dataGridView3[4, j].Value.ToString();
                            campo[4] = "NÃO CONSTA";
                            campo[6] = double.Parse(campo[5]).ToString("f2", CultureInfo.InvariantCulture);
                            dataGridView1.Rows.Add(campo);
                            break;
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var openDlg = new FolderBrowserDialog())
            {
                openDlg.SelectedPath = path;

                if (openDlg.ShowDialog() == DialogResult.OK)
                {
                    path = openDlg.SelectedPath;
                    textBox1.Text = path;
                }
            }
        }

        private void ConectarBanco(string query, bool consulta=false)
        {
            Conn.Server = textBox2.Text; 
            Conn.DataBase = textBox3.Text; 
            Conn.Password = textBox4.Text;

            using (SqlConnection cn = new SqlConnection(Conn.StrCon))
            {
                cn.Open();
                toolStripStatusLabel1.Text = "Conectado.";
                using (SqlDataAdapter da = new SqlDataAdapter(query, cn))
                {
                    using (DataTable dt = new DataTable())
                    {
                        da.Fill(dt);
                        if (consulta)
                        {
                            ConsultaXmlBanco(dt);
                        }
                        else
                        {
                            GerarXml(dt);  
                        }
                    }
                }
            }
        }

        private void GerarXml(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    string itemString = item.ToString();

                    if (itemString.Substring(178, 3) == "NFe")
                    {
                        string filePath = Path.Combine(path, itemString.Substring(182, 43) + "-nfe.xml");
                        CriarArquivoXml(filePath, itemString);
                    }
                    else if (itemString.Substring(55, 3) == "CFe")
                    {
                        string filePath = Path.Combine(path, "AD" + itemString.Substring(58, 44) + ".xml");
                        CriarArquivoXml(filePath, itemString);
                    }
                }
            }
        }

        private void CriarArquivoXml(string filePath, string itemString)
        {
            try
            {
                using (var file = File.CreateText(filePath))
                {
                    file.WriteLine(itemString);
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = "Falha.";
                MessageBox.Show($"Falha ao tentar conectar{ex.Message}");
            }
        }
    }
}
