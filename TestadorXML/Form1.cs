using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Linq;
using System.Globalization;

namespace IzzyXML
{
    public partial class Form1 : Form
    {
        string path = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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
                //Form2 form2 = new Form2();
                //form2.Show();
                label7.Visible = true; label8.Visible = true; label14.Visible = true; label15.Visible = true; label16.Visible = true;

                //Busca os arquivos XML
                BuscarArquivos(path);

                // Login Servidor
                string query = $"Select Modelo, Emissao as Emissão, NF as Número, Serie, TotalNota as Total from Fiscal.Documento " +
                    $"where CONVERT(DATE, Emissao) between CONVERT(DATE, '{dateTimePicker1.Value}') " +
                    $"and CONVERT(DATE, '{dateTimePicker2.Value}') and NF != '' order by NF; ";
                ConectarBanco(query, true);
                // Gera a tabela com a diferença entre XML e Banco
                if (dataGridView2.RowCount == 0)
                {
                    MessageBox.Show("Nenhum XML encontrado");
                    label12.Text = dataGridView3.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture)).ToString("N2");
                    label18.Text = dataGridView3.Rows.Count.ToString();
                    label18.Visible = true;
                }
                else if (dataGridView3.RowCount == 0)
                {
                    MessageBox.Show("Nenhum documento encontrado no banco de dados nesse período");
                    label11.Text = dataGridView2.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture)).ToString("N2");
                    label17.Text = dataGridView2.Rows.Count.ToString();
                    label17.Visible = true;
                }
                else
                {
                    label11.Text = dataGridView2.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture)).ToString("N2");
                    label12.Text = dataGridView3.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture)).ToString("N2");
                    label13.Text = Math.Abs(dataGridView2.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture)) -
                        dataGridView3.Rows.Cast<DataGridViewRow>().Sum(i => Convert.ToDecimal(i.Cells["TOTAL"].Value, CultureInfo.InvariantCulture))).ToString("F2");
                    label17.Text = dataGridView2.Rows.Count.ToString();
                    label17.Visible = true;
                    label18.Text = dataGridView3.Rows.Count.ToString();
                    label18.Visible = true;
                    DiferencaBancoXml();
                }
                //form2.Hide();
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = "Falha.";
                MessageBox.Show($"Falha ao tentar conectar\n\n{ex.Message}");
            }
            
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            
        }
        private void DocumentosDiferentes_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        private void LimparTabelas()
        {
            if (dataGridView1.RowCount != 0)
            {
                dataGridView1.Rows.Clear();
            }
            if (dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows.Clear();
            }
            
        }
        private void BuscarArquivos(string path)
        {
            // Puxa os arquivos da pasta
            string[] arquivos = Directory.GetFiles(path);
            foreach (string arq in arquivos)
            {
                LerXml(arq, path);
            }
        }
        private void LerXml(string arquivo, string path)
        {
            string[] campo = new string[5]; // Array onde ficam os dados da tabela
            string tipoArquivo = Path.GetFileNameWithoutExtension(path + "\\" + arquivo);
            tipoArquivo = tipoArquivo.Substring(tipoArquivo.Length - 3, 3);
            var xmlRed = false; // Verifica se o arquivo XML foi lido
            //MessageBox.Show(tipoArquivo);

            using (XmlReader meuXml = XmlReader.Create(arquivo))
            {
                double totalXML = 0;
                while (meuXml.Read())
                {
                    if (tipoArquivo != "nfe")
                    {
                        //CF-e
                        string data = "";
                        campo[0] = "59";
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "nserieSAT")
                        {
                            campo[3] = meuXml.ReadElementContentAsString(); // Série
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "nCFe")
                        {
                            campo[2] = meuXml.ReadElementContentAsString(); // Número do cupom
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "dEmi")
                        {
                            data = meuXml.ReadElementContentAsString(); // Data emissão
                            data = data.Insert(4, "-");
                            data = data.Insert(7, "-");
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "hEmi")
                        {
                            string hora = meuXml.ReadElementContentAsString(); // Hora
                            hora = hora.Insert(2, ":");
                            hora = hora.Insert(5, ":");
                            data = data + " " + hora;
                            DateTime dt = Convert.ToDateTime(data);
                            campo[1] = dt.ToString("g"); // Data completa
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "vCFe")
                        {
                            campo[4] = meuXml.ReadElementContentAsString(); // Valor do cupom
                            totalXML += double.Parse(campo[4]);
                            xmlRed = true; // Xml lido
                        }

                        if (xmlRed)
                        {
                            dataGridView2.Rows.Add(campo); // Adiciona linha na tabela
                            xmlRed = false;
                        }

                    }
                    else if (tipoArquivo == "nfe")
                    {
                        //NFC-e
                        campo[0] = "65";
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "serie")
                        {
                            campo[3] = meuXml.ReadElementContentAsString(); // Série
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "nNF")
                        {
                            campo[2] = meuXml.ReadElementContentAsString(); // Número da nota
                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "dhEmi")
                        {
                            DateTime dt = meuXml.ReadElementContentAsDateTime(); // Data da emissão
                            campo[1] = dt.ToString("g");

                        }
                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "vNF")
                        {
                            campo[4] = meuXml.ReadElementContentAsString(); // Valor da nota 
                            xmlRed = true;
                        }
                        if (xmlRed)
                        {
                            dataGridView2.Rows.Add(campo);  // Adiciona linha na tabela
                            totalXML += double.Parse(campo[4]);
                            xmlRed = false;
                        }
                    }
                }
                
            }
        }
        private void ConectarBanco(string query, bool consulta=false)
        {
            Conn.Server = textBox2.Text; Conn.DataBase = textBox3.Text; Conn.Password = textBox4.Text;

            // Conexão com o banco de dados
            using (SqlConnection cn = new SqlConnection(Conn.StrCon))
            {
                cn.Open();
                toolStripStatusLabel1.Text = "Conectado.";
                var sqlQuery = query;
                //var sqlQuery = $"Select Modelo, Emissao as Emissão, NF as Número, Serie, TotalNota as Total from Fiscal.Documento " +
                //    $"where CONVERT(DATE, Emissao) between CONVERT(DATE, '{dateTimePicker1.Value}') " +
                //    $"and CONVERT(DATE, '{dateTimePicker2.Value}') and NF != '' order by NF; ";

                using (SqlDataAdapter da = new SqlDataAdapter(sqlQuery, cn))
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
        private void DiferencaBancoXml()
        {
            // Calcula o menor número da nota
            int min1 = int.Parse((string)dataGridView2[2, 0].Value);
            int min2 = int.Parse(dataGridView3[2, 0].Value.ToString());
            int min = min1 > min2 ? min2 : min1;

            // Calcula o maior número da nota
            int max1 = int.Parse((string)dataGridView2[2, dataGridView2.Rows.Count - 1].Value);
            int max2 = int.Parse(dataGridView3[2, dataGridView3.Rows.Count - 1].Value.ToString());
            int max = max1 < max2 ? max2 : max1;

            
            string[] campo = new string[7]; // Array onde são armazenados os valores da linha
            List<string> listaDfeXml = new List<string>(); //Lista dos números das DFe dos XMLs
            List<string> listaDfeBanco = new List<string>(); //Lista dos números das DFe do banco

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
                campo[4] = "0"; campo[5] = "0"; // Por padrão o valor da nota do xml e do banco são colocados como 0
                campo[2] = (min + i).ToString(); // Número da nota

                if (listaDfeXml.Contains(campo[2]))
                {
                    for (int j = 0; j < dataGridView2.RowCount; j++)
                    {
                        if (int.Parse(dataGridView2[2, j].Value.ToString()).ToString() == int.Parse(campo[2]).ToString())
                        {
                            campo[0] = dataGridView2[0, j].Value.ToString();
                            campo[1] = dataGridView2[1, j].Value.ToString(); // Emissão
                            campo[3] = dataGridView2[3, j].Value.ToString(); // Série
                            campo[4] = dataGridView2[4, j].Value.ToString(); // Valor nota Xml
                            if (listaDfeBanco.Contains(campo[2]))
                            {
                                for (int k = 0; k < dataGridView3.RowCount; k++)
                                {
                                    if (int.Parse(dataGridView3[2, k].Value.ToString()).ToString() == int.Parse(campo[2]).ToString() && int.Parse(dataGridView3[3, k].Value.ToString()).ToString() == int.Parse(campo[3]).ToString())
                                    {
                                        campo[5] = dataGridView3[4, k].Value.ToString(); // Valor nota Banco
                                        campo[5] = double.Parse(campo[5]).ToString("f2", CultureInfo.InvariantCulture); // Valor nota Banco
                                        break;
                                    }
                                }
                            // Calcula a diferença dos dois valores de nota
                            campo[6] = (Math.Abs(double.Parse(campo[5], CultureInfo.InvariantCulture) - double.Parse(campo[4], CultureInfo.InvariantCulture))).ToString("f2",CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                campo[5] = "NÃO CONSTA";
                                campo[6] = campo[4];
                            }
                            dataGridView1.Rows.Add(campo); // Adiciona uma linha na tabela
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
                            campo[1] = dataGridView3[1, j].Value.ToString(); // Emissão
                            campo[3] = dataGridView3[3, j].Value.ToString(); // Série
                            campo[5] = dataGridView3[4, j].Value.ToString(); // Valor nota Banco
                            campo[4] = "NÃO CONSTA";
                            // Calcula a diferença dos dois valores de nota
                            campo[6] = double.Parse(campo[5]).ToString("f2", CultureInfo.InvariantCulture);
                            dataGridView1.Rows.Add(campo); // Adiciona uma linha na tabela
                            break;
                        }

                    }
                }
                
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // Abre uma janela de pesquisa de pasta
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
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //textBox4.Text = "";
            textBox4.PasswordChar = '*';
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            label7.Visible = false; label8.Visible = false; label14.Visible = false; label15.Visible = false; label16.Visible = false;

            string query = $"Select XMLDocumentoAutorizado from Fiscal.Documento " +
                    $"where CONVERT(DATE, Emissao) between CONVERT(DATE, '{dateTimePicker1.Value}') " +
                    $"and CONVERT(DATE, '{dateTimePicker2.Value}') AND XMLDocumentoAutorizado != ''; ";

            ConectarBanco(query);

        }

        private void ConsultaXmlBanco(DataTable dt)
        {
            dataGridView3.Columns.Clear();
            dataGridView3.DataSource = dt;
        }

        private void GerarXml(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    string itemString = item.ToString();

                    if (item.ToString().Substring(178, 3) == "NFe")
                    {
                        string filePath = path + "\\" + itemString.Substring(182, 43) + "-nfe.xml";

                        string date = itemString.Substring(350, 8) + " " + itemString.Substring(359,8);

                        try
                        {
                            using var file = File.CreateText(filePath);
                            File.SetCreationTime(filePath, DateTime.Parse(date));
                            file.WriteLine(item);
                            file.Close();

                        }
                        catch (Exception ex)
                        {
                            toolStripStatusLabel1.Text = "Falha.";
                            MessageBox.Show($"Falha ao tentar conectar\n\n{ex.Message}");

                        }
                    }
                    else if (item.ToString().Substring(55, 3) == "CFe")
                    {
                        string filePath = path + "\\AD" + item.ToString().Substring(58, 44) + ".xml";

                        string date = itemString.Substring(271, 2) + "\\" + itemString.Substring(269, 2) + "\\" + itemString.Substring(265, 4) +
                            " " + itemString.Substring(286, 2) + ":" + itemString.Substring(288, 2) + ":" + itemString.Substring(290, 2);

                        try
                        {
                            File.SetCreationTime(filePath, DateTime.Parse(date));
                            using var file = File.CreateText(filePath);
                            file.WriteLine(item);
                            file.Close();
                        }
                        catch (Exception ex)
                        {
                            toolStripStatusLabel1.Text = "Falha.";
                            MessageBox.Show($"Falha ao tentar conectar\n\n{ex.Message}");
                        }

                    }

                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }


}