using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Media;
using System.Net.Security;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace Ne_Excel_Entní
{
    public partial class Neexcelentni : Form
    {
        private PrintDocument printDocument;
        private PrintDocument documentToPrint;

        public Neexcelentni()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;


            printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

            documentToPrint = new PrintDocument();
            documentToPrint.PrintPage += new PrintPageEventHandler(DocumentToPrint_PrintPage);

            this.Text = "Ne Enxel Entní " + version;
        }
        int theme = 0;
        bool orderColumns = false;

        bool saved = false;
        string version = "1.0";
        bool openFile = false;
        string openedFile;
        System.Windows.Forms.NotifyIcon notifyIcon1 = new System.Windows.Forms.NotifyIcon();

        private void toolStripAddColumnButton_Click(object sender, EventArgs e)
        {
            Form requestDialog = new Form();
            requestDialog.Text = "Výzva";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 300;
            requestDialog.Height = 150;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Zadej název sloupce:",
                Location = new Point(10, 15),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox1 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 45),
                Width = 250
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 60,
                Height = 25,
                Location = new Point(210, 80),
                Text = "OK",
            };

            requestDialogTextBox1.KeyDown += (Nullable, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    requestDialogButton1.PerformClick();
                }
            };

            requestDialogButton1.Click += (Nullable, args) =>
            {
                string input = requestDialogTextBox1.Text;
                if (input != "")
                {
                    dataGrid1.Columns.Add(input, input);
                }
                else
                {
                    Notify("Nesprávně zadaný název!", "Nelze vytvořit sloupec bez názvu");
                }
                requestDialog.Close();
            };

            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogTextBox1);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialogButton1.Focus();

            requestDialog.ShowDialog();
        }
        private void toolStripRemoveColumnButton_Click(object sender, EventArgs e)
        {
            Form requestDialog = new Form();
            requestDialog.Text = "Výzva";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 300;
            requestDialog.Height = 150;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Zadej název sloupce k vymazání:",
                Location = new Point(10, 15),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox1 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 45),
                Width = 250
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 60,
                Height = 25,
                Location = new Point(210, 80),
                Text = "Vymazat",

            };

            requestDialogTextBox1.KeyDown += (Nullable, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    requestDialogButton1.PerformClick();
                }
            };

            requestDialogButton1.Click += (Nullable, args) =>
            {
                string input = requestDialogTextBox1.Text;
                if (input != "")
                {
                    try
                    {
                        dataGrid1.Columns.Remove(input);
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        requestDialog.Close();
                    }


                }
                requestDialog.Close();
            };


            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogTextBox1);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialogButton1.Focus();

            requestDialog.ShowDialog();
        }

        private void toolStripRenameColumnButton_Click(object sender, EventArgs e)
        {
            Form requestDialog = new Form();
            requestDialog.Text = "Výzva";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 300;
            requestDialog.Height = 250;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Zadej název sloupce k přejmenování:",
                Location = new Point(10, 15),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox1 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 45),
                Width = 250
            };

            System.Windows.Forms.Label requestDialogLabel2 = new System.Windows.Forms.Label()
            {
                Text = "Přejmenovat na:",
                Location = new Point(10, 100),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox2 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 130),
                Width = 250
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 80,
                Height = 25,
                Location = new Point(190, 175),
                Text = "Přejmenovat",

            };

            requestDialogTextBox1.KeyDown += (Nullable, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    requestDialogButton1.PerformClick();
                }
            };
            requestDialogTextBox2.KeyDown += (Nullable, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    requestDialogButton1.PerformClick();
                }
            };

            requestDialogButton1.Click += (Nullable, args) =>
            {
                try
                {

                    string rename = requestDialogTextBox1.Text;
                    string input = requestDialogTextBox2.Text;
                    dataGrid1.Columns[rename].HeaderText = input;
                    dataGrid1.Columns[rename].Name = input;
                    requestDialog.Close();
                }
                catch
                {
                    requestDialog.Close();
                    Notify("Nesprávně zadaný sloupec!", "Nenalezen sloupec pod tímto názvem.");
                }
            };

            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogLabel2);
            requestDialog.Controls.Add(requestDialogTextBox1);
            requestDialog.Controls.Add(requestDialogTextBox2);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialogButton1.Focus();

            requestDialog.ShowDialog();
        }

        private void informaceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Ne-Excel-Entní v" + version + "\n\nOpen source, freeware zmenšenina programu Excel pro jednoduché zapisování dat.\n\nVytvořil Bohuslav Janda 5/11/2024", "Informace");
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Form requestDialog = new Form();
            requestDialog.Text = "Výzva";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 300;
            requestDialog.Height = 250;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Zadej název sloupce A:",
                Location = new Point(10, 15),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox1 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 45),
                Width = 250
            };

            System.Windows.Forms.Label requestDialogLabel2 = new System.Windows.Forms.Label()
            {
                Text = "Zadej název sloupce B:",
                Location = new Point(10, 100),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox2 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 130),
                Width = 250
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 80,
                Height = 25,
                Location = new Point(190, 175),
                Text = "Vyměnit",
            };

            requestDialogTextBox1.KeyDown += (Nullable, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    requestDialogButton1.PerformClick();
                }
            };

            requestDialogButton1.Click += (Nullable, args) =>
            {
                try
                {
                    string columnA = requestDialogTextBox1.Text;
                    string columnB = requestDialogTextBox2.Text;

                    int indexA = dataGrid1.Columns[columnA].DisplayIndex;
                    int indexB = dataGrid1.Columns[columnB].DisplayIndex;

                    if (indexA == indexB || indexA < 0 || indexB < 0 ||
            indexA >= dataGrid1.Columns.Count || indexB >= dataGrid1.Columns.Count)
                        return;

                    foreach (DataGridViewRow row in dataGrid1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            var temp = row.Cells[indexA].Value;
                            row.Cells[indexA].Value = row.Cells[indexB].Value;
                            row.Cells[indexB].Value = temp;
                        }
                    }

                    string tempHeader = dataGrid1.Columns[indexA].HeaderText;
                    dataGrid1.Columns[indexA].HeaderText = dataGrid1.Columns[indexB].HeaderText;
                    dataGrid1.Columns[indexB].HeaderText = tempHeader;

                    requestDialog.Close();
                }
                catch
                {
                    requestDialog.Close();
                    Notify("Nesprávně zadaný sloupec!", "Nenalezen sloupec pod tímto názvem.");
                }



            };

            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogLabel2);
            requestDialog.Controls.Add(requestDialogTextBox1);
            requestDialog.Controls.Add(requestDialogTextBox2);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialogButton1.Focus();

            requestDialog.ShowDialog();

        }

        private void toolStripButton4_Click(object sender, EventArgs e)

        {
            Form requestDialog = new Form();
            requestDialog.Text = "Výzva";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 300;
            requestDialog.Height = 150;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Zadej termín k vyhledání:",
                Location = new Point(10, 15),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox1 = new System.Windows.Forms.TextBox()
            {
                Location = new Point(10, 45),
                Width = 250
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 60,
                Height = 25,
                Location = new Point(210, 80),
                Text = "OK",

            };

            requestDialogTextBox1.KeyDown += (Nullable, args) =>
{
    if (args.KeyCode == Keys.Enter)
    {
        requestDialogButton1.PerformClick();
    }
};

            requestDialogButton1.Click += (Nullable, args) =>
            {
                string searchValue = requestDialogTextBox1.Text;
                if (searchValue != "")
                {
                    FindValueInDataGrid(dataGrid1, searchValue);
                }
                requestDialog.Close();
            };

            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogTextBox1);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialogButton1.Focus();

            requestDialog.ShowDialog();

            void FindValueInDataGrid(DataGridView grid, object searchValue)
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null && cell.Value.Equals(searchValue))
                        {
                            cell.Style.BackColor = Color.Yellow;

                            grid.CurrentCell = cell;
                        }
                        else
                        {
                            cell.Style.BackColor = Color.White;
                        }
                    }
                }
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Opravdu chceš vymazat všechna data?", "Upozornění", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                dataGrid1.Columns.Clear();
            }
        }

        private void nápovědaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Pomocí kláves ALT a F4 ukončíte program.", "Nápověda");
        }

        void Save()
        {

            if (openFile == true)
            {
                try
                {
                    StringBuilder stringBuilder = new StringBuilder();

                    for (int i = 0; i < dataGrid1.Columns.Count; i++)
                    {
                        stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                        if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                    }
                    stringBuilder.AppendLine();

                    foreach (DataGridViewRow row in dataGrid1.Rows)
                    {
                        if (row.IsNewRow) continue;

                        for (int j = 0; j < dataGrid1.Columns.Count; j++)
                        {
                            var cellValue = row.Cells[j].Value?.ToString() ?? "";
                            stringBuilder.Append(cellValue);

                            if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                        }
                        stringBuilder.AppendLine();
                    }
                    File.WriteAllText(openedFile, stringBuilder.ToString());

                    MessageBox.Show("Tabulka úspěšně uložena", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                catch (Exception exception)
                {
                    MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            else
            {
                try
                {

                    StringBuilder stringBuilder = new StringBuilder();

                    for (int i = 0; i < dataGrid1.Columns.Count; i++)
                    {
                        stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                        if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                    }
                    stringBuilder.AppendLine();

                    foreach (DataGridViewRow row in dataGrid1.Rows)
                    {
                        if (row.IsNewRow) continue;

                        for (int j = 0; j < dataGrid1.Columns.Count; j++)
                        {
                            var cellValue = row.Cells[j].Value?.ToString() ?? "";
                            stringBuilder.Append(cellValue);

                            if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                        }
                    }
                    stringBuilder.AppendLine();

                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                        saveFileDialog.Title = "Uložit tabulku";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            saved = true;
                            openFile = true;
                            openedFile = saveFileDialog.FileName;
                            File.WriteAllText(saveFileDialog.FileName, stringBuilder.ToString());
                            MessageBox.Show("Tabulka úspěšně uložena", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }


                }
                catch (Exception exception)
                {
                    MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }


        private void LoadFile()

        {
            saved = false;
            string[] lines = new string[0];
            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.Title = "Načíst tabulku dat";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        dataGrid1.Rows.Clear();
                        dataGrid1.Columns.Clear();
                        lines = File.ReadAllLines(openFileDialog.FileName);
                    }



                    if (lines.Length == 0)
                    {
                        MessageBox.Show("Soubor je prázdný", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Add columns based on the first line (headers)
                    string[] headers = lines[0].Split('\t');
                    foreach (string header in headers)
                    {
                        dataGrid1.Columns.Add(header, header);
                    }

                    // Add rows starting from the second line
                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] cells = lines[i].Split('\t');
                        dataGrid1.Rows.Add(cells);
                    }
                    openFile = true;
                    openedFile = openFileDialog.FileName;
                    MessageBox.Show("Data se úspěšně načetla", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Chyba při načířání souboru: " + ex.Message, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void otevřítToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadFile();
        }

        private void uložitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void uložitJakoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                StringBuilder stringBuilder = new StringBuilder();

                for (int i = 0; i < dataGrid1.Columns.Count; i++)
                {
                    stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                    if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                }
                stringBuilder.AppendLine();

                foreach (DataGridViewRow row in dataGrid1.Rows)
                {
                    if (row.IsNewRow) continue;

                    for (int j = 0; j < dataGrid1.Columns.Count; j++)
                    {
                        var cellValue = row.Cells[j].Value?.ToString() ?? "";
                        stringBuilder.Append(cellValue);

                        if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                    }
                }
                stringBuilder.AppendLine();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                    saveFileDialog.Title = "Uložit tabulku";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        openFile = true;
                        openedFile = saveFileDialog.FileName;
                        File.WriteAllText(saveFileDialog.FileName, stringBuilder.ToString());
                        MessageBox.Show("Tabulka úspěšně uložena", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }


            }
            catch (Exception exception)
            {
                MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void uzavřítToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Neexcelentni_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataGrid1.ColumnCount != 0)
            {
                DialogResult result = MessageBox.Show("   Máte neuložená data! \n\n   Chcete je uložit?                           ", "Upozornění", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                e.Cancel = (result == DialogResult.Cancel);
                if (result == DialogResult.Yes)
                {
                    Save();
                    if (saved)
                    {
                        dataGrid1.Columns.Clear();
                        dataGrid1.Rows.Clear();
                        Close();
                    }
                    else
                    {
                        e.Cancel = (result == DialogResult.Yes);
                    }
                }
                if (result == DialogResult.No)
                {
                    dataGrid1.Columns.Clear();
                    dataGrid1.Rows.Clear();
                    Close();
                }
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            {
                foreach (DataGridViewCell cell in dataGrid1.SelectedCells)
                {
                    cell.Value = null; // Clear the cell content
                }
            }
        }

        private void nastaveníToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form requestDialog = new Form();
            requestDialog.Text = "Nastavení";
            requestDialog.ShowIcon = false;
            requestDialog.ShowInTaskbar = false;
            requestDialog.MaximizeBox = false;
            requestDialog.MinimizeBox = false;
            requestDialog.StartPosition = FormStartPosition.CenterParent;
            requestDialog.Width = 400;
            requestDialog.Height = 500;
            requestDialog.AutoSize = false;
            requestDialog.FormBorderStyle = FormBorderStyle.FixedDialog;

            System.Windows.Forms.Label requestDialogLabel1 = new System.Windows.Forms.Label()
            {
                Text = "Vzhled",
                Location = new Point(20, 15),
                AutoSize = true,
            };

            System.Windows.Forms.CheckBox requestDialogCheckBox1 = new System.Windows.Forms.CheckBox()
            {
                Location = new Point(30, 40),
                Text = "Zobrazit pás nástrojů",
                AutoSize = true,
            };

            System.Windows.Forms.CheckBox requestDialogCheckBox2 = new System.Windows.Forms.CheckBox()
            {
                Location = new Point(30, 65),
                Text = "Zobrazit pás stavu",
                AutoSize = true,
            };

            System.Windows.Forms.Label requestDialogLabel2 = new System.Windows.Forms.Label()
            {
                Text = "Motiv:",
                Location = new Point(30, 100),
                AutoSize = true,
            };

            System.Windows.Forms.ComboBox requestDialogComboBox = new System.Windows.Forms.ComboBox()
            {
                Location = new Point(120, 100),
                AutoSize = true,

            };

            System.Windows.Forms.Label requestDialogLabel3 = new System.Windows.Forms.Label()
            {
                Text = "Funkčnost",
                Location = new Point(20, 140),
                AutoSize = true,
            };

            System.Windows.Forms.Label requestDialogLabel4 = new System.Windows.Forms.Label()
            {
                Text = "Doba automatického ukládání (minuty):",
                Location = new Point(30, 165),
                AutoSize = true,
            };

            System.Windows.Forms.TextBox requestDialogTextBox = new System.Windows.Forms.TextBox()
            {
                Location = new Point(230, 164),
                AutoSize = true,
                MaxLength = 2

            };

            System.Windows.Forms.CheckBox requestDialogCheckBox3 = new System.Windows.Forms.CheckBox()
            {
                Location = new Point(30, 195),
                Text = "Povolit dynamické přesouvání sloupců ⚠︎",
                AutoSize = true,
            };

            System.Windows.Forms.Button requestDialogButton1 = new System.Windows.Forms.Button()
            {
                Width = 60,
                Height = 25,
                Location = new Point(305, 425),
                Text = "OK",

            };

            requestDialogTextBox.KeyPress += (Nullable, args) =>
            {
                if (!(args.KeyChar == 8 || (args.KeyChar >= 48 && args.KeyChar <= 57)))
                {
                    args.Handled = true;
                }
            };
            requestDialogTextBox.Text = (autoSaveTimer.Interval / 60000).ToString();


            requestDialogComboBox.Items.Add("Světlý");
            requestDialogComboBox.Items.Add("Světle tmavý");

            if (theme == 0)
            {
                requestDialogComboBox.SelectedItem = "Světlý";
            }
            if (theme == 1)
            {
                requestDialogComboBox.SelectedItem = "Světle tmavý";
            }

            if (orderColumns)
            {
                requestDialogCheckBox3.Checked = true;
            }
            else
            {
                requestDialogCheckBox3.Checked = false;
            }

            requestDialogComboBox.SelectedIndexChanged += (Nullable, args) =>
            {
                if (requestDialogComboBox.SelectedIndex == 1)
                {
                    theme = 1;
                    menuStrip1.BackColor = Color.Silver;
                    checkBox1.BackColor = Color.Silver;
                    BackColor = Color.LightGray;
                    toolStrip1.BackColor = Color.LightGray;
                    dataGrid1.BackColor = Color.DarkGray;
                    statusStrip1.BackColor = Color.Silver;
                }
                else if (requestDialogComboBox.SelectedIndex == 0)
                {
                    theme = 0;
                    menuStrip1.BackColor = DefaultBackColor;
                    checkBox1.BackColor = DefaultBackColor;
                    BackColor = DefaultBackColor;
                    toolStrip1.BackColor = DefaultBackColor;
                    dataGrid1.BackColor = DefaultBackColor;
                    statusStrip1.BackColor = DefaultBackColor;
                }
            };

            requestDialogCheckBox3.CheckedChanged += (Nullable, args) =>
            {
                if (!requestDialogCheckBox3.Checked)
                {
                    dataGrid1.AllowUserToOrderColumns = false;
                    orderColumns = false;
                }
                else
                {
                    dataGrid1.AllowUserToOrderColumns = true;
                    orderColumns = true;
                    MessageBox.Show("Přesouvání sloupců je pouze vizuální a pořadí \nsloupce se uloží podle jejich skutečného indexu.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            if (toolStrip1.Visible == true)
            {
                requestDialogCheckBox1.Checked = true;
            }
            else requestDialogCheckBox1.Checked = false;

            requestDialogCheckBox1.CheckedChanged += (Nullable, args) =>
            {
                statusStripToolStripMenuItem_Click(null, e);
            };

            if (statusStrip1.Visible == true)
            {
                requestDialogCheckBox2.Checked = true;
            }
            else requestDialogCheckBox2.Checked = false;

            requestDialogCheckBox2.CheckedChanged += (Nullable, args) =>
            {
                pásToolStripMenuItem_Click(null, e);
            };

            requestDialogButton1.Click += (Nullable, args) =>
            {
                int tempTimeValue = int.Parse(requestDialogTextBox.Text);
                tempTimeValue = tempTimeValue * 60000;
                autoSaveTimer.Interval = tempTimeValue;
                requestDialog.Close();
            };

            requestDialog.Controls.Add(requestDialogLabel1);
            requestDialog.Controls.Add(requestDialogLabel2);
            requestDialog.Controls.Add(requestDialogLabel3);
            requestDialog.Controls.Add(requestDialogLabel4);
            requestDialog.Controls.Add(requestDialogCheckBox1);
            requestDialog.Controls.Add(requestDialogCheckBox2);
            requestDialog.Controls.Add(requestDialogCheckBox3);
            requestDialog.Controls.Add(requestDialogButton1);
            requestDialog.Controls.Add(requestDialogComboBox);
            requestDialog.Controls.Add(requestDialogTextBox);

            requestDialog.ShowDialog();
        }

        private void statusStripToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (statusStripToolStripMenuItem.Checked)
            {
                statusStripToolStripMenuItem.Checked = false;
                toolStrip1.Hide();
                tabControl1.Top -= 74;
                tabControl1.Height += 74;
                dataGrid1.Height += 74;
            }
            else
            {
                statusStripToolStripMenuItem.Checked = true;
                toolStrip1.Show();
                tabControl1.Top += 74;
                tabControl1.Height -= 74;
                dataGrid1.Height -= 74;

            }
        }

        private void dataGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            toolStripStatusLabel1.Text = "Řádek: " + (e.RowIndex + 1).ToString();
            toolStripStatusLabel2.Text = "Sloupec: " + (e.ColumnIndex + 1).ToString();
        }

        private void refresh_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel3.Text = "Počet řádků: " + dataGrid1.RowCount.ToString() + "    Počet sloupců: " + dataGrid1.ColumnCount.ToString() + " ";
        }

        private void pásToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pásToolStripMenuItem.Checked)
            {
                pásToolStripMenuItem.Checked = false;
                statusStrip1.Hide();
                tabControl1.Height += 22;
                dataGrid1.Height += 22;
            }
            else
            {
                pásToolStripMenuItem.Checked = true;
                statusStrip1.Show();
                tabControl1.Height -= 22;
                dataGrid1.Height -= 22;

            }
        }

        async void Notify(string Title, string Message)
        {
            SystemSounds.Exclamation.Play();
            await Task.Delay(1000);
            System.Windows.Forms.ToolTip hint = new System.Windows.Forms.ToolTip();
            hint.IsBalloon = true;
            hint.ToolTipTitle = Title;
            hint.ToolTipIcon = ToolTipIcon.None;
            hint.Show(Message, statusStrip1, new Point(20, -75));
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Save();
                autoSaveTimer.Start();
            }
            else
            {
                autoSaveTimer.Stop();
            }
        }

        private void autoSaveTimer_Tick(object sender, EventArgs e)
        {
            Save();
        }

        private void programuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Bitmap bitmap = new Bitmap(this.Width, this.Height);
            this.DrawToBitmap(bitmap, new Rectangle(0, 0, this.Width, this.Height));
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures));
            string filename = Path.Combine(directory, $"Screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            bitmap.Save(filename, ImageFormat.Png);

            MessageBox.Show($"Snímek uložen v: {filename}", "Snímek pořízen", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void tabulkyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dimension = dataGrid1;
            using (var bmp = new Bitmap(dimension.Width, dimension.Height))
            {
                dimension.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures));
                string filename = Path.Combine(directory, $"Screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                bmp.Save(filename, ImageFormat.Png);

                MessageBox.Show($"Snímek uložen v: {filename}", "Snímek pořízen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void programuToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PrintForm();
        }
        private void PrintForm()
        {
            using (PrintDialog printDialog = new PrintDialog())
            {
                printDialog.Document = printDocument;

                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    printDocument.Print();
                }
            }
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            using (Bitmap bitmap = new Bitmap(this.Width, this.Height))
            {
                this.DrawToBitmap(bitmap, new Rectangle(0, 0, this.Width, this.Height));
                e.Graphics.DrawImage(bitmap, 0, 0);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintForm();
        }

        private void datToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintGridData();
        }
        private void PrintGridData()
        {
            using (PrintDialog dialogForPrint = new PrintDialog())
            {
                dialogForPrint.Document = documentToPrint;

                if (dialogForPrint.ShowDialog() == DialogResult.OK)
                {
                    documentToPrint.Print();
                }
            }
        }

        private void DocumentToPrint_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font textFont = new Font("Arial", 10);
            float rowHeight = textFont.GetHeight();
            float xPos = e.MarginBounds.Left;
            float yPos = e.MarginBounds.Top;

            foreach (DataGridViewColumn col in dataGrid1.Columns)
            {
                e.Graphics.DrawString(col.HeaderText, textFont, Brushes.Black, xPos, yPos);
                xPos += col.Width;
            }

            yPos += rowHeight;
            xPos = e.MarginBounds.Left;

            foreach (DataGridViewRow row in dataGrid1.Rows)
            {
                if (!row.IsNewRow)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        e.Graphics.DrawString(cell.Value?.ToString() ?? string.Empty, textFont, Brushes.Black, xPos, yPos);
                        xPos += dataGrid1.Columns[cell.ColumnIndex].Width;
                    }
                    yPos += rowHeight;
                    xPos = e.MarginBounds.Left;
                }
            }
        }

        private void ukončitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void uložitAUkončitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
            Close();
        }

        private void veFormátutxtneexToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGrid1.Columns.Count != 0)
            {
                DialogResult dialogResult = MessageBox.Show("   Máte neuloženou práci! \n\n   Chcete ji uložit?                            ", "   Upozornění", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    Save();
                    if (saved)
                    {
                        dataGrid1.Columns.Clear();
                        dataGrid1.Rows.Clear();

                        try
                        {

                            StringBuilder stringBuilder = new StringBuilder();

                            for (int i = 0; i < dataGrid1.Columns.Count; i++)
                            {
                                stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                                if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                            }
                            stringBuilder.AppendLine();

                            foreach (DataGridViewRow row in dataGrid1.Rows)
                            {
                                if (row.IsNewRow) continue;

                                for (int j = 0; j < dataGrid1.Columns.Count; j++)
                                {
                                    var cellValue = row.Cells[j].Value?.ToString() ?? "";
                                    stringBuilder.Append(cellValue);

                                    if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                                }
                            }
                            stringBuilder.AppendLine();

                            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                            {
                                saveFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                                saveFileDialog.Title = "Vytvořit novou tabulku";

                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    openFile = true;
                                    openedFile = saveFileDialog.FileName;
                                    File.WriteAllText(saveFileDialog.FileName, stringBuilder.ToString());
                                    MessageBox.Show("Úspěšně vytvořen nový soubor", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                    }
                }
                if (dialogResult == DialogResult.No)
                {
                    try
                    {

                        StringBuilder stringBuilder = new StringBuilder();

                        for (int i = 0; i < dataGrid1.Columns.Count; i++)
                        {
                            stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                            if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                        }
                        stringBuilder.AppendLine();

                        foreach (DataGridViewRow row in dataGrid1.Rows)
                        {
                            if (row.IsNewRow) continue;

                            for (int j = 0; j < dataGrid1.Columns.Count; j++)
                            {
                                var cellValue = row.Cells[j].Value?.ToString() ?? "";
                                stringBuilder.Append(cellValue);

                                if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                            }
                        }
                        stringBuilder.AppendLine();

                        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                        {
                            saveFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                            saveFileDialog.Title = "Vytvořit novou tabulku";

                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                openFile = true;
                                openedFile = saveFileDialog.FileName;
                                File.WriteAllText(saveFileDialog.FileName, stringBuilder.ToString());
                                MessageBox.Show("Úspěšně vytvořen nový soubor", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }
            }

            try
            {

                StringBuilder stringBuilder = new StringBuilder();

                for (int i = 0; i < dataGrid1.Columns.Count; i++)
                {
                    stringBuilder.Append(dataGrid1.Columns[i].HeaderText);
                    if (i < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                }
                stringBuilder.AppendLine();

                foreach (DataGridViewRow row in dataGrid1.Rows)
                {
                    if (row.IsNewRow) continue;

                    for (int j = 0; j < dataGrid1.Columns.Count; j++)
                    {
                        var cellValue = row.Cells[j].Value?.ToString() ?? "";
                        stringBuilder.Append(cellValue);

                        if (j < dataGrid1.Columns.Count - 1) stringBuilder.Append("\t");
                    }
                }
                stringBuilder.AppendLine();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Neexcelentní data soubor (*.neex)|*.neex|Textový soubor (*.txt)|*.txt|All files (*.*)|*.*";
                    saveFileDialog.Title = "Vytvořit novou tabulku";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        openFile = true;
                        openedFile = saveFileDialog.FileName;
                        File.WriteAllText(saveFileDialog.FileName, stringBuilder.ToString());
                        MessageBox.Show("Úspěšně vytvořen nový soubor", "Informace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Nastala neočekávaná chyba! \n\n" + exception.Message, "Chyba!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void lidéToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGrid1.ColumnCount != 0)
            {
                DialogResult result = MessageBox.Show("   Máte neuložená data! \n\n   Chcete je uložit?                           ", "Upozornění", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {

                }
                else if (result == DialogResult.Yes)
                {
                    Save();
                    if (saved)
                    {
                        dataGrid1.Columns.Clear();
                        dataGrid1.Rows.Clear();
                    }
                    else
                    {
                    }
                }
                else if (result == DialogResult.No)
                {
                    dataGrid1.Columns.Clear();
                    dataGrid1.Rows.Clear();
                }
            }

            dataGrid1.Columns.Add("Pořadí", "Pořadí");
            dataGrid1.Columns.Add("Jméno", "Jméno");
            dataGrid1.Columns.Add("Příjmení", "Příjmení");
            dataGrid1.Columns.Add("Celé jméno", "Celé jméno");
            dataGrid1.Columns.Add("Ročník", "Ročník");
            dataGrid1.Columns.Add("Datum 1", "Datum 1");
            dataGrid1.Columns.Add("Datum 2", "Datum 2");
            dataGrid1.Columns.Add("Bydliště 1", "Bydliště 1");
            dataGrid1.Columns.Add("Bydliště 2", "Bydliště 2");
            dataGrid1.Columns.Add("Plat", "Plat");
            dataGrid1.Columns.Add("Obor zaměstnání", "Obor zaměstnání");
            dataGrid1.Columns.Add("Vyučení", "Vyučení");
        }

        private void skladToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGrid1.ColumnCount != 0)
            {
                DialogResult result = MessageBox.Show("   Máte neuložená data! \n\n   Chcete je uložit?                           ", "Upozornění", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {

                }
                else if (result == DialogResult.Yes)
                {
                    Save();
                    if (saved)
                    {
                        dataGrid1.Columns.Clear();
                        dataGrid1.Rows.Clear();
                    }
                    else
                    {
                    }
                }
                else if (result == DialogResult.No)
                {
                    dataGrid1.Columns.Clear();
                    dataGrid1.Rows.Clear();
                }
            }

            dataGrid1.Columns.Add("Index", "Index");
            dataGrid1.Columns.Add("Pořadí", "Pořadí");
            dataGrid1.Columns.Add("Číslo", "Číslo");
            dataGrid1.Columns.Add("Krátký název", "Krátký název");
            dataGrid1.Columns.Add("Číslo modelu", "Číslo modelu");
            dataGrid1.Columns.Add("Sériové číslo", "Sériové číslo");
            dataGrid1.Columns.Add("Stav", "Stav");
            dataGrid1.Columns.Add("Výrobce", "Výrobce");
            dataGrid1.Columns.Add("Distributor", "Distributor");
        }

        private void zbožíToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGrid1.ColumnCount != 0)
            {
                DialogResult result = MessageBox.Show("   Máte neuložená data! \n\n   Chcete je uložit?                           ", "Upozornění", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {

                }
                else if (result == DialogResult.Yes)
                {
                    Save();
                    if (saved)                    {
                        dataGrid1.Columns.Clear();
                        dataGrid1.Rows.Clear();
                    }
                    else
                    {
                    }
                }
                else if (result == DialogResult.No)
                {
                    dataGrid1.Columns.Clear();
                    dataGrid1.Rows.Clear();
                }
            }
            dataGrid1.Columns.Add("Index", "Index");
            dataGrid1.Columns.Add("Pořadí", "Pořadí");
            dataGrid1.Columns.Add("Číslo", "Číslo");
            dataGrid1.Columns.Add("Krátký název", "Krátký název");
            dataGrid1.Columns.Add("Číslo modelu", "Číslo modelu");
            dataGrid1.Columns.Add("Sériové číslo", "Sériové číslo");
            dataGrid1.Columns.Add("Výrobce", "Výrobce");
            dataGrid1.Columns.Add("Distributor", "Distributor");
            dataGrid1.Columns.Add("Cena", "Cena");
            dataGrid1.Columns.Add("Prodejní cena", "Prodejní cena");
            dataGrid1.Columns.Add("Kontakt prodejce", "Kontakt prodejce");

        }
    }




}
