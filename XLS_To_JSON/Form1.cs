using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace XLS_To_JSON
{
	public partial class Form1 : Form
	{
		private DataTable Datos = null;
		private string JsonString = string.Empty;
		private string file_name = string.Empty;

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			try
			{

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message + ex.StackTrace, "Error Insesperado:", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		private void Form1_Shown(object sender, EventArgs e)
		{

		}

		private void abrirXLSToolStripMenuItem_Click(object sender, EventArgs e)
		{
			try
			{
				this.Cursor = Cursors.WaitCursor;

				OpenFileDialog OFDialog = new OpenFileDialog();
				OFDialog.Filter = "Documentos de Excel|*.xls;*.xlsx|Todos los archivos|*.*";
				OFDialog.FilterIndex = 0;
				OFDialog.DefaultExt = "xls";
				OFDialog.AddExtension = true;
				OFDialog.CheckPathExists = true;
				OFDialog.CheckFileExists = true;
				OFDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

				if (OFDialog.ShowDialog() == DialogResult.OK)
				{
					this.file_name = System.IO.Path.GetFileNameWithoutExtension(OFDialog.FileName); //<- Nombre sin Extension ni Path
					this.Datos = Util.Excel_To_DataTable(OFDialog.FileName);
					if (this.Datos != null)
					{
						this.dataGridView1.DataSource = this.Datos;

						if (MessageBox.Show("Deseas Convertir los datos a JSON ahora?", "Convertir a JSON?",
							MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
						{
							this.JsonString = Util.Serialize_ToJSON(this.Datos);
							if (!this.JsonString.IsNullOrEmpty())
							{
								this.textBox1.Text = this.JsonString;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message + ex.StackTrace, "Error Insesperado:", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally { this.Cursor = Cursors.Default; }
		}

		private void guardarJSONToolStripMenuItem_Click(object sender, EventArgs e)
		{
			try
			{
				this.Cursor = Cursors.WaitCursor;

				if (!this.JsonString.IsNullOrEmpty())
				{
					//Muestra el Cuadro de Dialogo Guardar Archivo:
					SaveFileDialog SFDialog = new SaveFileDialog();
					SFDialog.Filter = "Archivo JSON|*.json|Todos los archivos|*.*";
					SFDialog.FilterIndex = 0;
					SFDialog.DefaultExt = "json";
					SFDialog.AddExtension = true;
					SFDialog.CheckPathExists = true;
					SFDialog.OverwritePrompt = true;
					SFDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
					if (!this.file_name.IsNullOrEmpty()) SFDialog.FileName = this.file_name;

					if (SFDialog.ShowDialog() == DialogResult.OK)
					{
						Util.SaveTextFile(SFDialog.FileName, this.JsonString);
						if (File.Exists(SFDialog.FileName))
						{
							MessageBox.Show("Archivo Guardado Correctamente.", "Export to JSON Complete!",
								MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
				else
				{
					MessageBox.Show("No hay nada que Guardar!", "ERR_404 NOT FOUND!",
								MessageBoxButtons.OK, MessageBoxIcon.Hand);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message + ex.StackTrace, "Error Insesperado:", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally { this.Cursor = Cursors.Default; }
		}

		private void salirToolStripMenuItem_Click(object sender, EventArgs e)
		{
			try
			{
				if (this.JsonString.IsNullOrEmpty())
				{
					this.Close();
				}
				else
				{
					if (MessageBox.Show("Hay Datos sin Guardar, Seguro deseas Salir?", "Datos sin Guradar!",
							MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
					{
						this.Close();
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message + ex.StackTrace, "Error Insesperado:", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally { this.Cursor = Cursors.Default; }
		}

		private void toJSONToolStripMenuItem_Click(object sender, EventArgs e)
		{
			try
			{
				this.JsonString = Util.Serialize_ToJSON(this.Datos);
				if (!this.JsonString.IsNullOrEmpty())
				{
					this.textBox1.Text = this.JsonString;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message + ex.StackTrace, "Error Insesperado:", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally { this.Cursor = Cursors.Default; }
		}
	}
}
