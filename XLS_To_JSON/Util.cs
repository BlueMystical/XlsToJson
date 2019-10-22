using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace XLS_To_JSON
{
	public static class Util
	{
		#region Excel

		/* AUTOR:  Jhollman, 2017  */
		/* REFERENCIAS REQUERIDAS (en el proyecto 'DXCutcsa.VO.Common'):
		 * Ubicacion:	'C:\Cutcsa\DXComercial\Librerias\'  
		 * Librerias:	NPOI.dll, NPOI.OOXML.dll, NPOI.OpenXml4Net.dll, NPOI.OpenXmlFormats.dll  */

		/* MODO DE USO:
		 * Guardar Excel:
		 * -------------
		 *  DataTable datos_xls = Util.ObjectToDataTable(this.Datos.ToArray());
			if (datos_xls != null && datos_xls.Rows.Count > 0)
			{
				datos_xls.TableName = "Alertas_Niveles_Stock";

				SaveFileDialog SFDialog = new SaveFileDialog();
				SFDialog.Filter = "Excel|*.xls|Excel Nuevo|*.xlsx|Todos los archivos|*.*";
				SFDialog.FilterIndex = 0;
				SFDialog.DefaultExt = "xls";
				SFDialog.AddExtension = true;
				SFDialog.CheckPathExists = true;
				SFDialog.OverwritePrompt = true;
				SFDialog.FileName = string.Format("Alertas_Niveles_Stock_{0}.xls", Util.TurnoDesdeFecha(DateTime.Now));
				SFDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

				if (SFDialog.ShowDialog() == DialogResult.OK)
				{							
					Util.DataTable_To_Excel(datos_xls, SFDialog.FileName);
					if (System.IO.File.Exists(SFDialog.FileName))
					{
						//Abre el Archivo con el programa predeterminado:
						System.Diagnostics.Process.Start(SFDialog.FileName);
					}
				}
			}   */

		/// <summary>Abre un archivo de Excel (xls o xlsx) y lo convierte en un DataTable.
		/// LA PRIMERA FILA DEBE CONTENER LOS NOMBRES DE LOS CAMPOS.</summary>
		/// <param name="pRutaArchivo">Ruta completa del archivo a abrir.</param>
		/// <param name="pHojaIndex">Número (basado en cero) de la hoja que se desea abrir. 0 es la primera hoja.</param>
		public static DataTable Excel_To_DataTable(string pRutaArchivo, int pHojaIndex = 0)
		{
			// --------------------------------- //
			/* REFERENCIAS:
			 * NPOI.dll
			 * NPOI.OOXML.dll
			 * NPOI.OpenXml4Net.dll */
			// --------------------------------- //
			/* USING:
			 * using NPOI.SS.UserModel;
			 * using NPOI.HSSF.UserModel;
			 * using NPOI.XSSF.UserModel; */
			// AUTOR: Ing. Jhollman Chacon R. Cutcsa 2015
			// --------------------------------- //
			DataTable Tabla = null;
			try
			{
				if (System.IO.File.Exists(pRutaArchivo))
				{

					NPOI.SS.UserModel.IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx	 			
					ISheet worksheet = null;
					string first_sheet_name = "";

					using (FileStream FS = new FileStream(pRutaArchivo, FileMode.Open, FileAccess.Read))
					{
						workbook = WorkbookFactory.Create(FS);          //Abre tanto XLS como XLSX
						worksheet = workbook.GetSheetAt(pHojaIndex);    //Obtener Hoja por indice
						first_sheet_name = worksheet.SheetName;         //Obtener el nombre de la Hoja

						Tabla = new DataTable(first_sheet_name);
						Tabla.Rows.Clear();
						Tabla.Columns.Clear();

						// Leer Fila por fila desde la primera
						for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
						{
							DataRow NewReg = null;
							IRow row = worksheet.GetRow(rowIndex);
							IRow row2 = null;
							IRow row3 = null;

							if (rowIndex == 0)
							{
								row2 = worksheet.GetRow(rowIndex + 1); //Si es la Primera fila, obtengo tambien la segunda para saber el tipo de datos
								row3 = worksheet.GetRow(rowIndex + 2); //Y la tercera tambien por las dudas
							}

							if (row != null) //null is when the row only contains empty cells 
							{
								if (rowIndex > 0) NewReg = Tabla.NewRow();

								int colIndex = 0;
								//Leer cada Columna de la fila
								foreach (ICell cell in row.Cells)
								{
									object valorCell = null;
									string cellType = "";
									string[] cellType2 = new string[2];

									if (rowIndex == 0) //Asumo que la primera fila contiene los titulos:
									{
										for (int i = 0; i < 2; i++)
										{
											ICell cell2 = null;
											if (i == 0) { cell2 = row2.GetCell(cell.ColumnIndex); }
											else { cell2 = row3.GetCell(cell.ColumnIndex); }

											if (cell2 != null)
											{
												switch (cell2.CellType)
												{
													case CellType.Blank: break;
													case CellType.Boolean:
														cellType2[i] = "System.Boolean"; break;
													case CellType.String: cellType2[i] = "System.String"; break;
													case CellType.Numeric:
														if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
														else
														{
															cellType2[i] = "System.Double";  //valorCell = cell2.NumericCellValue;
														}
														break;

													case CellType.Formula:
														bool continuar = true;
														switch (cell2.CachedFormulaResultType)
														{
															case CellType.Boolean:
																cellType2[i] = "System.Boolean"; break;
															case CellType.String: cellType2[i] = "System.String"; break;
															case CellType.Numeric:
																if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
																else
																{
																	try
																	{
																		//DETERMINAR SI ES BOOLEANO
																		if (cell2.CellFormula == "TRUE()")
																		{ cellType2[i] = "System.Boolean"; continuar = false; }
																		if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
																		if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
																	}
																	catch { }
																}
																break;
														}
														break;
													default:
														cellType2[i] = "System.String"; break;
												}
											}
										}

										//Resolver las diferencias de Tipos
										if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
										else
										{
											if (cellType2[0] == null) cellType = cellType2[1];
											if (cellType2[1] == null) cellType = cellType2[0];
											if (cellType == "") cellType = "System.String";
										}

										//Obtener el nombre de la Columna
										string colName = "Column_{0}";
										try { colName = cell.StringCellValue; }
										catch { colName = string.Format(colName, colIndex); }

										//Verificar que NO se repita el Nombre de la Columna
										foreach (DataColumn col in Tabla.Columns)
										{
											if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
										}

										//Agregar el campos de la tabla:
										if (cellType == null) cellType = "System.String";
										DataColumn codigo = new DataColumn(colName, System.Type.GetType(cellType));
										Tabla.Columns.Add(codigo); colIndex++;
									}
									else
									{
										//Las demas filas son registros:
										switch (cell.CellType)
										{
											case CellType.Blank: valorCell = DBNull.Value; break;
											case CellType.Boolean:
												valorCell = cell.BooleanCellValue; break;
											case CellType.String: valorCell = cell.StringCellValue; break;
											case CellType.Numeric:
												if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
												else { valorCell = cell.NumericCellValue; }
												break;
											case CellType.Formula:
												if (cell.CellFormula == "FALSE()") valorCell = false;
												else if (cell.CellFormula == "TRUE()") valorCell = true;
												else
												{
													switch (cell.CachedFormulaResultType)
													{
														case CellType.Blank: valorCell = DBNull.Value; break;
														case CellType.String: valorCell = cell.StringCellValue; break;
														case CellType.Boolean:
															valorCell = cell.BooleanCellValue; break;
														case CellType.Numeric:
															if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
															else { valorCell = cell.NumericCellValue; }
															break;
													}
												}
												break;
											default: valorCell = cell.StringCellValue; break;
										}
										//Agregar el nuevo Registro
										if (cell.ColumnIndex <= Tabla.Columns.Count - 1) NewReg[cell.ColumnIndex] = valorCell;
									}
								}
							}
							if (rowIndex > 0) Tabla.Rows.Add(NewReg);
						}
						Tabla.AcceptChanges();
					}
				}
				else
				{
					throw new Exception("ERROR 404: El archivo especificado NO existe.");
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return Tabla;
		}

		/// <summary>Convierte un DataTable en un archivo de Excel (xls o Xlsx) y lo guarda en disco.</summary>
		/// <param name="pDatos">Datos de la Tabla a guardar. Usa el nombre de la tabla como nombre de la Hoja</param>
		/// <param name="pFilePath">Ruta del archivo donde se guarda.</param>
		public static void DataTable_To_Excel(DataTable pDatos, string pFilePath)
		{
			try
			{
				if (pDatos != null && pDatos.Rows.Count > 0)
				{
					IWorkbook workbook = null;
					ISheet worksheet = null;

					using (FileStream stream = new FileStream(pFilePath, FileMode.Create, FileAccess.ReadWrite))
					{
						string Ext = System.IO.Path.GetExtension(pFilePath); //<-Extension del archivo
						switch (Ext.ToLower())
						{
							case ".xls":
								HSSFWorkbook workbookH = new HSSFWorkbook();
								NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
								dsi.Company = "Cutcsa"; dsi.Manager = "Departamento Informatico";
								workbookH.DocumentSummaryInformation = dsi;
								workbook = workbookH;
								break;

							case ".xlsx": workbook = new XSSFWorkbook(); break;
						}

						worksheet = workbook.CreateSheet(pDatos.TableName); //<-Usa el nombre de la tabla como nombre de la Hoja

						//CREAR EN LA PRIMERA FILA LOS TITULOS DE LAS COLUMNAS
						int iRow = 0;
						if (pDatos.Columns.Count > 0)
						{
							int iCol = 0;
							IRow fila = worksheet.CreateRow(iRow);
							foreach (DataColumn columna in pDatos.Columns)
							{
								ICell cell = fila.CreateCell(iCol, CellType.String);
								cell.SetCellValue(columna.ColumnName);
								iCol++;
							}
							iRow++;
						}

						//FORMATOS PARA CIERTOS TIPOS DE DATOS
						ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
						_doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

						ICellStyle _intCellStyle = workbook.CreateCellStyle();
						_intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

						ICellStyle _boolCellStyle = workbook.CreateCellStyle();
						_boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

						ICellStyle _dateCellStyle = workbook.CreateCellStyle();
						_dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

						ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
						_dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

						//AHORA CREAR UNA FILA POR CADA REGISTRO DE LA TABLA
						foreach (DataRow row in pDatos.Rows)
						{
							IRow fila = worksheet.CreateRow(iRow);
							int iCol = 0;
							foreach (DataColumn column in pDatos.Columns)
							{
								ICell cell = null; //<-Representa la celda actual								
								object cellValue = row[iCol]; //<- El valor actual de la celda

								switch (column.DataType.ToString())
								{
									case "System.Boolean":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Boolean);

											if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
											else { cell.SetCellFormula("FALSE()"); }

											cell.CellStyle = _boolCellStyle;
										}
										break;

									case "System.String":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.String);
											cell.SetCellValue(Convert.ToString(cellValue));
										}
										break;

									case "System.Int32":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Numeric);
											cell.SetCellValue(Convert.ToInt32(cellValue));
											cell.CellStyle = _intCellStyle;
										}
										break;
									case "System.Int64":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Numeric);
											cell.SetCellValue(Convert.ToInt64(cellValue));
											cell.CellStyle = _intCellStyle;
										}
										break;
									case "System.Decimal":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Numeric);
											cell.SetCellValue(Convert.ToDouble(cellValue));
											cell.CellStyle = _doubleCellStyle;
										}
										break;
									case "System.Double":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Numeric);
											cell.SetCellValue(Convert.ToDouble(cellValue));
											cell.CellStyle = _doubleCellStyle;
										}
										break;

									case "System.DateTime":
										if (cellValue != DBNull.Value)
										{
											cell = fila.CreateCell(iCol, CellType.Numeric);
											cell.SetCellValue(Convert.ToDateTime(cellValue));

											//Si No tiene valor de Hora, usar formato dd-MM-yyyy
											DateTime cDate = Convert.ToDateTime(cellValue);
											if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
											else { cell.CellStyle = _dateCellStyle; }
										}
										break;
									default:
										break;
								}
								iCol++;
							}
							iRow++;
						}

						workbook.Write(stream);
						stream.Close();
					}
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		#endregion

		#region Serializacion

		/// <summary>Serializa y escribe el objeto indicado en un archivo JSON.
		/// <para>La Clase a Serializar DEBE tener un Constructor sin parametros.</para>
		/// <para>Only Public properties and variables will be written to the file. These can be any type, even other classes.</para>
		/// <para>If there are public properties/variables that you do not want written to the file, decorate them with the [JsonIgnore] attribute.</para>
		/// </summary>
		/// <typeparam name="T">El tipo de Objeto a guardar en el Archivo.</typeparam>
		/// <param name="filePath">Ruta completa al archivo donde se guardan.</param>
		/// <param name="objectToWrite">Instancia del Objeto a Serializar</param>
		/// <param name="append">'false'=Sobre-Escribe el Archivo, 'true'=Añade datos al final del archivo.</param>
		public static void Serialize_ToJSON<T>(string filePath, T objectToWrite, bool append = false) where T : new()
		{
			TextWriter writer = null;
			try
			{
				var contentsToWriteToFile = Newtonsoft.Json.JsonConvert.SerializeObject(objectToWrite);
				writer = new StreamWriter(filePath, append);
				writer.Write(contentsToWriteToFile);
			}
			finally
			{
				if (writer != null)
					writer.Close();
			}
		}

		/// <summary>Serializa y escribe el objeto indicado en una cadena JSON.
		/// <para>Object type must have a parameterless constructor.</para>
		/// <para>Only Public properties and variables will be written to the file. These can be any type though, even other classes.</para>
		/// <para>If there are public properties/variables that you do not want written to the file, decorate them with the [JsonIgnore] attribute.</para>
		/// </summary>
		/// <typeparam name="T">The type of object being written to the file.</typeparam>
		/// <param name="objectToWrite">The object instance to write to the file.</param>
		public static string Serialize_ToJSON<T>(T objectToWrite) where T : new()
		{
			string _ret = string.Empty;
			try
			{
				_ret = Newtonsoft.Json.JsonConvert.SerializeObject(objectToWrite);
			}
			catch { }
			return _ret;
		}


		/// <summary>Crea una instancia de un Objeto leyendo sus datos desde un archivo JSON.
		/// <para>Object type must have a parameterless constructor.</para></summary>
		/// <typeparam name="T">The type of object to read from the file.</typeparam>
		/// <param name="filePath">The file path to read the object instance from.</param>
		/// <returns>Returns a new instance of the object read from the Json file.</returns>
		public static T DeSerialize_FromJSON<T>(string filePath) where T : new()
		{
			TextReader reader = null;
			try
			{
				reader = new StreamReader(filePath);
				var fileContents = reader.ReadToEnd();
				return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(fileContents);
			}
			finally
			{
				if (reader != null)
					reader.Close();
			}
		}

		/// <summary>Crea una instancia de un Objeto leyendo sus datos desde una cadena JSON.
		/// <para>Object type must have a parameterless constructor.</para></summary>
		/// <typeparam name="T">The type of object to read from the file.</typeparam>
		/// <param name="JSONstring">Texto con formato JSON</param>
		/// <returns>Returns a new instance of the object</returns>
		public static T DeSerialize_FromJSON_String<T>(string JSONstring) where T : new()
		{
			if (JSONstring != null && JSONstring != string.Empty)
			{
				return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(JSONstring);
			}
			else
			{
				return default(T);
			}
		}


		/// <summary>Serializa un Obje.</summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="serializableObject"></param>
		/// <param name="fileName"></param>
		public static void Serialize_ToXML<T>(T serializableObject, string fileName)
		{
			if (serializableObject == null) { return; }

			try
			{
				System.Xml.XmlDocument xmlDocument = new System.Xml.XmlDocument();
				System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(serializableObject.GetType());
				using (System.IO.MemoryStream stream = new System.IO.MemoryStream())
				{
					serializer.Serialize(stream, serializableObject);
					stream.Position = 0;
					xmlDocument.Load(stream);
					xmlDocument.Save(fileName);
				}
			}
			catch (Exception ex)
			{
				//Log exception here
			}
		}

		/// <summary>
		/// Deserializes an xml file into an object list
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public static T DeSerialize_FromXML<T>(string fileName)
		{
			if (string.IsNullOrEmpty(fileName)) { return default(T); }

			T objectOut = default(T);

			try
			{
				System.Xml.XmlDocument xmlDocument = new System.Xml.XmlDocument();
				xmlDocument.Load(fileName);
				string xmlString = xmlDocument.OuterXml;

				using (System.IO.StringReader read = new System.IO.StringReader(xmlString))
				{
					Type outType = typeof(T);

					System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(outType);
					using (System.Xml.XmlReader reader = new System.Xml.XmlTextReader(read))
					{
						objectOut = (T)serializer.Deserialize(reader);
					}
				}
			}
			catch (Exception ex)
			{
				//Log exception here
			}

			return objectOut;
		}

		#endregion

		#region Extension del Lenguaje C#

		/* Estos metodos se pueden llamar directamente desde el tipo base:
			 * Ejem:  string X = "Hola Mundo ";
			 *  X = X.RemoveLastCharacter(); --> "Hola Mundo"
				var L = X.Left(2); --> "Ho"
				var R = X.Right(3); --> "ndo"
				var M = X.Mid(1, 4); --> "ola M"
				bool I = X.In("a", "b"); --> falso
			 * 
			 *  List<string> Cosas = null;
				if (Cosas.IsEmpty())
				{
					//reeplaza a ->  if (Cosas != null && Cosas.Count > 0) 
				}
			 * */

		#region Generales

		/// <summary>Verifica que la lista de objetos NO sea nula y tenga al menos 1 elemento.
		/// <para>Devuelve 'True' si la Lista contiene Elementos.</para>
		/// <para>Devuelve 'False' si la Lista es Nula o Vacia.</para></summary>
		public static bool IsNotEmpty(this ICollection elements)
		{
			return elements != null && elements.Count > 0;
		}

		/// <summary>Devuelve uno de dos objetos, dependiendo de la evaluación de una expresión.</summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="expression">Expresión que se desea evaluar.</param>
		/// <param name="truePart">Se devuelve si Expression se evalúa como True.</param>
		/// <param name="falsePart">Se devuelve si Expression se evalúa como False.</param>
		public static object IIf(bool expression, object truePart, object falsePart)
		{ return expression ? truePart : falsePart; }
		public static T IIf<T>(bool expression, T truePart, T falsePart)
		{ return expression ? truePart : falsePart; }

		/// <summary>Determina si el valor de la Variable se encuentra dentro del Rango especificado.</summary>
		/// <typeparam name="T">Tipo de Datos del  Objeto</typeparam>
		/// <param name="Valor">Valor (numerico) a comparar.</param>
		/// <param name="Desde">Rango Inicial</param>
		/// <param name="Hasta">Rango final</param>
		public static bool Between<T>(this T Valor, T Desde, T Hasta) where T : IComparable<T>
		{
			return Valor.CompareTo(Desde) >= 0 && Valor.CompareTo(Hasta) < 0;
		}


		/// <summary>Evalua si un determinado valor se encuentra entre una lista de valores.</summary>
		/// <param name="pVariable">Valor a Buscar.</param>
		/// <param name="pValores">Lista de Valores de Referencia. Ignora Mayusculas.</param>
		/// <returns>Devuelve 'True' si el valor existe en la lista al menos una vez.</returns>
		public static bool In(this String text, params string[] pValores)
		{
			bool retorno = false;
			try
			{
				foreach (string val in pValores)
				{
					if (text.Equals(val, StringComparison.InvariantCultureIgnoreCase))
					{ retorno = true; break; }
				}
			}
			catch { }
			return retorno;
		}
		/// <summary>Evalua si un determinado valor se encuentra entre una lista de valores.</summary>
		/// <param name="pVariable">Valor a Buscar.</param>
		/// <param name="pValores">Lista de Valores de Referencia.</param>
		/// <returns>Devuelve 'True' si el valor existe en la lista al menos una vez.</returns>
		public static bool In(this Int32 valor, params int[] pValores)
		{
			bool retorno = false;
			try
			{
				foreach (int val in pValores)
				{
					if (val == valor) { retorno = true; break; }
				}
			}
			catch { }
			return retorno;
		}


		#endregion

		#region Texto

		/// <summary>Devuelve la Cantidad de Palabras en una Cadena de Texto.</summary>
		/// <param name="str">Texto</param>
		public static int ContarPalabras(this String text)
		{
			return text.Split(new char[] { ' ', '.', '?', ',', '!', '-', '(', ')', '"', '\'' }, StringSplitOptions.RemoveEmptyEntries).Length;
		}

		/// <summary>Busca palabras en una cadena de Texto. Ignora Mayusculas.</summary>
		/// <param name="text">Cadena de Texto donde se Busca</param>
		/// <param name="pValores">Lista de Palabras a Buscar</param>
		/// <returns>Devuelve la cantidad de palabras encontradas.</returns>
		public static int Search(this String text, params string[] pValores)
		{
			int _ret = 0;
			try
			{
				var Palabras = text.Split(new char[] { ' ', '.', '?', ',', '!', '-', '(', ')', '"', '\'' },
					StringSplitOptions.RemoveEmptyEntries);

				foreach (string word in Palabras)
				{
					foreach (string palabra in pValores)
					{
						if (Regex.IsMatch(word, string.Format(@"\b{0}\b", palabra), RegexOptions.IgnoreCase))
						{
							_ret++;
						}
					}
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Devuelve una cadena que contiene un número especificado de caracteres desde el lado izquierdo de una cadena.</summary>
		/// <param name="str">Cadena de texto Original.</param>
		/// <param name="length">Indica cuántos caracteres se van a devolver. Si es 0, se devuelve una cadena de longitud cero (""). 
		/// Si es mayor o igual que el número de caracteres en 'text', se devuelve toda la cadena.</param>
		public static string Left(this String text, int length)
		{
			if (length < 0) return "";
			else if (length == 0 || text.Length == 0) return "";
			else if (text.Length <= length) return text;
			else return text.Substring(0, length);
		}

		/// <summary>Devuelve una cadena que contiene un número especificado de caracteres desde el lado derecho de una cadena.</summary>
		/// <param name="text">Cadena de texto Original.</param>
		/// <param name="length">Indica cuántos caracteres se van a devolver. Si es 0, se devuelve una cadena de longitud cero (""). 
		/// Si es mayor o igual que el número de caracteres en 'text', se devuelve toda la cadena.</param>
		public static string Right(this String text, int length)
		{
			if (length < 0) { return ""; }
			else if (length == 0 || text.Length == 0) { return ""; }
			else if (text.Length <= length) { return text; }
			else { return text.Substring(text.Length - length, length); }
		}

		/// <summary>Devuelve una porcion de texto dentro de una cadena.</summary>
		/// <param name="text">Cadena de texto Original.</param>
		/// <param name="startIndex">Posicion de inicio.</param>
		/// <param name="length">Cantidad de carácteres que se quieren extraer.</param>
		public static string Mid(this String text, int startIndex, int length)
		{
			string result = text.Substring(startIndex, length);
			return result;
		}
		public static string Mid(this String text, int startIndex)
		{
			string result = text.Substring(startIndex);
			return result;
		}

		/// <summary>Convierte la cadena especificada en Mayuscula Inicial.</summary>
		/// <param name="text">Texto a Cambiar</param>
		public static string ToTitleCase(this string text)
		{
			if (text == null) return text;

			System.Globalization.CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
			System.Globalization.TextInfo textInfo = cultureInfo.TextInfo;

			return textInfo.ToTitleCase(text.ToLower());
		}

		public static string RemoveLastCharacter(this String instr)
		{
			return instr.Substring(0, instr.Length - 1);
		}
		public static string RemoveLast(this String instr, int number)
		{
			return instr.Substring(0, instr.Length - number);
		}
		public static string RemoveFirstCharacter(this String instr)
		{
			return instr.Substring(1);
		}
		public static string RemoveFirst(this String instr, int number)
		{
			return instr.Substring(number);
		}


		/// <summary>Obtiene el valor de una cadena validando NULL, Elimina espacios sobrantes y la convierte a mayusculas.</summary>
		/// <param name="pString">Cadena de Texto a Obtener</param>
		/// <param name="ToUpper">Convierte la cadena a Mayusculas</param>
		/// <param name="LeadingSpace">Agrega un espacio al principio</param>		
		public static string GetString(this String pString, bool ToUpper = true, bool LeadingSpace = false)
		{
			string _ret = string.Empty;
			if (pString != null)
			{
				_ret = pString.Trim();

				if (ToUpper) _ret = _ret.ToUpper();

				if (LeadingSpace && _ret != string.Empty)
				{
					_ret = string.Format(" {0}", _ret);
				}
			}
			return _ret;
		}


		/// <summary>Reemplaza uno o varios elementos de la cadena especificada con los valores indicados.
		/// <para>Ejem: string X = "{0}{1}".Format(1, 2);</para> </summary>
		/// <param name="texto">Cadena de texto con Formato Compuesto.</param>
		/// <param name="args">Valores a insertar en la Cadena.</param>
		public static string Format(this String texto, params object[] args)
		{
			return string.Format(texto, args);
		}

		/// <summary>Formatea la Cadena a tipo Numerico entero con separadores de miles. ej: '0.000.000'</summary>
		/// <param name="texto">Cadena de texto a Formatear</param>
		public static string FormatNumber(this String texto, string Mask = "n0")
		{
			return string.Format("{0:" + Mask + "}", texto.ToInteger());
		}

		/// <summary>Formatea la Cadena a tipo Numerico entero con separadores de miles y 2 decimales. ej: '0.000.000,00'</summary>
		/// <param name="texto">Cadena de texto a Formatear</param>
		public static string FormatDecimal(this string texto, string Mask = "n2")
		{
			return string.Format("{0:" + Mask + "}", texto.ToDecimal());
		}

		/// <summary>Formatea la Cadena a tipo Numerico entero con separadores de miles y 2 decimales. ej: '0.000.000,00'</summary>
		/// <param name="texto">Cadena de texto a Formatear</param>
		public static string FormatCurrency(this string texto)
		{
			return string.Format("$ {0:n2}", texto.ToDecimal());
		}

		/// <summary>Formats the string according to the specified mask</summary>
		/// <param name="input">The input string.</param>
		/// <param name="mask">The mask for formatting. Like "A##-##-T-###Z"</param>
		/// <returns>The formatted string</returns>
		public static string FormatWithMask(this string input, string mask)
		{
			if (input.IsNullOrEmpty()) return input;
			var output = string.Empty;
			var index = 0;
			foreach (var m in mask)
			{
				if (m == '#')
				{
					if (index < input.Length)
					{
						output += input[index];
						index++;
					}
				}
				else
					output += m;
			}
			return output;
		}


		/// <summary>Devuelve la Cadena de texto con sus caracteres invertidos.</summary>
		/// <param name="input">Texto</param>
		public static string Reverse(this string input)
		{
			if (string.IsNullOrWhiteSpace(input)) return string.Empty;
			char[] chars = input.ToCharArray();
			Array.Reverse(chars);
			return new String(chars);
		}


		/// <summary>Valida que la cadena tenga el formato de una direccion de Correo Electronico.
		/// No intenta enviar ningun mensaje para verificar su existencia.</summary>
		/// <param name="input"></param>
		public static bool isEmail(this string input)
		{
			var match = Regex.Match(input,
			  @"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
			return match.Success;
		}

		public static bool isNumber(this string input)
		{
			var match = Regex.Match(input, @"^[0-9]+$", RegexOptions.IgnoreCase);
			return match.Success;
		}

		public static bool IsDate(this string input)
		{
			if (!string.IsNullOrEmpty(input))
			{
				DateTime dt;
				return (DateTime.TryParse(input, out dt));
			}
			else
			{
				return false;
			}
		}

		public static bool IsUnicode(this string value)
		{
			int asciiBytesCount = System.Text.Encoding.ASCII.GetByteCount(value);
			int unicodBytesCount = System.Text.Encoding.UTF8.GetByteCount(value);

			if (asciiBytesCount != unicodBytesCount)
			{
				return true;
			}
			return false;
		}

		/// <summary>Devuelve 'true' si la cadena es Nula o Vacia.</summary>
		/// <param name="input">Cadena de Texto a Validar</param>
		public static bool IsNullOrEmpty(this String input)
		{
			bool _ret = true;

			if (input != null && input != string.Empty) _ret = false;

			return _ret;
		}

		public static bool IsBoolean(this string value)
		{
			var val = value.ToLower().Trim();
			if (val == "false")
				return true;
			if (val == "f")
				return true;
			if (val == "true")
				return true;
			if (val == "t")
				return true;
			if (val == "yes")
				return true;
			if (val == "no")
				return true;
			if (val == "y")
				return true;
			if (val == "n")
				return true;
			if (val == "s")
				return true;
			if (val == "si")
				return true;

			return false;
		}

		/// <summary>Devuelve un Valor por Defecto si la cadena es Nula o Vacia.</summary>
		/// <param name="str">Cadena de Texto</param>
		/// <param name="defaultValue">Valor x Defecto</param>
		/// <param name="considerWhiteSpaceIsEmpty">Los Espacios se consideran como Vacio?</param>
		public static string DefaultIfEmpty(this string str, string defaultValue, bool considerWhiteSpaceIsEmpty = false)
		{
			return (considerWhiteSpaceIsEmpty ? string.IsNullOrWhiteSpace(str) : string.IsNullOrEmpty(str)) ? defaultValue : str;
		}

		/// <summary>Elimina todos los caracteres de formato de una cadena, tales como: ().,-$</summary>
		/// <param name="pCadenaFormateada">Cadena con Formato.</param>
		/// <returns>Devuelve la cadena sin formato.</returns>
		public static string LimpiarFormato(string pCadenaFormateada)
		{
			string retorno = "";
			try
			{
				retorno = pCadenaFormateada.Replace(".", null);
				retorno = retorno.Replace(",", null);
				retorno = retorno.Replace("-", null);
				retorno = retorno.Replace("_", null);
				retorno = retorno.Replace("$", null);
				retorno = retorno.Replace("(", null);
				retorno = retorno.Replace(")", null);
				retorno = retorno.Replace("/", null);
				retorno = retorno.Replace(":", null);
				retorno = retorno.Replace(";", null);
				retorno = retorno.Trim();
			}
			catch { }
			return retorno;
		}

		#endregion

		#region Numericos

		/// <summary>Trata de convertir una cadena de texto a un valor numerico entero.</summary>
		/// <param name="input">Texto a convertir</param>
		/// <param name="throwExceptionIfFailed"></param>
		public static int ToInteger(this string input, bool throwExceptionIfFailed = false)
		{
			int result;
			var valid = int.TryParse(LimpiarFormato(input), out result);
			if (!valid)
				if (throwExceptionIfFailed)
					throw new FormatException(string.Format("'{0}' cannot be converted as int", input));
			return result;
		}

		/// <summary>Trata de convertir una cadena de texto a un valor numerico entero.</summary>
		/// <param name="input">Texto a convertir</param>
		/// <param name="throwExceptionIfFailed"></param>
		public static long ToLong(this string input, bool throwExceptionIfFailed = false)
		{
			long result;
			var valid = long.TryParse(LimpiarFormato(input), out result);
			if (!valid)
				if (throwExceptionIfFailed)
					throw new FormatException(string.Format("'{0}' cannot be converted as int", input));
			return result;
		}

		/// <summary>Convierte una Cadena de Texto a su representacion Numerica Decimal.</summary>
		/// <param name="input">Texto a convertir</param>
		/// <param name="throwExceptionIfFailed"></param>
		public static decimal ToDecimal(this string input, bool throwExceptionIfFailed = false)
		{
			decimal result;
			var valid = decimal.TryParse(input, System.Globalization.NumberStyles.AllowDecimalPoint,
			  new System.Globalization.NumberFormatInfo { NumberDecimalSeparator = "." }, out result);
			if (!valid)
				if (throwExceptionIfFailed)
					throw new FormatException(string.Format("'{0}' cannot be converted as decimal", input));
			return result;
		}


		/// <summary>Genera Numeros Enteros Consecutivos en el Rango especificado.</summary>
		/// <param name="MinValue">Valor menor del Rango.</param>
		/// <param name="MaxValue">Valor Mayor del Rango.</param>
		public static List<int> NumerosEnRango(this Int32 MinValue, Int32 MaxValue)
		{
			List<int> _ret = new List<int>();
			try
			{
				int inicio = IIf(MaxValue > MinValue, MinValue, MaxValue);
				int final = IIf(MaxValue > MinValue, MaxValue, MinValue);

				for (int i = inicio; i <= final; i++)
				{
					_ret.Add(i);
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Generador de numeros aleatorios.
		/// Al ser declarado a nivel de clase, mejora la calidad de los numeros aleatorios generados.</summary>
		private static Random RND = new Random();

		/// <summary>Obtiene un Número aleatorio.
		/// Si el Numero Base es Diferente a Cero se usará como Semilla.</summary>
		/// <param name="Numero">(Seed) Un número usado para calcular un valor inicial para la secuencia numérica pseudoaleatoria. Si se especifica un número negativo, se usa el valor absoluto del número.</param>
		public static int RandomNumber(this Int32 Numero)
		{
			int _ret = 0;
			try
			{
				if (Numero != 0)
				{
					_ret = new Random(Numero).Next();
				}
				else
				{
					_ret = RND.Next();
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Obtiene un Número aleatorio entre 0.0 y 1.0
		/// Si el Numero Base es Diferente a Cero se usará como Semilla.
		/// Ejem: double DD = new double().RandomNumber();</summary>
		/// <param name="Numero">(Seed) Un número usado para calcular un valor inicial para la secuencia numérica pseudoaleatoria. Si se especifica un número negativo, se usa el valor absoluto del número.</param>
		public static double RandomNumber(this Double Numero)
		{
			double _ret = 0;
			try
			{
				if (Numero > 0)
				{
					_ret = new Random(Convert.ToInt32(Numero)).NextDouble();
				}
				else
				{
					_ret = RND.NextDouble();
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Obtiene un Número aleatorio entre el Rango especificado.
		/// Si el Numero Base es Diferente a Cero se usará como Semilla.
		/// Ejem: int G = new Int32().RandomBetween(0, 10);</summary>
		/// <param name="Numero">(Seed) Un número usado para calcular un valor inicial para la secuencia numérica pseudoaleatoria. Si se especifica un número negativo, se usa el valor absoluto del número.</param>
		/// <param name="MinValue">Valor Minimo del Rango</param>
		/// <param name="MaxValue">Valor Maximo del Rango</param>
		public static int RandomBetween(this Int32 Numero, int MinValue, int MaxValue)
		{
			int _ret = 0;
			try
			{
				if (Numero != 0)
				{
					_ret = new Random(Numero).Next(MinValue, MaxValue + 1);
				}
				else
				{
					_ret = RND.Next(MinValue, MaxValue + 1); //<- MaxValue No es Inclusivo
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Obtiene un Número aleatorio entre el Rango especificado.
		/// Si el Numero Base es Diferente a Cero se usará como Semilla.
		/// Ejem: double D = new double().RandomBetween(1, 10);</summary>
		/// <param name="Numero">(Seed) Un número usado para calcular un valor inicial para la secuencia numérica pseudoaleatoria. Si se especifica un número negativo, se usa el valor absoluto del número.</param>
		/// <param name="MinValue">Valor Minimo del Rango</param>
		/// <param name="MaxValue">Valor Maximo del Rango</param>
		public static double RandomBetween(this Double Numero, int MinValue, int MaxValue)
		{
			double _ret = 0;
			try
			{
				if (Numero != 0)
				{
					_ret = new Random(Convert.ToInt32(Numero)).NextDouble() * MaxValue;
				}
				else
				{
					_ret = RND.NextDouble() * MaxValue; //<- MaxValue No es Inclusivo
				}
			}
			catch { }
			return _ret;
		}

		/// <summary>Reordena al Azar la lista de elementos.</summary>
		/// <typeparam name="T">Cualquier tipo de Objeto.</typeparam>
		/// <param name="list">Lista de elementos a reordenar.</param>
		public static void RandomizeOrder<T>(this IList<T> list)
		{
			int n = list.Count;
			while (n > 1)
			{
				n--;
				int k = RND.Next(n + 1);
				T value = list[k];
				list[k] = list[n];
				list[n] = value;
			}
		}

		#endregion

		#region Fechas

		/// <summary>Convierte una cadena de texto a Fecha.</summary>
		/// <param name="input">Texto a Convertir</param>
		/// <param name="throwExceptionIfFailed"></param>
		public static DateTime ToDate(this string input, bool throwExceptionIfFailed = false)
		{
			DateTime result;
			var valid = DateTime.TryParse(input, out result);
			if (!valid)
				if (throwExceptionIfFailed)
					throw new FormatException(string.Format("'{0}' cannot be converted as DateTime", input));
			return result;
		}

		/// <summary>Determina si el valor de la Variable se encuentra dentro del Rango especificado.</summary>
		/// <typeparam name="T">Tipo de Datos del  Objeto</typeparam>
		/// <param name="Valor">Valor (Fecha) a comparar.</param>
		/// <param name="Desde">Rango Inicial</param>
		/// <param name="Hasta">Rango final</param>
		public static bool Between(this DateTime Fecha, DateTime Inicio, DateTime Final)
		{
			return Fecha.Ticks >= Inicio.Ticks && Fecha.Ticks <= Final.Ticks;
		}

		/// <summary>Compara la Fecha con el dia actual y Convierte la diferencia en un texto humanamente leible.</summary>
		/// <param name="value">Fecha a Convertir</param>
		public static string ToReadableTime(this DateTime value)
		{
			string _ret = string.Empty;
			string _prefix = string.Empty;
			var ts = new TimeSpan(DateTime.Now.Ticks - value.Ticks);

			if (value < DateTime.Now) //<- Fecha Anterior a Hoy
			{
				_prefix = "hace";
				ts = DateTime.Now.Subtract(value); //<- Tiempo Transcurrido    
			}
			else
			{
				_prefix = "dentro de";
				ts = value.Subtract(DateTime.Now); //<- Tiempo Transcurrido      
			}

			double delta = Math.Abs(ts.TotalSeconds);
			if (delta > 31104000) // 12 * 30 * 24 * 60 * 60
			{
				var years = Convert.ToInt32(Math.Floor((double)ts.Days / 365));
				_ret = years <= 1 ? string.Format("{0} un año", _prefix) : string.Format("{0} {1} años", _prefix, years);
			}
			if (delta <= 31104000) // 12 * 30 * 24 * 60 * 60
			{
				int months = Convert.ToInt32(Math.Floor((double)ts.Days / 30));
				_ret = months <= 1 ? string.Format("{0} un mes", _prefix) : string.Format("{0} {1} meses", _prefix, months);
			}
			if (delta < 2592000) // 30 * 24 * 60 * 60
			{
				_ret = string.Format("{0} {1} días", _prefix, ts.Days);
			}
			if (delta < 172800) // 48 * 60 * 60
			{
				_ret = "Ayer";
			}
			if (delta < 86400) // 24 * 60 * 60
			{
				_ret = "Hoy";
			}
			if (delta < 43200) // 12 * 60 * 60
			{
				_ret = string.Format("{0} {1} horas", _prefix, ts.Hours);
			}
			if (delta < 5400) // 90 * 60
			{
				_ret = string.Format("{0} una hora", _prefix);
			}
			if (delta < 2700) // 45 * 60
			{
				_ret = string.Format("{0} {1} minutos", _prefix, ts.Minutes);
			}
			if (delta < 2100) // 30 * 60
			{
				_ret = string.Format("{0} media hora", _prefix);
			}
			if (delta < 1300) // 15 * 60
			{
				_ret = string.Format("{0} {1} minutos", _prefix, ts.Minutes);
			}
			if (delta < 420) //7 * 60  
			{
				_ret = string.Format("{0} 5 minutos", _prefix);
			}
			if (delta < 180) //3 * 60  <- 3 minutos
			{
				_ret = string.Format("{0} {1} minutos", _prefix, ts.Minutes);
			}
			if (delta < 90) //1 minuto 
			{
				_ret = string.Format("{0} un minuto", _prefix);
			}
			if (delta < 40) //30 segundos
			{
				_ret = string.Format("{0} {1} segundos", _prefix, ts.Seconds);
			}
			if (delta < 20) //20 segundos
			{
				_ret = string.Format("{0} un momento", _prefix);
			}
			return _ret;
		}

		public static bool IsWorkingDay(this DateTime date)
		{
			//Determina si la fecha indicada es un dia laborable
			return date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday;
		}
		public static bool IsWeekend(this DateTime date)
		{
			//Determina si la Fecha indicada es en un fin de semana
			return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
		}
		public static DateTime NextWorkday(this DateTime date)
		{
			//Devuelve la Fecha del siguiente dia laborable
			var nextDay = date;
			while (!nextDay.IsWorkingDay())
			{
				nextDay = nextDay.AddDays(1);
			}
			return nextDay;
		}

		/// <summary>Obtiene el nombre del Dia de la Semana para la Fecha Indicada.</summary>
		public static string GetDayName(this DateTime date)
		{
			string _ret = string.Empty; //para .NET Framework 4++
			var culture = new System.Globalization.CultureInfo("es-419"); //<- 'es-419' = Spanish (Latin America), 'es-UY' = Spanish (Uruguay)
			_ret = culture.DateTimeFormat.GetDayName(date.DayOfWeek);
			//Convierte el texto a Mayuscula Inicial
			System.Globalization.TextInfo textInfo = culture.TextInfo;
			_ret = textInfo.ToTitleCase(_ret.ToLower());
			return _ret;
		}

		/// <summary>Obtiene el nombre del Mes para la Fecha Indicada.</summary>
		public static string GetMonthName(this DateTime date)
		{
			string _ret = string.Empty; //para .NET Framework 4++
			var culture = new System.Globalization.CultureInfo("es-419"); //<- 'es-419' = Spanish (Latin America), 'es-UY' = Spanish (Uruguay)
			_ret = culture.DateTimeFormat.GetMonthName(date.Month);
			//Convierte el texto a Mayuscula Inicial
			System.Globalization.TextInfo textInfo = culture.TextInfo;
			_ret = textInfo.ToTitleCase(_ret.ToLower());
			return _ret;
		}

		/// <summary>Carga los Feriados (Estaticos) Uruguayos para el Año de la Fecha indicada.</summary>
		/// <param name="date">Fecha Indicada</param>
		public static List<DateTime> CargarFeriados(this DateTime date)
		{
			List<DateTime> _ret = new List<DateTime>();

			_ret.Add(new DateTime(date.Year, 1, 1)); //<- Año Nuevo
			_ret.Add(new DateTime(date.Year, 1, 6)); //<- Dia de Reyes

			//Agregar la Semana Santa:
			DateTime Pascua = DomingoPascua(date.Year);
			_ret.Add(Pascua.AddDays(-3));   //<- Jueves Santo
			_ret.Add(Pascua.AddDays(-2));   //<- Viernes Santo
			_ret.Add(Pascua);               //<- Domingo de Pascua

			_ret.Add(new DateTime(date.Year, 4, 22)); //<- Desembarco de los 33 orientales
			_ret.Add(new DateTime(date.Year, 5, 1)); //<- Día de los Trabajadores
			_ret.Add(new DateTime(date.Year, 5, 18)); //<- Batalla de las Piedras
			_ret.Add(new DateTime(date.Year, 6, 19)); //<- Natalicio de Artigas
			_ret.Add(new DateTime(date.Year, 7, 18)); //<- Día de la Constitución
			_ret.Add(new DateTime(date.Year, 8, 25)); //<- Día de la Independencia
			_ret.Add(new DateTime(date.Year, 10, 12)); //<- Día de la Raza
			_ret.Add(new DateTime(date.Year, 11, 2)); //<-Día de los Difuntos
			_ret.Add(new DateTime(date.Year, 12, 25)); //<- Navidad

			return _ret;
		}

		#region Calcula la Semana Santa

		private struct ParConstantes
		{
			public int M { get; set; }
			public int N { get; set; }
		}
		private static DateTime DomingoPascua(int anio)
		{
			DateTime pascuaResurreccion;
			int a, b, c, d, e;

			ParConstantes p = new ParConstantes();
			if (anio < 1583)
			{
				throw new ArgumentOutOfRangeException("El año deberá ser superior a 1583");
			}
			else if (anio < 1700) { p.M = 22; p.N = 2; }
			else if (anio < 1800) { p.M = 23; p.N = 3; }
			else if (anio < 1900) { p.M = 23; p.N = 4; }
			else if (anio < 2100) { p.M = 24; p.N = 5; }
			else if (anio < 2200) { p.M = 24; p.N = 6; }
			else if (anio < 2299) { p.M = 25; p.N = 0; }
			else
			{
				throw new ArgumentOutOfRangeException("El año deberá ser inferior a 2299");
			}

			a = anio % 19;
			b = anio % 4;
			c = anio % 7;
			d = (19 * a + p.M) % 30;
			e = (2 * b + 4 * c + 6 * d + p.N) % 7;

			if (d + e < 10)
				pascuaResurreccion = new DateTime(anio, 3, d + e + 22);
			else
				pascuaResurreccion = new DateTime(anio, 4, d + e - 9);

			// Excepciones
			if (pascuaResurreccion == new DateTime(anio, 4, 26))
				pascuaResurreccion = new DateTime(anio, 4, 19);

			if (pascuaResurreccion == new DateTime(anio, 4, 25)
					&& d == 28 && e == 6 && a > 10)
				pascuaResurreccion = new DateTime(anio, 4, 18);

			return pascuaResurreccion;
		}
		#endregion

		#endregion

		#endregion

		#region Manejo de Archivos

		/// <summary>Constantes para los Codigos de Pagina al leer o guardar archivos de texto.</summary>
		public enum TextEncoding
		{
			/// <summary>CodePage:1252; windows-1252 ANSI Latin 1; Western European (Windows)</summary>
			ANSI = 1252,
			/// <summary>CodePage:850; ibm850; ASCII Multilingual Latin 1; Western European (DOS)</summary>
			DOS_850 = 850,
			/// <summary>CodePage:1200; utf-16; Unicode UTF-16, little endian byte order (BMP of ISO 10646);</summary>
			Unicode = 1200,
			/// <summary>CodePage:65001; utf-8; Unicode (UTF-8)</summary>
			UTF8 = 65001
		}

		/// <summary>Guarda Datos en un Archivo de Texto usando la Codificacion especificada.</summary>
		/// <param name="FilePath">Ruta de acceso al Archivo. Si no existe, se Crea. Si existe, se Sobreescribe.</param>
		/// <param name="Data">Datos a Grabar en el Archivo.</param>
		/// <param name="CodePage">[Opcional] Pagina de Codigos con la que se guarda el archivo. Por defecto se usa Unicode(UTF-16).</param>
		public static bool SaveTextFile(string FilePath, string Data, TextEncoding CodePage = TextEncoding.Unicode)
		{
			bool _ret = false;
			try
			{
				if (FilePath != null && FilePath != string.Empty)
				{
					/* ANSI code pages, like windows-1252, can be different on different computers, 
					 * or can be changed for a single computer, leading to data corruption. 
					 * For the most consistent results, applications should use UNICODE, 
					 * such as UTF-8 or UTF-16, instead of a specific code page. 
					 https://docs.microsoft.com/es-es/windows/desktop/Intl/code-page-identifiers  */

					System.Text.Encoding ENCODING = System.Text.Encoding.GetEncoding((int)CodePage); //<- Unicode Garantiza Maxima compatibilidad
					using (System.IO.FileStream FILE = new System.IO.FileStream(FilePath, System.IO.FileMode.Create))
					{
						using (System.IO.StreamWriter WRITER = new System.IO.StreamWriter(FILE, ENCODING))
						{
							WRITER.Write(Data);
							WRITER.Close();
						}
					}
					if (System.IO.File.Exists(FilePath)) _ret = true;
				}
			}
			catch (Exception ex) { throw ex; }
			return _ret;
		}

		/// <summary>Lee un Archivo de Texto usando la Codificacion especificada.</summary>
		/// <param name="FilePath">Ruta de acceso al Archivo. Si no existe se produce un Error.</param>
		/// <param name="CodePage">Pagina de Codigos con la que se Leerá el archivo.</param>
		public static string ReadTextFile(string FilePath, TextEncoding CodePage)
		{
			string _ret = string.Empty;
			try
			{
				if (FilePath != null && FilePath != string.Empty)
				{
					if (System.IO.File.Exists(FilePath))
					{
						System.Text.Encoding ENCODING = System.Text.Encoding.GetEncoding((int)CodePage);
						_ret = System.IO.File.ReadAllText(FilePath, ENCODING);
					}
					else { throw new Exception(string.Format("ERROR 404: Archivo '{0}' NO Encontrado!", FilePath)); }
				}
				else { throw new Exception("No se ha Especificado la Ruta de acceso al Archivo!"); }
			}
			catch (Exception ex) { throw ex; }
			return _ret;
		}

		/// <summary>Convierte el tamaño de un archivo a la unidad más adecuada.</summary>
		/// <param name="pFileBytes">Tamaño del Archivo en Bytes</param>
		/// <returns>"0.### XB", ejem. "4.2 KB" or "1.434 GB"</returns>
		public static string GetFileSize(long pFileBytes)
		{
			// Get absolute value
			long absolute_i = (pFileBytes < 0 ? -pFileBytes : pFileBytes);
			// Determine the suffix and readable value
			string suffix;
			double readable;
			if (absolute_i >= 0x1000000000000000) // Exabyte
			{
				suffix = "EB";
				readable = (pFileBytes >> 50);
			}
			else if (absolute_i >= 0x4000000000000) // Petabyte
			{
				suffix = "PB";
				readable = (pFileBytes >> 40);
			}
			else if (absolute_i >= 0x10000000000) // Terabyte
			{
				suffix = "TB";
				readable = (pFileBytes >> 30);
			}
			else if (absolute_i >= 0x40000000) // Gigabyte
			{
				suffix = "GB";
				readable = (pFileBytes >> 20);
			}
			else if (absolute_i >= 0x100000) // Megabyte
			{
				suffix = "MB";
				readable = (pFileBytes >> 10);
			}
			else if (absolute_i >= 0x400) // Kilobyte
			{
				suffix = "KB";
				readable = pFileBytes;
			}
			else
			{
				return pFileBytes.ToString("0 B"); // Byte
			}

			readable = System.Math.Round((readable / 1024), 2);
			return string.Format("{0:n1} {1}", readable, suffix);
		}
		
		#endregion

	}
}