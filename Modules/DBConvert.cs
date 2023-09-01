using System;
using System.Data;
using System.Threading.Tasks;
using System.Data.OleDb;
using SoftCircuits.IniFileParser;
using System.IO;
using System.Data.SQLite;
using System.Diagnostics;

namespace TSI.SIREP.Modules
{
    internal class DBConvert
    {
        #region Declares
        private readonly string _INI = AppDomain.CurrentDomain.BaseDirectory + @"settings.ini";
        private readonly string _SQLPath = AppDomain.CurrentDomain.BaseDirectory + @"sirep.db";
        private readonly string _XTableEjemplar, _XTableFichas;
        private string _FilePath, shaDB;
        private bool _IsFileOpen, _IsOutdated = false, _IsNew = false;
        private bool _forceConvert = false, _ignoreConvert = false;

        private DataTable[] SQLContents, SQLBackup;
        private DataTable _DataNew, _DataSplit;

        private OleDbConnection _db;
        private SQLiteConnection _sql, _sqlBackup;
        private readonly IniFile _INIFile;
        #endregion

        public DBConvert(IniFile Value)
        {
            _INIFile = Value;
            _FilePath = "";
            _IsFileOpen = false;
            _XTableEjemplar = _INIFile.GetSetting("SQLValues", "Sheet1");
            _XTableFichas = _INIFile.GetSetting("SQLValues", "Sheet2");
            if (_XTableEjemplar == null || _XTableFichas == null)
            {
                _XTableEjemplar = "Ejemplares";
                _XTableFichas = "FICHAS";
                _INIFile.SetSetting("SQLValues", "Sheet1", _XTableEjemplar);
                _INIFile.SetSetting("SQLValues", "Sheet2", _XTableFichas);
                _INIFile.Save(_INI);
            }

        }

        #region Methods
        /// <summary>
        /// Procedimiento que guarda la base de datos procesada en memoria en archivo local
        /// </summary>
        /// <returns></returns>
        public async Task<bool> FlushData()
        {
            if (!IsFileOpen) return false;
            if (!_IsOutdated) return false;
            if (_ignoreConvert) return false;
            int indexUpdate = 0, indexListUpdate = 0;
            bool contentFound = false, updatesFound = false, updateList = false;
            if (!_IsNew)
            {
                // Si el archivo no es nuevo, entonces reemplazamos las tablas contenidas en SQLContents por las de SQLBackup
                int i = 0;
                foreach (DataTable table in SQLContents)
                {
                    if (table.TableName == _XTableEjemplar) SQLContents[i] = SQLBackup[0];
                    if (table.TableName == _XTableFichas) SQLContents[i] = SQLBackup[1];
                    if (table.TableName == "UPDATELIST")
                    {
                        updateList = true;
                        indexListUpdate = i;
                    }
                    if (table.TableName == "UPDATES")
                    {
                        updatesFound = true;
                        indexUpdate = i;
                    }
                    if (table.TableName == "CONTENTS")
                    {
                        contentFound = true;
                        SQLContents[i] = _DataSplit;
                    }

                    i++;
                }
            }

            // Si no se encuentra la tabla de actualizaciones, creamos una nueva en el array
            // Mientras no sea nueva la base de datos
            if (!_IsNew)
            {
                if (!updatesFound)
                {
                    DataTable[] _d = new DataTable[SQLContents.Length + 1];
                    for (int i = 0; i < _d.Length; i++)
                    {
                        if (i == _d.Length - 1)
                        {
                            _d[i] = _DataNew;
                        }
                        else { _d[i] = SQLContents[i]; }
                    }
                    SQLContents = _d;
                }
                else
                {
                    for (int i = 0; i < _DataNew.Rows.Count; i++)
                    {
                        SQLContents[indexUpdate].Rows.Add(_DataNew.Rows[i].ItemArray);
                    }
                }
            }

            if (!contentFound)
            {
                DataTable[] _d = new DataTable[SQLContents.Length + 1];
                for (int i = 0; i < _d.Length; i++)
                {
                    if (i == _d.Length - 1)
                    {
                        _d[i] = _DataSplit;
                    }
                    else { _d[i] = SQLContents[i]; }
                }
                SQLContents = _d;
            }

            // Verificamos la existencia de la tabla deseada donde guardamos el registro de actualizaciones
            // Si no existe sólo la creamos
            if (updateList)
            {
                SQLContents[indexListUpdate].Rows.Add(shaDB, DateTime.Now.ToString(), _DataNew.Rows.Count);
            }
            else
            {
                DataTable dt = new DataTable("UPDATELIST");
                dt.Columns.Add("HASHUpdate", typeof(string));
                dt.Columns.Add("DateUpdated", typeof(string));
                dt.Columns.Add("RowsModified", typeof(int));

                DataTable[] _d = new DataTable[SQLContents.Length + 1];
                for (int i = 0; i < _d.Length; i++)
                {
                    if (i == _d.Length - 1)
                    {
                        _d[i] = dt;
                    }
                    else { _d[i] = SQLContents[i]; }
                }
                SQLContents = _d;
            }

            // Almacenamos el contenido del array DataTable en la conexión sql backup
            SQLiteConnection _temp = _sqlBackup;
            if (_IsNew) _temp = _sql;
            using (SQLiteCommand cmd = new SQLiteCommand(_temp))
            {
                foreach (DataTable table in SQLContents)
                {
                    cmd.CommandText = $"CREATE TABLE [{table.TableName}] ({GetCreateTableColumns(table)})";
                    await cmd.ExecuteNonQueryAsync();

                    int index = 0;
                    string letcmd = "";
                    foreach (DataRow row in table.Rows)
                    {
                        letcmd += $"INSERT INTO [{table.TableName}] VALUES ({GetInsertValues(row)});";
                        if (index == 1000)
                        {
                            cmd.CommandText = letcmd;
                            await cmd.ExecuteNonQueryAsync();
                            letcmd = "";
                            index = 0;
                        }
                        index++;
                    }

                    if (index > 0)
                    {
                        cmd.CommandText = letcmd;
                        await cmd.ExecuteNonQueryAsync();
                    }

                }
            }

            return true;
        }

        /// <summary>
        /// Separa los registros existentes en columnas específicas para poderse procesar después
        /// </summary>
        /// <returns>Devuelve true si se completó la acción</returns>
        public bool SplitData()
        {
            if (!IsFileOpen) return false;
            if (_IsOutdated)
            {
                _DataSplit = new DataTable("CONTENTS");

                _DataSplit.Columns.Add("MARC001", typeof(int));     // número de ficha
                _DataSplit.Columns.Add("MARC082", typeof(string));  // clasificación
                _DataSplit.Columns.Add("MARC100", typeof(string));  // autor por apellido
                _DataSplit.Columns.Add("MARC245", typeof(string));  // título
                _DataSplit.Columns.Add("MARC250", typeof(string));  // n° edición
                _DataSplit.Columns.Add("MARC260", typeof(string));  // país y editorial
                _DataSplit.Columns.Add("MARC300", typeof(string));  // # páginas, dimensiones
                _DataSplit.Columns.Add("MARC008", typeof(string));  // año
                _DataSplit.Columns.Add("MARC440", typeof(string));  // colección y número
                _DataSplit.Columns.Add("MARC500", typeof(string));  // contenidos, otros títulos, donante
                _DataSplit.Columns.Add("MARC020", typeof(string));  // ISBN
                _DataSplit.Columns.Add("MARC650", typeof(string));  // tags1
                _DataSplit.Columns.Add("MARC700", typeof(string));  // tags2

                foreach (DataRow row in SQLBackup[1].Rows)
                {
                    string EtiquetasMARC = row["EtiquetasMARC"].ToString().Substring(1).TrimEnd('Ì');
                    string[] data = EtiquetasMARC.Split('¦');

                    string[] insertRow = new string[13];
                    insertRow[0] = row["Ficha_No"].ToString();
                    foreach (string label in data)
                    {
                        try
                        {
                            string padleft = label.Substring(0, 3);
                            string content = label.Substring(3, label.Length - 3);
                            if (content.Length > 0)
                            {
                                content = content.TrimEnd(' ');
                                content = content.TrimStart(' ');
                            }
                            switch (padleft)
                            {
                                case "082": // Clasificación
                                    insertRow[1] = content;
                                    break;
                                case "100": // Autor por apellido
                                    insertRow[2] = content;
                                    break;
                                case "245": // Título
                                    insertRow[3] = content;
                                    break;
                                case "250": //N° edición
                                    insertRow[4] = content;
                                    break;
                                case "260": // País y Editorial
                                    insertRow[5] = content;
                                    break;
                                case "300": // # páginas, dimensiones
                                    insertRow[6] = content;
                                    break;
                                case "008": // Año
                                    insertRow[7] = content;
                                    break;
                                case "440": // Colección y número
                                    insertRow[8] = content;
                                    break;
                                case "500": // Contenidos, otros títulos, donante
                                    insertRow[9] = content;
                                    break;
                                case "020": // ISBN
                                    insertRow[10] = content;
                                    break;
                                case "650": // Tags1
                                    insertRow[11] = content;
                                    break;
                                case "700": // Tags2
                                    insertRow[12] = content;
                                    break;
                                case null:
                                    break;
                            }
                        }
                        catch { }
                    }
                    _DataSplit.Rows.Add(insertRow);
                }
            }
            return true;
        }

        /// <summary>
        /// Compara los datos nuevos, eliminados o modificados de la base de datos, si es que se haya actualizado la MDB de origen y se haya generado un backup de éste en SQLite
        /// </summary>
        /// <returns></returns>
        public bool CompareData()
        {
            if (!IsFileOpen) return false;
            if (_IsNew) return false;
            if (_IsOutdated)
            {
                // Este procedimiento compara los datos de la MDB en memoria
                // Con los existentes en la SQL, obteniendo los datos para almacenarlos en otra tabla
                
                // Asignamos los parámetros para la tabla de modificaciones
                _DataNew = new DataTable();
                bool added;

                _DataNew.TableName = "UPDATES";
                _DataNew.Columns.Add("id", typeof(int));
                _DataNew.Columns.Add("oldMARCdata", typeof(string));
                _DataNew.Columns.Add("state", typeof(string));
                _DataNew.Columns.Add("dateprocessed", typeof(DateTime));
                _DataNew.Columns.Add("hashupdate", typeof(string));

                foreach(DataTable table in SQLContents)
                {
                    if (table.TableName == _XTableFichas)
                    {
                        foreach (DataRow rowBackup in SQLBackup[1].Rows)
                        {
                            added = true;
                            foreach (DataRow row in table.Rows)
                            {
                                // Buscamos que se encuentre la ficha, y se ejecuta cuando se localiza
                                if (row["Ficha_No"].ToString() == rowBackup["Ficha_No"].ToString())
                                {
                                    added = false;
                                    if (row["EtiquetasMARC"].ToString() != rowBackup["EtiquetasMARC"].ToString())
                                    {
                                        _DataNew.Rows.Add(row["Ficha_No"].ToString(), row["EtiquetasMARC"].ToString(), "modified", DateTime.Now, shaDB);
                                    }
                                    break;
                                }
                            }
                            if (added)
                            {
                                // Significa que es nuevo el registro
                                _DataNew.Rows.Add(rowBackup["Ficha_No"].ToString(), rowBackup["EtiquetasMARC"].ToString(), "added", DateTime.Now, shaDB);
                            }
                        }
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Este procedimiento carga la base de datos MDB a memoria, para procesarla desde ahí
        /// </summary>
        /// <returns>Devuelve true si se completó la acción correctamente</returns>
        public async Task<bool> LoadDBtoMemory()
        {
            if (!IsFileOpen) return false;
            if (_IsOutdated)
            {
                string[] Tables = { _XTableEjemplar, _XTableFichas };
                int i = 0;

                // Redimensionamos a 2 el DataTableBackup, que será ahí donde cargaremos la MDB
                SQLBackup = new DataTable[2];
                // Cargamos las tablas de la MDB en DataTable Backup
                foreach (string table in Tables)
                {
                    using (OleDbCommand _cmd = new OleDbCommand($"SELECT * FROM [{table}]", _db))
                    {
                        using (OleDbDataAdapter _data = new OleDbDataAdapter(_cmd))
                        {
                            SQLBackup[i] = new DataTable(table);
                            _data.Fill(SQLBackup[i]);
                            i++;
                        }
                    }
                }
                // Cerramos la conexión de la MDB, para liberar recursos
                _db.Close();

                // Redimensionamos el DataTable global de acuerdo a las tablas contenidas
                if (_IsNew)
                {
                    // Como es nueva la base de datos, se procede a redimensionarlo a 2
                    SQLContents = new DataTable[2];
                    // Y le asignamos las tablas anteriores como parámetros, y salimos del procedimiento
                    SQLContents[0] = SQLBackup[0];
                    SQLContents[1] = SQLBackup[1];
                }
                else
                {
                    // Se procede a contar la cantidad de tablas existentes en la conexión
                    int _count = GetTables(_sql).Rows.Count;
                    SQLContents = new DataTable[_count];
                    i = 0;
                    SQLiteCommand cmd = new SQLiteCommand(_sql);

                    foreach (DataRow row in GetTables(_sql).Rows)
                    {
                        string table = row[2].ToString();
                        SQLContents[i] = new DataTable(table);
                        cmd.CommandText = $"SELECT * FROM [{table}]";
                        await cmd.ExecuteNonQueryAsync();
                        
                        SQLiteDataAdapter _ad = new SQLiteDataAdapter(cmd);
                        _ad.Fill(SQLContents[i]);
                        i++;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Inicializa el procedimiento para abrir la base de datos
        /// </summary>
        public async Task<bool> OpenDatabase()
        {
            if (_IsFileOpen) return false;
            // Primero verifica si por el HASH del archivo, es nuevo o es el mismo
            // Si ha cambiado, se procederá a su actualización completa
            // Y en caso contrario, no se hará
            _IsOutdated = _forceConvert; // Parámetro por si se fuerza la actualización
            if (shaDB != _INIFile.GetSetting("SQLValues", "HASHFile")) _IsOutdated = true;

            // Si el archivo está desactualizado, abrimos el Prometeo V
            if (_IsOutdated)
            {
                _db = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_FilePath};");
                await _db.OpenAsync();

                // Comprobamos si existe la base de datos en el directorio
                // Si existe, los datos se pasarán a memoria y en un archivo .bak
                if (File.Exists(_SQLPath))
                {
                    // Si existiera el archivo .bak, lo eliminamos
                    if (File.Exists(_SQLPath + ".bak")) { File.Delete(_SQLPath + ".bak"); }
                    // Creamos la conexión
                    _sql = new SQLiteConnection($"Data Source={_SQLPath};Version=3;");
                    _sqlBackup = new SQLiteConnection($"Data Source={_SQLPath}.bak;Version=3;");
                    await _sqlBackup.OpenAsync();
                }
                // En caso de que la base no exista, usamos la conexión normal y especificamos que es nuevo archivo
                // Por lo que las comprobaciones se ignorarán por completo
                else 
                { 
                    _sql = new SQLiteConnection($"Data Source={_SQLPath};Version=3;");
                    _IsNew = true;
                }

                // Abrimos la conexión
                await _sql.OpenAsync();
            }
            _IsFileOpen = true;
            return true;
        }

        /// <summary>
        /// Finaliza y libera el procedimiento de carga de la base de datos
        /// </summary>
        public bool CloseDatabase()
        {
            if (!_IsFileOpen) return false;
            if (_sql != null)
            {
                _sql.Close();
                _sql.Dispose();
            }
            SQLContents = null;
            SQLBackup = null;
            if (_IsOutdated)
            {
                // Reiniciamos la aplicación asignando parámetros de copia de la base de datos
                Process p = new Process();
                p.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "SIREPV.exe";
                p.StartInfo.Arguments = "--SwapDB";
                p.Start();
                Environment.Exit(0);
            }

            _db = null;
            _sql = null;
            _IsFileOpen = false;
            _DataSplit = null;
            _DataNew = null;
            return true;
        }
        #endregion
        
        #region Parameters
        /// <summary>
        /// Devuelve o establece el HASH del archivo a actualizar
        /// </summary>
        public string HASHMDB
        {
            get { return shaDB; }
            set { if (_IsFileOpen == false) shaDB = value;  }
        }

        /// <summary>
        /// Devuelve o establece la ruta de acceso al archivo
        /// </summary>
        public string FileLocated
        {
            get { return _FilePath; }
            set { if (_IsFileOpen == false) _FilePath = value; }
        }

        /// <summary>
        /// Indica si se deberá forzar el parámetro de conversión
        /// </summary>
        public bool ForceConvert
        {
            set { if (_IsFileOpen == false) _forceConvert = value; }
        }

        /// <summary>
        /// Indica si se deberá ignorar la conversión de inicio
        /// </summary>
        public bool IgnoreConvert
        {
            set { if (_IsFileOpen == false) _ignoreConvert = value; }
        }

        /// <summary>
        /// Indica si existe algún fichero abierto
        /// </summary>
        public bool IsFileOpen
        {
            get { return _IsFileOpen; }
        }
        #endregion

        #region PrivateMethods
        private static DataTable GetTables(SQLiteConnection connection)
        {
            DataTable tables = connection.GetSchema("Tables");
            return tables;
        }

        private string GetCreateTableColumns(DataTable dataTable)
        {
            string columns = "";
            foreach (DataColumn column in dataTable.Columns)
            {
                columns += $"{column.ColumnName} {GetSQLiteDataType(column.DataType)}, ";
            }
            columns = columns.TrimEnd(',', ' ');
            return columns;
        }

        private string GetSQLiteDataType(Type dataType)
        {
            if (dataType == typeof(int)) return "INTEGER";
            if (dataType == typeof(bool)) return "INTEGER";
            if (dataType == typeof(byte)) return "INTEGER";
            if (dataType == typeof(char)) return "TEXT";
            if (dataType == typeof(DateTime)) return "TEXT";
            if (dataType == typeof(decimal)) return "TEXT";
            if (dataType == typeof(string)) return "TEXT";
            if (dataType == typeof(double)) return "REAL";
            if (dataType == typeof(Single)) return "REAL";

            return "TEXT";
        }

        private string GetInsertValues(DataRow row)
        {
            string values = "";
            foreach (var item in row.ItemArray)
            {
                values += $"'{item.ToString().Replace("'", "''")}', ";
            }
            values = values.TrimEnd(',', ' ');
            return values;
        }
        #endregion
    }
}
