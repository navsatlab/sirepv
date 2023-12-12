using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TSI.SIREP.Modules
{
    internal class ExcelSheet
    {
        #region Declares
        private string _FilePath;     
        private bool _IsFileOpen;

        private Application _excel;
        private Workbook _wb;
        #endregion

        #region Methods
        /// <summary>
        /// Devuelve una lista de las tablas existentes en el documento
        /// </summary>
        /// <returns>Devuelve el nombre de las tablas que se encuentren</returns>
        public System.Data.DataTable GetTables()
        {
            if (!IsFileOpen) return null;
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("id", typeof(int)); 
            dt.Columns.Add("name", typeof(string));
            dt.Columns.Add("contents", typeof(string));

            for (int i = 0; i < _wb.Sheets.Count; i++)
            {
                Worksheet _ws = _wb.Worksheets[i + 1];
                string _list = "";
                foreach (object obj in _ws.ListObjects)
                {
                    if (obj is ListObject _x)
                    {
                        _list += _x.Name + ";";
                    }
                }
                dt.Rows.Add(i, _ws.Name, _list);
                Marshal.ReleaseComObject(_ws);
            }
            return dt;
        }

        /// <summary>
        /// Devuelve el contenido específico de una tabla, indicando pestaña y nombre de aquella
        /// </summary>
        /// <returns>Devuelve el resultado de la tabla buscada</returns>
        public System.Data.DataTable GetTableValues(string sheetName, string objectName)
        {
            if (!IsFileOpen) return null;
            System.Data.DataTable dt = new System.Data.DataTable(objectName);
            Worksheet _ws = _wb.Worksheets[sheetName];
            ListObject excelTable = _ws.ListObjects[objectName];

            Range dataRange = excelTable.DataBodyRange;
            object[,] values = dataRange.Value2;

            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);

            // Agregar las columnas al DataTable
            for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                dt.Columns.Add($"Column{columnIndex}");
            }

            // Agregar las filas al DataTable
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                System.Data.DataRow row = dt.NewRow();
                for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    row[columnIndex - 1] = values[rowIndex, columnIndex];
                }
                dt.Rows.Add(row);
            }

            dt.TableName = objectName;
            Marshal.ReleaseComObject(dataRange);
            Marshal.ReleaseComObject(excelTable);

            return dt;
        }

        /// <summary>
        /// Finaliza y libera los recursos utilizados por este módulo
        /// </summary>
        public bool CloseSheet()
        {
            if (!_IsFileOpen) return false;
            _wb.Close(false);
            _excel.Quit();
            Marshal.ReleaseComObject(_wb);
            Marshal.ReleaseComObject(_excel);
            _wb = null;
            _excel = null;
            _IsFileOpen = false;
            return true;
        }

        /// <summary>
        /// Inicializa el procedimiento de lectura del Excel
        /// </summary>
        public bool OpenSheet()
        {
            // Abrimos el archivo de Excel
            if (_IsFileOpen) return false;
            _excel = new Application();
            _wb = _excel.Workbooks.Open(_FilePath);

            _IsFileOpen = true;
            return true;
        }
        #endregion

        #region Parameters
        /// <summary>
        /// Devuelve o establece la ruta de acceso al archivo
        /// </summary>
        public string FileLocated
        {
            get { return _FilePath; }
            set { if (_IsFileOpen == false) _FilePath = value; }
        }

        /// <summary>
        /// Indica si existe algún fichero abierto
        /// </summary>
        public bool IsFileOpen
        {
            get { return _IsFileOpen; }
        }
        #endregion
    }
}
