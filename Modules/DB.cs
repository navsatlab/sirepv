using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Security.Cryptography;
using SoftCircuits.IniFileParser;
using System.Diagnostics;

namespace TSI.SIREP.Modules
{
    internal class DB
    {
        #region Declares
        private string _FilePath;
        private readonly string _INI = AppDomain.CurrentDomain.BaseDirectory + @"settings.ini";
        private string _XTableEjemplar, _XTableFichas;
        private bool _IsFileOpen;

        private SQLiteConnection _sql;
        private IniFile _INIFile;
        #endregion

        public DB(IniFile _ini)
        {
            _INIFile = _ini;
            _FilePath = "";
            _IsFileOpen = false;
            _INIFile.Load(_INI);
            _XTableEjemplar = _INIFile.GetSetting("SQLValues", "Sheet1");
            _XTableFichas = _INIFile.GetSetting("SQLValues", "Sheet2");
        }

        #region Methods
        /// <summary>
        /// Inicializa el procedimiento para abrir la base de datos
        /// </summary>
        public async Task<bool> OpenDatabase()
        {
            if (!_IsFileOpen)
            {
                _sql = new SQLiteConnection($"Data Source={_FilePath};Version=3;");
                await _sql.OpenAsync();

                _IsFileOpen = true;
            }
            return true;
        }

        /// <summary>
        /// Finaliza y libera el procedimiento de carga de la base de datos
        /// </summary>
        public bool CloseDatabase()
        {
            if (_IsFileOpen)
            {
                
                _IsFileOpen = false;
            }
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
