using System;
using System.IO;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using System.Net.Http;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SoftCircuits.IniFileParser;
using TSI.SIREP.Modules;
using System.Net.NetworkInformation;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Threading;

namespace SIREP_V1
{
    internal class Program
    {
        #region Declares
        private static readonly string _settings = AppDomain.CurrentDomain.BaseDirectory + @"settings.ini";

        private static long CommandState;
        private static string MDBPath, ExcelPath;
        private static string areaName, userName;
        private static string _urlform;
        private static bool awaitTask = false, _IsNetwork, _IsNew, _noCommand;

        private static DBConvert _dbc;
        private static DB _db;
        private static ExcelSheet _excel;
        private static IniFile _ini;
        private static DataTable DataExcel;
        private static GForms _gForm;
        private static PDFEdit _edit;

        private enum HelpWindow
        {
            NoWindow = 0,
            LoadWindow = 1,
            StartWindow = 2,
            InfoWindow = 3,
            NoImplementedWindow = 4,
            AboutWindow = 5,
            ListViewWindow = 6,
            MessageBoxWindow = 7,
            CriticalErrorWindow = 8,
            InputWindow = 9,
            MultilineInputWindow = 10,
            CredentialWindow = 11
        }
        #endregion

        [STAThread]
        static async Task Main(string[] args)
        {
            // Verificamos si existe el proceso abierto, si es así cerramos esta instancia
            Mutex mt = new Mutex(true, "SIREPV", out bool _inst);
            if (!_inst) Environment.Exit(0);
            
            // UNICODE Console
            Console.OutputEncoding = Encoding.UTF8;

            // Se ejecuta el procedimiento de carga inicial
            await StartProcess(args);

            bool _nocmd = false;
            do
            {
                SetCurrentWindow(HelpWindow.StartWindow);
                if (_nocmd)
                {
                    _nocmd = false;
                    
                    Console.ResetColor();
                    Console.BackgroundColor = ConsoleColor.DarkRed;
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.SetCursorPosition(0, Console.WindowHeight - 4);
                    Console.Write(" ".PadRight(Console.WindowWidth - 1)); 
                    Console.SetCursorPosition(0, Console.WindowHeight - 4);
                    Console.Write($"El comando «{args[0]}» no fué reconocido correctamente");
                    Console.SetCursorPosition(0, Console.WindowHeight - 3);
                    Console.Write(" ".PadRight(Console.WindowWidth - 1));
                    Console.SetCursorPosition(0, Console.WindowHeight - 3);
                    Console.Write("Compruébalo y vuelve a ejecutar el comando");
                    Console.ResetColor();
                }

                args = WaitCommand();
                int response = 0;
                switch (args[0])
                {
                    case "SALIR":
                        EndApp();
                        break;
                    case "RECONFIG":
                        response = int.Parse(SetCurrentWindow(HelpWindow.MessageBoxWindow,
                            "Estás a punto de eliminar todos los datos existentes de este programa",
                            "Eso quiere decir que se reiniciará el proceso desde cero, previo a la instalación de este programa\n",
                            "¿Deseas continuar?").ToString());
                        if (response == 1) StartProcess("--ResetAll").Wait();
                        break;
                    case "RECARGAR":
                        response = int.Parse(SetCurrentWindow(HelpWindow.MessageBoxWindow,
                            "Se comenzará el proceso de carga de datos desde la MDB de Prometeo V",
                            "Esto puede llevar un tiempo\n",
                            "¿Deseas continuar?").ToString());
                        if (response == 1) StartProcess("--ForceConvert").Wait();
                        break;
                    case "ADVANCED":
                        SetCurrentWindow(HelpWindow.InfoWindow, "TSI.SIREP.Help.Advanced.txt");
                        break;
                    case "BUSCAR":
                        SetCurrentWindow(HelpWindow.NoImplementedWindow);
                        break;
                    case "EXPORTAR":
                        SetCurrentWindow(HelpWindow.NoImplementedWindow);
                        break;
                    case "VER":
                        SetCurrentWindow(HelpWindow.NoImplementedWindow);
                        break;
                    case "ABOUT":
                        SetCurrentWindow(HelpWindow.AboutWindow);
                        break;
                    case "AYUDA":
                        if (args.Length == 1) break;
                        switch (args[1])
                        {
                            case "BUSCAR":
                                SetCurrentWindow(HelpWindow.InfoWindow, "TSI.SIREP.Help.HelpSearch.txt");
                                break;
                            case "INICIAR":
                                SetCurrentWindow(HelpWindow.InfoWindow, "TSI.SIREP.Help.HelpStart.txt");
                                break;
                            case "VER":
                                SetCurrentWindow(HelpWindow.InfoWindow, "TSI.SIREP.Help.HelpView.txt");
                                break;
                            case "EXPORTAR":
                                SetCurrentWindow(HelpWindow.InfoWindow, "TSI.SIREP.Help.HelpExport.txt");
                                break;
                        }
                        break;
                    case "TEST":
                        await TestMethod();
                        break;
                    case "INICIAR":
                        await InitRevision();
                        break;
                    case "CREDENCIAL":
                        SetCurrentWindow(HelpWindow.CredentialWindow);
                        break;
                    case "SQLCMD":
						// abrir el editor de comando para ejecutar sentencias SQLite, para no estar en Linux a cada rato extrayendo datos temporales
						
                    	break;
                    default:
                        _nocmd = true;
                        break;
                }
            }
            while (true);
        }

        #region EnviromentFunctions
        private static async Task<bool> TestMethod()
        {
            string Url = SetCurrentWindow(HelpWindow.InputWindow,
                            "Por favor ingrese la ruta de acceso web del formulario").ToString();
            string ApiKey = "AIzaSyAOEgSac7KGCfnGg0_9qIoePzF0vEbbSFc";
            _gForm.ApiKey = ApiKey;
            _gForm.UrlForm = Url;
            await _gForm.InitializeLoad();
            return true;
        }

        private static async Task<bool> InitRevision()
        {
            // procedimiento de inicio de revisión del documento
            _urlform = _ini.GetSetting("Forms", "Url");
            if (_urlform == null)
            {
                _urlform = SetCurrentWindow(HelpWindow.InputWindow,
                    "Actualmente no se ha definido una dirección web del formulario de Google",
                    "Por favor escribe la página web en el siguiente cuadro\n",
                    "Ejemplos: https://forms.gle/6N6qqy1J8x6QP5QE9",
                    "https://docs.google.com/forms/d/e/1FAIpQLSf6CsoRgNc3GwmaXHRq96IadOLRfuqM5Tj8mjqsMRzqKLz4Tw/viewform").ToString();

                _ini.SetSetting("Forms", "Url", _urlform);
                _ini.Save(_settings);
                _ini.Load(_settings);
            }
            userName = SetCurrentWindow(HelpWindow.InputWindow,
                "Se comenzará la revisión de las fichas de Prometeo",
                "Para esto se te pedirá que escribas tu nombre para continuar el proceso",
                "Este será utilizado en todos los formularios que se llenarán de las correcciones").ToString();


            return true;
        }

        [STAThread]
        private static async Task<bool> StartProcess(params string[] args)
        {
            Console.Title = "SIREP V1";
            awaitTask = true;
            CommandState = 0;
            bool _forceConvert = false, _reset = false, _ignoreConvert = false;
            // Comprobación si existen argumentos de inicio
            foreach (string arg in args)
            {
                switch (arg)
                {
                    case "--ForceConvert":
                        _forceConvert = true;
                        break;
                    case "--ResetAll":
                        _reset = true;
                        break;
                    case "--NoConvert":
                        _ignoreConvert = true;
                        break;
                    case "--SwapDB":
                        // Copiamos la base de datos
                        try
                        {
                            File.Delete(AppDomain.CurrentDomain.BaseDirectory + "sirep.db");
                            File.Copy(AppDomain.CurrentDomain.BaseDirectory + "sirep.db.bak", AppDomain.CurrentDomain.BaseDirectory + "sirep.db");
                            File.Delete(AppDomain.CurrentDomain.BaseDirectory + "sirep.db.bak");
                        }
                        catch { }
                        break;
                    case "--Debug":
                        // Ingresa al modo seguro de depuración, antes de poder cargar las variables al sistema
                        // Se usarán las determinadas por el depurador

                        break;
                }
            }

            // Comprobamos configuraciones iniciales, de los archivos a utilizar
            SetCurrentWindow(HelpWindow.LoadWindow, "Espera por favor", "Cargando configuración inicial");
            VerifyIfINI(_reset);

            _dbc = new DBConvert(_ini);
            _excel = new ExcelSheet();
            _dbc.FileLocated = MDBPath;
            _excel.FileLocated = ExcelPath;

            _dbc.ForceConvert = _forceConvert;
            _dbc.IgnoreConvert = _ignoreConvert;

            // Abrimos las conexiones de las bases de datos
            SetCurrentWindow(HelpWindow.LoadWindow, "Espera por favor", "Abriendo bases de datos");
            try
            {
                _dbc.HASHMDB = CalculateFileHash(MDBPath);
            }
            catch
            {
                // Por si se genera un error de acceso (por estar en uso), usamos el hash predeterminado
                _dbc.HASHMDB = _ini.GetSetting("SQLValues", "HASHFILE");
            }
            
            await _dbc.OpenDatabase();
            _excel.OpenSheet();
            _ini.SetSetting("SQLValues", "HASHFile", _dbc.HASHMDB);
            _ini.Save(_settings);
            _ini.Load(_settings);

            // Cargamos en memoria los datos necesarios de las bases de datos especificadas
            SetCurrentWindow(HelpWindow.LoadWindow, "Espera por favor", "Cargando datos");
            await _dbc.LoadDBtoMemory();

            DataTable _table = _excel.GetTables();
            string sheetName = "", objectName = "";
            if (_IsNew)
            {
                string[] str = new string[_table.Rows.Count + 1];
                str[0] = $"Se encontraron las siguientes pestañas dentro del archivo de Excel.\n\nPor favor indica con cuál se deberá trabajar:";
                for (int i = 0; i < _table.Rows.Count; i++)
                {
                    str[i + 1] = _table.Rows[i][1].ToString();
                }
                int index = int.Parse(SetCurrentWindow(HelpWindow.ListViewWindow, str).ToString());
                string[] _objEx = _table.Rows[index - 1][2].ToString().Split(';');

                sheetName = _table.Rows[index - 1][1].ToString();
                _ini.SetSetting("ExcelDB", "NameWorksheet", sheetName);

                str = new string[_objEx.Length + 1];
                if (_objEx.Length == 2)
                {
                    // Sólo cuenta con una tabla
                    index = 1;
                }
                else if (_objEx.Length < 2)
                {
                    // No hay objetos a listar, sale del procedimiento
                    SetCurrentWindow(HelpWindow.MessageBoxWindow, "No se encontraron objetos de tabla en la selección anterior.", "La aplicación se cerrará");
                    _dbc.CloseDatabase();
                    _excel.CloseSheet();
                    EndApp();
                }
                else
                {
                    str[0] = $"Dentro de la pestaña «{sheetName}» se encontraron las siguientes tablas\n\nPor favor selecciona la que contenga los valores a trabajar:";
                    for (int i = 0; i < _objEx.Length; i++)
                    {
                        str[i + 1] = _objEx[i].ToString();
                    }
                    index = int.Parse(SetCurrentWindow(HelpWindow.ListViewWindow, str).ToString());
                }
                objectName = _objEx[index - 1];
                _ini.SetSetting("ExcelDB", "NameObjectTable", objectName);

                _ini.Save(_settings);
            }
            else
            {
                sheetName = _ini.GetSetting("ExcelDB", "NameWorksheet");
                objectName = _ini.GetSetting("ExcelDB", "NameObjectTable");
            }
            // Descargamos la tabla específica en un DataTable
            DataExcel = _excel.GetTableValues(sheetName, objectName);
            if (_IsNew)
            {
                // Se define cuáles son las columnas que contienen datos
                // En este caso se escogerá { Columna y Charola }
                // { Título, Autor, País, Editorial, Año de edición, Clasificación, N° adquisición, Donante }
                // { Idiomas y notas }

                // Se mostrará un registro por defecto donde se le podrá asignar la columna deseada

                string[] select = new string[DataExcel.Rows[1].ItemArray.Count()];
                //_col = SetCurrentWindow(HelpWindow.ListViewWindow, );
                //DataExcel.Columns[0].ColumnName = 
                //string[] select =
            }
            
            // Comenzamos el procedimiento de comparación de datos existentes
            SetCurrentWindow(HelpWindow.LoadWindow, "Espera por favor", "Procesando cambios");
            _dbc.CompareData();
            _dbc.SplitData();

            SetCurrentWindow(HelpWindow.LoadWindow, "Espera por favor", "Realizando cambios");
            await _dbc.FlushData();

            SetCurrentWindow(HelpWindow.LoadWindow, "Liberando memoria");
            _excel.CloseSheet();
            _dbc.CloseDatabase();

            SetCurrentWindow(HelpWindow.LoadWindow, "Verificando acceso a internet");
            _IsNetwork = await CheckInternet();

            // Procedimiento de carga en memoria de la DB real para comenzar el cotejo
            SetCurrentWindow(HelpWindow.LoadWindow, "Iniciando aplicación");
            _db = new DB(_ini)
            {
                FileLocated = AppDomain.CurrentDomain.BaseDirectory + @"settings.ini"
            };
            await _db.OpenDatabase();
            _gForm = new GForms();

            return true;
        }

        [STAThread]
        private static string[] WaitCommand()
        {
            SetCursorToCommand();
            return Console.ReadLine().Split(' ');
        }

        private static async Task<bool> CheckInternet()
        {
            try
            {
                using (var _x = new Ping())
                {
                    var reply = await _x.SendPingAsync("www.google.com");
                    return (reply.Status == IPStatus.Success);
                }
            }
            catch { return false; }
        }

        private static void EndApp()
        {
            // Destruimos todas las variables involucradas
            try
            {
                _excel.CloseSheet();
                _dbc.CloseDatabase();
                _db.CloseDatabase();
                DataExcel = null;
                _ini = null;
            }
            catch { }
            Console.Clear();
            Console.ResetColor();
            Environment.Exit(0);
        }

        [STAThread]
        private static bool VerifyIfINI(bool reset = false)
        {
            if (reset)
            {
                try
                {
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"settings.ini");
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"sirep.db");
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"sirep.db.bak");
                }
                catch { }
            }
            _ini = new IniFile();
            if (!File.Exists(_settings))
            {
                // Como el INI no existe, se procede a buscar los archivos necesarios
                // En este caso se toma la ruta C:\\PROMETEO como el origen del archivo
                // Si no se encontrara, se procederá a abrir una ventana de búsqueda
                int _response = int.Parse(SetCurrentWindow(HelpWindow.MessageBoxWindow,
                    "Esta es la primera vez que se iniciará este programa",
                    "Recuerda que ésto fue diseñado para la sala de Narrativa de la sede de Av. Juárez",
                    "Por lo que puede variar si se implementa en otra área",
                    "Además, éste sólo se usará para esa sala únicamente", "",
                    "¿Deseas continuar?").ToString());
                if (_response == 2) EndApp();
                if (!File.Exists(@"C:\PROMETEO\PROMETEO.mdb"))
                {
                    SetCurrentWindow(HelpWindow.LoadWindow, "No se encontró la base de datos en la ruta especificada por default", "A continuación se abrirá una ventana donde deberás buscar la base de datos de Prometeo");
                    using (OpenFileDialog _op = new OpenFileDialog())
                    {
                        _op.Filter = "Bases de datos Access 2000 (*.mdb)|*.mdb";
                        _op.Title = "Abrir base de datos";
                        if (_op.ShowDialog() == DialogResult.OK)
                        {
                            MDBPath = _op.FileName;
                        }
                        else { EndApp(); }
                    }
                }
                else { MDBPath = @"C\PROMETEO\PROMETEO.mdb"; }

                SetCurrentWindow(HelpWindow.LoadWindow, "Deberás seleccionar dónde se encuentra el archivo de Excel que guarda tu Inventario");
                using (OpenFileDialog _op = new OpenFileDialog())
                {
                    _op.Filter = "Hoja de cálculo de Microsoft Excel (*.xlsx)|*.xlsx";
                    _op.Title = "Abrir hoja de cálculo";
                    if (_op.ShowDialog() == DialogResult.OK)
                    {
                        ExcelPath = _op.FileName;
                    }
                    else { EndApp(); }
                }

                areaName = SetCurrentWindow(HelpWindow.InputWindow,
                    "Por favor ingresa el nombre de la sala donde se está haciendo el cotejo",
                    "Toma en cuenta que éste nombre será utilizado en todos los formularios de Google Forms").ToString();

                _ini.SetSetting("Main", "MDBDatabase", MDBPath);
                _ini.SetSetting("Main", "ExcelFile", ExcelPath);
                _ini.SetSetting("Forms", "AreaName", areaName);
                _ini.Save(_settings);
                _IsNew = true;
            }
            else
            {
                _ini.Load(_settings);
                MDBPath = _ini.GetSetting("Main", "MDBDatabase");
                ExcelPath = _ini.GetSetting("Main", "ExcelFile");
                areaName = _ini.GetSetting("Forms", "AreaName");
            }
            return true;
        }

        private static string CalculateFileHash(string filePath)
        {
            using (var sha256 = SHA256.Create())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hashBytes = sha256.ComputeHash(stream);
                    return ByteArrayToHexString(hashBytes);
                }
            }
        }

        private static string ByteArrayToHexString(byte[] bytes)
        {
            StringBuilder hexBuilder = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++)
            {
                hexBuilder.Append(bytes[i].ToString("x2"));
            }
            return hexBuilder.ToString();
        }
        #endregion

        #region ConsoleOutput
        [STAThread]
        private static object SetCurrentWindow(HelpWindow windowtype, params string[] loadValue)
        {
            awaitTask = true;
            bool IsHelpWindow = false, IsLoadWindow = false, IsShowWindow = false, IsListWindow = false, IsMessageWindow = false, IsErrorWindow = false;
            bool IsInputWindow = false, IsMultilineInputWindow = false, IsCredentialWindow = false;
            _noCommand = false;
            var assly = Assembly.GetExecutingAssembly();
            var resnam = "";

            switch (windowtype)
            {
                case HelpWindow.NoWindow:
                    CommandState = 0;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.StartWindow:
                    resnam = "TSI.SIREP.Help.StartUp.txt";
                    IsHelpWindow = true;
                    CommandState = 1;
                    awaitTask = false;
                    break;
                case HelpWindow.InfoWindow:
                    resnam = loadValue[0];
                    IsShowWindow = true;
                    CommandState = 2;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.LoadWindow:
                    awaitTask = true;
                    IsLoadWindow = true;
                    CommandState = 0;
                    _noCommand = true;
                    break;
                case HelpWindow.NoImplementedWindow:
                    resnam = "TSI.SIREP.Help.NotImplemented.txt";
                    IsShowWindow = true;
                    CommandState = 2;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.AboutWindow:
                    resnam = "TSI.SIREP.Help.About.txt";
                    IsShowWindow = true;
                    CommandState = 2;
                    awaitTask = false;
                    _noCommand = true;
                    break;
				case HelpWindow.ListViewWindow:
                	IsListWindow = true;
                    CommandState = 3;
                    awaitTask = false;
                    _noCommand = true;
                    break;
				case HelpWindow.MessageBoxWindow:
                	IsMessageWindow = true;
                    CommandState = 4;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.CriticalErrorWindow:
                    IsErrorWindow = true;
                    CommandState = 2;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.InputWindow:
                    IsInputWindow = true;
                    CommandState = 5;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.MultilineInputWindow:
                    IsMultilineInputWindow = true;
                    CommandState = 5;
                    awaitTask = false;
                    _noCommand = true;
                    break;
                case HelpWindow.CredentialWindow:
                    IsCredentialWindow = true;
                    CommandState = 7;
                    awaitTask = false;
                    _noCommand = true;
                    break;
            }

            Console.CursorVisible = true;
            DrawWindow();
            if (IsHelpWindow)
            {
                using (Stream stream = assly.GetManifestResourceStream(resnam))
                using (StreamReader reader = new StreamReader(stream))
                {
                    DrawText(reader.ReadToEnd());
                }
                return true;
            }
            if (IsLoadWindow)
            {
                Console.CursorVisible = false;
                Console.ResetColor();
                int widthP = 0;
                int heighP = (Console.WindowHeight - loadValue.Length) / 2;
                foreach (string val in loadValue)
                {
                    widthP = (Console.WindowWidth - val.Length) / 2;
                    Console.SetCursorPosition(widthP, heighP);
                    Console.WriteLine(val);
                    heighP++;
                }

                widthP = Console.WindowWidth / 2;
                Console.SetCursorPosition(widthP, heighP);

                int i = 0;
                Task loadX = Task.Run(async () =>
                {
                    while (awaitTask)
                    {
                        switch (i % 4)
                        {
                            case 0:
                                Console.Write("/");
                                break;
                            case 1:
                                Console.Write("-");
                                break;
                            case 2:
                                Console.Write(@"\");
                                break;
                            case 3:
                                Console.Write("|");
                                break;
                        }
                        await Task.Delay(250);
                        Console.SetCursorPosition(heighP, widthP + 1);
                        i++;
                        if (!awaitTask)
                        {
                            SetCursorToCommand();
                            break;
                        }
                    }
                });
                return true;
            }
            if (IsShowWindow)
            {
                using (Stream stream = assly.GetManifestResourceStream(resnam))
                using (StreamReader reader = new StreamReader(stream))
                {
                    DrawText(reader.ReadToEnd());
                }
                while (true) { if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape) break; }
                return true;
            }
            if (IsListWindow)
            {
                Console.CursorVisible = false;
                Console.ResetColor();
            	Console.SetCursorPosition(0, 2);
                Console.WriteLine(loadValue[0]);
                int index = 1, heighT = (Console.WindowHeight - loadValue.Length - 4) / 2;
                bool _out = false;
                Console.SetCursorPosition(2, heighT);
                do
                {
                    for (int i = 1; i < loadValue.Length; i++)
                    {
                        Console.SetCursorPosition(2, heighT + i);
                        if (i == index)
                        {
                            Console.BackgroundColor = ConsoleColor.White;
                            Console.ForegroundColor = ConsoleColor.Black;
                        }
                        Console.WriteLine(loadValue[i]);
                        Console.ResetColor();
                    }
                    switch (Console.ReadKey(true).Key)
                    {
                        case ConsoleKey.UpArrow:
                            index = Math.Max(1, index - 1);
                            break;
                        case ConsoleKey.DownArrow:
                            index = Math.Min(loadValue.Length - 1, index + 1);
                            break;
                        case ConsoleKey.Enter:
                            _out = true;
                            break;
                    }
                    
                }
                while (!_out);
                return index;
            }
            if (IsMessageWindow)
            {
                Console.CursorVisible = false;
                Console.ResetColor();
                int heighP = (Console.WindowHeight - loadValue.Length - 4) / 2;
                int widthP = 0;
                foreach (string val in loadValue)
                {
                    widthP = (Console.WindowWidth - val.Length) / 2;
                    Console.SetCursorPosition(widthP, heighP);
                    Console.WriteLine(val);
                    heighP++;
                }
                heighP += 2;
                int index = 2;
                bool _out = false;
                do
                {
                    widthP = (Console.WindowWidth - 30) / 2;
                    Console.SetCursorPosition(widthP, heighP);
                    Console.Write("  ACEPTAR  ");
                    Console.SetCursorPosition(widthP + 19, heighP);
                    Console.Write("  CANCELAR  ");
                    Console.BackgroundColor = ConsoleColor.White;
                    Console.ForegroundColor = ConsoleColor.Black;

                    if (index == 1)
                    {
                        Console.SetCursorPosition(widthP, heighP);
                        Console.Write("  ACEPTAR  ");
                    }
                    else if (index == 2)
                    {
                        Console.SetCursorPosition(widthP + 19, heighP);
                        Console.Write("  CANCELAR  ");
                    }
                    Console.ResetColor();
                    Debug.WriteLine(index);
                    switch (Console.ReadKey(true).Key)
                    {
                        case ConsoleKey.RightArrow:
                            if (index != 2) index++;
                            break;
                        case ConsoleKey.LeftArrow:
                            if (index != 1) index--;
                            break;
                        case ConsoleKey.Enter:
                            _out = true;
                            break;
                    }
                }
                while (!_out);
                return index;
            }
            if (IsErrorWindow)
            {
                Console.CursorVisible = false;
                Console.SetCursorPosition(Console.WindowHeight - loadValue.Length - 4, 0);
                Console.ResetColor();
                Console.WriteLine("Se ha detectado un error en la aplicación y se ha cerrado");
                Console.WriteLine("A continuación se muestran los detalles del error:\n");
                foreach (string val in loadValue)
                {
                    Console.WriteLine(val);
                }

                while (true) { if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape) EndApp(); }
            }
            if (IsInputWindow)
            {
                Console.ResetColor();
                Console.SetCursorPosition(0, (Console.WindowHeight - loadValue.Length - 4) / 2);
                foreach (string val in loadValue)
                {
                    Console.WriteLine(val);
                }
                Console.WriteLine();

                Console.CursorVisible = true;
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.Write(" ".PadRight(Console.WindowWidth - 1));
                Console.SetCursorPosition(0, Console.CursorTop);
                return Console.ReadLine();
            }
            if (IsMultilineInputWindow)
            {
                
            }
            if (IsCredentialWindow)
            {
                // Comienza el procedimiento de asignación de credencial
                bool _out = false;

                _edit = new PDFEdit();
                _edit.ExpirationDate = SetCurrentWindow(HelpWindow.InputWindow, "Antes de poder realizar las credenciales, deberás ingresar la fecha de expiración",
                    "Usualmente se toma un año de más a partir de esta fecha",
                    "Deberá tener el siguiente formato: 12 de enero del 2024 (si es que hoy es 12 de enero del 2023)").ToString();
                SetCurrentWindow(HelpWindow.NoWindow);
                do
                {
                    CommandState = 7;
                    DrawWindow();
                    DrawText("Este es el módulo para poder generar credenciales de usuario de la biblioteca del IAGO",
                        $"Actualmente hay registrados «{_edit.UserCount}» usuarios para generarles credencial",
                        "Puedes comenzar agregando usuarios presionando la tecla [F8]",
                        "El sistema te irá guiando en los datos que deberás de introducir", "",
                        "Cuando termines de agregar los 4 (actualmente no se puede hacer más), presiona la tecla [F10] y la aplicación te mostrará un cuadro de diálogo",
                        "Ahí deberás seleccionar la carpeta donde quieres que se guarde el PDF para imprimir","",
                        "Recuerda que podrás editarlo después de terminar acá", "", "" +
                        "Si deseas finalizar la tarea sin guardar cambios, presiona la tecla [Esc]");
                    switch (Console.ReadKey(true).Key)
                    {
                        case ConsoleKey.Escape:
                            _out = true;
                            break;
                        case ConsoleKey.F8:
                            PDFEdit.UserInfo _dt = new PDFEdit.UserInfo();
                            _dt.name = SetCurrentWindow(HelpWindow.InputWindow, "Insertar nuevo usuario", "",
                                "Captura el nombre del usuario, este debe ser como el siguiente ejemplo:",
                                "Apellido paterno Apellido materno, Nombre(s)", "",
                                "Puedes copiar el que aparece en su anterior credencial, o de alguna identificación oficial",
                                "Como INE, pasaporte o licencia de conducir").ToString();
                            _dt.PO = SetCurrentWindow(HelpWindow.InputWindow, "Insertar nuevo usuario", "",
                                "Ahora escribe la dirección que aparece en su credencial, o en el comprobante de domicilio",
                                "Deberá seguir este formato como ejemplo",
                                "Av. Juárez #203 Col. Centro, C.P. 68000 Oaxaca de Juárez, Oax.").ToString();
                            _dt.cel = SetCurrentWindow(HelpWindow.InputWindow, "Insertar nuevo usuario", "",
                                $"Por favor captura el número telefónico de contacto del usuario «{_dt.name}»:").ToString();
                            _dt.mail = SetCurrentWindow(HelpWindow.InputWindow, "Insertar nuevo usuario", "",
                                $"Por favor captura algún correo electrónico del usuario «{_dt.name}» para contactarlo:").ToString();

                            _edit.AddUser(_dt);
                            break;
                        case ConsoleKey.F10:
                            int _i = int.Parse(SetCurrentWindow(HelpWindow.MessageBoxWindow, $"Estás a punto de guardar los cambios, ya tienes registrados «{_edit.UserCount}» usuarios",
                                "A continuación se te pedirá seleccionar una carpeta donde quieras guardar el PDF", "",
                                "¿Deseas continuar?").ToString());
                            if (_i == 1)
                            {
                                using (SaveFileDialog _op = new SaveFileDialog())
                                {
                                    _op.Filter = "Documento de Adobe Acrobat (*.pdf)|*.pdf";
                                    _op.Title = "Guardar PDF";
                                    //_op.ShowDialog();

                                    _edit.SavePDF(@"C:\PROMETEO\CREDENCIAL.PDF");  //_op.FileName);
                                }
                                _out = true;
                            }
                            break;
                    }
                }
                while (!_out);
            }

            return false;
        }

        private static void DrawCommands()
        {
            Console.SetCursorPosition(0, Console.WindowHeight - 1);
            switch (CommandState)
            {
                case 0:
                    Console.Write("Espera por favor");
                    break;
                case 1:
                    Console.Write("[Entrar] Ejecutar el comando");
                    break;
                case 2:
                    Console.Write("[Esc] Salir de esta ventana");
                    break;
				case 3:
                	Console.Write("[▲] Subir  [▼] Bajar  [Enter] Seleccionar el elemento");
                    break;
				case 4:
                	Console.Write("[◄] Izquierda  [►] Derecha  [Enter] Seleccionar el elemento");
                    break;
                case 5:
                    Console.Write("[Entrar] Guardar y continuar");
                    break;
                case 6:
                    Console.Write("[Shift] + [Entrar] Guardar y continuar [Esc] Salir de esta ventana");
                    break;
                case 7:
                    Console.Write("[F8] Agrega un nuevo usuario [F10] Guardar PDF [Esc] Cancelar tarea");
                    break;
            }
        }

        private static void DrawText(params string[] args)
        {
            Console.ResetColor();
            Console.SetCursorPosition(0, 2);
            foreach (string arg in args)
            {
                Console.WriteLine(arg);
            }
            SetCursorToCommand();
        }

        private static void SetCursorToCommand()
        {
            Console.CursorVisible = true;
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.SetCursorPosition(1, Console.WindowHeight - 2);
        }

        private static void DrawWindow()
        {
            Console.ResetColor();
            Console.Clear();
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.SetCursorPosition(0, 0);
            Console.Write("SIREP V1 | Sistema de revisión para Prometeo 2023");
            Console.Write(" ".PadRight(Console.WindowWidth - 49));
            Console.WriteLine();
            Console.ResetColor();
            Console.SetCursorPosition(0, Console.WindowHeight - 2);
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            if (!_noCommand) Console.Write(">");
            Console.Write(" ".PadRight(Console.WindowWidth - 1));
            Console.WriteLine();
            Console.ResetColor();
            DrawCommands();
            SetCursorToCommand();
        }
        #endregion
    }
}
