using System;
using System.IO;
using System.Reflection;
using iTextSharp.text.pdf;
using System.Data;

namespace TSI.SIREP.Modules
{
    internal class PDFEdit
    {
        #region Declares
        private int _users;
        private string _expirateDate;

        public struct UserInfo
        {
            public string name;
            public string PO;
            public string cel;
            public string mail;
        }

        private UserInfo[] _usr;
        private PdfReader reader;
        private PdfStamper pstp;
        #endregion

        public PDFEdit()
        {
            var str = Assembly.GetExecutingAssembly().GetManifestResourceStream("TSI.SIREP.Resources.editable.pdf");
            reader = new PdfReader(str);
            _usr = new UserInfo[4];
        }

        #region Methods
        /// <summary>
        /// Agrega un usuario nuevo al registro del PDF
        /// </summary>
        /// <param name="_user">Parámetros del nuevo usuario</param>
        /// <returns>Devuelve true si se completa la tarea</returns>
        public bool AddUser(UserInfo _user)
        {
            if (_users < 4)
            {
                _usr[_users] = _user;
                _users++;
            }
            return true;
        }

        /// <summary>
        /// Guarda los cambios del PDF en un archivo
        /// </summary>
        /// <param name="_file">Ruta de acceso del nuevo archivo</param>
        /// <returns>Devuelve true si se completa la tarea</returns>
        public bool SavePDF(string _file)
        {
            pstp = new PdfStamper(reader, new FileStream(_file, FileMode.Create));
            AcroFields std = pstp.AcroFields;

            for (int i = 1; i < 5; i++)
            {
                std.SetField($"_{i}name", _usr[i - 1].name);
                std.SetField($"_{i}locate", _usr[i - 1].PO);
                std.SetField($"_{i}number", _usr[i - 1].cel);
                std.SetField($"_{i}mail", _usr[i - 1].mail);
                std.SetField($"_{i}date", _expirateDate);
            }

            pstp.FormFlattening = false;
            pstp.Close();

            return true;
        }
        #endregion

        #region Parameters
        /// <summary>
        /// Devuelve o establece la fecha de expiración para las credenciales
        /// </summary>
        public string ExpirationDate
        {
            get { return _expirateDate; }
            set { _expirateDate = value; }
        }

        /// <summary>
        /// Devuelve la cantidad de usuarios registrados
        /// </summary>
        public int UserCount
        {
            get { return _users; }
        }
        #endregion
    }
}
