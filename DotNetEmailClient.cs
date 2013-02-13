using Microsoft.Win32;
using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.Permissions;

namespace DM.Core.Net
{
    public class DotNetEmailClient : IDisposable
    {
        //Prueba, esto lo agrego Eugenio

        #region VARIABLES LOCALES

        private System.ComponentModel.IContainer components = null;

        private NetworkCredential _networkCredential;
        private MailMessage _mailMessage;
        private MailAddress _mailFrom;
        private SmtpClient _smtpClient;
        private AlternateView _avHtmlBody = null;
        private AlternateView _avTextBody = null;
        private EmailImages _emailImages = null;

        #endregion

        #region CONSTRUCTOR Y DESTRUCTOR DEL COMPONENTE

        #region ctor
        public DotNetEmailClient()
        {
            InitializeComponent();
        }
        #endregion

        

        #region InitializeComponent
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            try
            {
                _networkCredential = new NetworkCredential(this.SmtpUser, this.SmtpPass);

                _smtpClient = new SmtpClient(this.SmtpServer, this.SmtpPort);
                _smtpClient.Credentials = _networkCredential;
                _smtpClient.SendCompleted += new SendCompletedEventHandler(_smtpClient_SendCompleted);

                _mailMessage = new MailMessage();
                _mailMessage.IsBodyHtml = true;
                this.Encoding = System.Text.Encoding.UTF8;

                _emailImages = new EmailImages();
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion

        #endregion

        #region PROPIEDADES PUBLICAS

        #region Message
        [Browsable(false)]
        public MailMessage Message
        {
            get { return _mailMessage; }
            set { _mailMessage = value; }
        }
        #endregion
        #region Credential
        [Browsable(false)]
        public NetworkCredential Credential
        {
            get { return _networkCredential; }
            set { _networkCredential = value; }
        }
        #endregion
        #region Headers
        [Browsable(false)]
        public NameValueCollection Headers
        {
            get { return _mailMessage.Headers; }
        }
        #endregion
        #region Encoding
        [Browsable(false)]
        public System.Text.Encoding Encoding
        {
            get { return _mailMessage.BodyEncoding; }
            set 
            { 
                _mailMessage.BodyEncoding = value;
                _mailMessage.SubjectEncoding = value;
            }
        }
        #endregion
        #region Priority
        [DefaultValue(MailPriority.Normal)]
        public MailPriority Priority
        {
            get { return _mailMessage.Priority; }
            set { _mailMessage.Priority = value; }
        }
        #endregion
        #region DeliveryNotificationOptions
        [DefaultValue(DeliveryNotificationOptions.OnFailure)]
        public DeliveryNotificationOptions DeliveryNotificationOptions
        {
            get { return _mailMessage.DeliveryNotificationOptions; }
            set { _mailMessage.DeliveryNotificationOptions = value; }
        }
        #endregion

        #region From
        [Browsable(false)]
        public MailAddress From
        {
            get { return _mailFrom; }
            set { _mailFrom = value; }
        }
        #endregion
        #region ReplyTo
        [Browsable(false)]
        public MailAddress ReplyTo
        {
            get { return _mailMessage.ReplyTo; }
            set { _mailMessage.ReplyTo = value; }
        }
        #endregion
        [Browsable(false)]
        public string DispositionNotificationTo
        {
            get
            {
                if (_mailMessage.Headers["Disposition-Notification-To"] != null
                    && _mailMessage.Headers["Disposition-Notification-To"] != "")
                    return _mailMessage.Headers["Disposition-Notification-To"];
                else
                    return string.Empty;
            }
            set
            {
                try { _mailMessage.Headers.Remove("Disposition-Notification-To"); }
                catch { /*nada*/ }
                _mailMessage.Headers.Add("Disposition-Notification-To", value.Trim());
            }
        }

        #region HtmlBody
        private string __HtmlBody = string.Empty;
        [Browsable(false), Bindable(true)]
        public String HtmlBody
        {
            get { return __HtmlBody; }
            set { __HtmlBody = value.Trim(); }
        }
        #endregion
        #region TextBody
        private string __TextBody = string.Empty;
        [Browsable(false), Bindable(true)]
        public String TextBody
        {
            get { return __TextBody; }
            set { __TextBody = value.Trim(); }
        }
        #endregion
        #region Subject
        [Bindable(true)]
        public string Subject
        {
            get { return _mailMessage.Subject; }
            set { _mailMessage.Subject = value; }
        }
        #endregion
        #region IsHtmlBody
        public bool IsHtmlBody
        {
            get { return _mailMessage.IsBodyHtml; }
            set { _mailMessage.IsBodyHtml = value; }
        }
        #endregion

        #region SmtpServer
        string __SmtpServer = "";
        [Bindable(true)]
        public string SmtpServer
        {
            get { return __SmtpServer; }
            set 
            { 
                __SmtpServer = value; 
                _smtpClient.Host = __SmtpServer; 
            }
        }
        #endregion
        #region SmtpPort
        int __SmtpPort = 25;
        [Bindable(true)]
        public int SmtpPort
        {
            get { return __SmtpPort; }
            set 
            { 
                __SmtpPort = value;
                _smtpClient.Port = __SmtpPort;
            }
        }
        #endregion
        #region SmtpUser
        string __SmtpUser = "";
        [Bindable(true)]
        public string SmtpUser
        {
            get { return __SmtpUser; }
            set 
            { 
                __SmtpUser = value;
                _networkCredential.UserName = __SmtpUser;
                _smtpClient.Credentials = _networkCredential;
            }
        }
        #endregion
        #region SmtpPass
        string __SmtpPass = "";
        [Bindable(true)]
        public string SmtpPass
        {
            get { return __SmtpPass; }
            set
            {
                __SmtpPass = value;
                _networkCredential.Password = __SmtpPass;
                _smtpClient.Credentials = _networkCredential;
            }
        }
        #endregion
        #region SmtpSenderEmail
        string __SmtpSenderEmail = "";
        [Bindable(true)]
        public string SmtpSenderEmail
        {
            get { return __SmtpSenderEmail; }
            set { __SmtpSenderEmail = value; }
        }
        #endregion
        #region SmtpSenderName
        string __SmtpSenderName = "";
        [Bindable(true)]
        public string SmtpSenderName
        {
            get { return __SmtpSenderName; }
            set { __SmtpSenderName = value; }
        }
        #endregion

        #endregion

        #region EVENTOS Y DELEGADOS
        public delegate void SendErrorHandler(object sender, Exception Ex);

        #region SendCompleted
        public event EventHandler SendCompleted;
        private void OnSendCompleted()
        {
            if (SendCompleted != null)
                SendCompleted(this, null);
        }
        #endregion

        #region SendStarted
        public event EventHandler SendStarted;
        private void OnSendStarted()
        {
            if (SendStarted != null)
                SendStarted(this, null);
        }
        #endregion

        #region SendError
        public event SendErrorHandler SendError;
        private void OnSendError(Exception ex)
        {
            if (SendError != null)
                SendError(this, ex);
        }
        #endregion

        #endregion

        #region PUBLIC METHODS

        #region SetFrom
        /// <summary>
        /// Define la dirección de quien envía el correo
        /// </summary>
        public void SetFrom()
        {
            try
            {
                _mailMessage.From = new MailAddress(this.SmtpSenderEmail, this.SmtpSenderName);

                // Establecer ReplyTo si no está definido
                if (_mailMessage.ReplyTo == null)
                    this.SetReplyTo(this.SmtpSenderEmail, this.SmtpSenderName);
            }
            catch (Exception ex) { throw ex; }
        }
        public void SetFrom(String emailAddress, String nombre)
        {
            try
            {
                _mailMessage.From = new MailAddress(emailAddress, nombre);

                // Establecer ReplyTo si no está definido
                if (_mailMessage.ReplyTo == null)
                    this.SetReplyTo(emailAddress, nombre);
            }
            catch (Exception ex) { throw ex; }
        }
        public void SetFrom(String emailAddress)
        {
            try
            {
                _mailMessage.From = new MailAddress(emailAddress);

                // Establecer ReplyTo si no está definido
                if (_mailMessage.ReplyTo == null)
                    this.SetReplyTo(emailAddress);
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
        #region SetReplyTo
        /// <summary>
        /// Establece la dirección de respuesta para el correo electrónico
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <param name="nombre"></param>
        public void SetReplyTo(String emailAddress, String nombre)
        {
            try { _mailMessage.ReplyTo = new MailAddress(emailAddress, nombre); }
            catch (Exception ex) { throw ex; }
        }
        public void SetReplyTo(String emailAddress)
        {
            try { _mailMessage.ReplyTo = new MailAddress(emailAddress); }
            catch (Exception ex) { throw ex; }
        }
        /// <summary>
        /// Agrega la dirección de respuesta por omisión en el Header ReplyTo
        /// </summary>
        public void SetReplyTo()
        {
            try { _mailMessage.ReplyTo = new MailAddress(this.SmtpSenderEmail, this.SmtpSenderName); }
            catch (Exception ex) { throw ex; }
        }
        #endregion

        #region Send
        /// <summary>
        /// Envía el correo electrónico
        /// </summary>
        public void Send()
        {
            try 
            {
                this.SetMailMessageValues();
                this._smtpClient.Send(_mailMessage);
                this.OnSendCompleted();
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
        #region SendAsync
        /// <summary>
        /// Enviar correo electrónico de forma asíncrona, al finalizar correctamente se lanzará el evento 'SendCompleted'.
        /// En caso de ocurrir un error se lanzará el evento 'SendError'
        /// </summary>
        public void SendAsync()
        {
            this.SetMailMessageValues();
            this._smtpClient.SendAsync(_mailMessage, null);
            this.OnSendStarted();
        }
        #endregion

        #region Rutinas para agregar destinatarios del correo electrónico

        #region AddTo

        /// <summary>
        /// Agrega una dirección a la lista de destinatarios directos
        /// </summary>
        /// <param name="emailAddress">Dirección de correo electrónico</param>
        /// <param name="nombre">Nombre o descripción del destinatario</param>
        public void AddTo(String emailAddress, String nombre)
        {
            try
            {
                _mailMessage.To.Add(new MailAddress(emailAddress, nombre));
            }
            catch (Exception ex) { throw ex; }
        }

        public void AddTo(String emailAddress)
        {
            try
            {
                _mailMessage.To.Add(new MailAddress(emailAddress));
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion

        #region AddCC

        /// <summary>
        /// Agrega una dirección a la lista de correos de copia
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <param name="nombre"></param>
        public void AddCC(String emailAddress, String nombre)
        {
            try
            {
                _mailMessage.CC.Add(new MailAddress(emailAddress, nombre));
            }
            catch (Exception ex) { throw ex; }
        }

        public void AddCC(String emailAddress)
        {
            try
            {
                _mailMessage.CC.Add(new MailAddress(emailAddress));
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion

        #region AddBcc

        /// <summary>
        /// Agrega una dirección a la lista de correos de copia oculta
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <param name="nombre"></param>
        public void AddBcc(String emailAddress, String nombre)
        {
            try
            {
                _mailMessage.Bcc.Add(new MailAddress(emailAddress, nombre));
            }
            catch (Exception ex) { throw ex; }
        }

        public void AddBcc(String emailAddress)
        {
            try
            {
                _mailMessage.Bcc.Add(new MailAddress(emailAddress));
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion

        #endregion

        #region Rutinas para realizar la carga del contenido HTML o Text
        #region loadBodyFromHtmlFile
        /// <summary>
        /// Carga un archivo en la parte de HTML en el cuerpo del email.
        /// </summary>
        /// <param name="filePath">Ubicación del archivo HTML.</param>
        public void loadBodyFromHtmlFile(string filePath)
        {
            try
            {
                this.HtmlBody = File.ReadAllText(filePath);
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
        #region loadBodyFromTextFile
        /// <summary>
        /// Carga un archivo en la parte de TEXTO PLANO en el cuerpo del email.
        /// </summary>
        /// <param name="filePath">Ubicación del archivo de texto plano.</param>
        public void loadBodyFromTextFile(string filePath)
        {
            try
            {
                this.TextBody = File.ReadAllText(filePath);
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
        #endregion

        #region Rutinas para adjuntar archivos.

        #region addImage
        /// <summary>
        /// Incrusta una imagen al correo electrónico
        /// <remarks>Se tomará como tipo de imagen JPEG por omisión</remarks>
        /// </summary>
        /// <param name="filePath">Archivo de imagen que se incrustará</param>
        /// <param name="ContentID">Id de la imagen dentro del contenido</param>
        public void addImage(String filePath, String ContentID)
        {
            if(File.Exists(filePath))
                _emailImages.Add(new EmailImage(filePath, ImageFormat.Jpeg, ContentID));
        }
        
        /// <summary>
        /// Incrusta una imagen al correo electrónico permitiendo especificar el tipo de imagen.
        /// </summary>
        /// <param name="filePath">Archivo de imagen que se incrustará</param>
        /// <param name="ContentID">Id de la imagen dentro del contenido</param>
        /// <param name="mediaTypeName">Tipo de imagen a incrustar, soporta (JPEG,GIF,TIFF)</param>
        public void addImage(String filePath, ImageFormat mediaTypeName, String ContentID)
        {
            if (File.Exists(filePath))
                _emailImages.Add(new EmailImage(filePath, mediaTypeName, ContentID));
        }

        public void addImage(Bitmap image, String ContentID)
        {
            this.addImage(image, ImageFormat.Jpeg, ContentID);
        }

        public void addImage(Bitmap image, ImageFormat mediaTypeName, String ContentID)
        {
            string tmp = System.IO.Path.GetTempFileName();
            image.Save(tmp, mediaTypeName);
            this.addImage(tmp, mediaTypeName, ContentID);
        }


        #endregion

        #region addAttachment
        /// <summary>
        /// Adjunta un archivo al correo electrónico
        /// </summary>
        /// <param name="filePath">Ubicación del archivo a adjuntar</param>
        public void AddAttachment(String filePath)
        {
            if (File.Exists(filePath))
            {
                FileInfo fi = new FileInfo(filePath);
                Attachment attch;
                if (fi.Extension.EndsWith("pdf"))
                    attch = new Attachment(filePath, "application/pdf");
                else
                    attch = new Attachment(filePath, getMimeType(fi));
                attch.ContentId = fi.Name;
                attch.TransferEncoding = TransferEncoding.Base64;

                _mailMessage.Attachments.Add(attch);
            }
        }
        /// <summary>
        /// Adjuntar un archivo
        /// </summary>
        /// <param name="fileStream">Stream contenedor del archivo a adjuntar</param>
        /// <param name="fileName">Nombre del archivo. Ej. "factura.pdf"</param>
        /// <param name="contentType">MIME que especifica el tipo de archivo</param>
        public void AddAttachment(byte[] fileBuffer, string fileName, string contentType)
        {

            using (MemoryStream ms = new MemoryStream(fileBuffer))
            {
                Attachment attch = new Attachment(ms, new ContentType(contentType));
                attch.ContentId = fileName;
                attch.ContentDisposition.FileName = fileName;
                attch.TransferEncoding = TransferEncoding.Base64;

                _mailMessage.Attachments.Add(attch);
                ms.Close();
            }

        }

        #endregion

        #endregion

        #region SetKeyValue
        /// <summary>
        /// Reemplaza el "Key" del Template cargado en Body
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="Value"></param>
        public void SetKeyValue(string Key, string Value)
        {

            this.HtmlBody = this.HtmlBody.Replace(Key.Trim().ToLower(), Value.Trim());
            this.TextBody = this.TextBody.Replace(Key.Trim().ToLower(), Value.Trim());

            this.HtmlBody = this.HtmlBody.Replace(Key.Trim().ToUpper(), Value.Trim());
            this.TextBody = this.TextBody.Replace(Key.Trim().ToUpper(), Value.Trim());
        }
        #endregion

        #region getMimeType
        /// <summary>
        /// Obtener el MIME Type del archivo
        /// </summary>
        public static string getMimeType(FileInfo file)
        {
            try
            {
                String ext = file.Extension.ToLower();

                RegistryPermission regPerm = new RegistryPermission(RegistryPermissionAccess.Read, "\\HKEY_CLASSES_ROOT");
                RegistryKey classesRoot = Registry.ClassesRoot;
                RegistryKey typeKey = classesRoot.OpenSubKey(@"MIME\Database\Content Type");

                foreach (String keyName in typeKey.GetSubKeyNames())
                {
                    RegistryKey currKey = classesRoot.OpenSubKey(String.Concat(@"MIME\Database\Content Type\", keyName));
                    if (currKey.GetValue("Extension").ToString().ToLower() == ext)
                        return keyName;
                }
            }
            catch { /* Ignorar errores */  }

            return string.Empty;

        }
        #endregion

        #endregion

        #region PRIVATE METHODS

        private void SetMailMessageValues()
        {
            if (_mailMessage.From == null)
                throw new Exception("Debe definir el correo 'From' (Obj.SetFrom)");

            if (_mailMessage.To.Count == 0 &&
                _mailMessage.Bcc.Count == 0 &&
                _mailMessage.CC.Count == 0)
                throw new Exception("Debe agregar por lo menos a un destinatario (Obj.Add[To|CC|BCC])");

            

            if (this.HtmlBody.Length > 0)
            {
                _avHtmlBody = AlternateView.CreateAlternateViewFromString(this.HtmlBody, null, "text/html");
                _avHtmlBody.TransferEncoding = TransferEncoding.Base64;
                
                // Si hay imágenes...
                foreach(EmailImage im in _emailImages)
                    _avHtmlBody.LinkedResources.Add(im.LinkedResource);

                _mailMessage.AlternateViews.Add(_avHtmlBody);
            }

            if (this.TextBody.Length > 0)
            {
                _avTextBody = AlternateView.CreateAlternateViewFromString(this.TextBody, null, "text/plain");
                _avTextBody.TransferEncoding = TransferEncoding.Base64;

                _mailMessage.AlternateViews.Add(_avTextBody);
            }
            
        }

        private void _smtpClient_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
                this.OnSendError(e.Error);
            else
                this.OnSendCompleted();
        }

        private byte[] GetBufferFromBitmap(Bitmap img, ImageFormat format)
        {
            byte[] buffer = new byte[0];
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, format);
                ms.Position = 0;
                buffer = new byte[ms.Length];
                ms.Read(buffer, 0, buffer.Length);
                ms.Close();
            }
            return buffer;
        }


        #endregion



        public void Dispose()
        {
            foreach (EmailImage im in _emailImages)
            {
                try
                {
                    if (im.Filename.EndsWith(".tmp"))
                    {
                        try { File.Delete(im.Filename); }
                        catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* no se logró eliminar el archivo temporal */ }
                    }
                }
                catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /*Algo anda mal en el object... */ }
            }


            try
            {
                _mailMessage.Attachments.Clear();
                _mailMessage.Attachments.Dispose();
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}


            try
            {
                _mailMessage.AlternateViews.Clear();
                _mailMessage.AlternateViews.Dispose();
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}

            try
            {
                _emailImages.Dispose();
                _emailImages = null;
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}


            try
            {
                _avHtmlBody.Dispose();
                _avHtmlBody = null;
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}


            try
            {
                _avTextBody.Dispose();
                _avTextBody = null;
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}


            try
            {
                _mailMessage.Dispose();
                _mailMessage = null;
            }
            catch (Exception ex) { System.Diagnostics.Debug.Print("{0}: DM.Core.Net.DotNetEmailClient->Dispose()::Exception => {1}", DateTime.Now.ToString("YYYY-MM-dd HH:mm:ss"), ex.Message); /* nafin */}
        }
    }


}
