using ExportDXF.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;

namespace ExportDXF.Services
{
    /// <summary>
    /// Helper class to get COM objects from Running Object Table (ROT) in .NET Core/5+
    /// </summary>
    internal static class ComHelper
    {
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx bindCtx);

        public static object GetActiveObject(string progId)
        {
            IRunningObjectTable rot = null;
            IEnumMoniker enumMoniker = null;
            IBindCtx bindCtx = null;

            try
            {
                if (GetRunningObjectTable(0, out rot) != 0 || rot == null)
                    return null;

                rot.EnumRunning(out enumMoniker);
                if (enumMoniker == null)
                    return null;

                if (CreateBindCtx(0, out bindCtx) != 0 || bindCtx == null)
                    return null;

                IMoniker[] monikers = new IMoniker[1];
                IntPtr fetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, fetched) == 0)
                {
                    monikers[0].GetDisplayName(bindCtx, null, out string displayName);

                    if (displayName != null && displayName.IndexOf(progId, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        rot.GetObject(monikers[0], out object obj);
                        return obj;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
                if (rot != null) Marshal.ReleaseComObject(rot);
                if (bindCtx != null) Marshal.ReleaseComObject(bindCtx);
            }
        }
    }

    /// <summary>
    /// Service for managing SolidWorks application connection and document lifecycle.
    /// </summary>
    public interface ISolidWorksService : IDisposable
    {
        /// <summary>
        /// Connects to the SolidWorks application asynchronously.
        /// </summary>
        Task ConnectAsync();

        /// <summary>
        /// Gets the currently active SolidWorks document.
        /// </summary>
        /// <returns>The active document or null if no document is open.</returns>
        SolidWorksDocument GetActiveDocument();

        /// <summary>
        /// Enables or disables user control of the SolidWorks application.
        /// </summary>
        /// <param name="enable">True to enable user control, false to disable.</param>
        void EnableUserControl(bool enable);

        /// <summary>
        /// Gets the SolidWorks application instance.
        /// </summary>
        /// <returns>The SldWorks instance.</returns>
        SldWorks GetNativeSldWorks();

        /// <summary>
        /// Closes the document with the specified title.
        /// </summary>
        /// <param name="documentTitle"></param>
        /// <returns></returns>
        bool CloseDocument(string documentTitle);

        /// <summary>
        /// Event raised when the active document changes.
        /// </summary>
        event EventHandler ActiveDocumentChanged;
    }

    /// <summary>
    /// Service for managing SolidWorks application connection and document lifecycle.
    /// </summary>
    public class SolidWorksService : ISolidWorksService
    {
        private SldWorks _sldWorks;
        private bool _disposed;

        /// <summary>
        /// Event raised when the active document changes in SolidWorks.
        /// </summary>
        public event EventHandler ActiveDocumentChanged;

        /// <summary>
        /// Connects to the SolidWorks application asynchronously.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when connection to SolidWorks fails.</exception>
        public async Task ConnectAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    // Try to get running instance first
                    _sldWorks = GetRunningInstance();

                    // If no running instance, create new one
                    if (_sldWorks == null)
                    {
                        _sldWorks = CreateNewInstance();
                    }

                    if (_sldWorks == null)
                    {
                        throw new InvalidOperationException(
                            "Failed to connect to SolidWorks. Please ensure SolidWorks is installed.");
                    }

                    // Make SolidWorks visible
                    _sldWorks.Visible = true;

                    // Subscribe to document change events
                    _sldWorks.ActiveModelDocChangeNotify += OnSolidWorksActiveDocChanged;
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException(
                        $"COM error while connecting to SolidWorks: {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// Gets the currently active SolidWorks document.
        /// </summary>
        /// <returns>The active document wrapper, or null if no document is open.</returns>
        public SolidWorksDocument GetActiveDocument()
        {
            if (_sldWorks == null)
                return null;

            var model = _sldWorks.ActiveDoc as ModelDoc2;
            if (model == null)
                return null;

            return CreateDocumentWrapper(model);
        }

        /// <summary>
        /// Enables or disables user control of the SolidWorks application.
        /// When disabled, SolidWorks won't show dialogs or allow user interaction.
        /// </summary>
        /// <param name="enable">True to enable user control, false to disable.</param>
        public void EnableUserControl(bool enable)
        {
            if (_sldWorks != null)
            {
                _sldWorks.UserControl = enable;
            }
        }

        /// <summary>
        /// Gets the native SolidWorks application instance.
        /// Use this when you need direct access to the SolidWorks API.
        /// </summary>
        /// <returns>The SldWorks instance, or null if not connected.</returns>
        public SldWorks GetNativeSldWorks()
        {
            return _sldWorks;
        }

        /// <summary>
        /// Checks if SolidWorks is connected and ready.
        /// </summary>
        public bool IsConnected => _sldWorks != null;

        /// <summary>
        /// Gets the version of the connected SolidWorks instance.
        /// </summary>
        public string GetSolidWorksVersion()
        {
            if (_sldWorks == null)
                return null;

            try
            {
                return _sldWorks.RevisionNumber();
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// Closes a document by its title.
        /// </summary>
        /// <param name="documentTitle">The title of the document to close.</param>
        /// <returns>True if successfully closed, false otherwise.</returns>
        public bool CloseDocument(string documentTitle)
        {
            if (_sldWorks == null || string.IsNullOrWhiteSpace(documentTitle))
                return false;

            try
            {
                _sldWorks.CloseDoc(documentTitle);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Creates a new document from a template.
        /// </summary>
        /// <param name="templatePath">Path to the template file.</param>
        /// <param name="paperSize">Paper size for drawings.</param>
        /// <param name="width">Width dimension.</param>
        /// <param name="height">Height dimension.</param>
        /// <returns>The created document, or null if creation failed.</returns>
        public ModelDoc2 CreateDocument(string templatePath, int paperSize, double width, double height)
        {
            if (_sldWorks == null)
                return null;

            try
            {
                return _sldWorks.NewDocument(templatePath, paperSize, width, height) as ModelDoc2;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Opens an existing document.
        /// </summary>
        /// <param name="filePath">Full path to the document.</param>
        /// <param name="options">Open options.</param>
        /// <param name="configuration">Configuration to open (empty for default).</param>
        /// <returns>The opened document, or null if open failed.</returns>
        public ModelDoc2 OpenDocument(string filePath, int options = 0, string configuration = "")
        {
            if (_sldWorks == null || string.IsNullOrWhiteSpace(filePath))
                return null;

            try
            {
                int errors = 0;
                int warnings = 0;

                var doc = _sldWorks.OpenDoc6(
                    filePath,
                    (int)GetDocumentType(filePath),
                    options,
                    configuration,
                    ref errors,
                    ref warnings);

                return doc as ModelDoc2;
            }
            catch
            {
                return null;
            }
        }

        #region Private Methods

        private SldWorks GetRunningInstance()
        {
            try
            {
                return ComHelper.GetActiveObject("SldWorks.Application") as SldWorks;
            }
            catch (COMException)
            {
                // No running instance
                return null;
            }
        }

        private SldWorks CreateNewInstance()
        {
            try
            {
                var type = Type.GetTypeFromProgID("SldWorks.Application");
                if (type == null)
                    return null;

                return Activator.CreateInstance(type) as SldWorks;
            }
            catch
            {
                return null;
            }
        }

        private SolidWorksDocument CreateDocumentWrapper(ModelDoc2 model)
        {
            return new SolidWorksDocument
            {
                Title = model.GetTitle(),
                FilePath = model.GetPathName(),
                DocumentType = DetermineDocumentType(model),
                NativeDocument = model
            };
        }

        private DocumentType DetermineDocumentType(ModelDoc2 model)
        {
            if (model is PartDoc) return DocumentType.Part;
            if (model is AssemblyDoc) return DocumentType.Assembly;
            if (model is DrawingDoc) return DocumentType.Drawing;
            return DocumentType.Unknown;
        }

        private swDocumentTypes_e GetDocumentType(string filePath)
        {
            var extension = System.IO.Path.GetExtension(filePath)?.ToLowerInvariant();

            switch (extension)
            {
                case ".sldprt":
                    return swDocumentTypes_e.swDocPART;
                case ".sldasm":
                    return swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw":
                    return swDocumentTypes_e.swDocDRAWING;
                default:
                    return swDocumentTypes_e.swDocNONE;
            }
        }

        private int OnSolidWorksActiveDocChanged()
        {
            try
            {
                ActiveDocumentChanged?.Invoke(this, EventArgs.Empty);
            }
            catch
            {
                // Swallow exceptions from event handlers
            }

            return 1; // SolidWorks expects 1 to be returned
        }

        #endregion

        #region IDisposable Implementation

        /// <summary>
        /// Releases all resources used by the SolidWorksService.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases unmanaged and optionally managed resources.
        /// </summary>
        /// <param name="disposing">True to release both managed and unmanaged resources; 
        /// false to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Unsubscribe from events
                if (_sldWorks != null)
                {
                    try
                    {
                        _sldWorks.ActiveModelDocChangeNotify -= OnSolidWorksActiveDocChanged;
                    }
                    catch
                    {
                        // Ignore errors during cleanup
                    }
                }
            }

            // Release COM object
            if (_sldWorks != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_sldWorks);
                }
                catch
                {
                    // Ignore errors during cleanup
                }

                _sldWorks = null;
            }

            _disposed = true;
        }

        public SldWorks GetSldWorks()
        {
            if (_sldWorks == null)
            {
                _sldWorks = GetRunningInstance();
                return _sldWorks;
            }
            else
            {
                _sldWorks = CreateNewInstance();
                return _sldWorks;
            }
        }

        /// <summary>
        /// Finalizer to ensure COM objects are released.
        /// </summary>
        ~SolidWorksService()
        {
            Dispose(false);
        }

        #endregion
    }
}


