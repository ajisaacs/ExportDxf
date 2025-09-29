using ExportDXF.ViewFlipDeciders;
using System;
using System.Drawing;
using System.Threading;

namespace ExportDXF.Services
{
    /// <summary>
    /// Context object containing all information needed for an export operation.
    /// </summary>
    public class ExportContext
    {
        /// <summary>
        /// The document to be exported.
        /// </summary>
        public SolidWorksDocument ActiveDocument { get; set; }

        /// <summary>
        /// The view flip decider to determine if views should be flipped.
        /// </summary>
        public IViewFlipDecider ViewFlipDecider { get; set; }

        /// <summary>
        /// Prefix to prepend to exported filenames.
        /// </summary>
        public string FilePrefix { get; set; }

        /// <summary>
        /// Cancellation token for canceling the export operation.
        /// </summary>
        public CancellationToken CancellationToken { get; set; }

        /// <summary>
        /// Callback for reporting progress and status messages.
        /// </summary>
        public Action<string, Color?> ProgressCallback { get; set; }
    }
}