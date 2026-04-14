using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FabWorks.Core.PressBrake
{
    public class Program
    {
        public Program()
        {
            UpperToolSets = new List<ToolSetup>();
            LowerToolSets = new List<ToolSetup>();
            Steps = new List<Step>();
        }

        public int Version { get; set; }

        public string ProgName { get; set; }

        public string FilePath { get; set; }

        public double MatThick { get; set; }

		public MatType MatType { get; set; }

		public double KFactor { get; set; }

        public string TeachName { get; set; }

        public string PartName { get; set; }

        public string SetupNotes { get; set; }

        public string ProgNotes { get; set; }

        public bool RZEnabled { get; set; }

        public List<ToolSetup> UpperToolSets { get; set; }

        public List<ToolSetup> LowerToolSets { get; set; }

        public List<Step> Steps { get; set; }

        public static Program Load(string file)
        {
            var reader = new ProgramReader();
            reader.Read(file);
            return reader.Program;
        }

        public static Program Load(Stream stream)
        {
            var reader = new ProgramReader();
            reader.Read(stream);
            return reader.Program;
        }
    }
}
