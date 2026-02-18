using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace FabWorks.Core.PressBrake
{
    public class ProgramReader
    {
        public Program Program { get; set; }

        public ProgramReader()
        {
            Program = new Program();
        }

        public void Read(string file)
        {
            var xml = XDocument.Load(file);
            Program.FilePath = file;
            Read(xml);
        }

        public void Read(Stream stream)
        {
            var xml = XDocument.Load(stream);
            Read(xml);
        }

        private void Read(XDocument doc)
        {
            var data = doc.Root.Element("PressBrakeProgram");

            Program.Version = data.Attribute("Version").ToInt();
            Program.ProgName = data.Attribute("ProgName")?.Value;
            Program.MatThick = data.Attribute("MatThick").ToDouble();
            Program.MatType = GetMaterialType(data.Attribute("MatType")?.Value);
            Program.KFactor = data.Attribute("KFactor").ToDouble();
            Program.TeachName = data.Attribute("TeachName")?.Value;
            Program.PartName = data.Attribute("PartName")?.Value;
            Program.SetupNotes = data.Attribute("SetupNotes")?.Value;
            Program.ProgNotes = data.Attribute("ProgNotes")?.Value;
            Program.RZEnabled = Convert.ToBoolean(data.Attribute("RZEnabled").ToInt());

            foreach (var item in data.Element("UpperToolSets").Descendants("ToolSetup"))
            {
                var setup = ReadToolSetup(item);
                Program.UpperToolSets.Add(setup);
            }

            foreach (var item in data.Element("LowerToolSets").Descendants("ToolSetup"))
            {
                var setup = ReadToolSetup(item);
                Program.LowerToolSets.Add(setup);
            }

            foreach (var item in data.Element("StepData").Descendants("Step"))
            {
                var step = ReadStep(item);
                step.UpperTool = Program.UpperToolSets.FirstOrDefault(t => t.Id == step.UpperID);
                step.LowerTool = Program.LowerToolSets.FirstOrDefault(t => t.Id == step.LowerID);

                Program.Steps.Add(step);
            }
        }

        private ToolSetup ReadToolSetup(XElement x)
        {
            var setup = new ToolSetup();

            setup.Name = x.Attribute("Name").Value;
            setup.Id = x.Attribute("ID").ToInt();
            setup.Length = x.Attribute("Length").ToDouble();
			setup.StackedHolderType = x.Attribute("StackedHolderType").ToInt();
			setup.HolderHeight = x.Attribute("HolderHeight").ToDouble();

			foreach (var item in x.Descendants("SegEntry"))
            {
                var entry = new SegEntry();
                entry.SegValue = item.Attribute("SegValue").ToDouble();
                setup.Segments.Add(entry);
            }

            return setup;
        }

        private Step ReadStep(XElement x)
        {
            var step = new Step();

            step.RevMode = x.Attribute("RevMode").ToInt();
            step.RevTons = x.Attribute("RevTons").ToDouble();
            step.MaxTons = x.Attribute("MaxTons").ToDouble();
            step.RevAbsPos = x.Attribute("RevAbsPos").ToDouble();
            step.ActualAng = x.Attribute("ActualAng").ToDouble();
            step.AngleAdj = x.Attribute("AngleAdj").ToDouble();
            step.BendLen = x.Attribute("BendLen").ToDouble();
            step.StrokeLen = x.Attribute("StrokeLen").ToDouble();
            step.UpperID = x.Attribute("UpperID").ToInt();
            step.LowerID = x.Attribute("LowerID").ToInt();
            step.SpdChgDwn = x.Attribute("SpdChgDwn").ToDouble();
            step.SpdChgUp = x.Attribute("SpdChgUp").ToDouble();
            step.Tilt = x.Attribute("Tilt").ToDouble();
            step.FormSpeed = x.Attribute("FormSpeed").ToDouble();
            step.XLeft = x.Attribute("XLeft").ToDouble();
            step.XRight = x.Attribute("XRight").ToDouble();
            step.RLeft = x.Attribute("RLeft").ToDouble();
            step.RRight = x.Attribute("RRight").ToDouble();
            step.ZLeft = x.Attribute("ZLeft").ToDouble();
            step.ZRight = x.Attribute("ZRight").ToDouble();
            step.FLeft = x.Attribute("FLeft").ToDouble();
            step.FRight = x.Attribute("FRight").ToDouble();
            step.SSLeft = x.Attribute("SSLeft").ToDouble();
            step.SSRight = x.Attribute("SSRight").ToDouble();
            step.ReturnSpd = x.Attribute("ReturnSpd").ToDouble();
            step.SideFlgHeight = x.Attribute("SideFlgHeight").ToDouble();

            return step;
        }

		private MatType GetMaterialType(string value)
		{
			if (value == null)
				return MatType.MildSteel;

			int i;

			if (!int.TryParse(value, out i))
				return MatType.MildSteel;

			switch (i)
			{
				case 0:
					return MatType.MildSteel;
				case 1:
					return MatType.HighStrengthSteel;
				case 2:
					return MatType.Stainless;
				case 3:
					return MatType.SoftAluminum;
				case 4:
					return MatType.HardAluminum;
			}

			return MatType.MildSteel;
		}
    }
}
