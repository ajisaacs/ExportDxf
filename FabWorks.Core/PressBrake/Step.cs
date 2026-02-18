namespace FabWorks.Core.PressBrake
{
    public class Step
    {
        public int RevMode { get; set; }
        public double RevTons { get; set; }
        public double MaxTons { get; set; }
        public double RevAbsPos { get; set; }
        public double ActualAng { get; set; }
        public double AngleAdj { get; set; }
        public double BendLen { get; set; }
        public double StrokeLen { get; set; }
        public double Tilt { get; set; }
        public int UpperID { get; set; }
        public int LowerID { get; set; }
        public double SpdChgDwn { get; set; }
        public double SpdChgUp { get; set; }
        public double FormSpeed { get; set; }
        public double XLeft { get; set; }
        public double XRight { get; set; }
        public double RLeft { get; set; }
        public double RRight { get; set; }
        public double ZLeft { get; set; }
        public double ZRight { get; set; }
        public double FLeft { get; set; }
        public double FRight { get; set; }

        public double SSLeft { get; set; }
        public double SSRight { get; set; }
        public double ReturnSpd { get; set; }
        public double SideFlgHeight { get; set; }

        public ToolSetup UpperTool { get; set; }
        public ToolSetup LowerTool { get; set; }
    }
}
