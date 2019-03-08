namespace ExportDXF
{
    internal class Bend
    {
        public BendDirection Direction { get; set; }

        public double ParallelBendAngle { get; set; }

        public double Angle { get; set; }

        public double X { get; set; }
        public double Y { get; set; }
    }
}