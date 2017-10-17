namespace ExportDXF
{
	class Bounds
	{
		public double X { get; set; }
		public double Y { get; set; }
		public double Width { get; set; }
		public double Height { get; set; }

		public double Left
		{
			get { return X; }
		}

		public double Right
		{
			get { return X + Width; }
		}

		public double Bottom
		{
			get { return Y; }
		}

		public double Top
		{
			get { return Y + Height; }
		}
	}
}
