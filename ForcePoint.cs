namespace LateralLoadApp
{
    public class ForcePoint
    {
        public string Joint { get; set; }
        public double? X { get; set; }
        public double? Y { get; set; }
        public double? Z { get; set; }
        public double Fx { get; set; }
        public double Fy { get; set; }
        public double Fz { get; set; }
        public double Mx { get; set; }
        public double My { get; set; }
        public double Mz { get; set; }

        // Constructor to initialize with default values
        public ForcePoint()
        {
            Fx = 0.0;
            Fy = 0.0;
            Fz = 0.0;
            Mx = 0.0;
            My = 0.0;
            Mz = 0.0;
        }

        // Optional method to display the values of the point
        public override string ToString()
        {
            return $"Joint: {Joint}, X: {X}, Y: {Y}, Z: {Z}, Fx: {Fx}, Fy: {Fy}, Fz: {Fz}, Mx: {Mx}, My: {My}, Mz: {Mz}";
        }
    }
}
