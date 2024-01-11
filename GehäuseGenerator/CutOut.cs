namespace GehäuseGenerator
{
    public class CutOut
    {
        #region Properties
        public bool Top { get; set; }       //true if top
        public bool Connector { get; set; } //true if connector, false if LED or Display

        public double XP { get; set; }      //Coordianates of centerpoint of cutout from center of the Board
        public double YP { get; set; }
        public double ZP { get; set; }


        public double XS { get; set; }      //Size of the Cutout in X, Y and Z direction
        public double YS { get; set; }
        public double ZS { get; set; }

        public string Name { get; set; }

        public int Index { get; set; }
        #endregion

        #region Constructors
        public CutOut()
        {
        }
        #endregion

        #region Overrides
        public override string ToString()
        {
            return string.Format("{0}: {1}", Index, Name);
        }
        #endregion

    }

}
    

