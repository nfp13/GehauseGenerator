using System;


namespace GehäuseGenerator
{
    public class Status
    {
        public int Progress = 0; // 1 -> 100%

        public string Name;

        public event EventHandler Progressed;
        public void OnProgess()
        {
            EventHandler handler = Progressed;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }

}
    

