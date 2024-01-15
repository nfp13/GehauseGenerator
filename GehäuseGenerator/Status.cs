using System;


namespace GehäuseGenerator
{
    public class Status
    {
        public int Progress = 0;    // 0 = 0% ; 100 = 100%

        public string Name;         //Name des aktuellen Prozesses
            
        public event EventHandler Progressed;
        public void OnProgess()     //ruft event Progressed auf
        {
            EventHandler handler = Progressed;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }
}