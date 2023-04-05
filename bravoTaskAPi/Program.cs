using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Synthesis;


namespace bravoTaskAPi
{
    internal class Program
    {
        static void Main(string[] args)
        {
            SpeechSynthesizer s=new SpeechSynthesizer();
            s.Volume=50;
            s.Speak("Where is Haitham...........   Haitham Where");

        }
    }
}
