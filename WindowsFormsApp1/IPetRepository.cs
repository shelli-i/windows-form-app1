using System;
using System.Collections.Generic;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1
{
    public interface IPetRepository
    {
        IList<PetData> FindAllFiles();
        void PrintExcelFile(int datasrc, string fname);
    }
}
