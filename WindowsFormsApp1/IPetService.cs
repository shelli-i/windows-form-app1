using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1
{
    public interface IPetService
    {
        Task<IList<PetData>> FindAllFiles();
        void PrintExcelFile(int datasrc, string fname);
    }
}
