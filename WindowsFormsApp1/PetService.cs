using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1
{
    public class PetService : IPetService
    {
        private readonly IPetRepository _rep;

        public PetService(
            IPetRepository rep)
            : base()
        {
            var msg = $"{this.GetType().Name} expects ctor injection.";

            this._rep = rep ?? throw new ArgumentNullException(
                msg);
        }
        public Task<IList<PetData>> FindAllFiles()
        {
            return Task.FromResult(
                this._rep
                    .FindAllFiles());
        }

        public void PrintExcelFile(int DataId, string fname)
        {
               Task.Run(() => _rep.PrintExcelFile(DataId, fname));
        }
    }
}
