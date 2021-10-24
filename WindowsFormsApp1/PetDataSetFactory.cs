using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1
{
    internal class PetDataSetFactory
    {
        internal PetData Create(IDataReader dr)
        {
            return new PetData
            {
                DataSrcID = Int32.Parse(dr["DataSrcID"].ToString()),
                fileName = dr["File Name"].ToString(),
                userID = dr["UserID"].ToString(),
                DateEntry = dr["DateEntry"].ToString(),
                Count = Int32.Parse(dr["Count"].ToString())

            };
        }
    }
}
