using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SuppliesReport
{
    public class ActOfTestingAndInspection
    {
        //private int _id;

        //public int ID
        //{
        //    get { return _id; }
        //    set { _id = value; }
        //}

        private string _consecutiveNumber;

        public string ConsecutiveNumber
        {
            get { return _consecutiveNumber; }
            set { _consecutiveNumber = value; }
        }

        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        private string _inventoryNumber;

        public string InventoryNumber
        {
            get { return _inventoryNumber; }
            set { _inventoryNumber = value; }
        }

        private string _count;

        public string Count
        {
            get { return _count; }
            set { _count = value; }
        }

        private string _cost;

        public string Cost
        {
            get { return _cost; }
            set { _cost = value; }
        }

        private string _summaryCost;

        public string SummaryCost
        {
            get { return _summaryCost; }
            set { _summaryCost = value; }
        }

        public ActOfTestingAndInspection(PetitionGeneral petition)
        {
            ConsecutiveNumber = petition.ConsecutiveNumber;
            Name = petition.Name;
            Count = petition.Count;
            Cost = petition.Cost;
            SummaryCost = (float.Parse(petition.Cost) * float.Parse(Count)).ToString("F2");
        }

    }
}
