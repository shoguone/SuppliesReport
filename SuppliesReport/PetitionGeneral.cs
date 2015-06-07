using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SuppliesReport
{
    public class PetitionGeneral
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

        private string _cost;

        public string Cost
        {
            get { return _cost; }
            set { _cost = value; }
        }

        public float CostFloat
        {
            get
            {
                float cost;
                float.TryParse(Cost.Replace('.', ','), out cost);
                return cost;
            }
        }

        public float SummaryCost
        {
            get
            {
                //float cost, count, mult;
                //float.TryParse(Cost.Replace('.', ','), out cost);
                //float.TryParse(Count.Replace('.', ','), out count);
                //mult = cost * count;
                return CostFloat * CountFloat;
            }
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

        public float CountFloat
        {
            get
            {
                float count;
                float.TryParse(Count.Replace('.', ','), out count);
                return count;
            }
        }

        private string _manufacturingYear;

        public string ManufacuringYear
        {
            get { return _manufacturingYear; }
            set { _manufacturingYear = value; }
        }

        private DateTime _acceptedDate;

        public DateTime AcceptedDate
        {
            get { return _acceptedDate; }
            set { _acceptedDate = value; }
        }

        private string _lifeTime;

        public string LifeTime
        {
            get { return _lifeTime; }
            set { _lifeTime = value; }
        }

        private string _actualPeriod;

        public string ActualPeriod
        {
            get { return _actualPeriod; }
            set { _actualPeriod = value; }
        }


    }
}
