using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mimeo.MimeoConnect
{
	public class Address
	{
		public string lastName
		{
			get;
			set;
		}
		public string firstName
		{
			get;
			set;
		}
		public string street
		{
			get;
			set;
		}

        public string apartmentOrSuite
        {
            get;
            set;
        }

        public string careOf
        {
            get;
            set;
        }
		public string city
		{
			get;
			set;
		}

		public string country
		{
			get;
			set;
		}

		public string state
		{
			get;
			set;
		}
		public string postalCode
		{
			get;
			set;
		}

		public string telephone
		{
			get;
			set;
		}
        public string email
        {
            get;
            set;
        }
		public string shipping
		{
			get;
			set;
		}

		public int additionalProcessing
		{
			get;
			set;
		}
	}
}
