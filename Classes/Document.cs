using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mimeo.MimeoConnect
{
	public class Document
	{
		public Guid id
		{
			get;
			set;
		}
		public string Name
		{
			get;
			set;
		}
		public int Quantity
		{
			get;
			set;
		}
	}
}
