using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rederi.Functions
{
	public class GroupRequest
	{
		[Required]
		public int ID { get; set; }
		[Required]
		public string ListId { get; set; }
		[Required]
		public string WebUrl { get; set; }
	}
}
