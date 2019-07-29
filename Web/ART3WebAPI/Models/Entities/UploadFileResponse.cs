using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ Class UploadFileResponse ]
	public class UploadFileResponse
	{
		public string Status { get; set; }
		public string Message { get; set; } = "";
		public UploadFileItemResponse[] files { get; set; }
	}
	#endregion

	#region [ Class UploadFileItemResponse ]
	public class UploadFileItemResponse
	{
		public string field_name { get; set; }
		public string original_file_name { get; set; }
		public string stored_file_guid { get; set; }
		public string stored_file_name { get; set; }
		public string folder_name { get; set; }
	}
	#endregion
}