using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class DownloadCompressedFileResponse
	{
		public string Status { get; set; }
		public string Message { get; set; } = "";
		public int CompressedFilesQuantity = 0;
		public string CompressedFileName = "";
	}
}