using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class UploadStoredFileInfo
	{
		public int id { get; set; }
		public string guid { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public byte st_temporary_file { get; set; }
		public byte st_confirmation_required { get; set; }
		public string original_file_name { get; set; }
		public string original_full_file_name { get; set; }
		public string stored_file_name { get; set; }
		public string stored_full_file_name { get; set; }
		public string stored_relative_path { get; set; }
		public int id_module_folder_name { get; set; }
		public long file_size { get; set; }
		public string remote_IP { get; set; }
		public byte[] file_content { get; set; }
		public string file_content_text { get; set; }
	}
}