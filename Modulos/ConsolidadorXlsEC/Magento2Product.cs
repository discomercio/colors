using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	#region [ Magento2Product ]
	public class Magento2Product
	{
		public string id { get; set; }
		public string sku { get; set; }
		public string name { get; set; }
		public string attribute_set_id { get; set; }
		public string price { get; set; }
		public string status { get; set; }
		public string visibility { get; set; }
		public string type_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string weight { get; set; }
		public Magento2ProductExtensionAttributes extension_attributes { get; set; }
		// product_links (array): campo ignorado por ter estrutura desconhecida
		public List<Magento2ProductOptions> options { get; set; } = new List<Magento2ProductOptions>();
		public List<Magento2ProductMediaGalleryEntries> media_gallery_entries { get; set; } = new List<Magento2ProductMediaGalleryEntries>();
		// tier_prices (array): campo ignorado por ter estrutura desconhecida
		public List<Magento2ProductCustomAttributes> custom_attributes { get; set; } = new List<Magento2ProductCustomAttributes>();

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			int iCounter;
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2Product).GetProperties())
			{
				#region [ extension_attributes ]
				if (prop.Name.Equals("extension_attributes"))
				{
					if (this.extension_attributes != null)
					{
						sbResp.AppendLine(margem + "extension_attributes");
						sbResp.Append(this.extension_attributes.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "extension_attributes = null");
					}

					continue;
				}
				#endregion

				#region [ options ]
				if (prop.Name.Equals("options"))
				{
					if (this.options.Count == 0)
					{
						sbResp.AppendLine(margem + "options = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductOptions item in this.options)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "options (" + iCounter.ToString() + "/" + this.options.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "options [" + this.options.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ media_gallery_entries ]
				if (prop.Name.Equals("media_gallery_entries"))
				{
					if (this.media_gallery_entries.Count == 0)
					{
						sbResp.AppendLine(margem + "media_gallery_entries = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductMediaGalleryEntries item in this.media_gallery_entries)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "media_gallery_entries (" + iCounter.ToString() + "/" + this.media_gallery_entries.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "media_gallery_entries [" + this.media_gallery_entries.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ custom_attributes ]
				if (prop.Name.Equals("custom_attributes"))
				{
					if (this.custom_attributes.Count == 0)
					{
						sbResp.AppendLine(margem + "custom_attributes = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductCustomAttributes item in this.custom_attributes)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "custom_attributes (" + iCounter.ToString() + "/" + this.custom_attributes.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "custom_attributes [" + this.custom_attributes.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductExtensionAttributes ]
	public class Magento2ProductExtensionAttributes
	{
		public List<string> website_ids { get; set; }
		public List<Magento2ProductExtensionAttributesCategoryLinks> category_links { get; set; } = new List<Magento2ProductExtensionAttributesCategoryLinks>();
		public Magento2ProductExtensionAttributesStockItem stock_item { get; set; }
		public List<Magento2ProductExtensionAttributesConfigurableProductOptions> configurable_product_options { get; set; } = new List<Magento2ProductExtensionAttributesConfigurableProductOptions>();
		public List<string> configurable_product_links { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			int iCounter;
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductExtensionAttributes).GetProperties())
			{
				#region [ website_ids ]
				if (prop.Name.Equals("website_ids"))
				{
					if (this.website_ids == null)
					{
						sbResp.AppendLine(margem + "website_ids = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.website_ids)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "website_ids");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ category_links ]
				if (prop.Name.Equals("category_links"))
				{
					if (this.category_links.Count == 0)
					{
						sbResp.AppendLine(margem + "category_links = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductExtensionAttributesCategoryLinks item in this.category_links)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "category_links (" + iCounter.ToString() + "/" + this.category_links.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "category_links [" + this.category_links.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ stock_item ]
				if (prop.Name.Equals("stock_item"))
				{
					if (this.stock_item != null)
					{
						sbResp.AppendLine(margem + "stock_item");
						sbResp.Append(this.stock_item.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "stock_item = null");
					}

					continue;
				}
				#endregion

				#region [ configurable_product_options ]
				if (prop.Name.Equals("configurable_product_options"))
				{
					if (this.configurable_product_options.Count == 0)
					{
						sbResp.AppendLine(margem + "configurable_product_options = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductExtensionAttributesConfigurableProductOptions item in this.configurable_product_options)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "configurable_product_options (" + iCounter.ToString() + "/" + this.configurable_product_options.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "configurable_product_options [" + this.configurable_product_options.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ configurable_product_links ]
				if (prop.Name.Equals("configurable_product_links"))
				{
					if (this.configurable_product_links == null)
					{
						sbResp.AppendLine(margem + "configurable_product_links = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.configurable_product_links)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "configurable_product_links");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductExtensionAttributesCategoryLinks ]
	public class Magento2ProductExtensionAttributesCategoryLinks
	{
		public string position { get; set; }
		public string category_id { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductExtensionAttributesCategoryLinks).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductExtensionAttributesStockItem ]
	public class Magento2ProductExtensionAttributesStockItem
	{
		public string item_id { get; set; }
		public string product_id { get; set; }
		public string stock_id { get; set; }
		public string qty { get; set; }
		public string is_in_stock { get; set; }
		public string is_qty_decimal { get; set; }
		public string show_default_notification_message { get; set; }
		public string use_config_min_qty { get; set; }
		public string min_qty { get; set; }
		public string use_config_min_sale_qty { get; set; }
		public string min_sale_qty { get; set; }
		public string use_config_max_sale_qty { get; set; }
		public string max_sale_qty { get; set; }
		public string use_config_backorders { get; set; }
		public string backorders { get; set; }
		public string use_config_notify_stock_qty { get; set; }
		public string notify_stock_qty { get; set; }
		public string use_config_qty_increments { get; set; }
		public string qty_increments { get; set; }
		public string use_config_enable_qty_inc { get; set; }
		public string enable_qty_increments { get; set; }
		public string use_config_manage_stock { get; set; }
		public string manage_stock { get; set; }
		public string low_stock_date { get; set; }
		public string is_decimal_divided { get; set; }
		public string stock_status_changed_auto { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductExtensionAttributesStockItem).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductExtensionAttributesConfigurableProductOptions ]
	public class Magento2ProductExtensionAttributesConfigurableProductOptions
	{
		public string id { get; set; }
		public string attribute_id { get; set; }
		public string label { get; set; }
		public string position { get; set; }
		public List<Magento2ProductExtensionAttributesConfigurableProductOptionsValues> values { get; set; } = new List<Magento2ProductExtensionAttributesConfigurableProductOptionsValues>();
		public string product_id { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			int iCounter;
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductExtensionAttributesConfigurableProductOptions).GetProperties())
			{
				#region [ values ]
				if (prop.Name.Equals("values"))
				{
					if (this.values.Count == 0)
					{
						sbResp.AppendLine(margem + "values = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductExtensionAttributesConfigurableProductOptionsValues item in this.values)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "values (" + iCounter.ToString() + "/" + this.values.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "values [" + this.values.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductExtensionAttributesConfigurableProductOptionsValues ]
	public class Magento2ProductExtensionAttributesConfigurableProductOptionsValues
	{
		public string value_index { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductExtensionAttributesConfigurableProductOptionsValues).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductOptions ]
	public class Magento2ProductOptions
	{
		public string product_sku { get; set; }
		public string option_id { get; set; }
		public string title { get; set; }
		public string type { get; set; }
		public string sort_order { get; set; }
		public string is_require { get; set; }
		public string price { get; set; }
		public string price_type { get; set; }
		public string sku { get; set; }
		public string max_characters { get; set; }
		public string image_size_x { get; set; }
		public string image_size_y { get; set; }
		public List<Magento2ProductOptionsValues> values { get; set; } = new List<Magento2ProductOptionsValues>();

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			int iCounter;
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductOptions).GetProperties())
			{
				#region [ values ]
				if (prop.Name.Equals("values"))
				{
					if (this.values.Count == 0)
					{
						sbResp.AppendLine(margem + "values = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ProductOptionsValues item in this.values)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "values (" + iCounter.ToString() + "/" + this.values.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "values [" + this.values.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductOptionsValues ]
	public class Magento2ProductOptionsValues
	{
		public string title { get; set; }
		public string sort_order { get; set; }
		public string price { get; set; }
		public string price_type { get; set; }
		public string option_type_id { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductOptionsValues).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductMediaGalleryEntries ]
	public class Magento2ProductMediaGalleryEntries
	{
		public string id { get; set; }
		public string media_type { get; set; }
		public string label { get; set; }
		public string position { get; set; }
		public string disabled { get; set; }
		public List<string> types { get; set; }
		public string file { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbAux;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductMediaGalleryEntries).GetProperties())
			{
				#region [ types ]
				if (prop.Name.Equals("types"))
				{
					if (this.types == null)
					{
						sbResp.AppendLine(margem + "types = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.types)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "types");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ProductCustomAttributes ]
	public class Magento2ProductCustomAttributes
	{
		public string attribute_code { get; set; }
		// Tratamento com conversor customizado porque o campo 'value' às vezes retorna como string e às vezes como um array de string.
		[JsonProperty("value")]
		[JsonConverter(typeof(JsonSingleOrArrayConverter<string>))]
		public List<string> value { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbAux;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ProductCustomAttributes).GetProperties())
			{
				#region [ value ]
				if (prop.Name.Equals("value"))
				{
					if (this.value == null)
					{
						sbResp.AppendLine(margem + "value = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.value)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "value");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion
}
