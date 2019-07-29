using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	class ProductInfo
	{
		public string product_id { get; set; }
		public string sku { get; set; }
		public string set { get; set; }
		public string type { get; set; }
		public List<string> categories { get; set; } = new List<string>();
		public List<string> websites { get; set; } = new List<string>();
		public string type_id { get; set; }
		public string name { get; set; }
		public string titulo_ml { get; set; }
		public string weight { get; set; }
		public string cubagem { get; set; }
		public string old_id { get; set; }
		public string news_from_date { get; set; }
		public string news_to_date { get; set; }
		public string status { get; set; }
		public string url_key { get; set; }
		public string visibility { get; set; }
		public string url_path { get; set; }
		public string country_of_manufacture { get; set; }
		public string volume_comprimento { get; set; }
		public string volume_altura { get; set; }
		public List<string> category_ids { get; set; } = new List<string>();
		public string required_options { get; set; }
		public string volume_largura { get; set; }
		public string vender_buscape { get; set; }
		public string has_options { get; set; }
		public string image_label { get; set; }
		public string enviado_buscape { get; set; }
		public string small_image_label { get; set; }
		public string fretegratis { get; set; }
		public string thumbnail_label { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string price { get; set; }
		public string parcelamento_cartao { get; set; }
		public string parcela_cartao { get; set; }
		public string forma_pagamento { get; set; }
		public string group_price { get; set; }
		public string special_price { get; set; }
		public string special_from_date { get; set; }
		public string special_to_date { get; set; }
		public string minimal_price { get; set; }
		public string tier_price { get; set; }
		public string msrp_enabled { get; set; }
		public string msrp_display_actual_price_type { get; set; }
		public string msrp { get; set; }
		public string tax_class_id { get; set; }
		public string markup { get; set; }
		public string preco_calculado { get; set; }
		public string procel { get; set; }
		public string marca { get; set; }
		public string codigo_fabricante { get; set; }
		public string ean { get; set; }
		public string inverter { get; set; }
		public string voltagem { get; set; }
		public string temperatura { get; set; }
		public string inmetro { get; set; }
		public string consumo_energia { get; set; }
		public string trifasico { get; set; }
		public string medida_unidade_interma { get; set; }
		public string medida_unidade_externa { get; set; }
		public string short_description { get; set; }
		public string description { get; set; }
		public string detalhes { get; set; }
		public string additional_1 { get; set; }
		public string additional_2 { get; set; }
		public string detalhes_tecnicos_comparacao { get; set; }
		public string multi_split { get; set; }
		public string tipo { get; set; }
		public string capacidade { get; set; }
		public string meta_title { get; set; }
		public string meta_keyword { get; set; }
		public string meta_description { get; set; }
		public string cjm_imageswitcher { get; set; }
		public string cjm_moreviews { get; set; }
		public string cjm_useimages { get; set; }
		public string is_recurring { get; set; }
		public string recurring_profile { get; set; }
		public string custom_design { get; set; }
		public string custom_design_from { get; set; }
		public string custom_design_to { get; set; }
		public string custom_layout_update { get; set; }
		public string page_layout { get; set; }
		public string options_container { get; set; }
		public string gift_message_available { get; set; }
		public string package_height { get; set; }
		public string package_width { get; set; }
		public string package_length { get; set; }
		public string integra_anymarket { get; set; }
		public string id_anymarket { get; set; }
		public string garantia { get; set; }
		public string tempo_garantia { get; set; }
		public string categoria_anymarket { get; set; }
		public string origem { get; set; }
		public string modelo { get; set; }
		public string nbm { get; set; }
		public string intelipost_altura { get; set; }
		public string intelipost_largura { get; set; }
		public string intelipost_comprimento { get; set; }
		public string intelipost_peso { get; set; }
		public string intelipost_prazo_produto { get; set; }
		public List<string> UnknownFields { get; set; } = new List<string>();
	}
}
