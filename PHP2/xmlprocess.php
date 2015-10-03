<?PHP
	class Records{
		var $parser;
		var $rows = array();
		var $record;
		var $data_tag;
		
		function Records ($xml) {			
			$this->parser = xml_parser_create();
			xml_set_object($this->parser, &$this);
			xml_parser_set_option($this->parser, XML_OPTION_CASE_FOLDING, true);
			xml_set_element_handler($this->parser, 'start_tag', 'end_tag');
			xml_set_character_data_handler($this->parser, 'character_data');		
			
			xml_parse($this->parser, $xml);
			xml_parser_free($this->parser);
		}
		
		function start_tag($p, $tag, &$attributes) {
			//echo "&lt;$tag&gt;";
			switch ($tag) {
				case 'RESULTINFO':
					break;
				case 'ROW':		
					$this->record = array();
					break;
				default:
					$this->data_tag = $tag;
					$this->record[$tag]='';
					break;
			}
		}
		
		function end_tag($p, $tag) {
			//echo "&lt;/$tag&gt;<BR />";
			switch ($tag) {
				case 'RESULTINFO':
					$this->data_tag = '';
					break;
				case 'ROW':
					$this->rows[] = $this->record;
					$this->data_tag = '';
					break;
				default:
					$this->data_tag = '';
					break;
			}
		}
		
		function character_data ($p, $text) {
			if ($this->data_tag != '') {
				$this->record[$this->data_tag] = $this->record[$this->data_tag].$text;
			}
		}		
		
		function show_blank ($text) {
			return ($text==''?'&nbsp;':$text);
		}
	}
?>