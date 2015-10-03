<?php
	function GetTaxonomyData($language_code) {
		$mysqli = new mysqli("localhost","root","","croogo14");
		
		$languages = array("chn","deu","jpn");
		$term_title = in_array($language_code, $languages)?"title_$language_code":"title";
	
		
		$terms_results = $mysqli->query("SELECT id, $term_title, slug FROM terms");
		$taxonomies_results = $mysqli->query("SELECT id, vocabulary_id, parent_id, term_id FROM taxonomies");
		$vocabularies_results = $mysqli->query("SELECT id, title FROM vocabularies");
	
		$terms=array();
		for ($row=0; $row<$terms_results->num_rows; $row++) {
			$terms_results->data_seek($row);
			$row_data = $terms_results->fetch_assoc();
			
			$id = $row_data['id'];
			$title = $row_data[$term_title];
			$slug = $row_data['slug'];
			
			$terms[$id] = array('id'=>$id,'title'=>$title,'slug'=>$slug);
		}	
		
		$taxonomies = array();
		for ($row=0; $row<$taxonomies_results->num_rows; $row++) {
			$taxonomies_results->data_seek($row);
			$row_data = $taxonomies_results->fetch_assoc();
			
			$id = $row_data['id'];
			$term_id = $row_data['term_id'];
			$vocabulary_id = $row_data['vocabulary_id'];
			$parent_id = $row_data['parent_id'];
			
			$taxonomies[$id] = array('id'=>$id,'parent_id'=>$parent_id,'term_id'=>$term_id,'vocabulary_id'=>$vocabulary_id);
		}
	
		$vocabularies = array();
		for ($row=0; $row<$vocabularies_results->num_rows; $row++) {
			$vocabularies_results->data_seek($row);
			$row_data = $vocabularies_results->fetch_assoc();
			
			$id = $row_data['id'];
			$title = $row_data['title'];
		
			$vocabularies[$id] = array('id'=>$id,'title'=>$title,'slug'=>"");
		}	
		
		return array('terms'=>$terms, 'taxonomies'=>$taxonomies, 'vocabularies'=>$vocabularies);
	}	
	

	
	function GetTaxonomyTree($terms, $taxonomies, $vocabularies) {
		$add_node =
		function ($parent_id,&$tree,$id,$data) use (&$add_node) {
			foreach ($tree[0] as $sub_id=>&$sub_node) {
				if ($sub_id==$parent_id) {
					$sub_node[0][$id] = array(array(),$data);
					return $sub_node;
				} else {
					$node = $add_node($parent_id,$sub_node,$id,$data);
					if ($node!==false) {
						return $node;
					}
				}
			}
			return false;
		};
				
		$tree = array(0=>array(),array());
		
		foreach ($vocabularies as $vocabulary) {
			$id = $vocabulary['id'];
			$title = $vocabulary['title'];
	
			$tree[0][$id+10000] = array(array(),array('id'=>$id,'title'=>$title,'slug'=>''));
		}
	
		$taxonomy_tree = array();
	
		foreach ($taxonomies as $taxonomy) {
			$id = $taxonomy['id'];
			$term_id = $taxonomy['term_id'];
			$vocabulary_id = $taxonomy['vocabulary_id'];
			$parent_id = $taxonomy['parent_id']===null?$vocabulary_id+10000:$taxonomy['parent_id'];
			
			$node = $add_node($parent_id,$tree,$id,$terms[$term_id]);
		}

		return $tree;
	}
	
	
	function RenderTreeHTML($tree,$levels_html,$level=0) {
		$output_text = "";
		$html = $levels_html[$level];
		foreach ($tree[1] as $key=>$value) {
			$key_text = "<!--$key-->";
			$html = str_replace($key_text,$value,$html);
		}
		$split = explode('<!--haschild-->',$html);
		if (count($tree[0])>0) {
			$html= implode('',$split);
		} else {
			$join='';
			$ok=true;
			foreach ($split as $part) {
				if ($ok) {
					$join.=$part;
				}
				$ok = !$ok;	
			}
			$html = $join;
		}
		
		$split = explode('<!---->',$html);
		$head = $split[0];
		if (count($split)>1) {
			$foot = $split[1];
		} else {
			$foot='';
		}
		$output_text.=$head;
		foreach ($tree[0] as $branch) {
			$output_text.=RenderTreeHTML($branch,$levels_html,$level+1);
		}
		$output_text.=$foot;
		
		return $output_text;
	}

?>