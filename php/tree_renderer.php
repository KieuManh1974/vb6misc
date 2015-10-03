<?php
	// Convert a list of nodes into a tree of nodes (with one blank root node)
	// Each node in list should have both an id and parent_id
	// Each tree node has structure: array(sub_nodes,data)
	// Tree nodes with parent_id of NULL is attached to root node
	function GetTree($rows) {
		// Depth first search and attach node to parent_id given
		$add_node = function (&$tree,$data) use (&$add_node) {
			if ($data['parent_id']===null) {
				$tree[0][$data['id']] = array(array(),$data);
				return $tree;
			} else {
				foreach ($tree[0] as $sub_id=>&$sub_node) {
					if ($sub_id==$data['parent_id']) {
						$sub_node[0][$data['id']] = array(array(),$data);
						return $sub_node;
					} else {
						$node = $add_node($sub_node,$data);
						if ($node!==false) {
							return $node;
						}
					}
				}
				return false;
			}
		};

		$tree = array(array(),array());

		foreach ($rows as $row) {
			$add_node($tree,$row);
		}

		return $tree;
	}	


	// Convert a tree (see GetTree) into HTML
	// Each level in the tree is given distinct HTML
	// Data is inserted using tags e.g. <!--title--> or <!--id--> corresponding to keys in node data
	// A sub-level is indicated with <!---->
	// If a sub-level may or may not have children then use e.g. <!--haschild-->xyz<!-haschild-->. This makes the HTML conditional on having children
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