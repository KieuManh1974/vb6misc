<?php
	class LinkNode {
		public $char;
		public $links=array();
		public $value;

		public function __construct($char) {
			$this->char = $char;
		}

		public function AddLink($char) {
			foreach ($this->links as &$link) {
				if ($link->char==$char) {
					return $link;
				}
			}
			$new_link = new LinkNode($char);
			$this->links[] = $new_link;
			return $new_link;

		}

		public function RemoveLink($char) {
			foreach ($this->links as &$link) {
				if ($link->char==$char) {
					unset($link);
				}
			}			
		}

		public function FindLink($char) {
			foreach ($this->links as &$link) {
				if ($link->char==$char) {
					return $link;
				}
			}	
			return null;		
		}

	}


	class Dictionary {
		private $base_link;

		function __construct() {
			$this->base_link = new LinkNode('');
		}

		public function AddWord($word, $value=null) {
			$length = strlen($word);
			$current_link = $this->base_link;
			for ($index=0; $index<$length; $index++) {
				$current_link = $current_link->AddLink(substr($word,$index,1));
			}
			$current_link = $current_link->AddLink('');
			$current_link->value = $value;
		}

		public function RemoveWord($word) {

		}

		public function FindWord($word) {
			$length = strlen($word);
			$current_link = $this->base_link;
			for ($index=0; $index<=$length; $index++) {
				if ($index<$length) {
					$char = substr($word,$index,1);
				} else {
					$char = '';
				}

				$current_link = $current_link->FindLink($char);
				if ($current_link==null) {
					return null;
				}
			}
			return $current_link;
		}

		public function Serialize() {

		}

		public function Deserialize() {

		}

		public function x() {

		}
	}

/*
	$d = new Dictionary();

	$d->AddWord('camber',123);
	$d->AddWord('chamber',321);

	echo $d->FindWord('chamber');

*/
?> 