<?php

	class Stream {
		var $text='';
		var $position = 0;
		var $end = 0;
		
		function __construct($text) {
			$this->text = $text;
			$this->end = strlen($text);
		}
		
		public function Move($step) {
			$this->position += $step;
		}
		
		public function AtEnd($size=1) {
			return ($this->position+$size-1)>=$this->end;
		}
	}

	class Result {
		public $start = -1;
		public $end = -1;
		public $index = -1;
		public $omit = false;
		public $sub_results = array();
		public $leaf = false;
		public $ok = false;
		
		function __construct($start=-1,$end=-1,$index=-1) {
			$this->start = $start;
			$this->end = $end;
			$this->index = $index;
		}
		
		function text(Stream &$stream) {
			return substr($stream->text,$this->start,$this->end-$this->start+1);
		}
		
		function full_text(Stream &$stream) {
			if (count($this->sub_results)>0) {
				$full_text = '';
				foreach ($this->sub_results as &$sub_result) {
					$full_text .= $sub_result->full_text($stream);
				}
				return $full_text;
			} else {
				if ($this->leaf) {
					return substr($stream->text,$this->start,$this->end-$this->start+1);
				} else {
					return '';
				}
			}
		}
	}
	
	class ParseText {
		private $compare_text = '';
		private $length;
		public $omit = false;
		private $non_consuming = false;
		private $case_insensitive = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;
		}
		
		function def($case_insensitive, $text) {
			$this->case_insensitive = $case_insensitive;
			if ($case_insensitive) {
				$this->compare_text = strtoupper($text);
			} else {
				$this->compare_text = $text;
			}
			$this->length = strlen($text);			
		}
		
		public function Parse(Stream &$stream) {
			$result = new Result($stream->position);
			$result->omit = $this->omit;
			$result->leaf = true;
			if ($stream->AtEnd($this->length)) {
			} elseif (!$this->case_insensitive && (substr($stream->text, $stream->position, strlen($this->compare_text))==$this->compare_text)) {
				$result->end = $stream->position+$this->length-1;
				$result->ok = true;
				if (!$this->non_consuming) {
					$stream->Move($this->length);
				}
			} elseif ($this->case_insensitive && (strtoupper(substr($stream->text, $stream->position, strlen($this->compare_text)))==$this->compare_text)) {
				$result->end = $stream->position+$this->length-1;
				$result->ok = true;
				if (!$this->non_consuming) {
					$stream->Move($this->length);
				}				
			}
			
			return $result;
		}
	}
	
	class ParseSet {
		private $compare_set = '';
		private $length = 0;
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;

		}

		function def($case_insensitive, $text_set) {
			$this->case_insensitive = $case_insensitive;
			if ($case_insensitive) {
				$this->compare_set = strtoupper($text_set);
			} else {
				$this->compare_set = $text_set;
			}					
			$this->length = strlen($text_set);				
		}
		
		public function Parse(Stream &$stream) {
			$result = new Result($stream->position);
			$result->omit = $this->omit;
			$result->leaf = true;
			
			if ($stream->AtEnd(1)) {
				return $result;
			}
			
			if ($this->case_insensitive) {
				$stream_char = strtoupper(substr($stream->text,$stream->position,1));
			} else {
				$stream_char = substr($stream->text,$stream->position,1);	
			}
			for ($index=0; $index<$this->length; $index++) {
				$char = substr($this->compare_set,$index,1);
				
				if ($char==$stream_char) {
					$result->end = $stream->position;
					$result->index = $index+1;
					$result->ok = true;
					
					if (!$this->non_consuming) {
						$stream->Move(1);
					}					
					return $result;
				}
			}
			return $result;
		}
	}
	
	class ParseAny {
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;
		}
		
		public function Parse(Stream &$stream) {
			$result = new Result($stream->position);
			$result->omit = $this->omit;
			$result->leaf = true;
			
			if (!$stream->AtEnd()) {
				$result->end = $stream->position;
				if (!$this->non_consuming) {
					$stream->Move(1);
				}
				$result->ok = true;
			}
			return $result;
		}
	}	
	
	class ParseEOS {
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;
		}
		
		public function Parse(Stream &$stream) {
			$result = new Result($stream->position);
			$result->omit = $this->omit;
			$result->ok = $stream->AtEnd();
			return $result;
		}
	}
	
	class ParseNot {
		private $omit = false;
		private $non_consuming = false;
		private $condition;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;		
		}

		function def($condition) {
			$this->condition = $condition;
		}
		
		public function Parse(Stream &$stream) {
			$result = new Result($stream->position);
			$result->omit = $this->omit;
			$sub_result = $this->condition->Parse($stream);
			if ($sub_result->ok) {
				$result->ok = false;
				if ($this->non_consuming) {
					$stream->position = $start;
				}				
			} else {
				$result->ok = true;
			}
			return $result;
		}
	}
		
	class ParseAnd {
		public $set = array();
		public $length = 0;
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;			
			for ($arg=2; $arg<func_num_args(); $arg++) {
				$this->set[]=func_get_arg($arg);
			}
			$this->length = func_num_args()-2;
		}
		
		function def() {
			$this->set = array();
			for ($arg=0; $arg<func_num_args(); $arg++) {
				$this->set[]=func_get_arg($arg);
			}
			$this->length = func_num_args();			
		}
		
		public function Parse(Stream &$stream) {
			$start = $stream->position;
			$result = new Result($start);
			$result->omit = $this->omit;
			foreach ($this->set as $object) {
				$sub_result = $object->Parse($stream);
				if ($sub_result->ok===false) {
					$stream->position = $start;
					return $result;
				} else {
					if (!$sub_result->omit) {
						$result->sub_results[] = $sub_result;
					}
				}
			}
			$result->end = $stream->position-1;
			$result->ok = true;

			if ($this->non_consuming) {
				$stream->position = $start;
			}			
			return $result;
		}
	}
	
	class ParseOr {
		public $set = array();
		public $length = 0;
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;			
			for ($arg=2; $arg<func_num_args(); $arg++) {
				$this->set[]=func_get_arg($arg);
			}
			$this->length = func_num_args();
		}

		function def() {
			$this->set = array();
			for ($arg=0; $arg<func_num_args(); $arg++) {
				$this->set[]=func_get_arg($arg);
			}
			$this->length = func_num_args();			
		}
		
		public function Parse(Stream &$stream) {
			$start = $stream->position;
			$result = new Result($start);
			$result->omit = $this->omit;
			$index = 1;
			foreach ($this->set as $object) {
				$sub_result = $object->Parse($stream);
				if ($sub_result->ok) {
					if (!$sub_result->omit) {
						$result->sub_results[] = $sub_result;
					}
					$result->end = $sub_result->end;
					$result->index = $index;
					$result->ok = true;

					if ($this->non_consuming) {
						$stream->position = $start;
					}
					return $result;
				}
				$index++;
			}
			$stream->position = $start;
			return $result;
		}
	}
	
	class ParseOptional {
		private $condition;
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;
		}

		function def($condition) {
			$this->condition = $condition;	
		}
		
		public function Parse(Stream &$stream) {
			$start = $stream->position;
		
			$result = new Result($stream->position);
			$result->omit = $this->omit;

			$sub_result=$this->condition->Parse($stream);
			if ($sub_result->ok) {
				if (!$sub_result->omit) {
					$result->sub_results[] = $sub_result;
				}
				$result->index=1;
				$result->end = $stream->position-1;
			} else {
				$result->index=0;
			}

			$result->ok = true;

			if ($this->non_consuming) {
				$stream->position = $start;
			}			
			return $result;
		}
	}
	
	class ParseWrapper {
		private $object;
		private $omit = false;
		private $non_consuming = false;

		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;
		}

		function def($object) {
			$this->object = $object;	
		}
		
		public function Parse(Stream &$stream) {
			$start = $stream->position;
		
			$result=$this->object->Parse($stream);
			$result->omit = $result->omit || $this->omit;
			
			if ($this->non_consuming) {
				$stream->position = $start;
			}
			return $result;
		}
	}
	
	class ParseList {
		public $condition;
		public $delimiter;
		public $terminator;
		public $min;
		public $max;
		private $omit = false;
		private $non_consuming = false;
		
		function __construct($omit, $non_consuming) {
			$this->omit = $omit;
			$this->non_consuming = $non_consuming;			

		}

		function def($condition, $delimiter=null, $terminator=null, $min=1, $max=-1) {
			$this->condition = $condition;
			$this->delimiter = $delimiter;
			$this->terminator = $terminator;
			$this->min = $min;
			$this->max = $max;				
		}
		
		public function Parse(Stream &$stream) {
			$start = $stream->position;
			$count = 0;
			$terminator_found = false;

			$result = new Result($stream->position);
			$result->omit = $this->omit;
			
			while (true) {
				if ($this->terminator!==null) {
					$sub_result=$this->terminator->Parse($stream);
					if ($sub_result->ok) {
						if (!$sub_result->omit) {
							$result->sub_results[] = $sub_result;
						}
						$terminator_found = true;
						break;
					}
				}
				if ($this->max!=-1 && $count==$this->max) {
					break;
				}			

				if ($this->delimiter===null || $count===0) {
					$sub_result=$this->condition->Parse($stream);
					if ($sub_result->ok) {
						if (!$sub_result->omit) {
							$result->sub_results[] = $sub_result;
						}
					} else {
						break;
					}
				} else {
					$start_delimiter = $stream->position;
					$delimiter_result=$this->delimiter->Parse($stream);
					if (!$delimiter_result->ok) {
						break;
					}		

					$sub_result=$this->condition->Parse($stream);
					if ($sub_result->ok) {
						if (!$delimiter_result->omit) {
							$result->sub_results[] = $delimiter_result;
						}
						if (!$sub_result->omit) {
							$result->sub_results[] = $sub_result;
						}
					} else {
						$stream->position = $start_delimiter;
						break;
					}
						
				}
				
				$count++;
			}
			if ($this->terminator!==null && !$terminator_found) {
				$stream->position = $start;
				$result->ok = false;
				return $result;		
			}
			
			if ($count>=$this->min) {
				$result->end = $stream->position-1;
				$result->index = $count;
				$result->ok = true;
			} else {
				$stream->position = $start;
				$result->ok = false;
			}
			
			if ($this->non_consuming) {
				$stream->position = $start;
			}			
			return $result;
	
		}
	}

?>