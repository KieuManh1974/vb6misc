<?php


	class Counter {

		private $start_ranges = array();
		private $end_ranges = array();
		private $step_sizes = array();

		public $index=0;
		private $final_index;
		public $finished=false;

		public $values = array();
		private $total_counters = 0;

		function __construct($ranges) {

			$final_index=0;
			foreach ($ranges as $range) {			
				$this->start_ranges[]=$range[0];
				$this->end_ranges[]=$range[1];
				if ($this->total_counters > 0) {
					$this->step_sizes[]=$range[2]-($this->end_ranges[$this->total_counters-1]-$this->start_ranges[$this->total_counters-1])*($this->step_sizes[$this->total_counters-1])-($this->step_sizes[$this->total_counters-1]);
				} else {
					$this->step_sizes[]=$range[2];
				}
				$this->values[]=$range[0];
				$this->final_index+=$range[2]*$range[1];
				$this->total_counters++;
			}
$this->final_index--;

echo $this->final_index;
		}

		public function CountUp(){
			for ($column=0; $column<$this->total_counters; $column++) {
				$this->values[$column]++;
				$this->index+=$this->step_sizes[$column];
				if ($this->values[$column]<=$this->end_ranges[$column]) {
					break;
				}
				$this->values[$column]=$this->start_ranges[$column];
			}

			if ($column==$this->total_counters) {
				$this->finished = true;
			} else {
				$this->finished = false;
			}

		}

		public function CountDown(){
			for ($column=0; $column<$this->total_counters; $column++) {
				$this->values[$column]--;
				$this->index-=$this->step_sizes[$column];
				if ($this->values[$column]>=$this->start_ranges[$column]) {
					break;
				}
				$this->values[$column]=$this->end_ranges[$column];
			}

			if ($column==$this->total_counters) {
				$this->finished = true;
			} else {
				$this->finished = false;
			}

		}

		public function Reset() {
			$this->values = $this->start_ranges;
			$this->index = 0;
		}
		public function ResetEnd() {
			$this->values = $this->end_ranges;
			$this->index = $this->final_index;
		}

	}

$pretty = function($v='',$c="&nbsp;&nbsp;&nbsp;&nbsp;",$in=-1,$k=null)use(&$pretty){$r='';if(in_array(gettype($v),array('object','array'))){$r.=($in!=-1?str_repeat($c,$in):'').(is_null($k)?'':"$k: ").'<br>';foreach($v as $sk=>$vl){$r.=$pretty($vl,$c,$in+1,$sk).'<br>';}}else{$r.=($in!=-1?str_repeat($c,$in):'').(is_null($k)?'':"$k: ").(is_null($v)?'&lt;NULL&gt;':"<strong>$v</strong>");}return$r;};




	$c = new Counter(array(array(1,5,1),array(0,3,10)));
	$c->ResetEnd();

do {
	echo $pretty($c->index); echo '<br>';
	$c->CountDown();
} while (!$c->finished);


?>