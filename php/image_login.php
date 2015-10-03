<html>
	<head>
		<link rel="stylesheet" type="text/css" href="mystyle.css">
		<style type="text/css">
			body {
				background-color: black;
			}

			#choose {
				width:100%;
				/*height:50%;*/
			}

			td {
				height:100px;
			}

			tr {
				height:100px;
			}

			.chosen {
				background-color: cyan;
				cursor:hand;				
				visibility: normal;
			}

			.unchosen {
				background-color: grey;
			}

			#chosen {
				width:100%;
			}

		</style>
		<script src="jquery.js"></script>

	</head>
	<?php

		class Counter {
			public $index=0;
			private $columns = array();
			private $counter = array();
			private $total_columns = 0;
			public $finished = false;
			public $event = 'event';
			
			function __construct($columns) {
				$this->columns = $columns;
				foreach ($columns as $index=>$column) {
					$this->counter[$index]=0;
					$this->total_columns++;
				}
			}
			function reset() {
				foreach ($this->columns as $index=>$column) {
					$this->counter[$index]=0;
				}
				$this->finished = false;
			}

			function tick() {
				$this->finished = false;
				$column = -1;
							
				do {
					$column++;
					if ($column<$this->total_columns) {
						$this->counter[$column] = ($this->counter[$column]+1) % $this->columns[$column];
						call_user_func($this->event, $column, $this->index);
					}
				} while ($column<$this->total_columns && $this->counter[$column]==0);
				if ($column == $this->total_columns) {
					call_user_func($this->event, $column, $this->index);
					$this->index = -1;
					$this->finished = true;
				}
				$this->index++;
			}

			function tick_all() {
				$this->reset();
				while (!$this->finished) {
					$this->tick();
				}
			}
		}

		function tag($tag, $attributes, $content) {
			$html = '';
			$html .= "<$tag";
			foreach ($attributes as $attribute=>$value) {
				$html .= " $attribute='$value'";
			}
			$html .= '>';
			$html .= $content;
			$html .= "</$tag>";
			return $html;
		}
				
		class build_table {
			public $html = '';
			private $levels = array('','');
			private $data = array();
			private $id = '';

			function __construct($id, $width, $height, $data) {
				$this->data = $data;
				$this->id = $id;
				$counter = new Counter(array($width,$height));
				$counter->event = array($this, 'element');
				$counter->tick_all();			
			}

			function element($column, $index) {
				switch ($column) {
					case 0:
						$this->levels[0] .= tag('td',array('id'=>$index),$index);
						break;
					case 1:
						$this->levels[1] .= tag('tr',array(),$this->levels[0]);
						$this->levels[0] = '';
						break;
					case 2:
						$this->html = tag('table',array('id'=>$this->id),$this->levels[1]);
						$this->levels[1] = '';
						break;
				}

			}
		}

		$width = 3;
		$height = 3;
		$choices = 4;

		$table = new build_table('choose',$width, $height, array());
		$chosen_table = new build_table('chosen',$choices,1,array());

	?>
	<body>
		<?= $table->html ?>
		<?= $chosen_table->html ?>
	</body>		
	<script>
		var added = 0;

		$(
			function() {
				$('#choose>tbody>tr>td').addClass("button");
				$('#choose>tbody>tr>td').addClass("chosen");

				$('.button').click(
					function () {
						$(this).toggleClass('chosen');
						$(this).toggleClass('unchosen');

						if ($(this).hasClass('unchosen')) {
							alert($(this).val());
							$('#chosen td[id='+added+']').addClass("chosen");
							$('#chosen td[id='+added+']').text($(this).text());
							added++;
						}
					}
				)

			}
		)

	</script>
</html>