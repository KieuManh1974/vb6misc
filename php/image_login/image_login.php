<html>
	<head>
		<link rel="stylesheet" type="text/css" href="mystyle.css">
		<style type="text/css">
			body {
				background-color: black;
			}

			#choose {
				/*width:100%;*/
				/*height:50%;*/
			}

			#chosen {
				/*width:100%;*/
			}

			td {
				width:180px;
				height:150px;
				background-color: grey;
			}

			tr {
				height:100px;
			}

			.chosen {
				background-size:100%;
				background-repeat:no-repeat;	
				background-position:center; 			
				background-color: white;
				cursor:hand;				
				visibility: normal;
			}

			.unchosen {
				background-color: grey;
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
						$this->levels[0] .= tag('td',array('id'=>"image$index"),'');
						break;
					case 1:
						$this->levels[1] .= tag('tr',array(),$this->levels[0]);
						$this->levels[0] = '';
						break;
					case 2:
						$this->html = tag('table',array('id'=>$this->id,'cellspacing'=>'12'),$this->levels[1]);
						$this->levels[1] = '';
						break;
				}

			}
		}

		$width = 4;
		$height = 6;
		$choices = 4;

		$table = new build_table('choose',$width, $height, array());
		$chosen_table = new build_table('chosen',$choices,1,array());

	?>
	<body>		
		<div id='slideup'>
		<?= $table->html ?>
		</div>
		<?= $chosen_table->html ?>
		<form action="auth.php" method="POST">
			<input type="hidden" id="submit0" name="submit[0]" value="">
			<input type="hidden" id="submit1" name="submit[1]" value="">
			<input type="hidden" id="submit2" name="submit[2]" value="">
			<input type="hidden" id="submit3" name="submit[3]" value="">
		</form>

	
	</body>		
	<script>
		var added = 0;

		$(
			function() {
				$('#choose td').addClass("button");
				$('#choose td').addClass("chosen");

				<?php
					$images = array(0=>'dog',1=>'cat',2=>'rabbit',3=>'snake',4=>'orangutan',5=>'butterfly',6=>'cow',7=>'eagle',8=>'fish');
					$keys = array_keys($images);
					shuffle($keys);
					$shuffled_images = array();
					foreach ($keys as $key) {
						$shuffled_images[$key] = $images[$key];
					}
				?>	

				var images = [<?= "'".implode("','",$shuffled_images)."'"; ?>];
				var image_ids = [<?= "'".implode("','",array_keys($shuffled_images))."'"; ?>];;

				$('#choose>tbody>tr>td').each(
					function(index) {
						$(this).css("background-image", 'url('+images[index]+'.jpg'+')');
					}
				);


				$('.button').click(
					function () {
						if ($(this).hasClass('chosen')) {
							chosen_element = $(this);
							chosen_element_id = chosen_element.attr('id').substring(5);

							$(this).toggleClass('chosen');
							$(this).toggleClass('unchosen');
							$(this).css('background-image','none');

							$("#submit"+added).val(image_ids[chosen_element_id]);

							$target = $('#chosen td[id=image'+added+']');
							$target.addClass("chosen");
							$target.css('background-image', 'url('+images[chosen_element_id]+'.jpg)');
							added++;
							if (added==4) {
								$('#slideup').slideUp(400);
								document.forms[0].submit();
							}
						}

					}
				)

			}
		)

	</script>
</html>