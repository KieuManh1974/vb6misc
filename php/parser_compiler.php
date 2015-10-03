<?php
	include 'parser_classes.php';

	class Parser {
		private $rules_parser;
		private $stream;
		public $rules = array();
		
		function __construct() {
			$and = new ParseText(false,false);$and->def(false,"and");
			$or = new ParseText(false,false);$or->def(false,"or");
			$set = new ParseText(false,false);$set->def(false,"set");
			$not = new ParseText(false,false);$not->def(false,"not");
			$opt = new ParseText(false,false);$opt->def(false,"opt");
			$any = new ParseText(false,false);$any->def(false,"any");
			$eos = new ParseText(false,false);$eos->def(false,"eos");
			$list = new ParseText(false,false);$list->def(false,"list");
			$min = new ParseText(false,false);$min->def(false,"min");
			$max = new ParseText(false,false);$max->def(false,"max");
			$del = new ParseText(false,false);$del->def(false,"del");
			$until = new ParseText(false,false);$until->def(false,"until");
			$non = new ParseText(false,false);$non->def(false,"non");
			$omit = new ParseText(false,false);$omit->def(false,"omit");
			$case = new ParseText(false,false);$case->def(false,"case");
			$end = new ParseText(false,false);$end->def(false,"end");
			$pipe = new ParseText(false,false);$pipe->def(false,"|");
			$pipe_omit = new ParseText(true,false);$pipe_omit->def(false,"|");
			$double_pipe = new ParseAnd(false,false);$double_pipe->def($pipe,$pipe_omit);
			$double_pipe_omit = new ParseAnd(true,false);$double_pipe_omit->def($pipe,$pipe);
			$space = new ParseSet(false,false);$space->def(false," ");
			$ws = new ParseList(true,false);$ws->def($space,null,null,0);
			$anychar = new ParseAny(false,false);
			
			$digit = new ParseSet(false,false);$digit->def(false,"0123456789");

			$literal_char = new ParseOr(false,false);$literal_char->def($double_pipe,$anychar);
			$not_pipe = new ParseNot(false,false);$not_pipe->def($pipe);
			$single_pipe = new ParseAnd(true,false);$single_pipe->def($pipe,$not_pipe);
			
			$not_identifier_char = new ParseOr(true,true);$not_identifier_char->def($space,$single_pipe);
			$identifier = new ParseList(false,false);$identifier->def($literal_char,null,$not_identifier_char); 	
			
			$literal_text = new ParseList(false,false);$literal_text->def($literal_char,null,$single_pipe,0);
			$literal = new ParseAnd(false,false);$literal->def($double_pipe_omit,$literal_text);
			
			$case_opt = new ParseOptional(false,false);$case_opt->def($case);
			$omit_opt = new ParseOptional(false,false);$omit_opt->def($omit);
			$non_opt = new ParseOptional(false,false);$non_opt->def($non);
			
			$omit_non = new ParseAnd(false,false);$omit_non->def($omit_opt,$ws,$non_opt,$ws);
			
			$text = new ParseOr(false,false);$text->def($literal,$identifier);
			$text_case = new ParseAnd(false,false);$text_case->def($case_opt,$ws,$text);
			
			$and_ommited_expression = new ParseAnd(false,false);
			$and_expression = new ParseAnd(false,false);
			$or_expression = new ParseAnd(false,false);
			$not_expression = new ParseAnd(false,false);
			$set_expression = new ParseAnd(false,false);
			$opt_expression = new ParseAnd(false,false);
			$list_expression = new ParseAnd(false,false);
			
			$expression = new ParseOr(false,false);$expression->def($any,$eos,$set_expression,$not_expression,$opt_expression,$and_expression,$or_expression,$list_expression,$text_case);
			$full_expression = new ParseAnd(false,false);$full_expression->def($omit_non,$expression);
			$expression_list = new ParseList(false,false);$expression_list->def($full_expression,$ws);
			
			$and_ommited_expression->def($expression_list,$ws,$pipe);
			$and_expression->def($and,$ws,$expression_list,$ws,$pipe);
			$or_expression->def($or,$ws,$expression_list,$ws,$pipe);
			$not_expression->def($not,$ws,$expression,$ws,$pipe);
			$set_expression->def($set,$ws,$text_case,$ws,$pipe);
			$opt_expression->def($opt,$ws,$expression,$ws,$pipe);
			
			$number = new ParseList(false,false);$number->def($digit);
			
			$delimit_clause = new ParseAnd(false,false);$delimit_clause->def($ws,$del,$ws,$full_expression);
			$until_clause = new ParseAnd(false,false);$until_clause->def($ws,$until,$ws,$full_expression);
			$min_clause = new ParseAnd(false,false);$min_clause->def($ws,$min,$ws,$number);
			$max_clause = new ParseAnd(false,false);$max_clause->def($ws,$max,$ws,$number);
			
			$delimit_opt = new ParseOptional(false,false);$delimit_opt->def($delimit_clause);
			$until_opt = new ParseOptional(false,false);$until_opt->def($until_clause);
			$min_opt = new ParseOptional(false,false);$min_opt->def($min_clause);
			$max_opt = new ParseOptional(false,false);$max_opt->def($max_clause);
			
			$list_expression->def($list,$ws,$full_expression,$delimit_opt,$until_opt,$min_opt,$max_opt,$ws,$pipe);

			$end = new ParseEOS(false,false);
			$rule = new ParseAnd(false,false);$rule->def($identifier,$ws,$full_expression,$ws,$pipe,$ws);
			
			$this->rules_parser = new ParseList(false,false);$this->rules_parser->def($rule,null);
			
			
		}

		function CreateParser($definition) {
			$this->stream = new Stream($definition);

			$this->rules = array();
			
			$result = $this->rules_parser->Parse($this->stream);

			echo $result->ok?'':'no';
			
			$result->ok or exit;
			echo $result->text($this->stream); echo "*<br><br>";

			return $this->CompileDefinition($result);

		}

		private function CompileDefinition($tree) {
			foreach ($tree->sub_results as $tree_rule) {
				$this->InitialiseRule($tree_rule);
			}
			
			foreach ($tree->sub_results as $tree_rule) {
				$this->CompileRule($tree_rule);
			}
		}
		
		private function InitialiseRule($tree) {
			$rules = array();
			
			$omit = $tree->sub_results[1]->sub_results[0]->sub_results[0]->index==1;
			$non = $tree->sub_results[1]->sub_results[0]->sub_results[1]->index==1;
			$rule_name = $tree->sub_results[0]->text($this->stream);
			$rule_index = $tree->sub_results[1]->sub_results[1]->index;

			switch ($rule_index) {
				case 1: // any
					$object = new ParseAny($omit,$non);
					break;
				case 2: // eos
					$object = new ParseEOS($omit,$non);
					break;					
				case 3: // set
					$object = new ParseSet($omit,$non);
					break;
				case 4: // not
					$object = new ParseNot($omit,$non);
					break;
				case 5: // opt
					$object = new ParseOptional($omit,$non);
					break;
				case 6: // and
					$object = new ParseAnd($omit,$non);
					break;
				case 7: // or
					$object = new ParseOr($omit,$non);
					break;
				case 8: // list 
					$object = new ParseList($omit,$non);
					break;
				case 9: // literal
					$object = new ParseText($omit,$non);
					break;
			}
			
			$this->rules[$rule_name]=$object;
		}

		private function CompileRule($tree) {
			$rule_name = $tree->sub_results[0]->text($this->stream);
			$this->CompileExpression($tree->sub_results[1], $rule_name);
		}
		
		private function CompileExpression($tree, $rule_name='') {
			$omit = $tree->sub_results[0]->sub_results[0]->index==1;
			$non = $tree->sub_results[0]->sub_results[1]->index==1;
						
			$expression_type = $tree->sub_results[1]->index;
			
			if ($rule_name==='') {
				switch ($expression_type) {
					case 1: // any
						$object = new ParseAny($omit,$non);
						break;
					case 2: // eos
						$object = new ParseSet($omit,$non);
						break;
					case 3: // set
						$object = new ParseSet($omit,$non);
						break;
					case 4: // not
						$object = new ParseNot($omit,$non);
						break;
					case 5: // opt
						$object = new ParseOptional($omit,$non);
						break;
					case 6: // and
						$object = new ParseAnd($omit,$non);
						break;
					case 7: // or
						$object = new ParseOr($omit,$non);
						break;
					case 8: // list 
						$object = new ParseList($omit,$non);
						break;
					case 9: // literal
						$object = new ParseText($omit,$non);
						break;
				}	
			} else {
				$object = $this->rules[$rule_name];
			}

			switch ($expression_type) {
				case 1: // any
					break;
				case 2: // eos
					break;					
				case 3: // set
					$case = $tree->sub_results[1]->sub_results[0]->sub_results[1]->sub_results[0]->index==1;
					$object->def($case,$tree->sub_results[1]->sub_results[0]->sub_results[1]->sub_results[1]->full_text($this->stream));
					break;
				case 4: // not
				case 5: // opt	
					$condition = $this->CompileExpression($tree->sub_results[1]->sub_results[0]);
					$object->def($condition);
					break;
				case 6: // and 
				case 7:	// or				
					$objects = array();
					foreach ($tree->sub_results[1]->sub_results[0]->sub_results[1]->sub_results as $sub_result) {
						$objects[] = $this->CompileExpression($sub_result);
					}
					$object->set = $objects;
					$object->length = count($objects);
					break;
				case 8: // list
					$object->condition = $this->CompileExpression($tree->sub_results[1]->sub_results[0]->sub_results[1]);
					$object->delimiter = $tree->sub_results[1]->sub_results[0]->sub_results[2]->index==1?$this->CompileExpression($tree->sub_results[1]->sub_results[0]->sub_results[2]->sub_results[0]->sub_results[1]):null;
					$object->terminator = $tree->sub_results[1]->sub_results[0]->sub_results[3]->index==1?$this->CompileExpression($tree->sub_results[1]->sub_results[0]->sub_results[3]->sub_results[0]->sub_results[1]):null;
					$object->min = $tree->sub_results[1]->sub_results[0]->sub_results[4]->index==1?$tree->sub_results[1]->sub_results[0]->sub_results[4]->sub_results[0]->sub_results[1]->text($this->stream):1;
					$object->max = $tree->sub_results[1]->sub_results[0]->sub_results[5]->index==1?$tree->sub_results[1]->sub_results[0]->sub_results[5]->sub_results[0]->sub_results[1]->text($this->stream):-1;
					break;
				case 9: // literal
					$case = $tree->sub_results[1]->sub_results[0]->sub_results[0]->index==1;
					$text = $tree->sub_results[1]->sub_results[0]->sub_results[1]->sub_results[0]->full_text($this->stream);
					switch ($tree->sub_results[1]->sub_results[0]->sub_results[1]->index) {
						case 1: // escaped
							$object->def($case,$text);
							break; 
						case 2: // non escaped
							if (array_key_exists($text, $this->rules)) {
								$object = unserialize(serialize($this->rules[$text]));
								$object->__construct($omit,$non);
							} else {
								$object->def($case,$text);
							}
							break;
					}
					break;
			}
			
			return $object;
		}
		
		function Parse($rule, $stream) {
			$stream->position=0;
			return $this->rules[$rule]->Parse($stream);
		}

	}


	$x = new Parser();
	$x->CreateParser("z case cat | |");
	
	$file= new Stream("CATabcdefghijklmn");
	$result = $x->rules["z"]->Parse($file);
	
	echo $result->full_text($file);
	
?>