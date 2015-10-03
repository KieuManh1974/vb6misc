<?php

	class database_tree_api {
		private $mysqli;  
		private $prepared_add;
		private $prepared_delete;

		function __construct ($mysqli) {
			$this->mysqli = $mysqli;
			$this->prepared_add = $this->mysqli->prepare("INSERT INTO tree (parent_id, foreign_id, foreign_table, child_table) VALUES (?,?,?,?)");
			$this->prepared_delete = $this->mysqli->prepare("DELETE FROM tree WHERE id = ?");
		} 

		function AddNode($parent, $foreign_id = NULL, $foreign_table = NULL, $child_table = NULL) {
			$this->prepared_add->bind_param("iiss", $parent, $foreign_id, $foreign_table, $child_table);
			$this->prepared_add->execute();
		}

		function RemoveNode($id) {
			$this->prepared_delete->bind_param("i", $id);
			$this->prepared_delete->execute();
		}

		
	}

	$mysqli = new mysqli("localhost","root","","test_tree");
	$test = new database_tree_api($mysqli);
	$test->AddNode(123, 100, 'users');
	$test->RemoveNode(2);

/*
	$result = $mysqli->query("SELECT mpn FROM article LIMIT $limit OFFSET $offset");

	for ($row=0; $row<$result->num_rows; $row++) {
		$result->data_seek($row);
		$row_data = $result->fetch_assoc();
		$mpn = $row_data['mpn'];

	}
*/

	$mysqli->close();



?>