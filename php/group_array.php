<?
    /**
     * Similar to an SQL Group By clause
     * For each group in $groups there is a list of keys
     * The keys list the fields which are unchanging, these are aggregated together and form the parent
     * The child (group) is then inserted with its given key under this aggregated parent
     * This allows a hierarchical array to be built up given linear data (usually returned by SQL)
     *
     * @return Grouped array
     * @author Guillermo Phillips
     **/
    /**
     * Similar to an SQL Group By clause
     * For each group in $groups there is a list of keys
     * The keys list the fields which are unchanging, these are aggregated together and form the parent
     * The child (group) is then inserted with its given key under this aggregated parent
     * This allows a hierarchical array to be built up given linear data (usually returned by SQL)
     *
     * @return Grouped array
     * @author Guillermo Phillips
     **/
    private function GroupArray($array, $groups) {
        $first_group = array_shift(array_values($groups));
        $group_key = array_shift(array_keys($groups));
        $groups_tail = array_slice($groups,1);
        
        $same_as_previous = false;
        
        $grouped_array = array();
        $sub_group = array();
        $previous_row = array();
		$sub_group = array();
		
		$array[] = array(); // force addition of last group
		$first_group_flag = true;
		
        foreach ($array as $row) {
	        if (count($row)>0) {
	            $same_as_previous = true;
	            foreach ($first_group as $field) {
	            	if (isset($row[$field])) {
		                if (!isset($previous_row[$field]) || $row[$field]!=$previous_row[$field]) {
		                    $same_as_previous = false;
		                    break;
		                }
	            	}
	            }
	        } else {
	        	$same_as_previous = false;
	        }

            if (!$same_as_previous) {
            	if (!$first_group_flag) {
            		if (count($groups_tail)==0) {
	            		$group_row[$group_key] = $sub_group;
            		} else {
            			$group_row[$group_key] = $this->GroupArray($sub_group, $groups_tail);
            		}
	            	$grouped_array[] = $group_row;
            	}
            	$first_group_flag = false;
            	$group_row = array();
                foreach ($first_group as $field) {
                    $group_row[$field] = $row[$field];
                }

                $sub_group = array();
            }

            $sub_group_row = array();
            foreach ($row as $key=>$datum) {
                if (!in_array($key, $first_group)) {
                    $sub_group_row[$key] = $datum;
                }
            }
            $sub_group[] = $sub_group_row;

            $previous_row = $row;
        }

        return $grouped_array;
    }

?>