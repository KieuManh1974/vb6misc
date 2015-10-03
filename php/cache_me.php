<?php
	
	// Check the file with name $identifier exists and return contents if so, otherwise return FALSE

	function ReadCacheContent($identifier) {
		if (file_exists($identifier)) {
			$handle = fopen($identifier, 'r');
			$text = fread($handle, filesize($identifier));
			fclose($handle);

			return $text;
		} 
		return false;
	}

	// Write the $content to the file called $identifier
	// return the $content for convenience
	function WriteCacheContent($identifier, &$content) {
		$handle = fopen($identifier, 'w');
		fwrite($handle, $content);
		fclose($handle);

		return $content;
	}

?>