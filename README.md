Excel-2-Elasticsearch
====

Small & quick perl script to inject records from MS Excel (.xlsx) directly into Elasticsearch. Some inbuilt automation is done to directly map field types and index action using field names. For e.g. Field with name like **Author_NS** implies ~ Field name => **Author**, **N** => Not_analyzed index and **S** => String data type.

***WARNING:***

	To get the best results please follow instructions as mentioned in Formatting Excel Data. 


***USAGE:***

	$ xl2es.pl [Options] -x <ExcelFilename.xlsx>

	Elasticsearch

	   -i | --index <index name>   		Index name (default: xl2es)
	   -t | --type <data type>     		Type name (default: xldata)
	   -s | --es_server_port <host|IP:Port> (default: localhost:9200)

	Excel File (Ref. README for fields header requirements)

	   -x | --xl_filename           	Excel file name (required)							

	Help

	   -h | --help           		This help message
	   -v | --verbose          		Verbose while parsing (defaut: off)
 


**Formatting Excel Data**
--

* Ensure the first row in the Excel file has field names and first worksheet is the data worksheet
* To automate field mapping and data type, append each field name with an underscore "_" followed by one character each for Index Analysis and Data type. For e.g. a string field with name Author could be named as Author_NS
* Index analysis character could be **N** => Not Analyzed and **A** => Analyzed
* Data type character could be **I** => Integer, **D** => Date, **S** => String, **B** => Double 
* If not provided default field mapping is Not_Analyzed and String (_NS)
* For date fields choose custom format "dd-mmm-yyyy hh:mm:ss"

**License**
--
Copyright (C) 2014 Chetan Ganatra - Chetan.Ganatra~at~gmail.com

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details. <http://www.gnu.org/licenses>



