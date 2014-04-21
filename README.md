Excel-2-Elasticsearch
====

Small and quick perl script to inject records from MS Excel (.xlsx as well as .xls) directly into Elasticsearch. Useful for doing quick demos by importing existing Excel data and charting graphs using Kibana. 

Some inbuilt automation is done to directly map field types and index action using field names. For e.g. Field with name like **Author_NS** implies ~ Field name => **Author**, **N** => Not_analyzed index and **S** => String data type. In case you find it difficult to map fields, just go ahead and try your existing Excel files.


***WARNING:***

	To get the best results please follow instructions as mentioned in Formatting Excel Data. 


***USAGE:***

	$ xl2es.pl [Options] -x <ExcelFilename.xlsx>

	Elasticsearch

	   -i | --index <index name>   			Index name (default: xl2es)
	   -t | --type <data type>     			Type name (default: xldata)
	   -s | --es_server_port <host|IP:Port> ES Host:Port (default: localhost:9200)

	Excel File  

	   -x | --xl_filename           		Excel file name (required)							
	   -a | --all_worksheets				Parse all worksheets (default: off)
   
	Help

	   -h | --help           				This help message
	   -v | --verbose          				Verbose while parsing (defaut: off)
 

**Installation / Dependencies**
--
* Perl v5.6 or later 
* Requires [Perl client for Elasticsearch] (https://metacpan.org/pod/Search::Elasticsearch)
* Perl packages for parsing Excel files
	- for Excel XLSX [Parsing .xlsx] (http://search.cpan.org/~doy/Spreadsheet-ParseXLSX-0.05/lib/Spreadsheet/ParseXLSX.pm)
	- for Excel 97 - 2003 [Parsing .xls] (http://search.cpan.org/~dougw/Spreadsheet-ParseExcel-0.65/lib/Spreadsheet/ParseExcel.pm)

	
**Formatting Excel Data**
--

* Ensure the first row in the Excel file has field names and first worksheet is the data worksheet
* Append each field name with an underscore "_" followed by one character each for Index Analysis and Data type. For e.g. a string field with name Author could be named as Author_NS. i.e. Field is a String and Not_Analyzed index.
* If not provided, default field mapping is Not_Analyzed and String (_NS)
* Index analysis character could be **N** => Not Analyzed and **A** => Analyzed
* Data type character could be **I** => Integer, **D** => Date, **S** => String, **B** => Double 
* For Date fields, choose custom cell format "dd-mmm-yyyy hh:mm:ss". In case you wish to use a different Date format in Excel, appropriate changes needs to be done in the perl code.

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


**Next Steps [TODO]**
--
On similar lines generate Kibana dashboards using Field name automation!


