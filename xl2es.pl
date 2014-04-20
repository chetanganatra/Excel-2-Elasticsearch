#!/usr/bin/perl -w
# xl2es.pl - Version 1.0
# Small&Quick script to inject records from MS Excel into Elasticsearch
# Copyright (C) 2014 Chetan Ganatra - Chetan.Ganatra~at~gmail.com
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details. <http://www.gnu.org/licenses>
#

use strict;
use v5.10; 																			 
no warnings 'experimental::smartmatch';												# use Spreadsheet::ParseExcel;
use Time::Piece;
use Getopt::Long;
use Scalar::Util qw(looks_like_number);												
use feature qw{ switch };

use Spreadsheet::ParseXLSX;
use Elasticsearch;																	# ElasticSearch vs Elasticsearch matters! 
	
my $xl2es_version = "1.0";
	
my $id=1;
my $result="";

my $index = "xl2es";
my $type = "xldata";
my $es_server_port = "localhost:9200";
my $xl_filename = "";
my $verbose = 0;

GetOptions(
		'i|index=s' => \$index,
		't|type=s' => \$type,
		's|es_server_port=s' => \$es_server_port,
		'x|xl_filename=s' => \$xl_filename,
		'v|verbose' => \$verbose,
		'h|help'  => sub { usage() },
	);

sub usage {
print STDERR <<USAGE;
Excel-2-Elasticsearch version $xl2es_version - Data injection script.
Maintained by Chetan Ganatra (chetan.ganatra\@gmail.com) - Licensed under GPL.

Important Usage Guidelines:
    To run the script with the default options atleast Excel filename 
	needs to be provided.

Usage: $0 [Options] -x <ExcelFilename.xlsx>

Elasticsearch

   -i | --index <index name>   		Index name (default: xl2es)
   -t | --type <data type>     		Type name (default: xldata)
   -s | --es_server_port <host|IP:Port> (default: localhost:9200)

Excel File (Ref. README for fields header requirements)

   -x | --xl_filename           	Excel file name (required)							

Help

   -h | --help           		This help message
   -v | --verbose          		Verbose while parsing (defaut: off)
 
USAGE
    exit(2);
}

if ($xl_filename ne "") { chomp($xl_filename); } else { print STDERR "\nExcel file name not provided!\n\n"; usage(); }
	
print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " Parsing started...\n" if $verbose;
print "\nParsing with $index :: $type :: $es_server_port ::  $xl_filename" if $verbose;
	
my $e = Elasticsearch->new( nodes => "$es_server_port" );
my $parser   = Spreadsheet::ParseXLSX->new;
my $workbook = $parser->parse($xl_filename);

die $parser->error(), ".\n" if ( !defined $workbook );

print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " EL up and XL parsed!" if $verbose;
	
# Parse through worksheets -- currently interested only in the 1st!
for my $worksheet ( $workbook->worksheets() ) {

        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $col_min, $col_max ) = $worksheet->col_range();
																						# Capture the header line 1st Row
	    my @fields="";
		my %map_analysis=();
		my %map_type=();
		my %mapping=();
		for my $hdr ( $col_min .. $col_max ) {
			my $curr = $worksheet->get_cell( $row_min, $hdr );
			next unless $curr;
			my $fld = $curr->value();		
			print "\nProcessing field: $fld"  if $verbose;
			$fields[$hdr] = substr $fld, 0, -3;
			my $tmp_fldname=$fields[$hdr];
			unless (substr $tmp_fldname, 1, 1 eq "_")									# if field name has additional details 
			{
			$map_analysis{$tmp_fldname} = substr $fld, -2, 1;						# _ A|N . => (-2) Analyzed..or..not 
			$map_type{$tmp_fldname} = substr $fld, -1, 1;							# _ . S|D|I|B => (-1) Data type B=> double 
			my $tmp01=$map_type{$tmp_fldname};
			given($tmp01) {
				when("S") { $mapping{$tmp_fldname}{"type"} = "string"; }
				when("D") { 
							$mapping{$tmp_fldname}{"type"} = "date"; 
							$mapping{$tmp_fldname}{"format"} = "dd-MMM-YYYY HH:mm:ss";   
						}
				when("I") { $mapping{$tmp_fldname}{"type"} = "integer"; }
				when("B") { $mapping{$tmp_fldname}{"type"} = "double"; }
				default { $mapping{$tmp_fldname}{"type"} = "string"; }
				}
			given($map_analysis{$tmp_fldname}) {
				when("A") { $mapping{$tmp_fldname}{"index"}="analyzed"; }
				when("N") { $mapping{$tmp_fldname}{"index"}="not_analyzed"; }
				default { $mapping{$tmp_fldname}{"index"}="not_analyzed"; }
				}
			}
			else																	# else defaults
			{
			 $mapping{$tmp_fldname}{"type"} = "string"; 
			 $mapping{$tmp_fldname}{"index"} = "not_analyzed"; 
			}
		} # Close Header
		
		print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " ...$xl_filename field headers parsed."  if $verbose;
			
		if(! $e->indices->exists(index => $index)) {
			$result = $e->indices->create(index => $index);
			# $result = $e->indices->delete(index => $index);
		}
		
		print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " ...Index created/reused $index."  if $verbose;
		
		$result = $e->indices->put_mapping(index => $index, type  => $type, body => {$type => { "properties" => {%mapping}}});
		
		print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " ...mapping done.\n" if $verbose;
		
		for my $row ( 1 .. $row_max ) {
			my %kivalu = ();															
			my $ErrFlg=0;																			
			for my $col ( $col_min .. $col_max ) {
				my $tmp = $fields[$col];
				my $cell = $worksheet->get_cell( $row, $col );
                next unless $cell;
				my $valu = $cell->value();
				if($mapping{$tmp}{"type"} =~ "string") {
					$valu =~ s/[^[:print:]]+/./gi;
					$valu =~ s/[\"\-\\\/#,']/./gi;
					$valu =~ s/\^/,/gi;
					$kivalu{$tmp}=$valu;
					}
				elsif (($cell->type() =~ "Numeric") && ($mapping{$tmp}{"type"} =~ "integer" || $mapping{$tmp}{"type"} =~ "double"))
					{ $kivalu{$tmp}=$valu; }
				elsif ($mapping{$tmp}{"type"} =~ "date" && $cell->type() =~ "Numeric")
					{ $kivalu{$tmp}=$valu; }
				else { 
					print "\nErr: ",$row,":",$col," Field: ", $tmp, " MapType: ",$mapping{$tmp}{"type"}," CellType: ",$cell->type()," CellValue: ",$valu  if $verbose;
					$ErrFlg=1;
					}								
			} # Close individual Row	
			
			if($ErrFlg == 0) {
				# my $doc = { index => $index, type => $type, id => $id++, body => \%kivalu };     # incase need to replace ES _id with XL row no!
				$id++;
				my $doc = { index => $index, type => $type, body => \%kivalu };
				$e->index($doc);
				# print STDERR "\nIndexed: $id" if $verbose;										# incase needs to be too verbose
				}
			else
				{ print STDERR "XL2ES: Could not parse row:$id" if $verbose }
			if($id%500 == 0) { print "\nXL2ES :> ", localtime->strftime('%m/%d/%Y %H:%M'), " ...$id records done!"  if $verbose; }			
			
		} # Close All Rows
		
} # Close Worksheets


	
# that's it.. 	


	
# CG -- ones and for all -- an Excel parser for all XLS data to be imported into ES !
# comments sanit for clairty .. chk older ver for details.
# readme - pending
# todo - auto kibana dashbaord
# Aug.13 -- xl2es.pl :>
# more fun added on 6th feb... 7-8th.feb...

# --

#if ($opt{index} ne 0) {	chomp($opt{index}); } else { $opt{index} = "xl2es"; }
#if ($opt{type} ne 0) {	chomp($opt{type}); } else { $opt{type} = "xldata"; }
#if ($opt{es_server} ne 0) {	chomp($opt{es_server}); } else { $opt{es_server} = "localhost"; }
#$opt{es_port} = ($opt{es_port} eq 0)? 9200 : $opt{es_port} ;
