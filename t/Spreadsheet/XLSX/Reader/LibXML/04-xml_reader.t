#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLReader 7#########8#########9
#!/usr/bin/env perl
my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/xl/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../' . $lib;
		$test_file	= '../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 54;
use	Test::Moose;
use IO::File;
#~ use XML::LibXML::Reader;
use	MooseX::ShortCut::BuildInstance qw( build_class build_instance );
use	lib
		'../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'info',
###LogSD							},
###LogSD							Test =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'debug',
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'sharedStrings.xml';
my  ( 
			$class_instance, $test_instance, $capture, @answer, $error_instance, $file_handle,
	);
my 			@class_attributes = qw(
				file
				error_inst
				xml_version
				xml_encoding
				xml_header
			);
my  		@class_methods = qw(
				get_file
				set_file
				clear_file
				has_file
				error
				set_error
				clear_error
				set_warnings
				if_warn
				encoding
				start_the_file_over
				get_text_node
				get_attribute_hash_ref
				advance_element_position
				location_status
				version
				has_encoding
				get_header
				copy_current_node
			);
				#~ where_am_i
				#~ has_position
my			$answer_ref = [
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Hello</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>World</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>my</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t xml:space="preserve"> </t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Category</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Total</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Date</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Red</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Blue</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Omaha</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Row Labels</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Grand Total</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Superbowl Audibles</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Column Labels</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>2/6/2011</t></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><r><t>He</t></r><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">llo </t></r><r><rPr><b/><sz val="20"/><color rgb="FF0070C0"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>World</t></r></si>',
				],
				[
					'<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>Red</t></si>',
				],
			];

###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
ok			$class_instance = build_class(
								superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader' ],
								package		 => 'ReaderInstance',
							);
map{ 
has_attribute_ok
			$class_instance, $_,
										"Check that $class_instance has the -$_- attribute"
} 			@class_attributes;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$file_handle	=	IO::File->new( $test_file, "<");
			$test_instance	=	$class_instance->new(
									file	=> $file_handle,
									#~ xml_reader 	=> XML::LibXML::Reader->new( IO => $file_handle ),
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep a new Reader instance";
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
map{
			$test_instance->advance_element_position( 'si' );
			my ( $x, $row ) = ( 0, $_ );
			@answer = split "\n", $test_instance->copy_current_node( 1 )->toString;
map{
is			$_, $answer_ref->[$row]->[$x],
										'Test matching line -' . (1 + $x++) . "- of 'si' position: $row";
}			@answer;
}( 0..10);
ok			$test_instance->start_the_file_over,
										'Test re-starting the file';
my 			$row = 0;
while( 	$test_instance->advance_element_position( 'si' ) ){#$capture = 
			my $x = 0;
			@answer = split "\n", $test_instance->copy_current_node( 1 )->toString;
map{
is			$_, $answer_ref->[$row]->[$x],
										'Test matching line -' . (1 + $x++) . "- of 'si' position: $row";
}			@answer;
			$row++;
}
explain 								"...Test Done";
done_testing();

###LogSD	package Print::Log;
###LogSD	use Data::Dumper;
###LogSD	sub new{
###LogSD		bless {}, shift;
###LogSD	}
###LogSD	sub add_line{
###LogSD		shift;
###LogSD		my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
###LogSD						@{$_[0]->{message}} : $_[0]->{message};
###LogSD		my ( @print_list, @initial_list );
###LogSD		no warnings 'uninitialized';
###LogSD		for my $value ( @input ){
###LogSD			push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
###LogSD		}
###LogSD		for my $line ( @initial_list ){
###LogSD			$line =~ s/\n$//;
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;