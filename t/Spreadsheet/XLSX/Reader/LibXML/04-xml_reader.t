#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLReader 7#########8#########9
#!env perl
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

use	Test::Most tests => 47;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
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
			$test_instance, $capture, @answer, $error_instance,
	);
my 			@class_attributes = qw(
				file_name
				error_inst
			);
my  		@class_methods = qw(
				get_file_name
				copy_current_node
				byte_consumed
				start_reading
				encoding
				next_element
				next_sibling
				next_sibling_element
				get_attribute
				read_state
				node_name
				where_am_i
				has_position
				start_the_file_over
				get_log_space
				set_log_space
				
			);
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
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLReader', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLReader has the -$_- attribute"
} 			@class_attributes;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance =	Spreadsheet::XLSX::Reader::LibXML::XMLReader->new(
									file_name		=> $test_file,
									log_space	=> 'Test',
									error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
								);
}										"Prep a new Reader instance";
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
map{
			$test_instance->next_element( 'si' );
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
while( 	$test_instance->next_element( 'si' ) ){#$capture = 
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
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{name_space}, $_[0]->{level},
###LogSD					$_[0]->{filename}, $_[0]->{line},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;