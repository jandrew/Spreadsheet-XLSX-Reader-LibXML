#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings #####9
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
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 66;
use	Test::Moose;
use IO::File;
use XML::LibXML::Reader;
use Data::Dumper;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	lib
		'../../../../../../../Log-Shiras/lib',
		'../../../../../..//lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD							Test =>{
###LogSD								XMLReader =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								parse_element =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'sharedStrings.xml';
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance, $file_handle,
	);
my 			@class_attributes = qw(
				file
				error_inst
			);
my  		@class_methods = qw(
				get_shared_string_position
				get_file
				set_file
				has_file
				clear_file
				where_am_i
				has_position
			);
my			$answer_ref = [
				16,
				'UTF-8',
				[ 0, { 	
					raw_text => 'Hello',
				} ],
				[ 15, { 	
					raw_text => 'Hello World',
					rich_text =>[
						2,
						{	
							'color' => {
								'rgb' => 'FFFF0000'
							},
							'sz' => '11',
							#~ 'b' => 1,
							'b' => undef,
							'scheme' => 'minor',
							'rFont' => 'Calibri',
							'family' => '2'
						},
						6,
						{
							'color' => {
								'rgb' => 'FF0070C0'
							},
							'sz' => '20',
							#~ 'b' => 1,
							'b' => undef,
							'scheme' => 'minor',
							'rFont' => 'Calibri',
							'family' => '2'
						}
					] 
				} ],
				[ 10, { 	
					raw_text => 'Row Labels',
				} ],
				[ 1, { 	
					raw_text => 'World',
				} ],
				[ 2, { 	
					raw_text => 'my',
				} ],
				[ 3, { 	
					raw_text => ' ',
				} ],
				[ 4, { 	
					raw_text => 'Category',
				} ],
				[ 5, { 	
					raw_text => 'Total',
				} ],
				[ 6, { 	
					raw_text => 'Date',
				} ],
				[ 7, { 	
					raw_text => 'Red',
				} ],
				[ 8, { 	
					raw_text => 'Blue',
				} ],
				[ 9, { 	
					raw_text => 'Omaha',
				} ],
				[ 14, { 	
					raw_text => '2/6/2011',
				} ],
				[ 0, 'Hello', ],
				[ 15, 'Hello World', ],
				[ 10, 'Row Labels', ],
				[ 1, 'World', ],
				[ 2, 'my', ],
				[ 3, ' ', ],
				[ 4, 'Category', ],
				[ 5, 'Total', ],
				[ 6, 'Date', ],
				[ 7, 'Red', ],
				[ 8, 'Blue', ],
				[ 9, 'Omaha', ],
				[ 14, '2/6/2011', ],
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
									file		=> $test_file,
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep a new SharedStrings instance - cache_positions => 1";

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			my	$answer_row = 0;
is			$test_instance->_get_unique_count, $answer_ref->[$answer_row++],
										"Check for correct unique_count";
is			$test_instance->encoding, $answer_ref->[$answer_row++],
										"Check for correct encoding";
###LogSD	my $trigger = 3;
			for my $x ( 2..14 ){
###LogSD	print "-----------$x---------\n";
###LogSD	if( $x == $trigger ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'trace',
###LogSD			},
###LogSD		} );
###LogSD	}
###LogSD	elsif( $x > $trigger + 1 ){
###LogSD		exit 1;
###LogSD	}
			#~ print "position -$answer_ref->[$x]->[0]- " .Dumper( $test_instance->get_shared_string_position( $answer_ref->[$x]->[0] ) );
is_deeply	$test_instance->get_shared_string_position( $answer_ref->[$x]->[0] ), $answer_ref->[$x]->[1],
										"Get the -$answer_ref->[$x]->[0]- sharedStrings 'si' position as:" . Dumper( $answer_ref->[$x]->[1] );
			}
lives_ok{	$capture = $test_instance->get_shared_string_position( 20 ); 
}										"Attempt an element past the end of the list";
is		$capture, undef,				'Make sure it returns undef';
lives_ok{	$capture = $test_instance->get_shared_string_position( 16 ); 
}										"Attempt a different element past the end of the list";
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'warn',
###LogSD			},
###LogSD		} );
###LogSD		$phone->talk( level => 'info', message => [ "Rerun the whole thing without caching" ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
									file			=> $test_file,
									cache_positions	=> 0,
									#~ no_formats		=> 1,
									error_inst 		=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn		=> 1,
										should_warn		=> 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep a new SharedStrings instance - cache_positions => 0";
			$answer_row = 0;
is			$test_instance->_get_unique_count, $answer_ref->[$answer_row++],
										"Check for correct unique_count";
is			$test_instance->encoding, $answer_ref->[$answer_row++],
										"Check for correct encoding";
###LogSD	$trigger = 15;
			for my $x ( 2..14 ){
###LogSD	print "-----------$x---------\n";
###LogSD	if( $x == $trigger ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'trace',
###LogSD			},
###LogSD		} );
###LogSD	}
###LogSD	elsif( $x > $trigger + 1 ){
###LogSD		exit 1;
###LogSD	}
is_deeply	$test_instance->get_shared_string_position( $answer_ref->[$x]->[0] ), $answer_ref->[$x]->[1],
										"Get the -$answer_ref->[$x]->[0]- sharedStrings 'si' position as:" . Dumper( $answer_ref->[$x]->[1] );
			}
lives_ok{	$capture = $test_instance->get_shared_string_position( 20 ); 
}										"Attempt an element past the end of the list";
is		$capture, undef,				'Make sure it returns undef';
lives_ok{	$capture = $test_instance->get_shared_string_position( 16 ); 
}										"Attempt a different element past the end of the list";
is		$capture, undef,				'Make sure it returns undef';
###LogSD		$phone->talk( level => 'info', message => [ "Rerun the whole as values only" ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
									group_return_type	=> 'xml_value',
									file			=> $test_file,
									error_inst 		=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn		=> 1,
										should_warn		=> 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep a new SharedStrings instance - no_formats => 1";
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'trace',
###LogSD			},
###LogSD		} );

###LogSD		$phone->talk( level => 'trace', message => [ "The instance", $test_instance ] );
			for my $x ( 15..27 ){
is_deeply	$test_instance->get_shared_string_position( $answer_ref->[$x]->[0] ), $answer_ref->[$x]->[1],
										"Get the -$answer_ref->[$x]->[0]- sharedStrings 'si' position as:" . $answer_ref->[$x]->[1];
###LogSD		$phone->talk( level => 'trace', message => [ "The instance", $test_instance ] );
			}
lives_ok{	$capture = $test_instance->get_shared_string_position( 20 ); 
}										"Attempt an element past the end of the list";
is		$capture, undef,				'Make sure it returns undef';
lives_ok{	$capture = $test_instance->get_shared_string_position( 16 ); 
}										"Attempt a different element past the end of the list";
is		$capture, undef,				'Make sure it returns undef';
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