#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData #####9
#!/usr/bin/env perl
my ( $lib, $test_file, $test_fil2 );
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

use	Test::Most tests => 30;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	use Data::Dumper;
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
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
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_fil2 = $test_file . 'worksheets/sheet3_test.xml';
$test_file .= 'sharedStrings.xml';
#~ print "$lib\n$test_file\n$test_fil2\n";
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file
				error_inst
			);
my  		@instance_methods = qw(
				parse_element
				get_file
				set_file
				has_file
				clear_file
				where_am_i
				has_position
				parse_element
				error
				set_error
				clear_error
				set_warnings
				if_warn
				node_name
				byte_consumed
				move_to_first_att
				move_to_next_att
				node_depth
				node_value
				node_type
				has_value	
				start_reading
			);
my			$answer_ref = [
				{
					'list' => [
						{
							't' => {
								'raw_text' => 'He'
							}
						},
						{
							'rPr' => {
								'color' => {
									'rgb' => 'FFFF0000'
								},
								'sz' => '11',
								'b' => 1,
								'scheme' => 'minor',
								'rFont' => 'Calibri',
								'family' => '2'
							},
							't' => {
								'raw_text' => 'llo '
							}
						},
						{
							'rPr' => {
								'color' => {
									'rgb' => 'FF0070C0'
								},
								'sz' => '20',
								'b' => 1,
								'scheme' => 'minor',
								'rFont' => 'Calibri',
								'family' => '2'
							},
							't' => {
								'raw_text' => 'World'
							}
						}
					]
		        },
				{
					'r' => 'A11',
					'v' => {
						'raw_text' => '1'
					},
					's' => '8'
				},
				{
					'r' => 'B12',
					'v' => {
						'raw_text' => ''
					}
		        }
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
			$test_instance	=	build_instance(
									package => 'TestIntance',
									superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader', ],
									add_roles_in_sequence =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData', ],
									file	=> $test_file,
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep a new TestIntance to test XMLToPerlData";
map{ 
has_attribute_ok
			$test_instance, $_,
										"Check that the " . ref( $test_instance ) . " has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@instance_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
explain		"index to position 15";
			map{ $test_instance->next_element( 'si' ) }( 0..15 );
			#~ print Dumper( $test_instance->parse_element );
is_deeply	$test_instance->parse_element, $answer_ref->[0],
										"Check that the output matches expectations.";
lives_ok{
			$test_instance	=	TestIntance->new(
									file	=> $test_fil2,
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
			###LogSD				log_space	=> 'Test',
								);
}										"Prep another TestIntance to test XMLToPerlData";
explain		"Index to position 12";
			map{ $test_instance->next_element( 'c' ) }( 0..12 );
			#~ print Dumper( $test_instance->parse_element );
			#~ exit 1;
is_deeply	$test_instance->parse_element, $answer_ref->[1],
										"Check that the next output matches expectations.";
ok			$test_instance->next_element( 'c' ),
										"Advance to the next cell";
is_deeply	$test_instance->parse_element, $answer_ref->[2],
										"Check that the next output matches expectations.";
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