#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain8#########9
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
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 17;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'debug',
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'calcChain.xml';
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				error_inst
			);
my  		@instance_methods = qw(
				get_calc_chain_position
				get_file_name
				where_am_i
				has_position
				get_log_space
				set_log_space
				parse_element
			);
my			$answer_ref = [
				{
		          r => 'D12',
		          i => '1'
		        },
				{
		          r => 'D14',
		          s => '1'
		        },
				{
		          r => 'E14',
		          s => '1'
		        },
				{
		          r => 'B9',
		        },
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain has the -$_- attribute"
} 			@class_attributes;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance =	Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain->new(
									file_name	=> $test_file,
									log_space	=> 'Test',
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
								);
}										"Prep a new CalcChain instance";
map{
can_ok		$test_instance, $_,
} 			@instance_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
is_deeply	$test_instance->get_calc_chain_position( 0 ), $answer_ref->[0],
										"Get the zeroth calcChain 'c' element";
lives_ok{	$capture = $test_instance->get_calc_chain_position( 20 );
}										"Attempt an element past the end of the list";
is			$capture, undef,			'Show that undef is returned';
			for my $x ( 0..$#$answer_ref ){
is_deeply	$test_instance->get_calc_chain_position( $x ), $answer_ref->[$x],
										"Get the calcChain 'c' element at position: $x";
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