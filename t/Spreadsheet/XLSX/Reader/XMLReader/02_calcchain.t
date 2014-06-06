#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::CalcChain        8#########9
#!perl
$| = 1;

use	Test::Most;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
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
use	Spreadsheet::XLSX::Reader::XMLReader::CalcChain;
use	Spreadsheet::XLSX::Reader::Error;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/';
	$test_file .= 'calcChain.xml';
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				epoch_year
				error_inst
			);
my  		@instance_methods = qw(
				get_position
				get_file_name
				where_am_i
				has_position
				get_log_space
				set_log_space
				get_core_element
			);
my			$answer_ref = [
				[
					'<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="D11" i="1"/>',
				],
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::XMLReader::CalcChain', $_,
										"Check that Spreadsheet::XLSX::Reader::XMLReader::CalcChain has the -$_- attribute"
} 			@class_attributes;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance =	Spreadsheet::XLSX::Reader::XMLReader::CalcChain->new(
									file_name	=> $test_file,
									log_space	=> 'Test',
									error_inst	=> Spreadsheet::XLSX::Reader::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
									epoch_year	=> 1904,
								);
}										"Prep a new CalcChain instance";
map{
can_ok		$test_instance, $_,
} 			@instance_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
ok			$capture = $test_instance->get_position( 0 ),
										"Get the zeroth calcChain 'c' element";
			$x = 0;
			@answer = split "\n", $capture;
map{
is			$_, $answer_ref->[0]->[$x],
										'Test matching line -' . (1 + $x++) . "- of 'c' position: 0";
}			@answer;
lives_ok{	$capture = $test_instance->get_position( 20 );
}										"Attempt an element past the end of the list";
is			$capture, undef,			'Show that undef is returned';
ok			$capture = $test_instance->get_position( 0 ),
										"Get the zeroth calcChain 'c' element";
			$x = 0;
			@answer = split "\n", $capture;
map{
is			$_, $answer_ref->[0]->[$x],
										'Test matching line -' . (1 + $x++) . "- of 'c' position: 0";
}			@answer;
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