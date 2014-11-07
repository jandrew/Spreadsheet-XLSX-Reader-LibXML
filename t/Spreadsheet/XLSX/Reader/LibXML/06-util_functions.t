#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::UtilFunctions       8#########9
#!evn perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }##### $ENV{ Smart_Comments } = '### ####';
$| = 1;

use	Test::Most tests => 5;
use	Test::Moose;
use Data::Dumper;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );#
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							main =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'trace',
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::UtilFunctions v0.5;
use	Spreadsheet::XLSX::Reader::LibXML::LogSpace v0.5;
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
my  ( 
			$test_instance,
	);
my 			$row = 0;
my  		@class_methods = qw(
				_add_integer_separator				_continuous_fraction
			);
my			$question_list =[
				[ 123456789, ',', 3 ],
				[ 0.12345678, 20, 6 ],
			];
my			$answer_list =[
				'123,456,789',
				'10/81',
			];
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
			$test_instance	=	build_instance(
									package	=> 'FmtDefaultTest',
									roles	=>[ 
										'Spreadsheet::XLSX::Reader::LibXML::LogSpace'
									],
									add_roles_in_sequence =>[
										'Spreadsheet::XLSX::Reader::LibXML::UtilFunctions',
									],
									log_space => 'Test',
								);
}										"Prep a test UtilFunctions instance";
map{
can_ok		$test_instance, $_,
} 			@class_methods;
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			no warnings 'uninitialized';
			my $test_position = 0;
is			$test_instance->_add_integer_separator( @{$question_list->[$test_position]} ), $answer_list->[$test_position],
										,"Check the integer separator function for -$question_list->[$test_position]->[0]- with result: $answer_list->[$test_position++]";
is			$test_instance->_continuous_fraction( @{$question_list->[$test_position]} ), $answer_list->[$test_position],
										,"Check the continuous fraction function for -$question_list->[$test_position]->[0]- with result: $answer_list->[$test_position++]";
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