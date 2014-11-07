#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::Error     7#########8#########9
#!env perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }
$| = 1;

use	Test::Most tests => 9;
use	Test::Moose;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use Capture::Tiny qw( capture_stderr );
use	lib 
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
#~ use Log::Shiras::Switchboard v0.21 qw( :debug );#
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
###LogSD		$operator->add_skip_up_caller( qw( Carp __ANON__ ) );
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::LogSpace v0.4;
use Spreadsheet::XLSX::Reader::LibXML::Types v0.4;
my  ( 
			$test_instance, $capture,
	);
my 			@class_attributes = qw(
				log_space
			);
my  		@class_methods = qw(
				get_log_space
				set_log_space
				
			);
my			$answer_ref = [
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "harder question first ..." ] );
lives_ok{
			#~ use Spreadsheet::XLSX::Reader::Types;
			$test_instance = build_instance(
				package => 'LogSpace::Instance',
				roles =>[ 'Spreadsheet::XLSX::Reader::LibXML::LogSpace' ],
				log_space => 'Test',
			);
}										"Prep a test LogSpace instance";
###LogSD		$phone->talk( level => 'info', message => [ "now easy questions ..." ] );
map{ 
has_attribute_ok
			$test_instance, $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::LogSpace has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
is			$test_instance->get_log_space, 'Test',
										"Check that the log_space can be retrieved";
is			$Spreadsheet::XLSX::Reader::LibXML::Types::log_space, 'Test::Types',
										"...and that the Types log_space was set accordingly";
is			$test_instance->set_log_space( 'New::Space' ), 1,
										"Change the log_space";
is			$test_instance->get_log_space, 'New::Space',
										"Check that the new log_space can be retrieved";
is			$Spreadsheet::XLSX::Reader::LibXML::Types::log_space, 'New::Space::Types',
										"...and that the Types log_space was set accordingly";
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