#!perl
$|=1;
use lib '../../../../lib';
use MooseX::ShortCut::BuildInstance qw( build_instance );
use Spreadsheet::XLSX::Reader::Error;

my 	$action = build_instance(
		add_attributes =>{ 
			error_inst =>{
				handles =>[ qw( error set_error clear_error set_warnings if_warn ) ],
			},
		},
		error_inst => Spreadsheet::XLSX::Reader::Error->new(
			should_warn => 1,
			# should_warn => 0,# to turn off cluck when the error is set
		),
	);
print	$action->dump;
		$action->set_error( "You did something wrong" );
print	$action->dump;
print	$action->error . "\n";