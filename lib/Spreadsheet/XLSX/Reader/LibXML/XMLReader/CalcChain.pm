package Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain;
use version; our $VERSION = qv('v0.10.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
with	'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData';

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_calc_chain_position{
	my( $self, $position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::get_calc_chain_position', );
	if( !defined $position ){
		$self->set_error( "Requested calc chain position required - none passed" );
		return undef;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Getting the calcChain position: $position" ] );
	
	# Initiate the read or reset the file if needed
	if( $self->has_position and $self->where_am_i > $position ){
		$self->start_the_file_over;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Finished resetting the file" ] );
	}
	if( !$self->has_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the first cell" ] );
		my $found_it = $self->next_element( 'c' );
		if( $found_it < 1 ){
			$self->set_error( "No strings stored in the sharedStrings file" );
			return undef;
		}
		$self->_i_am_here( 0 );
	}
	
	while( $self->where_am_i < $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the next cell: " . ($self->where_am_i + 1) ] );
		my $result = $self->next_element( 'c' );
		if( !$result ){
			$self->_clear_location;
			$self->start_the_file_over;
			return undef;
		}
		$self->_i_am_here( $self->where_am_i + 1 );
	}
	
	my $calc_chain_node = $self->parse_element;
	$self->_i_am_here( $self->where_am_i + 1 );
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Returning shared strings node:",
	###LogSD		(($calc_chain_node) ? $calc_chain_node : '') ] );
	return $calc_chain_node;
}



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain - LibXML::Reader for the calcChain file
    
=head1 DESCRIPTION

POD not written yet!

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9