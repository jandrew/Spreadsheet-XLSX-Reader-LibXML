package Spreadsheet::XLSX::Reader::XMLReader::SharedStrings;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
    );
use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::XMLReader';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

sub get_shared_string_position{
	my( $self, $position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::get_shared_string_position', );
	if( !defined $position ){
		$self->set_error( "Requested shared string position required - none passed" );
		return undef;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Getting the sharedStrings position: $position" ] );
	
	#checking if the reqested position is too far
	if( $position > $self->_get_unique_count - 1 ){
		$self->set_error( "Asking for position -$position- (from 0) but the shared string " .
							"max cell position is: " . ($self->_get_unique_count - 1) );
		return undef;#  fail
	}
	
	# Initiate the read or reset the file if needed
	if( $self->has_position and $self->where_am_i > $position ){
		$self->start_the_file_over;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Finished resetting the file" ] );
	}
	if( !$self->has_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the first cell" ] );
		my $found_it = $self->next_element( 'si' );
		if( $found_it < 1 ){
			$self->set_error( "No strings stored in the sharedStrings file" );
			return undef;
		}
		$self->_i_am_here( 0 );
	}
	
	while( $self->where_am_i < $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the next cell: " . ($self->where_am_i + 1) ] );
		$self->next_element( 'si' );
		$self->_i_am_here( $self->where_am_i + 1 );
	}
	
	my $shared_strings_node = $self->copy_current_node( 1 );
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Returning shared strings node:",
	###LogSD		(($shared_strings_node) ? $shared_strings_node->toString : '') ] );
	return $shared_strings_node;
}



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	reader		=> '_get_unique_count',
	clearer		=> '_clear_unique_count',
);

has _count =>(
	isa			=> Int,
	writer		=> '_set_count',
	reader		=> '_get_count',
	clearer		=> '_clear_count',
	predicate	=> '_has_count',
);
	
has +_core_element =>(
		default => 'si',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the sharedStrings unique bits" ] );
	my	$node_name	= $self->name;
	my	$found_it	= 1;
	if( $node_name ne 'sst' ){
		$found_it = $self->next_element( 'sst' );
	}
	if( $found_it ){
		$self->_set_unique_count( $self->get_attribute( 'uniqueCount' ) );#
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The unique count is: " . $self->_get_unique_count ] );
		$self->_set_count( $self->get_attribute( 'count' ) );#
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The count is: " . $self->_get_count ] );
		return undef;
	}else{
		$self->set_error( "No 'sst' element found - can't parse this as a shared strings file" );
		$self->_clear_unique_count;
		$self->_clear_count;
		return undef;
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::XMLReader::SharedStrings - sharedStrings parsing with XML::LibXML::Reader
    
=head1 DESCRIPTION

This is the XMLReader version of the Shared strings parser.  Both the XMLReader and DOM 
versions of this parser should function in the background of this package. There shouldn't 
be any need to manage this class by the user.

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

=back

=head1 TODO

=over

B<1.> Nothing L<yet|/SUPPORT>

=back

=head1 AUTHOR

=over

=item Jed Lund

=item jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Spreadsheet::XLSX::Reader::XMLReader>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9