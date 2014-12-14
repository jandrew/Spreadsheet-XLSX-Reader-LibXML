package Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use version; our $VERSION = qv('v0.28.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				HashRef			is_HashRef
    );
use Carp qw( confess );
use lib	'../../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
with	'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

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
	
	# checking if the reqested position is stored
	if( $self->_has_last_position and $position == $self->_get_last_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already built the answer for position: $position", 
		###LogSD		$self->_get_last_position_ref						] );
		return $self->_get_last_position_ref;
	}
	
	# Initiate the read and reset the file if needed
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
	
	# Advance to the proper position
	while( $self->where_am_i < $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the next cell: " . ($self->where_am_i + 1) ] );
		$self->next_element( 'si' );
		$self->_i_am_here( $self->where_am_i + 1 );
	}
	
	# Read the data
	$self->_set_last_position( $position );
	my $init_ref = $self->parse_element;
	$self->_i_am_here( $position + 1 );
	###LogSD	$phone->talk( level => 'trace',  message =>[
	###LogSD		"Element parse resulted in:", $init_ref ] );
	if( is_HashRef( $init_ref ) ){
		if( exists $init_ref->{list} ){
			my ( $raw_text, $rich_text );
			for my $element( @{$init_ref->{list}} ){
				push( @$rich_text, length( $raw_text ), $element->{rPr} ) if exists $element->{rPr};
				$raw_text .= $element->{t}->{raw_text};
			}
			@$init_ref{qw( raw_text rich_text )} = ( $raw_text, $rich_text  );
			delete $init_ref->{list};
		}else{
			$init_ref = $init_ref->{t};
		}
	}else{
		set_error( "Unable to parse the shared string position: $position" );
		return undef;
	}
	$self->_set_last_position_ref( $init_ref );
	###LogSD	$phone->talk( level => 'trace',  message =>[
	###LogSD		"Final element rearranging result:", $init_ref ] );
	return $init_ref;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	reader		=> '_get_unique_count',
	clearer		=> '_clear_unique_count',
);

has _last_position =>(
		isa		=> Int,
		writer	=> '_set_last_position',
		reader	=> '_get_last_position',
		predicate => '_has_last_position',
		trigger	=> \&_clear_last_position_ref,
	);

has _last_position_ref =>(
		isa		=> HashRef,
		writer	=> '_set_last_position_ref',
		reader	=> '_get_last_position_ref',
		clearer => '_clear_last_position_ref',
		predicate => '_has_last_position_ref',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the sharedStrings unique bits" ] );
	my	$node_name	= $self->node_name;
	my	$found_it	= 1;
	if( $node_name ne 'sst' ){
		$found_it = $self->next_element( 'sst' );
	}
	if( $found_it ){
		$self->_set_unique_count( $self->get_attribute( 'uniqueCount' ) );#
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The unique count is: " . $self->_get_unique_count ] );
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

B<This documentation is written to explain ways to extend this package.  To use the data 
extraction of Excel workbooks, worksheets, and cells please review the documentation for  
L<Spreadsheet::XLSX::Reader::LibXML>,
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Spreadsheet::XLSX::Reader::LibXML::Cell>>

POD not written yet!

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

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

L<Spreadsheet::XLSX::Reader::LibXML>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::XLSX> - 2007+

L<Spreadsheet::ParseXLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9