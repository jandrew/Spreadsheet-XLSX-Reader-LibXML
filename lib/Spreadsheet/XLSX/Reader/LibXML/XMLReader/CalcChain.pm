package Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain;
use version; our $VERSION = qv('v0.38.16');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain-$VERSION";

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw( Enum	Bool ArrayRef );
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has empty_return_type =>(
		isa		=> Enum[qw( empty_string undef_string )],
		reader	=> 'get_empty_return_type',
		writer	=> 'set_empty_return_type',
	);
with	'Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData';

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_calc_chain_position{
	my( $self, $position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::get_calc_chain_position', );
	if( !defined $position ){
		$self->set_error( "Requested calc chain position required - none passed" );
		return undef;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Getting the calcChain position: $position" ] );
	my $result = $self->_get_cc_position( $position );
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning the calcChain result: " . ($result//'') ] );
	return $result;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has _calc_chain_positions =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		default	=> sub{ [] },
		handles	=>{
			_get_cc_position => 'get',
			_set_cc_position => 'set',
		},
		reader => '_get_all_cache',
		writer => '_set_all_cache',
	);
	
###LogSD	has '+class_space' =>( default => 'CalcChain' );

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

###LogSD	sub BUILD {
###LogSD	    my $self = shift;
###LogSD			$self->set_class_space( 'CalcChain' );
###LogSD	}

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the sharedStrings unique bits" ] );
	$self->start_the_file_over;
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Currently at libxml2 level: $node_depth",
	###LogSD		"Current node name: $node_name",
	###LogSD		"..for type: $node_type", ] );
	my	$result = 1;
	if( $node_name eq 'calcChain' ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"already at the calcChain node" ] );
	}else{
		$result = $self->advance_element_position( 'calcChain' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"attempt to get to the sst element result: $result" ] );
	}
	if( $result ){
		my $calcChain_ref = $self->parse_element;
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Perl CalcChain is:", $calcChain_ref ] );
			
		$self->_set_all_cache( $calcChain_ref->{list} );
		return undef;
	}else{
		$self->set_error( "No 'calcChain' element found - can't parse this as a calc chain file" );
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

Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain - A LibXML::Reader calcChain base class

=head1 SYNOPSIS

See General  L<SYNOPSIS|Spreadsheet::XLSX::Reader::LibXML::CalcChain/SYNOPSIS>
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module.  To use the general 
package for excel parsing out of the box please review the documentation for L<Workbooks
|Spreadsheet::XLSX::Reader::LibXML>, L<Worksheets
|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>.

This class is used to access the sub file calcChain.xml from an unzipped .xlsx file.  
The file to be read is generally found in the xl/ sub folder after the master file is 
unzipped.  The calcChain.xml file contains the calculation sequence and data source(s) 
used when building the value for cells with formulas.  (The formula presented in the 
L<Cell|Spreadsheet::XlSX::Reader::LibXML::Cell> instance is collected from somewhere 
else.)

This particular class uses the L<XML::LibXML::Reader> to parse the file.  The 
intent with that reader is to maintain the minimum amount of the file in working memory.  
As a consequence two things are sacrificed.  First this implementation will read the file 
serially.  This means that to read a previous element the file must start over at the 
beginning.  Second, the connection between the XMLReader and the file is broken to 
rewind the file pointer.  In an effort to minimize this pain the system file 'open' 
function is handled in a separate file handle variable so that a new system open is 
not required on each rewind.  For information regarding necessary steps to replace 
this package review the L<general|Spreadsheet::XLSX::Reader::LibXML::CalcChain/DESCRIPTION> 
documentation

This is a child class that both inherits from a parent and collates a role for full 
functional implementation.  Read the documentation for each of them as well as this 
documentation to gain a complete picture of this class.

=head2 extends

This is the parent class of this class

=head3 L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

=head2 with

These are attached roles for additional (re-used) functionality

=head3 L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData>

=head2 Primary Functions

These are the primary way(s) to use this class.

=head3 get_calc_chain_position( $integer )

=over

B<Definition:> This will return the calcChain information from the identified $integer position.  
(Counting from zero).  The information is returned as a perl hash ref.

B<Accepts:> an $integer for the calcChain position

B<Returns:> a hash ref of data

=back

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

This software is copyrighted (c) 2014, 2015 by Jed Lund

=head1 DEPENDENCIES

=over

L<version>

L<perl 5.010|perl/5.10.0>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults>

L<lib>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData>

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