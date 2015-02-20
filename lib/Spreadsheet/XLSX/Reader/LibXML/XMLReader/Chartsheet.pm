package Spreadsheet::XLSX::Reader::LibXML::XMLReader::Chartsheet;
use version; our $VERSION = qv('v0.34.4');


use	5.010;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use Carp qw( confess );
use Types::Standard qw(
		Int				Str				ArrayRef
		HashRef			HasMethods		Bool
		Enum
    );
use lib	'../../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
#~ with	'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow',
		#~ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData',
		#~ ;
#~ ###LogSD	use Log::Shiras::UnhideDebug;
#~ with	'Spreadsheet::XLSX::Reader::LibXML::GetCell';

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has sheet_type =>(
		isa		=> Enum[ 'chartsheet' ],
		default	=> 'chartsheet',
		reader	=> 'get_sheet_type',
	);

has sheet_rel_id =>(
		isa		=> Str,
		reader	=> 'rel_id',
	);

has sheet_id =>(
		isa		=> Int,
		reader	=> 'sheet_id',
	);

has sheet_position =>(# XML position
		isa		=> Int,
		reader	=> 'position',
	);

has sheet_name =>(
		isa		=> Str,
		reader	=> 'get_name',
	);

has drawing_rel_id =>(
		isa		=> Str,
		writer	=> '_set_drawing_rel_id',
		reader	=> 'get_drawing_rel_id',
	);
	
has workbook_instance =>(
		isa		=> HasMethods[qw(
						counting_from_zero			boundary_flag_setting
						change_boundary_flag		_has_shared_strings_file
						get_shared_string_position	_has_styles_file
						get_format_position			set_empty_is_end
						is_empty_the_end			_starts_at_the_edge
						get_group_return_type		set_group_return_type
						get_epoch_year				change_output_encoding
						get_date_behavior			set_date_behavior
						get_empty_return_type		set_error
						get_values_only				set_values_only
					)],
		handles	=> [qw(
						counting_from_zero			boundary_flag_setting
						change_boundary_flag		_has_shared_strings_file
						get_shared_string_position	_has_styles_file
						get_format_position			set_empty_is_end
						is_empty_the_end			_starts_at_the_edge
						get_group_return_type		set_group_return_type
						get_epoch_year				change_output_encoding
						get_date_behavior			set_date_behavior
						get_empty_return_type		set_error
						get_values_only				set_values_only
					)],
		required => 1,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

#~ sub min_row{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::min_row' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning the minimum row: " . $self->_min_row ] );
	#~ return $self->_min_row;
#~ }

#~ sub max_row{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::max_row' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning the maximum row: " . $self->_max_row ] );
	#~ return $self->_max_row;
#~ }

#~ sub min_col{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::min_col' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning the minimum column: " . $self->_min_col ] );
	#~ return $self->_min_col;
#~ }

#~ sub max_col{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::max_col' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning the maximum column: " . $self->_max_col ] );
	#~ return $self->_max_col;
#~ }

#~ sub row_range{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::row_range' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning row range( " . $self->_min_row . ", " . $self->_max_row . " )" ] );
	#~ return( $self->_min_row, $self->_max_row );
#~ }

#~ sub col_range{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::col_range' ), );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Returning col range( " . $self->_min_col . ", " . $self->_max_col . " )" ] );
	#~ return( $self->_min_col, $self->_max_col );
#~ }


#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

#~ has _sheet_min_col =>(
		#~ isa			=> Int,
		#~ default		=> 0,
		#~ reader		=> '_min_col',
		#~ predicate	=> 'has_min_col',
	#~ );

#~ has _sheet_min_row =>(
		#~ isa			=> Int,
		#~ default		=> 0,
		#~ reader		=> '_min_row',
		#~ predicate	=> 'has_min_row',
	#~ );

#~ has _sheet_max_col =>(
		#~ isa			=> Int,
		#~ default		=> 0,
		#~ reader		=> '_max_col',
		#~ predicate	=> 'has_max_col',
	#~ );

#~ has _sheet_max_row =>(
		#~ isa			=> Int,
		#~ default		=> 0,
		#~ reader		=> '_max_row',
		#~ predicate	=> 'has_max_row',
	#~ );

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_load_unique_bits' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the Chartsheet unique bits", "Byte position: " . $self->byte_consumed ] );
	
	#collect the drawing rel_id
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the relID" ] );
	if( $self->next_element('drawing') ){
		my	$rel_id = $self->get_attribute( 'r:id' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The relID is: $rel_id", ] );
		$self->_set_drawing_rel_id( $rel_id );
	}else{
		confess "Couldn't find the drawing relID for this chart";
	}
	#~ $self->start_the_file_over;# not needed yet
	return 1;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet - A LibXML::Reader worksheet base class

=head1 SYNOPSIS

See the SYNOPSIS in L<Spreadsheet::XLSX::Reader::LibXML>
    
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

#########1 Documentation End  3#########4#########5#########6#########7#########8#########9