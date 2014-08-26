package Spreadsheet::XLSX::Reader::LibXML::Cell;
use version; our $VERSION = qv('v0.4.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		Bool
		InstanceOf
		ArrayRef
		HashRef
		HasMethods
    );

use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
with	'Spreadsheet::XLSX::Reader::LibXML::LogSpace';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has value_encoding =>(
		isa		=> Str,
		reader	=> 'encoding',
	);

has value_type =>(
		isa			=> Str,
		reader		=> 'type',
		required	=> 1,
	);

has	cell_column =>(
		isa			=> Int,
		reader		=> 'column',
		writer		=> '_set_column',
		required	=> 1,
	);

has	cell_row =>(
		isa			=> Int,
		reader		=> 'row',
		writer		=> '_set_row',
		required	=> 1,
	);

has raw_value =>(
		isa			=> Str,
		reader		=> 'unformatted',
		predicate	=> 'is_not_empty',
		required	=> 1,
	);

has cell_formula =>(
		isa			=> Str,
		reader		=> 'formula',
		predicate	=> 'has_formula',
	);

has merge_range	=>(
		isa			=> Str,
		reader		=> '_get_merge_range',
		predicate	=> 'is_merged',
	);

has rich_text =>(
		isa			=> ArrayRef,
		reader		=> 'get_rich_text',
		predicate	=> 'has_rich_text',
	);
	
has font =>(
		isa		=> HashRef,
		reader	=> 'get_font',
	);

has fill =>(
		isa		=> HashRef,
		reader	=> 'get_fill',
	);

has borderId =>(
		isa		=> Int,
		reader	=> 'get_borderId',
	);

has fillId =>(
		isa		=> Int,
		reader	=> 'get_fillId',
	);

has fontId =>(
		isa		=> Int,
		reader	=> 'get_fontId',
	);

has applyFont =>(
		isa		=> Bool,
		reader	=> 'get_applyFont',
	);

has applyNumberFormat =>(
		isa		=> Bool,
		reader	=> 'get_applyNumberFormat',
	);

has border =>(
		isa		=> HashRef,
		reader	=> 'get_border',
	);

has alignment =>(
		isa		=> HashRef,
		reader	=> 'get_alignment',
	);
	
has numFmtId =>(
		isa		=> Int,
		reader	=> 'get_numFmtId',
	);
	
has xfId =>(
		isa		=> Int,
		reader	=> 'get_xfId',
	);

has NumberFormat =>(
		isa			=> HasMethods[ 'coerce', 'display_name' ],
		reader		=> 'get_format',
		writer		=> 'set_format',
		predicate	=> 'has_format',
		clearer		=> 'clear_format',
		handles		=>{
			format_name => 'display_name',
		},
	);

has error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
		handles 	=>[ qw( error set_error clear_error set_warnings if_warn ) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_instance',
		required	=> 1,
	);
with	'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow'; #Here to load set_error first

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub value{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::value', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Attempting to return the value of the cell formatted to " .
	###LogSD			$self->format_name ] );
	###LogSD		$phone->talk( level => 'trace', message => [ "Cell:", $self ] );
	my	$formatted;
	my	$unformatted	= $self->unformatted;
	if( !$self->has_format ){
		return $unformatted;
	}elsif( !defined $unformatted ){
		$self->set_error( "The cell does not have a value" );
	}elsif( $unformatted eq '' ){
		$self->set_error( "The cell has the empty string for a value" );
		$formatted = '';
	}else{
		eval '$formatted = $self->get_format->coerce( $unformatted )';
		if( $@ ){
			$self->set_error( $@ );
		}
	}
	$formatted =~ s/\\//g;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Format is:", $self->get_format,
	###LogSD		"Returning the formated value: " . $formatted ] );
	return $formatted;
}

sub get_merge_range{
	my( $self, $modifier ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::get_merge_range', );
	if( !$self->is_merged ){
		$self->set_error( 
			"Attempted to collect merge range but the cell is not merged with any others" 
		);
		return undef;
	}
	my	$merge_range = $self->_get_merge_range;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning merge_range:  $merge_range",
	###LogSD		(( $modifier ) ? "Modified by: $modifier" : ''),
	###LogSD	] );
	if( $modifier ){
		if( $modifier eq 'array' ){
			my ( $start, $end ) = split /:/, $merge_range;
			my ( $start_col, $start_row, $end_col, $end_row ) =
				( $self->parse_column_row( $start ), $self->parse_column_row( $end ) );
			$merge_range = [ [ $start_col, $start_row ], [ $end_col, $end_row ] ];
		}else{
			$self->set_error( 
				"Un-recognized modifier -$modifier- passed to 'get_merge_range' - it only accepts 'array'" 
			);
		}
	}
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"Final merge range:", $merge_range ] );
	return $merge_range;
}

sub cell_id{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::cell_id', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting cell ID for row -" . $self->row . '- column -' . 
	###LogSD			$self->column . '-', ] );
	my $cell_id = $self->build_cell_label( $self->column, $self->row );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD			"Cell ID is: $cell_id" ] );
	return $cell_id;
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

Spreadsheet::XLSX::Reader::LibXML::Cell - A class for Cell data and formatting

=head1 SYNOPSIS

See the SYNOPSIS in L<Spreadsheet::XLSX::Reader::LibXML>
    
=head1 DESCRIPTION

This is the class that contains cell data.  There are no XML actions taken in the 
background of this class.  All data has been pre-coalated/built from the L<Workbook
|Spreadsheet::XLSX::Reader::LibXML::Workbook> class.  See the Workbook class for 
creation of this class.  Accessing the data is done through the L<Methods|/Methods>.

=head2 Attributes

Attributes of this cell are not included in the documentation because 'new' should be 
called by other classes in this package.

=head2 Methods

These are ways to access the data and formats in the cell.  They also provide a 
way to modifiy the output of the format.

=head3 unformatted

=over

B<Definition:> Returns the unformatted value of the cell in whatever encoding 
it was stored in.

B<Accepts:>Nothing

B<Returns:> the unformatted (raw) cell value in whatever encoding it was stored in

=back

=head3 value

=over

B<Definition:> Returns the formatted value of the cell. Excel really only tries to 
manipulate numbers.  If the sheet has some pre-defined number manipulation this will 
attempt to implement it prior to returning the value.  If there is no format set then 
this will return $cell->unformatted.For adjustment of the conversion method see 
L<set_format|/set_format>.  Any failures to process this value can be retrieved 
L<$self-E<gt>error|/error>.

B<Accepts:>Nothing

B<Returns:> the cell value processed by the set format

=back

=head3 encoding

=over

B<Definition:> the libxml2 library will attempt to convert everything into UTF-8 
so the output from unformatted should be in UTF-8.  However, for strings the 
encoding of the strings (sub)file sharedStrings.xml may be stored in a different 
encoding.  This method returns the registered encoding of the sharedStrings.xml 
(sub)file.

B<Accepts:> Nothing

B<Returns:> the 'encoding' attribute of the sharedStrings.xml (sub)file

=back

=head3 type

=over

B<Definition:> Excel 2007 and newer only recognizes two types of data, strings 
or numbers.  All additional formating or other manipulation is done when the 
data is presented through the format layer.  This method identifies how the 
data in this cell was stored specifically as it is in XML.  For more information 
on the format applied see L<set_format|/set_format>

B<Accepts:> Nothing

B<Returns:> (s|number) s = string

=back

=head3 column

=over

B<Definition:> This method returns the column number of the cell counting 
either from zero or from one depending on the setting from the initial parser build.  
To check the setting see L<counting_from_zero|/counting_from_zero>..

B<Accepts:> Nothing

B<Returns:> cell column number

=back

=head3 row

=over

B<Definition:> This method returns the row number of the cell counting 
either from zero or from one depending on the setting from the initial parser build.  
To check the setting see L<counting_from_zero|/counting_from_zero>..

B<Accepts:> Nothing

B<Returns:> cell row number

=back

=head3 formula

=over

B<Definition:> For cells calculated from a formula they will have both the 
formula used to create the end result and the most recent calculated result.  
This method returns the formula string from Excel used to obtain the result.  
To see the most recent calculated result see L<value|/value> or L<unformatted
|/unformatted>..

B<Accepts:> Nothing

B<Returns:> the formula used in the excel spreadsheet to calculate the cell 
value.

=back

=head3 has_formula

=over

B<Definition:> This will indicate if the cell has a L<formula|/formula>
associated with it..

B<Accepts:> Nothing

B<Returns:> the formula used in the excel spreadsheet to calculate the cell 
value.

=back

=head3 get_merge_range

=over

B<Definition:> Any cell that is merged with another cell will have a merge range.  
in excel.  Only the top left cell will actually contain the value of the merged cell.  
This follows the excel precedent.

B<Accepts:> (undef|array)

B<Returns:> If the method is called with no arguments then the merge range is 
provided in the format. "A1:D3".  If the value 'array' (and only array) is passed then 
the range is returned as an array ref in the format [[$start_column, $start_row],
[$end_column,$end_row] ]

=back

=head3 is_merged

=over

B<Definition:> This method returns a boolean value that indicates if this cell has 
been L<merged|/get_merge_range> with other cells.

B<Accepts:> Nothing

B<Returns:> $bool

=back

=head3 get_rich_text

=over

B<Definition:> This method returns a rich text data structure like the same method 
in L<Spreadsheet::ParseExcel::Cell> with the exception that it doesn't bless each 
hashref into an object.  The hashref's are also organized per the Excel xlsx 
information the the sharedStrings.xml file.  In general this is an arrayref of 
arrayrefs where the second level contains two positions.  The first position is the 
place (from zero) where the formatting is implemented.  The second position is a 
hashref of the formatting values.

B<note:> It is important to understand that Excel can store two formats for the 
same cell and often they don't agree.  For example using the L<get_font|/get_font> 
method in class will not always yield the same value as get_rich_text.

B<Accepts:> Nothing

B<Returns:> an arrayref of rich text data

=back

=head3 has_rich_text

=over

B<Definition:> This method returns a boolean value that indicates if this cell has 
L<rich text data|/get_rich_text>.

B<Accepts:> Nothing

B<Returns:> $bool

=back

=head3 get_font

=over

B<Definition:> This method returns the font assigned to this cell.

B<Accepts:> Nothing

B<Returns:> $font

=back

=head3 get_fill

=over

B<Definition:> This method returns the fill assigned to this cell.

B<Accepts:> Nothing

B<Returns:> $fill

=back

=head3 get_border

=over

B<Definition:> This method returns the border assigned to this cell.

B<Accepts:> Nothing

B<Returns:> $border

=back

=head3 get_alignment

=over

B<Definition:> This method returns the alignment assigned to this cell.

B<Accepts:> Nothing

B<Returns:> $alignment

=back

=head3 set_format

=over

B<Definition:> To set a format object it must pass two criteria.  The ref must 
have the method 'coerce' and it must have the method 'display_name'.  This is the 
object that will be used to convert the unformatted value.  For another way to 
apply formats to the cell see the L<Spreadsheet::XLSX::Reader::Worksheet> 
'custom_formats' attribute.

B<Accepts:> a ref that can 'coerce' and can 'display_name'

B<Returns:> Nothing

=back

=head3 get_format

=over

B<Definition:> When excel talks about 'format' it is closer to the perl 
function sprintf.  This returns the object used to turn the L<unformatted
|/unformtted> value into a formatted L<value|/value>.

B<Accepts:> Nothing

B<Returns:> Object instance that can 'coerce'

=back

=head3 has_format

=over

B<Definition:> This method returns a boolean value that indicates if this cell has an 
L<assigned format|/get_format>.

B<Accepts:> Nothing

B<Returns:> $bool

=back

=head3 clear_format

=over

B<Definition:> This method clears any format set for the cell.  After this action the 
L<value|/value> function will return the equvalent of L<unformatted|/unformatted>

B<Accepts:> Nothing

B<Returns:> $bool

=back

=head3 format_name

=over

B<Definition:> This method calls -E<gt>display_name on the format instance.

B<Accepts:> Nothing

B<Returns:> $display_name

=back

=head3 error

=over

B<Definition:> This method gets the latest stored string from the error in 
L<Spreadsheet::XLSX::Reader::LibXML::Error> I could change in the future but currently 
the error instance is shared across all instances of the 
L<Spreadsheet::XLSX::Reader::LibXML> classes that have been created.

B<Accepts:> Nothing

B<Returns:> $error_string

=back

=head3 set_error( $string )

=over

B<Definition:> This method sets a new error $string to the 
L<Spreadsheet::XLSX::Reader::LibXML::Error> instance.

B<Accepts:> An error string $string

B<Returns:> Nothing

=back

=head3 clear_error

=over

B<Definition:> This method clears the error $string in the 
L<Spreadsheet::XLSX::Reader::LibXML::Error> instance.

B<Accepts:> Nothing

B<Returns:> Nothing

=back

=head3 set_warnings( $bool )

=over

B<Definition:> When the error string is set / changed for the  
L<Spreadsheet::XLSX::Reader::LibXML::Error> instance the instance can cluck/warn 
the message at that time based on this setting.

B<Accepts:> A $boolean value

B<Returns:> Nothing

=back

=head3 if_warn

=over

B<Definition:> This is the predicate for the set_warnings setting.  It will 
show whether warnings are turned on or off.

B<Accepts:> Nothing

B<Returns:> A $boolean value

=back

=head3 is_not_empty

=over

B<Definition:> This is a predicate method to tell if the cell has an unformatted value

B<Accepts:> Nothing

B<Returns:> a boolean value to indicate if the cell is empty or not

=back

=head3 cell_id

=over

B<Definition:> This returns the excel alphanumeric identifier for the 
cell location.

B<Accepts:> Nothing

B<Returns:> cell position ex. B15

=back

=head3 get_log_space

=over

B<Definition:> This returns the stored log space for this module

B<Accepts:> Nothing

B<Returns:> $log_space_string

=back

=head3 set_log_space

=over

B<Definition:> This changes the stored log space for this module

B<Accepts:> $log_space_string

B<Returns:> Nothing

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues
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

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Types::Standard>

L<Spreadsheet::XLSX::Reader::LibXML::LogSpace>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::ParseExcel::Cell>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9