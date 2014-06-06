package Spreadsheet::XLSX::Reader::XMLDOM::Cell;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		InstanceOf
    );

use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
with 'Spreadsheet::XLSX::Reader::LogSpace';
use		Spreadsheet::XLSX::Reader::Types v0.1 qw( NumberFormat );

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has	error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::Error' ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);
with 'Spreadsheet::XLSX::Reader::CellToColumnRow'; #here to load 'set_error' first

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
		isa		=> Int,
		reader	=> 'column',
		writer	=> '_set_column',
	);

has	cell_row	=>(
		isa		=> Int,
		reader	=> 'row',
		writer	=> '_set_row',
	);

has number_format =>(
		isa			=> NumberFormat,
		writer		=> 'set_format',
		reader		=> 'get_format',
		clearer		=> 'clear_format',
		predicate	=> 'has_format',
		handles	=>{
			format_name => 'display_name',
		},
	);
	
has cell_element =>(
		isa			=> InstanceOf[ 'XML::LibXML::Element' ],
		clearer		=> '_clear_cell_element',
		writer		=> '_set_cell_element',
		reader		=> 'get_xml',
		trigger		=> \&_set_chunk_data,
		handles	=>{
			get_cell_position 	=>[ getAttribute => 'r' ],
			#~ has_text_attribute	=>[ hasAttribute => 't' ],
			#~ get_text_attribute	=>[ getAttribute => 't' ],
			#~ has_a_format		=>[ hasAttribute => 's' ],
			#~ get_format_position	=>[ getAttribute => 's' ],
			unformatted			=> 'textContent',
			_get_formula_nodes	=>[ getChildrenByTagName => 'f' ],
			_get_v_nodes		=>[ getChildrenByTagName => 'v' ],
			_get_merge_nodes	=>[ getChildrenByTagName => 'mergeCell' ],
			_replace_child		=> 'replaceChild',
			_remove_child		=> 'removeChild',
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub value{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::value', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Attempting to return the value of the cell formatted as needed" ] );
	###LogSD		$phone->talk( level => 'debug', message => [ "Cell:", $self ] );
	
	my	$unformatted = $self->unformatted;
	if( !$self->has_format ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"No format to apply returning: $unformatted" ] );
		return $unformatted;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Format is:", $self->get_format,
	###LogSD		"Returning the formated value: " . $self->get_format->{translation}->( $unformatted ) ] );
	return $self->get_format->{translation}->( $unformatted );
}

sub formula{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::formula', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Looking for any formula to calculate the cell in the worksheet" ] );
	
	my	$chunk_ref 	= ($self->_get_formula_nodes)[0];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The formula node contains:", $chunk_ref->textContent, $chunk_ref ] );
	my	$formula	= $chunk_ref->textContent;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The formula node contains:", $chunk_ref->textContent, $chunk_ref ] );
	return $formula;
}

sub is_merged{
	my( $self, $return_type ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::is_merged', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Checking if the cell is merged with other cells" .
	###LogSD			(( !$return_type ) ? '' : " - returning type: $return_type" ) ] );
	
	my ( $top_left, $bottom_right ) = split( /:/, ($self->_get_merge_nodes)[0] );
	###LogSD	no warnings 'uninitialized';	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The merge node contains: ( $top_left, $bottom_right )" ] );
	if( !$top_left ){	
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"This cell is not merged with another" ] );
		return undef;
	}elsif( !$return_type or $return_type eq 'position' ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"This cell is within the merged range of( $top_left, $bottom_right )" ] );
		return( $top_left, $bottom_right );
	}elsif( $return_type eq 'array' ){
		$top_left 		= $self->parse_column_row( $top_left );
		$bottom_right	= $self->parse_column_row( $bottom_right );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"This cell is within the merged range of ( $top_left, $bottom_right )" ] );
		return( $top_left, $bottom_right );
	}
	$self->_set_error( "is_merged was called with -$return_type- but only accepts position or array" );
	return undef;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _set_chunk_data{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_set_chunk_data', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Pulling the defined worksheet position from the XML" ] );
	
	# Get the cell position
	my ( $column, $row ) = $self->parse_column_row( $self->get_cell_position );
	if( $self->column and $self->column != $column ){
		$self->_set_error(
			"XML column -$column- does not agree with the instance column set: " . $self->column );
		$self->clear_cell_element;
	}else{
		$self->_set_column( $column );
	}
	if( $self->row and $self->row != $row ){
		$self->_set_error(
			"XML row -$row- does not agree with the instance row set: " . $self->row );
		$self->clear_cell_element;
	}else{
		$self->_set_row( $row );
	}
	my ( $node ) = $self->_get_formula_nodes;
	if( $node ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The first formula node is: " . $node->textContent ] );
		$self->_set_formula( $node->textContent );
		$self->_remove_child( $node );
	}
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD				"Clearing the open cell element and disconnecting " . 
	###LogSD				"the workbook link" ] );
	$self->_clear_cell_element;
	$self->_clear_error_inst;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable(
	inline_constructor => 0,
);
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Excel::Reader::XLSX::Shiras::DOM::Styles - LibXML DOM parser of Styles
    
=head1 DESCRIPTION

This is the class that contains cell data.  The Worksheet and L<Workbook
|Spreadsheet::XLSX::Reader> classes work together to extract the relevant information 
for this cell and build an XML data set representing the cell contents and formatting.  
The information in the cell can be extracted using the L</Methods> outlined below.

=head2 Caveat utilitor

This class is not designed to be used by itself and requires hooks back into the 
L<Workbook|Spreadsheet::XLSX::Reader> class.  The information coallated by the 
L<Worksheet|Spreadsheet::XLSX::Reader::XMLReader::Worksheet> class and placed in an 
instance of the cell does not exist in one location for most .xlsx files.  Rather,
to build this cell instance the the Worksheet file must coalate formats, styles, and 
contents from several locations within the xlsx file.

=head2 Attributes

All attributes of this cell are private because 'new' should be L<called
|/Caveat utilitor> by other classes in this package.

=head2 Methods

These are ways to access the data and formats in the cell.  They also provide a 
way to modifiy the output of the format.  They are used as follows;

	$cell_instance->$method( @arg_list );

=head3 value

=over

B<Definition:> Returns the formatted value of the cell.  Note that to get the 
formatted value of the cell the unformatted content is transformed by the 
_content_conversion attribute.  The worksheet file will attempt to set this 
automatically based on the information that it already has using the predefined 
conversions available in L<Spreadsheet::XLSX::Reader::Types>.  If you want to 
define your own formatting then you need to set it using L</set_format>.  
If there is no format set then this will return $cell->unformatted.

B<Accepts:>Nothing

B<Returns:> the cell value processed by the '_content_conversion' action

=back

=head3 unformatted

=over

B<Definition:> Returns the unformatted value of the cell in whatever encoding 
it was stored in.

B<Accepts:>Nothing

B<Returns:> the unformatted cell value in whatever encoding it was stored in

=back

=head3 get_format

=over

B<Definition:> Returns the L<Type::Coercion> instance stored for use with the 
cell.  I<This is different than L<Spreadsheet::ParseExcel> which uses 
L<Spreadsheet::ParseExcel::Format> >

B<Accepts:>Nothing

B<Returns:> the stored formatting instance for this cell which is a 
L<Type::Coercion> class

=back

=head3 set_format( InstanceOf[ 'Types::Coercion' ] )

=over

B<Definition:> This is how you can set a user-defined format conversion.  
It must be a L<Type::Coercion> instance.  The instance will be used internally 
for the L</value> command as follows;

	$self->get_format->( $self->unformatted );
	
If you want to use pre-built formats then see L<Spreadsheet::XLSX::Reader::Types> 
for available options.

B<Accepts:> a L<Type::Coercion> instance

B<Returns:> Nothing

=back

=head3 has_format

=over

B<Definition:> This is a check to see if the format is set for the cell

B<Accepts:> Nothing

B<Returns:> $bool a true/false value to say if the format is set

=back

=head3 format_name

=over

B<Definition:> This just exposes the display_name method of the L<Type::Coercion> 
instance

B<Accepts:> Nothing

B<Returns:> $display_name of the stored coercion

=back

=head3 type

=over

B<Definition:> Excel 2007 and new only recognizes two types of data, strings 
or numbers.  All additional formating or other manipulation is done when the 
data is presented through the format layer.  This method identifies how the 
data in this cell was stored.  For more information on the format applied 
see L</get_format>

B<Accepts:> Nothing

B<Returns:> (string|number)

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

=head3 formula

=over

B<Definition:> For cells calculated from a formula they will have both the 
formula used to create the end result and the most recent calculated result.  
This method returns the formula used to obtain the result.  To see the most 
recent calculated result see L</value>.

B<Accepts:> Nothing

B<Returns:> the formula used in the excel spreadsheet to calculate the 
L</value>

=back

=head3 is_not_empty

=over

B<Definition:> This is a predicate method to tell if the cell has a value

B<Accepts:> Nothing

B<Returns:> a boolean value to indicate if the cell is empty or not

=back

=head3 get_cell_position

=over

B<Definition:> This returns the excel alphanumeric identifier for the 
cell location.  This is what the cell thinks it is.

B<Accepts:> Nothing

B<Returns:> cell position ex. B15

=back

=head3 column

=over

B<Definition:> This method returns the column number of the cell counting 
from one.  This is what the worksheet thinks it is.

B<Accepts:> Nothing

B<Returns:> cell column number

=back

=head3 row

=over

B<Definition:> This method returns the row number of the cell counting 
from one.  This is what the worksheet thinks it is.

B<Accepts:> Nothing

B<Returns:> cell row number

=back

=head3 get_xml

=over

B<Definition:> This method returns an XML::LibXML::Element instance containing 
information about the cell.  It includes all the by-position formatting provided 
by the 'get_rich_text' method found in L<Spreadsheet::ParseExcel>.  B<Since excel 
xlsx spreadsheets are not stored in one complete subfile in the master file this 
xml will not represent the actual organization of the xml stored in the workbook.>  T
his xml node is a coalated set of values that represent cell values as seen by a 
user of Microsoft Excel.

B<Accepts:> Nothing

B<Returns:> an L<XML::LibXML::Element>

=back

=head3 is_merged( $return_type )

=over

B<Definition:> This method returns undef if the cell is not merged.  If the 
cell is merged it returns an array of the top left cell and the bottom right 
cell.  These can either be in Excel L<cell position|/get_cell_position> format or 
two array refs of column, row integer lists.  The return type can be set by 
sending undef or position to receive the list of cell positions or you can send 
array to receive the array ref of integers.

B<Accepts:> (undef|position|array)

B<Returns:> [ 'Top Left Position', 'Bottom Right Position' ]

=back

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

L<Types::Standard>

L<Spreadsheet::XLSX::Reader::LogSpace>

L<Spreadsheet::XLSX::Reader::CellToColumnRow>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9