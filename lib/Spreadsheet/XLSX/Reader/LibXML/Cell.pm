package Spreadsheet::XLSX::Reader::LibXML::Cell;
use version 0.77; our $VERSION = qv('v0.34.2');

$| = 1;
use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Str					InstanceOf				HashRef
		Enum				HasMethods				ArrayRef
		Int					Maybe					CodeRef
		is_Object
    );
use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
###LogSD	with 'Log::Shiras::LogSpace';
use	Spreadsheet::XLSX::Reader::LibXML::Types qw(
		CellID
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has	error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);

has cell_unformatted =>(
		isa			=> Maybe[Str],
		reader		=> '_unformatted',
		predicate	=> 'has_unformatted',
		#~ default		=> '',
	);

has rich_text =>(
		isa		=> ArrayRef,
		reader	=> 'get_rich_text',
		predicate	=> 'has_rich_text',
	);

has cell_font =>(
		isa		=> HashRef,
		reader	=> 'get_font',
		predicate	=> 'has_font',
	);

has cell_border =>(
		isa		=> HashRef,
		reader	=> 'get_border',
		predicate	=> 'has_border',
	);
	
has cell_style =>(
		isa		=> HashRef,
		reader	=> 'get_style',
		predicate	=> 'has_style',
	);
	
has cell_fill =>(
		isa		=> HashRef,
		reader	=> 'get_fill',
		predicate	=> 'has_fill',
	);

has cell_type =>(
		isa		=> Enum[qw( Text Numeric Date Custom )],
		reader	=> 'type',
		writer	=> '_set_cell_type',
		predicate	=> 'has_type',
	);

has cell_encoding =>(
		isa		=> Str,
		reader	=> 'encoding',
		predicate	=> 'has_encoding',
	);

has cell_merge =>(
		isa			=> Str,
		reader		=> 'merge_range',
		predicate 	=> 'is_merged',
	);

has cell_formula =>(
		isa			=> Str,
		reader		=> 'formula',
		predicate	=> 'has_formula',
	);
	
has cell_row =>(
		isa			=> Int,
		reader		=> 'row',
		predicate	=> 'has_row',
	);
	
has cell_col =>(
		isa			=> Int,
		reader		=> 'col',
		predicate	=> 'has_col',
	);

has r =>(
		isa		=> CellID,
		reader	=> 'cell_id',
		predicate	=> 'has_cell_id',
	);

has cell_hyperlink =>(
		isa		=> ArrayRef,
		reader	=> 'get_hyperlink',
		predicate	=> 'has_hyperlink',
	);

has unformatted_converter =>(
		isa			=> CodeRef,
		reader		=> '_convert_output',
		required	=> 1,
	);

has cell_coercion =>(
		isa			=> HasMethods[ 'assert_coerce', 'display_name' ],
		reader		=> 'get_coercion',
		writer		=> 'set_coercion',
		predicate	=> 'has_coercion',
		clearer		=> 'clear_coercion',
		handles		=>{
			coercion_name => 'display_name',
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub unformatted{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::unformatted', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Attempting to return the unformatted value of the cell" ] );
	###LogSD		$phone->talk( level => 'trace', message => [ "Cell:", $self ] );
	
	# check if empty
	return undef if !$self->has_unformatted;
	
	# get the value
	my	$unformatted	= $self->_unformatted;
	###LogSD	$phone->talk( level => 'debug', message => [ "unformatted:", $unformatted ] );
	my	$converter = $self->_convert_output;
	$unformatted = $converter->( $unformatted );
	###LogSD	$phone->talk( level => 'debug', message => [ "converted:", $unformatted ] );
	return $unformatted;
}

sub value{
	my( $self, ) 	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::value', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached the -value- function' ] );
	###LogSD		$phone->talk( level => 'trace', message => [ "Cell:", $self ] );
	my	$unformatted = $self->unformatted;
	my	$formatted = $unformatted;
	if( !$self->has_coercion ){
		return $unformatted;
	}elsif( !defined $unformatted ){
		$self->set_error( "The cell does not have a value" );
	}elsif( $unformatted eq '' ){
		$self->set_error( "The cell has the empty string for a value" );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Attempting to return the value of the cell formatted to " .
		###LogSD		(($self->has_coercion) ? $self->coercion_name : 'No conversion available' ) ] );
		eval '$formatted = $self->get_coercion->assert_coerce( $unformatted )';
		$self->set_error( $@ ) if( $@ );
	}
	$formatted =~ s/\\//g if $formatted;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Format is:", $self->coercion_name,
	###LogSD		"Returning the formated value: " . 
	###LogSD		( $formatted ? $formatted : '' ), ] );
	return $formatted;
}

#~ sub get_merge_range{
	#~ my( $self, $modifier ) 	= @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> $self->get_log_space .  '::get_merge_range', );
	#~ if( !$self->is_merged ){
		#~ $self->set_error( 
			#~ "Attempted to collect merge range but the cell is not merged with any others" 
		#~ );
		#~ return undef;
	#~ }
	#~ my	$merge_range = $self->merge_range;
	#~ ###LogSD	$phone->talk( level => 'debug', message => [
	#~ ###LogSD		"Returning merge_range:  $merge_range",
	#~ ###LogSD		(( $modifier ) ? "Modified by: $modifier" : ''),
	#~ ###LogSD	] );
	#~ if( $modifier ){
		#~ if( $modifier eq 'array' ){
			#~ my ( $start, $end ) = split /:/, $merge_range;
			#~ my ( $start_col, $start_row, $end_col, $end_row ) =
				#~ ( $self->parse_column_row( $start ), $self->parse_column_row( $end ) );
			#~ $merge_range = [ [ $start_col, $start_row ], [ $end_col, $end_row ] ];
		#~ }else{
			#~ $self->set_error( 
				#~ "Un-recognized modifier -$modifier- passed to 'get_merge_range' - it only accepts 'array'" 
			#~ );
		#~ }
	#~ }
	#~ ###LogSD	$phone->talk( level => 'info', message => [
	#~ ###LogSD		"Final merge range:", $merge_range ] );
	#~ return $merge_range;
#~ }

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

after 'set_coercion' => sub{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::set_coercion' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Setting 'cell_type' to custom since the coercion has been set" ] );
	$self->_set_cell_type( 'Custom' );
};

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Cell::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the cell for cell ID:" . $self->cell_id, ] );
	#~ print "Clearing coercion\n";
	$self->clear_coercion;
	#~ print "Clearing error instance\n";
	$self->_clear_error_inst;
}



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Cell - XLSX Cell data class

=head1 SYNOPSIS

See the SYNOPSIS in the L<Workbook class|Spreadsheet::XLSX::Reader::LibXML/SYNOPSIS>
    
=head1 DESCRIPTION

This is the class that contains cell data.  There are no XML parsing actions taken in the 
background of this class.  All data has been pre-coalated/built from the L<Worksheet
|Spreadsheet::XLSX::Reader::LibXML::Worksheet> class.  In general the Worksheet class 
will populate the attributes of this class when it is generated.  If you want to use it 
as a standalone class just fill in the L<Attributes|/Attributes> below.

=head2 Primary Methods

These are methods used to transform data stored in the L<Attributes|/Attributes> 
(not just return it directly).  All methods are object methods and should be implemented 
on the instance.

B<Example:>

	my $unformatted_value = $cell_intance->unformatted;

=head3 unformatted

=over

B<Definition:> Returns the unformatted value of the cell transformed with the 
L<change_output_encoding|Spreadsheet::XLSX::Reader::LibXML::FmtDefault/change_output_encoding( $string )> 
method.

B<Accepts:>Nothing

B<Returns:> the cell value processed by the encoding conversion

=back

=head3 has_unformatted

=over

B<Definition:> This is a predicate method to determine if the cell had any value stored in it.  
Sometimes this class will be generated by the 
L<Worksheet|Spreadsheet::XLSX::Reader::LibXML::Worksheet> class when there is cell formatting but 
no value.  Ex. Merged cells store the value in the left upper corner of the merged area but have 
cell specific formatting for all cells in the merge area.

B<Accepts:>Nothing

B<Returns:> True if the cell holds an unformatted value (even if it is just a space or empty string)

=back

=head3 value

=over

B<Definition:> Returns the formatted value of the cell transformed from the 
L<unformatted|/unformatted> string. This method uses the conversion stored in the
L<cell_coercion|/cell_coercion> attribute.  If there is no format/conversion set 
then this will return the unformatted value. Any failures to process this value can be 
retrieved with L<$self-E<gt>error|/error>.

B<Accepts:>Nothing

B<Returns:> the cell value processed by the set conversion

=back

=head2 Attributes

This class is just a storage of coallated information about the requested cell stored 
in the following attributes. For more information on attributes see 
L<Moose::Manual::Attributes>.  The meta data about the cell can be retrieved from each 
attribute using the 'attribute methods'.

=head3 error_inst

=over

B<Definition:> This attribute holds an 'error' object instance.  In general 
the package shares a reference for this instance accross the workbook with all 
worksheets and all cells so any set or get action should return the latest error state 
from anywhere. (not just the instance you are working on)

B<Default:> a L<Spreadsheet::XLSX::Reader::LibXML::Error> instance with the 
attributes set as;
	
	( should_warn => 0 )

B<Range:> The minimum list of methods to implement for your own instance is;

	error set_error clear_error set_warnings if_warn

B<attribute methods> Methods provided to adjust this attribute

=over

B<get_error_inst>

=over

B<Definition:> returns this instance

=back

B<error>

=over

B<Definition:> Used to get the most recently logged error

=back

B<set_error>

=over

B<Definition:> used to set a new error string
		
=back

B<clear_error>

=over

B<Definition:> used to clear the current error string in this attribute
		
=back

B<set_warnings>

=over

B<Definition:> used to turn on or off real time warnings when errors are set.  
This is a delegated method from the error instance.
		
=back

B<if_warn>

=over

B<Definition:> a method mostly used to extend this package and see if warnings 
should be emitted.
		
=back

=back

=back

=head3 cell_unformatted

=over

B<Definition:> This holds the unformatted value of the cell (if any)

B<Default:> undef

B<Range:> a string

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<unformatted>

=over

B<Definition:> returns the attribute value
		
=back
		
=over

B<has_unformatted>

=over

B<Definition:> a predicate method to determine if any value is in the cell
		
=back

=back

=back

=back

=head3 rich_text

=over

B<Definition:> This attribute hold a rich text data structure like 
L<Spreadsheet::ParseExcel::Cell/get_rich_text()> with the exception that it 
doesn't bless each hashref into an object.  The hashref's are also organized 
per the Excel xlsx information in the the sharedStrings.xml file.  In general 
this is an arrayref of arrayrefs where the second level contains two positions.  
The first position is the place (from zero) where the formatting is implemented.  
The second position is a hashref of the formatting values.  The format is inforce 
until the next start place is identified.

=over

B<note:> It is important to understand that Excel can store two formats for the 
same cell and often they don't agree.  For example using the attribute L<cell_font
|/cell_font> will not always contain the same value as specific fonts (or any font) 
listed in the rich text array.
f
=back

B<Default:> undef = no rich text defined for this cell

B<Range:> an array ref of rich_text positions and definitions

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_rich_text>

=over

B<Definition:> returns the attribute value
		
=back

B<has_rich_text>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back

=back

=back

=back

=head3 cell_font

=over

B<Definition:> This holds the font assigned to the cell

B<Default:> undef

B<Range:> a hashref of definitions for the font

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_font>

=over

B<Definition:> returns the attribute value
		
=back

B<has_font>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back

=back

=back

=head3 cell_border

=over

B<Definition:> This holds the border settings assigned to the cell

B<Default:> undef

B<Range:> a hashref of border definitions

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_border>

=over

B<Definition:> returns the attribute value
		
=back

B<has_border>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back

=back

=back

=head3 cell_style

=over

B<Definition:> This holds the border settings assigned to the cell

B<Default:> undef

B<Range:> a hashref of style definitions

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_style>

=over

B<Definition:> returns the attribute value
		
=back

B<has_style>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_fill

=over

B<Definition:> This holds the fill settings assigned to the cell

B<Default:> undef

B<Range:> a hashref of style definitions

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_fill>

=over

B<Definition:> returns the attribute value
		
=back

B<has_fill>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_type

=over

B<Definition:> This holds the type of data stored in the cell.  In general it 
follows the convention of L<ParseExcel
|Spreadsheet::ParseExcel/ChkType($self, $is_numeric, $format_index)> (Date, Numeric, 
or Text) however, since custom coercions will change data to some possible non excel 
standard state this also allows a 'Custom' type representing any cell with a custom 
conversion assigned to it (by you either at the worksheet level or here).

B<Default:> Text

B<Range:> Text = Strings, Numeric = Real Numbers, Date = Real Numbers with an 
assigned Date conversion, Custom = any stored value with a custom conversion

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<type>

=over

B<Definition:> returns the attribute value
		
=back

B<has_type>

=over

B<Definition:> Indicates if the attribute has anything stored (Always true)
		
=back
		
=back
		
=back

=head3 cell_encoding

=over

B<Definition:> This holds the byte encodeing of the data stored in the cell

B<Default:> Unicode

B<Range:> Traditional encoding options

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<encoding>

=over

B<Definition:> returns the attribute value
		
=back

B<has_encoding>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_merge

=over

B<Definition:> if the cell is part of a group of merged cells this will 
store the upper left and lower right cell ID's in a string concatenated 
with a ':'

B<Default:> undef

B<Range:> two cell ID's

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<merge_range>

=over

B<Definition:> returns the attribute value
		
=back

B<is_merged>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_formula

=over

B<Definition:> if the cell value (unformatted) is calculated based on a 
formula the Excel formula string is stored in this attribute.

B<Default:> undef

B<Range:> Excel formula string

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<formula>

=over

B<Definition:> returns the attribute value
		
=back

B<has_formula>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_row

=over

B<Definition:> This is the sheet row that the cell was read from

B<Range:> the minimum row to the maximum row

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<row>

=over

B<Definition:> returns the attribute value
		
=back

B<has_row>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_col

=over

B<Definition:> This is the sheet column that the cell was read from

B<Range:> the minimum column to the maximum column

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<col>

=over

B<Definition:> returns the attribute value
		
=back

B<has_col>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 r

=over

B<Definition:> This is the cell ID of the cell

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<cell_id>

=over

B<Definition:> returns the attribute value
		
=back

B<has_cell_id>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_hyperlink

=over

B<Definition:> This stores any hyperlink from the cell

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_hyperlink>

=over

B<Definition:> returns the attribute value
		
=back

B<has_hyperlink>

=over

B<Definition:> Indicates if the attribute has anything stored
		
=back
		
=back
		
=back

=head3 cell_coercion

=over

B<Definition:> This attribute holds the tranformation code to turn an 
unformatted  value into a formatted value.

B<Default:> a L<Type::Tiny> instance with sub types set to assign different 
inbound data types to different coercions for the target outcome of formatted 
data.

B<Range:> If you wish to set this with your own code it must have two 
methods.  First, 'assert_coerce' which will be applied when transforming 
the unformatted value.  Second, 'display_name' which will be used to self 
identify.  For an example of how to build a custom format see 
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet/custom_formats>.

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_coercion>

=over

B<Definition:> returns this instance
		
=back

B<clear_coercion>

=over

B<Definition:> used to clear the this attribute
		
=back

B<set_coercion>

=over

B<Definition:> used to set a new coercion instance.  Implementation of this 
method will also switch the cell type to 'Custom'.
		
=back

B<has_coercion>

=over

B<Definition:> Indicate if any coecion code is applied
		
=back

B<coercion_name>

=over

B<Definition:> calls 'display_name' on the code in the background
		
=back
		
=back
		
=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Return the merge range in array and hash formats

B<2.> Add calc chain values

B<3.> Have unformatted return '' (the empty string) rather than undef for null?

=back

=head1 AUTHOR

=over

Jed Lund

jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014 by Jed 

=head1 DEPENDENCIES

=over

L<version> 0.77

L<perl 5.010|perl/5.10.0>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Type::Tiny> - 1.000>

L<Spreadsheet::XLSX::Reader::LibXML::Types>

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

#########1#########2 main pod documentation end  5#########6#########7#########8#########9