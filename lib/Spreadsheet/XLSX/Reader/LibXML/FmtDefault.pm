package Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use version; our $VERSION = qv('v0.38.6');
###LogSD	warn "You uncovered logging statements for Spreadsheet::XLSX::Reader::LibXML::FmtDefault-$VERSION";

use	5.010;
use	Moose;
use	Carp 'confess';
use	Encode qw(decode);
use Types::Standard qw( HashRef is_ArrayRef is_HashRef	Str is_StrictNum HasMethods Bool );
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
###LogSD	with 'Log::Shiras::LogSpace';

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has error_inst =>(
		isa	=> 	HasMethods[qw(
							error set_error clear_error
						) ],
		handles =>[ qw(
			error set_error clear_error
		) ],
		writer	=> 'set_error_inst',
	);

has excel_region =>(
		isa		=> Str,
		default	=> 'en',
		reader	=> 'get_excel_region',
		writer	=> 'set_excel_region',
	);
with 'Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings';#<--- NOTE PLACED HERE FOR PREREQS
	
has	target_encoding =>(
		isa			=> Str,
		reader		=> 'get_target_encoding',
		writer		=> 'set_target_encoding',
		predicate	=> 'has_target_encoding',
	);

has dont_inherit =>(
		isa		=> Bool,
		reader	=> 'block_inherit',
		default => 0,
	);

has defined_excel_translations =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		default	=> sub{ {
			0x00 => 'General',
			0x01 => '0',
			0x02 => '0.00',
			0x03 => '#,##0',
			0x04 => '#,##0.00',
			0x05 => '$#,##0_);($#,##0)',
			0x06 => '$#,##0_);[Red]($#,##0)',
			0x07 => '$#,##0.00_);($#,##0.00)',
			0x08 => '$#,##0.00_);[Red]($#,##0.00)',
			0x09 => '0%',
			0x0A => '0.00%',
			0x0B => '0.00E+00',
			0x0C => '# ?/?',
			0x0D => '# ??/??',
			0x0E => 'yyyy-mm-dd',      # Was 'm-d-yy', which is bad as system default
			0x0F => 'd-mmm-yy',
			0x10 => 'd-mmm',
			0x11 => 'mmm-yy',
			0x12 => 'h:mm AM/PM',
			0x13 => 'h:mm:ss AM/PM',
			0x14 => 'h:mm',
			0x15 => 'h:mm:ss',
			0x16 => 'm-d-yy h:mm',
			0x1F => '#,##0_);(#,##0)',
			0x20 => '#,##0_);[Red](#,##0)',
			0x21 => '#,##0.00_);(#,##0.00)',
			0x22 => '#,##0.00_);[Red](#,##0.00)',
			0x23 => '_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
			0x24 => '_($*#,##0_);_($*(#,##0);_($*"-"_);_(@_)',
			0x25 => '_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
			0x26 => '_($*#,##0.00_);_($*(#,##0.00);_($*"-"??_);_(@_)',
			0x27 => 'mm:ss',
			0x28 => '[h]:mm:ss',
			0x29 => 'mm:ss.0',
			0x2A => '##0.0E+0',
			0x2B => '@',
			0x31 => '@',
		} },
		handles =>{
			_get_defined_excel_format => 'get',
			_set_defined_excel_format => 'set',
			total_defined_excel_formats	=> 'count',
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub	get_defined_excel_format{
	my ( $self, $position, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::get_defined_excel_format', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Getting the defined excel format for position: $position", ] );
	my	$int_value = ( $position =~ /0x/ ) ? hex( $position ) : $position;
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"..after int conversion: $int_value", ] );
	return $self->_get_defined_excel_format( $int_value );
}

sub	set_defined_excel_formats{
	my ( $self, @args, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::set_defined_excel_format', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Setting the defined excel format for the elements in ref: ", @args, ] );
	my $position_ref;
	if( @args > 1 and @args % 2 == 0 ){
		$position_ref = { @args };
	}else{
		$position_ref = $args[0];
	}
	if( is_ArrayRef( $position_ref ) ){
		my $x = -1;
		for my $format_string ( @$position_ref ){
			$x++;
			next if !$format_string;
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Setting position -$x- to format string: $format_string", ] );
			$self->_set_defined_excel_format( $x => $format_string );
		}
	}elsif( is_HashRef( $position_ref ) ){
		for my $key ( keys %$position_ref ){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD			"Setting the defined excel format for position -$key- to : ", $position_ref->{$key}, ] );
			my	$int_value = ( $key =~ /0x/ ) ? hex( $key ) : $key;
			confess "The key -$key- must translate to a number!" if !is_StrictNum( $int_value );
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Initial -$key- translated to position: " . $int_value, ] );
			$self->_set_defined_excel_format( $int_value => $position_ref->{$key} );
		}
	}else{
		confess "Unrecognized format passed: " . join( '~|~', @$position_ref );
	}
	return 1;
}

sub	change_output_encoding{
	my ( $self, $string, ) = @_;
	return undef if !$string;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::change_output_encoding', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Changing the encoding of: $string",
	###LogSD				($self->has_target_encoding ? ('..to encoding type: ' . $self->get_target_encoding) : ''), ] );
	my $output = $self->has_target_encoding ? decode( $self->get_target_encoding, $string ) : $string;
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Final output: $output", ] );
	return $output;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

 #change_output_encoding( $string ) to TextFmt

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::FmtDefault - Default xlsx number formats and localization

=head1 PACKAGE SYNOPSIS

    #!/usr/bin/env perl
	use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
	use	Spreadsheet::XLSX::Reader::LibXML;
	my	$formatter = 	Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
								target_encoding => 'latin1',
								datetime_dates	=> 1,
								dont_inherit	=> 1,
							);
		$formatter->set_defined_excel_formats( 0x2C => 'MyCoolFormatHere' );
	my	$parser	= Spreadsheet::XLSX::Reader::LibXML->new;
	my	$workbook = $parser->parse( $file, $formatter );
		$workbook = Spreadsheet::XLSX::Reader::LibXML->new(# This is an alternate way
						file_name		=> $file,
						formatter_inst	=> $formatter,
					);
    
=head1 PACKAGE DESCRIPTION

If you wish to set attributes for the formatter class that are fixed and not inherited by the package 
then you can build the formatter class with the attribute L<dont_inherit|/dont_inherit> set to 1 and then 
pass that instance as either the second argument to the 
L<Spreadsheet::XLSX::Reader::LibXML/parse( $file_nameE<verbar>$file_handle, $formatter )> command or to the 
attribute L<Spreadsheet::XLSX::Reader::LibXML/formatter_inst> when calling new.  The following formatter 
attributes are inherited if 'dont_inherit' is not set.  It should be noted that datetime_dates can be 
set via a method L<inherited by the workbook|Spreadsheet::XLSX::Reader::LibXML/set_date_behavior>

	package                 => formatter
	################           ################
	error_inst              => error_inst
	get_epoch_year (method) => epoch_year
	cache_positions         => cache_formats
	
If you want to modify the default output formats you can do that here and they won't be 
overwritten even if 'dont_inherit => 0'.
	

=head1 CLASS SYNOPSIS

	#!/usr/bin/env perl
	use		Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
	my		$formatter = Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new;
	my 		$excel_format_string = $formatter->get_defined_excel_format( 0x0E );
	print 	$excel_format_string . "\n";
			$excel_format_string = $formatter->get_defined_excel_format( '0x0E' );
	print 	$excel_format_string . "\n";
			$excel_format_string = $formatter->get_defined_excel_format( 14 );
	print	$excel_format_string . "\n";
			$formatter->set_defined_excel_formats( '0x17' => 'MySpecialFormat' );#Won't really translate!
			$excel_format_string = $formatter->get_defined_excel_format( 23 );
	print 	$excel_format_string . "\n";
	my		$conversion	= $formatter->parse_excel_format_string( '[$-409]dddd, mmmm dd, yyyy;@' );
	print 	'For conversion named: ' . $conversion->name . "\n";
	for my	$unformatted_value ( '7/4/1776 11:00.234 AM', 0.112311 ){
		print "Unformatted value: $unformatted_value\n";
		print "..coerces to: " . $conversion->assert_coerce( $unformatted_value ) . "\n";
	}

	###########################
	# SYNOPSIS Screen Output
	# 01: yyyy-mm-dd
	# 02: yyyy-mm-dd
	# 03: yyyy-mm-dd
	# 04: MySpecialFormat	
	# 05: For conversion named: DATESTRING_0
	# 06: Unformatted value: 7/4/1776 11:00.234 AM
	# 07: ..coerces to: Thursday, July 04, 1776
	# 08: Unformatted value: 0.112311
	# 09: ..coerces to: Friday, January 01, 1904
	###########################
    
=head1 CLASS DESCRIPTION

This documentation is written to explain the lesser used options and features of this class.  
The general use of the main package is explained in the documentation for 
L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>.  In general replacement of this class 
requires replacement of the methods and attributes documented here and in 
L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings>.  The private methods 
and attributes here (if any) are up to you.  It is possible to replace only this class 
and use the role ~::ParseExcelFormatStrings with it.  Additionally while this class has 
been changed from a role to better match the flow of L<Spreadsheet::ParseExcel> The actual 
implementation is very different since the underlying architecture is different as 
well.  (The ~::ParseExcel equivalent is therefore not interchangeable)

This class is the tool for number and string localization.  It stores the number conversion 
format strings and the code of the defined region.  In this particular case this module is 
set for the base L<english conversion
|http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/02/16/dates-in-spreadsheetml.aspx> 
set.  It does rely on L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings> 
to build the actual coercions used to transform numbers for each format string.  However, 
the ParseExcelFormatStrings transformation should work for all region specific strings.  
When changing the number formats (output) for this class you could just set a different 
L<hash ref|/defined_excel_translations> when calling a new or update the values with the 
method L<set_defined_excel_formats|/set_defined_excel_formats( %args )>.  This package also 
uses L<Encode> to provide for encoding changes for strings.  In general the package 
L<XML::LibXML> pulls the defined encoding for the Excel file from the underlying XML encoding 
and auto parses it to perl coding for storage.  This package does provide a way to export the 
data out to your L<target_encoding|/target_encoding>.
	
=head2 Primary Methods

These are the primary ways to use this class.  For additional FmtDefault options see the 
L<Attributes|/Attributes> section.

=head3 get_defined_excel_format( $position )

=over

B<Definition:> This will return the preset excel format string for the stored position. 
The positions are actually stored in a hash where the keys are integers representing a 
position in an order list.

B<Accepts:> an integer or an octal number or octal string for the for the format string 
$position

B<Returns:> an excel format string

B<Delegated to the workbook class:> yes

=back

=head3 set_defined_excel_formats( %args )

=over

B<Definition:> This will set the excel format strings for the stored positions. 
The positions are actually stored in a hash where the keys are integers representing a 
position in an order list.

B<Accepts:> a Hash list, a hash ref (both with keys representing positions), or an array 
of strings with the update strings in the equivalent position.  All empty positions are 
ignored meaning that the defalt value is left in force.  To erase the default value send 
'@' (passthrough) as the format string for that position.  This function does not do any 
string validation.  The validation is done when the coercion is generated.

B<Returns:> 1 for success

=back

=head3 total_defined_excel_formats

=over

B<Definition:> This returns the current count of excel formats that are defined

B<Accepts:> nothing

B<Returns:> $count (an integer)

B<Delegated to the workbook class:> yes

=back

=head3 change_output_encoding( $string )

=over

B<Definition:> This is always called by the L<Worksheet
|Spreadsheet::XLSX::Reader::LibXML::Worksheet> when a cell value is retreived in order to allow 
for encoding adjustments on the way out.  See 
L<XML::LibXML/ENCODINGS SUPPORT IN XML::LIBXML> for an explanation of how the input encoding 
is handled.  This conversion out is done prior to any number formatting.  If you are replacing 
this role you need to have the function and you can use it to mangle your output string any 
way you want.

B<Accepts:> a perl coded string

B<Returns:> the converted $string decoded to the L<defined format|/target_encoding>

B<Delegated to the workbook class:> yes

=back

=head3 get_defined_conversion( $position )

Defined in L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/get_defined_conversion( $position )>

=head3 parse_excel_format_string( $format_string, $name )

Defined in L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/parse_excel_format_string( $string, $name )>

B<Delegated to the workbook class:> no

=head2 Attributes

Data passed to new when creating an instance of this class. For modification of these attributes 
see the listed 'attribute methods'.  For more information on attributes see 
L<Moose::Manual::Attributes>.  The easiest way to modify these attributes are when a class
instance is created and before it is passed to the workbook or parser.

=head3 error_inst

=over

B<Definition:> This is mostly a place to store the shared error reporting instance from 
the workbook.  It is not necessary to define this attribute since it will be overwritten 
by the workbook when the Spreadsheet::XLSX::Reader::LibXML instance is loaded to the 
parser / workbook.

B<Required Methods:> (of the instance) error, set_error, clear_error

B<Default:> none

B<Attribute required:> no

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_error_inst>

=over

B<Definition:> use to set a new error instance.

=back

=back

=back

=head3 excel_region

=over

B<Definition:> This records the target region of this localization role (Not the region of the 
Excel workbook being parsed).  It's mostly a reference value.

B<Default:> en = english

B<Attribute required:> no

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_excel_region>

=over

B<Definition:> returns the value of the attribute (en)

=back

B<set_excel_region>

=over

B<Definition:> sets the value of the attribute.

=back

=back

=back

=head3 target_encoding

=over

B<Definition:> This is the target output encoding.  If it is not defined the string 
transformation step becomes a passthrough.  When the value is loaded it is used as a 
'decode' target by L<Encode> to transform the internally stored perl string to some 
target 'output' formatting.  The testing for the results of entering actual values 
here is weak.

B<Attribute required:> no

B<Default:> none

B<Range:> Any encoding recognized by L<Encode>

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_target_encoding( $encoding )>

=over

B<Definition:> This should be recognized by L<Encode/Listing available encodings>

=back

B<get_target_encoding>

=over

B<Definition:> Returns the currently set attribute value

=back

=back

=back

=head3 defined_excel_translations

=over

B<Definition:> In Excel part of localization is the way numbers are displayed.  
Excel manages that with a default list of format strings that make the numbers appear 
in a familiar way.  This is where you store / change that list for this package.  
In this case the numbers are stored as hash key => value pairs where the keys are 
integers representing the format string position and the values are the Excel 
readable format strings (definitions).  Beware that if you change the list your 
reader may break if you don't supply replacements for all the values in the default 
list.  If you just want to replace some of the values use the method 
L<set_defined_excel_formats|/set_defined_excel_formats( %args )>.

B<Attribute required:> yes

B<Default:> see the code

B<Range:> Any hashref of formats recognized by 
L<Spreadsheet::XLSX::Reader::::LibXML::ParseExcelFormatStrings>

B<attribute methods> Methods provided to by the attribute to adjust it.
		
=over

B<total_defined_excel_formats>

=over

B<Definition:> get the count of the current key => value pairs

=back

See L<get_defined_excel_format|/get_defined_excel_format( $position )> and 
L<set_defined_excel_formats|/set_defined_excel_formats( %args )>

=back

=back

=head3 dont_inherit

=over

B<Definition:> If for some reason you wish to set the formatter instance with 
settings that do not inherit from the workbook then this attribute if for you.  
The package to formtter inheritances are identified in the L<PACKAGE DESCRIPTION
|/PACKAGE DESCRIPTION>

B<Attribute required:> no

B<Default:> 0 = inherit from the package

B<Range:> 1|0

B<attribute methods> Methods provided to by the attribute to adjust it.
		
=over

B<block_inherit>

=over

B<Definition:> This is the reader for the attribute value

=back

=back

=back

=head3 epoch_year

=over

Defined at L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/epoch_year>

=back

=head3 cache_formats

=over

Defined at L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/cache_formats>

=back

=head3 datetime_dates

=over

Defined at L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/datetime_dates>

=back

=head2 Replacement

Any replacement of this class must provide the following methods in order to interact 
correctly with the workbook

=over

L<set_error_inst|/error_inst>

L<set_excel_region|/excel_region>

L<set_target_encoding|/target_encoding>

L<get_defined_excel_format|/get_defined_excel_format( $position )>

L<set_defined_excel_formats|/set_defined_excel_formats( %args )>

L<change_output_encoding|/change_output_encoding( $string )>

L<block_inherit|/dont_inherit>

L<set_epoch_year|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/epoch_year>

L<set_cache_behavior|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/cache_formats>

L<set_date_behavior|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/datetime_dates>

L<get_date_behavior|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/datetime_dates>

L<get_defined_conversion|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/get_defined_conversion( $position )>

L<parse_excel_format_string|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/parse_excel_format_string( $string, $name )>

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

Nothing L<yet|/SUPPORT>.

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

This software is copyrighted (c) 2014, 2015 by Jed Lund

=head1 DEPENDENCIES

=over

L<version> - 0.77

L<perl 5.010|perl/5.10.0>

L<Moose>

L<Types::Standard>

L<Carp> - confess

L<Encode> - decode

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