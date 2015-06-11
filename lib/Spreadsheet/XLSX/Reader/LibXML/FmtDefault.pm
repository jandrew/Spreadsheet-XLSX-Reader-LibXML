package Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use version; our $VERSION = qv('v0.36.24');

use	5.010;
use	Moose::Role;
###LogSD	requires qw(
###LogSD		get_log_space
###LogSD	);

use Types::Standard qw( InstanceOf ArrayRef Str );
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has excel_region =>(
		isa		=> Str,
		default	=> 'en',
		reader	=> 'get_excel_region',
	);
	
has	target_encoding =>(
		isa			=> Str,
		reader		=> 'get_target_encoding',
		writer		=> 'set_target_encoding',
		default		=> 'UTF-8',
		required	=> 1,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub	change_output_encoding{
	my ( $self, $string, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::change_output_encoding', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Changing the encoding of: $string",
	###LogSD				'..to encoding type: ' . $self->get_target_encoding ] );
	return $string;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _defined_excel_translations =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		default	=> sub{ [
						'General',
						'0',
						'0.00',
						'#,##0',
						'#,##0.00',
						'$#,##0_);($#,##0)',
						'$#,##0_);[Red]($#,##0)',
						'$#,##0.00_);($#,##0.00)',
						'$#,##0.00_);[Red]($#,##0.00)',
						'0%',
						'0.00%',
						'0.00E+00',
						'# ?/?',
						'# ??/??',
						'yyyy-m-d',      # Was 'm-d-yy', which is bad as system default
						'd-mmm-yy',
						'd-mmm',
						'mmm-yy',
						'h:mm AM/PM',
						'h:mm:ss AM/PM',
						'h:mm',
						'h:mm:ss',
						'm-d-yy h:mm',
						undef, undef, undef, undef, undef, undef, undef, undef,
						'#,##0_);(#,##0)',
						'#,##0_);[Red](#,##0)',
						'#,##0.00_);(#,##0.00)',
						'#,##0.00_);[Red](#,##0.00)',
						'_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
						'_($*#,##0_);_($*(#,##0);_($*"-"_);_(@_)',
						'_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
						'_($*#,##0.00_);_($*(#,##0.00);_($*"-"??_);_(@_)',
						'mm:ss',
						'[h]:mm:ss',
						'mm:ss.0',
						'##0.0E+0',
						'@'
					]
		},
		reader => 'get_defined_excel_format_list',
		writer => 'set_defined_excel_format_list',
		handles =>{
			get_defined_excel_format => 'get',
			total_defined_excel_formats => 'count',
		},
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::FmtDefault - Default xlsx number formats and localization

=head1 SYNOPSIS

    #!/usr/bin/env perl
    package MyPackage;
    use Moose;
    with 'Spreadsheet::XLSX::Reader::LibXML::FmtDefault';
    
    package main;
    
    my $parser = MyPackage->new;
    print '(' . join( ', ', $parser->get_defined_excel_format( 14 ) ) . ")\n";
	
	###########################
	# SYNOPSIS Screen Output
	# 01: (yyyy-m-d)
	###########################
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your 
own excel parser.  To use the general package for excel parsing out of the box please 
review the documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This L<Moose Role|Moose::Manual::Roles> is the primary tool for localization.  It 
stores the number conversion format strings for the set region.  In this particular 
case this module  is the base L<english conversion
|http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/02/16/dates-in-spreadsheetml.aspx> 
set.  It does rely on L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings> 
to build the actual coercions used to transform numbers for each string.  However, 
the ParseExcelFormatStrings transformation should work for all regions strings.  
When writing a drop in replacement for this module you should be able to just replace 
the strings in the '_defined_excel_translations' attribute and then set the 
L<Spreadsheet::XLSX::Reader::LibXML/default_format_list> with your module when 
creating a new workbook parser.  (Don't forget to rename the module!)

The role also includes a string conversion function that is implemented after the data is 
extracted by libxml2 from the xml file.  Specifically libxml2 attempts to determine the input 
encoding from the xml header and convert whatever format the file is in to unicode so the 
conversion out should be from unicode to your L<target_encoding|/target_encoding>.   
L<For now|/TODO> no encoding (output) conversion is actually provided and the function is 
essentially a pass-through of standard perl unicode.
	
=head2 Primary Methods

These are the primary ways to use this Role.  For additional FmtDefault options see the 
L<Attributes|/Attributes> section.

=head3 change_output_encoding( $string )

=over

B<Definition:> Currently this is a placeholder that is always called by the L<Worksheet
|Spreadsheet::XLSX::Reader::LibXML::Worksheet> when a cell value is retreived in order to allow 
for I<future> encoding adjustments on the way out.  See 
L<XML::LibXML/ENCODINGS SUPPORT IN XML::LIBXML> for an explanation of how the input encoding 
is handled.  This conversion out is done prior to any number formatting.  If you are replacing 
this role you need to have the function and you can use it to mangle your output string any 
way you want.

B<Accepts:> a unicode string

B<Returns:> the converted string I<currently with no changes>

=back

=head3 get_defined_excel_format( $integer )

=over

B<Definition:> This will return the preset excel format string for the stored position.  
This role is used in the L<Styles|Spreadsheet::XLSX::Reader::LibXML::Styles> class but 
this method is actually implemented through the the attribute 
L<Spreadsheet::XLSX::Reader::LibXML/default_format_list> using L<delegation
|Moose::Manual::Delegation>.

B<Accepts:> an $integer for the format string position

B<Returns:> an excel format string

=back

=head3 total_defined_excel_formats

=over

B<Definition:> This will return the count of all defined Excel format strings for this 
role.  The primary value is to understand if the format string falls in the range of a 
pre-set value or if the general .xlsx sheet reader should look in the 
L<Styles|Spreadsheet::XLSX::Reader::LibXML::Styles> sheet for the format string.

B<Accepts:> nothing

B<Returns:> the total count of the pre-defined number coercion formats

=back

=head3 get_defined_excel_format_list

=over

B<Definition:> This will return the complete list of defined formats as an array ref

B<Accepts:> nothing

B<Returns:> an array ref of all pre-defined format strings

=back

=head3 set_defined_excel_format_list

=over

B<Definition:> If you don't want to re-write this role you can just set a new 
array ref of format strings that you want excel to use.  The strings need to comply with 
the capabilities of L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings>.  With 
any luck that means they also comply with the Excel L<format string definitions
|https://support.office.com/en-us/article/Create-or-delete-a-custom-number-format-83657ca7-9dbe-4ee5-9c89-d8bf836e028e?ui=en-US&rs=en-US&ad=US>.  
This role is consumed by the L<Styles|Spreadsheet::XLSX::Reader::LibXML::Styles> class but 
this method is actually exposed all the way up to the L<Workbook
|Spreadsheet::XLSX::Reader::LibXML> class through L<Delegation|Moose::Manual::Delegation>.

B<Accepts:> an array ref of format strings

B<Returns:> nothing

=back

=head2 Attributes

Data passed to new when creating the L<Styles|Spreadsheet::XLSX::Reader::LibXML::Styles> 
instance.   (or other class instance consuming this role) For modification of these attributes 
see the listed 'attribute methods'.  For more information on attributes see 
L<Moose::Manual::Attributes>.  Most of these attributes and methods are not exposed to the top 
level of L<Spreadsheet::XLSX::Reader::LibXML>.

=head3 excel_region

=over

B<Definition:> This records the target region of this localization role (Not the region of the 
Excel workbook being parsed)

B<Default:> en = english

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_excel_region>

=over

B<Definition:> returns the value of the attribute (en)

=back

=back

=back

=head3 target_encoding

=over

B<Definition:> This is the target output encoding

B<Default:> UTF-8

B<Range:> No real options here (since it currently is a No Op)

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_target_encoding( $encoding )>

=over

B<Definition:> Changing this won't affect anything

=back

B<get_target_encoding>

=over

B<Definition:> Returns the currently set attribute value

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

B<1.> Actually make the L<change_output_encoding|/change_output_encoding> method do 
something useful.

B<2.> Add more roles like this for othere regions and allow them to be selected 
by a region attribute setting in L<Spreadsheet::XLSX::Reader::LibXML>

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

L<Moose::Role>

L<Types::Standard>

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