package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = qv('v0.36.8');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				Str				HasMethods
		Bool
    );
use XML::LibXML::Reader;
use lib	'../../../../../lib',;
###LogSD	with 'Log::Shiras::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types qw(
		IOFileType
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file =>(
		isa			=> IOFileType,
		reader		=> 'get_file',
		writer		=> 'set_file',
		predicate	=> 'has_file',
		clearer		=> 'clear_file',
		coerce		=> 1,
		required	=> 1,
		trigger		=> \&_build_xml_reader,
	);

has	error_inst =>(
		isa			=> 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


sub start_the_file_over{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space => $self->get_log_space . '::start_the_file_over', );
	###LogSD		$phone->talk( level => 'debug', message =>[ "Resetting the XML file" ] );
	$self->_go_to_the_end;
	$self->_close_the_sheet;
	#~ $self->_clear_xml_parser;
	$self->_clear_location;
	my $fh = $self->get_file;
	$fh->seek( 0, 0 );
	$self->_set_xml_parser( XML::LibXML::Reader->new( IO => $fh ) );
	$self->start_reading;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _xml_reader =>(
	isa			=> 'XML::LibXML::Reader',
	reader		=> '_get_xml_parser',
	writer		=> '_set_xml_parser',
	predicate	=> '_has_xml_parser',
	clearer		=> '_clear_xml_parser',
	handles	=>{
		copy_current_node	=> 'copyCurrentNode',
		byte_consumed		=> 'byteConsumed',
		start_reading		=> 'read',
		next_element		=> 'nextElement',
		node_type			=> 'nodeType',
		node_name			=> 'name',
		node_value			=> 'value',
		has_value			=> 'hasValue',
		node_depth			=> 'depth',
		move_to_first_att	=> 'moveToFirstAttribute',
		move_to_next_att	=> 'moveToNextAttribute',
		encoding			=> 'encoding',
		_go_to_the_end		=> 'finish',
		_close_the_sheet	=> 'close',
	},
	trigger => \&_reader_init,
);

has _position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> '_i_am_here',
		clearer		=> '_clear_location',
		predicate	=> 'has_position',
	);

has _read_unique_bits =>(
		isa		=> Bool,
		reader	=> '_get_unique_bits',
		writer	=> '_need_unique_bits',
		clearer	=> '_clear_read_unique_bits',
		default	=> 1,
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _build_xml_reader{
	my( $self, $file_handle ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_build_xml_reader', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"turning a file handle into an xml reader", ] );
	
	# Build the reader
	$file_handle->seek( 0, 0 );
	my	$xml_reader = XML::LibXML::Reader->new( IO => $file_handle );
		$xml_reader->read;
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Loading reader:', $xml_reader ], );
	$self->_set_xml_parser( $xml_reader );
	###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading XML reader" ], );
}

sub _reader_init{
	my( $self, $reader ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_reader_init', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"loading any file specific settings", ] );
	
	# Set the file unique bits
	if( $self->_get_unique_bits ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Opening the file for the first time" ], );
		$self->_need_unique_bits( 0 );
		if( $self->can( '_load_unique_bits' ) ){
			###LogSD	$phone->talk( level => 'debug', message =>[ "Loading unique bits" ], );
			$self->_load_unique_bits;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Finished loading unique bits" 			], );
		}
		$self->start_the_file_over;
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ 
		###LogSD		"This is not the first time the file has been opened - don't seek unique" ], );
	}
	###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading unique bits BLOCK" ], );
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for log space:" . $self->get_log_space, ] );
	if( $self->_get_xml_parser ){
		#~ print "Disconnecting the sheet file handle from the parser\n";
		###LogSD	$phone->talk( level => 'debug', message =>[ "Disconnecting the file handle from the xml parser", ] );
		$self->_close_the_sheet;
		$self->_clear_xml_parser;
	}
	if( my $fh = $self->get_file ){
		#~ print "Clearing file handle\n";
		###LogSD	$phone->talk( level => 'debug', message =>[ "Closing the system file handle" ] );#, $self->dump(2)
		$fh->close;
		###LogSD	$phone->talk( level => 'debug', message =>[ "Clearing the system file handle" ] );#, $self->dump(2)
		$self->clear_file;
		###LogSD	$phone->talk( level => 'debug', message =>[ 'All done' ] );
		###LogSD	$phone->talk( level => 'trace', message =>[ "Final self", $self->dump(2) ] );
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader - A LibXML::Reader xlsx base class

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This module provides a generic way to open an xml file or xml file handle and then extract 
information using the L<XML::LibXML::Reader> parser.  The class does this by using 
L<delegation|Moose::Manual::Delegation> from the ~::Reader module to import useful functions 
to this module.  Aside from the delegation piece this module provides four other useful 
elements.  First, the module has an attribute to load the file or file handle and uses coercion 
to turn a file into a file handle from the L<Types
|Spreadsheet::XLSX::Reader::LibXML::Types/IOFileType> library.  Second, the module has an 
attribute to store an L<error handler|Spreadsheet::XLSX::Reader::LibXML::Error>.  Third, 
the module provides a L<rewind|/start_the_file_over> function since that is not available in  
the L<XML::LibXML::Reader> parser. Finally, this module has a hook for classes that extend 
this functionality during the initial build.  The initialization of the file will also attempt 
to call '_load_unique_bits'.  It will only call that method once on initialization.

Further use of the module or specialization of the reader can be done by L<extending|/SYNOPSIS> 
the class.

=head2 Attributes

Data passed to new when creating an instance.  For modification of these attributes see the 
listed 'attribute methods'. For general information on attributes see 
L<Moose::Manual::Attributes>.  For ways to manage the instance when opened see the 
L<Methods|/Methods>.

=head3 file

=over

B<Definition:> This attribute holds the file handle for the file being read.  If the full 
file name and path is passed to the attribute it is coerced to an IO::File file handle.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any unencrypted xml file name and path or file handle

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_file>

=over

B<Definition:> change the file value in the attribute (this will reboot 
the file instance and lock the file)

=back

B<get_file>

=over

B<Definition:> Returns the file handle of the file even if a file name 
was passed

=back

B<has_file>

=over

B<Definition:> this is used to see if the file loaded correctly.

=back

B<clear_file>

=over

B<Definition:> this clears (and unlocks) the file handle

=back

=back

=back

=head3 error_inst

=over

B<Definition:> This attribute holds the L<error handler
|Spreadsheet::XLSX::Reader::LibXML::Error>.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any object instance that can provide the required delegated methods.

B<delegated methods> Methods provided delegated by the attribute
		
=over

B<error>

=over

B<Definition:> returns the stored error string

=back

B<set_error>

=over

B<Definition:> Sets the error string

=back

B<clear_error>

=over

B<Definition:> clears the error string

=back

B<set_warnings>

=over

B<Definition:> Sets the state that determins if the instance pro-activly 
warns with the error string when the error string is set.

=back

B<if_warn>

=over

B<Definition:> Returns the current state of the state value from 'set_warnings'

=back

=back

=back

=head2 Class Methods

These are the methods provided by this class.  They most likely should be agumented 
with file specific methods when extending this module.

=head3 start_the_file_over

=over

B<Definition:> This will disconnect the L<XML::LibXML::Reader> from the file handle,  
rewind the file handle, and then reconnect the L<XML::LibXML::Reader> to the file handle.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head2 Delegated Methods

These are the methods delegated to this class from L<XML::LibXML::Reader>.  For more 
general parsing of subsections of the xml file also see L<Spreadsheet::XLSX::Reader::LibXML>.

=head3 copy_current_node

B<Delegated from:> L<XML::LibXML::Reader/copyCurrentNode (deep)>

=head3 byte_consumed

B<Delegated from:> L<XML::LibXML::Reader/byteConsumed ()>

=head3 start_reading

B<Delegated from:> L<XML::LibXML::Reader/read ()>

=head3 next_element

B<Delegated from:> L<XML::LibXML::Reader/nextElement>

=head3 node_type

B<Delegated from:> L<XML::LibXML::Reader/nodeType>

=head3 node_name

B<Delegated from:> L<XML::LibXML::Reader/name>

=head3 node_value

B<Delegated from:> L<XML::LibXML::Reader/value>

=head3 has_value

B<Delegated from:> L<XML::LibXML::Reader/hasValue>

=head3 node_depth

B<Delegated from:> L<XML::LibXML::Reader/depth>

=head3 move_to_first_att

B<Delegated from:> L<XML::LibXML::Reader/moveToFirstAttribute>

=head3 move_to_next_att

B<Delegated from:> L<XML::LibXML::Reader/moveToNextAttribute>

=head3 encoding

B<Delegated from:> L<XML::LibXML::Reader/encoding ()>

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing currently

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

L<XML::LibXML::Reader>

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