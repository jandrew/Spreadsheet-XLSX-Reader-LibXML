package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = qv('v0.4.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		InstanceOf
		FileHandle
    );
use XML::LibXML::Reader;
use lib	'../../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LibXML::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types v0.1 qw(
		XMLFile
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_name =>(
		isa			=> XMLFile,
		reader		=> 'get_file_name',
		trigger		=> \&_set_file_name,
		required	=> 1,
	);

has	error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
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
	$self->_clear_xml_parser;
	$self->_clear_location;
	my $fh = $self->_get_file_handle;
	seek( $fh, 0, 0 );#SEEK_SET
	$self->_set_xml_parser( XML::LibXML::Reader->new( IO => $fh ) );
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
		get_attribute		=> 'getAttribute',
		read_state			=> 'readState',
		name				=> 'name',
		_encoding			=> 'encoding',
		_go_to_the_end		=> 'finish',
		_close_the_sheet	=> 'close',
	}
);

has _file_handle =>(
		isa			=> FileHandle,
		reader		=> '_get_file_handle',
		writer		=> '_set_file_handle',
		predicate	=> '_has_file_handle',
		handles	=>{
			_close_file_handle	=> 'close',
		},
	);

has _file_encoding =>(
		isa		=> Str,
		reader	=> 'encoding',
		writer	=> '_set_encoding',
	);

has _position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> '_i_am_here',
		clearer		=> '_clear_location',
		predicate	=> 'has_position',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _set_file_name{
	my( $self, $new_file, $old_file, $mapped ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_set_file_name', );
	###LogSD	no warnings 'uninitialized';
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"(Re)setting the file to: $new_file",
	###LogSD			"From the file: $old_file",
	###LogSD			"With mapped setting: $mapped", ] );
	###LogSD	use warnings 'uninitialized';
	
	if( $self->_has_xml_parser ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pre-existing reader in place - clearing it" ] );
		$self->_go_to_the_end;
		$self->_close_the_sheet;
		$self->_clear_xml_parser;
		$self->_clear_location;
		$self->_close_file_handle if $self->_has_file_handle;
	}
	
	# Set the reader file
	open my $fh, '<', $new_file;
	binmode( $fh );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"File handle open;", $fh ] );
	my	$reader		= XML::LibXML::Reader->new( IO => $fh );#'XMLFILElocation => $new_file )', );#recover => 2, 
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"XML parser open;", $reader,
	###LogSD		'Read state: ' . $reader->readState ] );
	if( !$reader ){
		$self->_clear_xml_parser;
		return undef;
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ 'Success - Loading file handle: ' . $fh ], );
		$self->_set_file_handle( $fh );
		###LogSD	$phone->talk( level => 'debug', message =>[ 'Loading XML reader: ' . $reader ], );
		$self->_set_xml_parser( $reader );
		if( $self->byte_consumed == 0 ){
			###LogSD	$phone->talk( level => 'debug', message =>[ 'Starting the read' ], );
			$self->start_reading;
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Good reader built", "Byte position: " . $self->byte_consumed ], );#$reader->byteConsumed
		return 1 if $mapped;
	}
	
	#Set the file unique bits
	# Get file encoding
	my	$encoding	= $self->_encoding;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Encoding of file is: $encoding" ], );
	$self->_set_encoding( $encoding );
	if( $self->can( '_load_unique_bits' ) ){
		###LogSD	$phone->talk( level => 'debug', message => [ "Loading unique bits" ], );
		$self->_load_unique_bits;
		###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading unique bits" ], );
	}
	return 1;#$reader;
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for file_name:" . $self->get_file_name,
	###LogSD			"clearing the error instance",									] );
	$self->_clear_xml_parser,
	$self->_clear_error_inst,
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader - XLSX - LibXML Reader base class

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
    
=head1 DESCRIPTION

L<XML::LibXML> supports TIMTOWDI by providing multiple ways to parse a file.  This package 
is built to support general pull parsing using L<XML::LibXML::XMLReader>.  All specific 
pull parsers are built on this.  If you want to use this to write your own reader just load 
the file name and use the methods to explore the file.  B<Since this is a pull parser you 
have to rewind to the beginning to go back.>

This sheet has the role L<Spreadsheet::XLSX::Reader::::LibXML::LogSpace> and all it's 
functionality added.

=head2 Attributes

Data passed to new when creating an instance (pull_parser).  For modification of 
these attributes see the listed L<Methods|/Methods> of the instance.  All role attributes 
and methods are documented in the role documentation.

=head3 file_name

=over

B<Definition:> This is the file to be read using L<XML::LibXML::Reader> techniques.

B<Default> none

B<Range> any complete file name
		
=back

=head3 error_inst

=over

B<Definition:> This package can share a single error instance so that an error registered in one 
place can be read in another place.  The documentation for the instance is found in 
L<Spreadsheet::XLSX::Reader::Error>

B<Default> none

B<Range> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ]
		
=back

=head2 Methods

These include methods to adjust attributes as well as providing methods to navigate the file.

=head3 get_system_type

=over

B<Definition:> This is the way to see whether the conversion is Windows or Apple based

B<Accepts:>Nothing

B<Returns:> win_excel|apple_excel

=back

=head3 _load_unique_bits

=over

B<Definition:> _load_unique_bits isn't a function with and inner() call in it.  I could never 
get Moose to call inner() in a trigger function and I'm sure there is good theory behind not doinng it.  
However, when the file is loaded to the parser there may be elements of the file that are not 
found in the _core_element(s) and should be loaded in the subclass for reference.  When this class
loads the file and checks if the subclass has a _load_unique_bits function.  The subclass can then 
load any relevent data to the instance that is unique to the subclass.  The subclass function 
_load_unique_bits is expected to either return undef|0 or '1'.  Returning '1' will force the file to 
reload.  This is useful when the _load_unique_bits function in the sublclass needs to iterate 
through the _core_elements section to find non-core data.  For the reader to then not be lost it 
needs to reset at the beginning of the file.  The relevant section in this class looks like this:

	my	$reload = 0;
	if( !$mapped and $self->can( '_load_unique_bits' ) ){
		###LogSD Loading unique bits
		$reload = $self->_load_unique_bits( $reader, $encoding );
	}
	if( !$reload ){
		###LogSD All ready
	}elsif( $reload == 1 ){
		$reader		= XML::LibXML::Reader::LibXML->new( location => $self->get_file_name );
	}else{
		return undef;
	}
		
=back

=head2 ATTRIBUTES

Data passed to new when creating an instance (parser).  For modification of 
these attributes see the listed L</METHODS> of the instance.

=head3 file_name

=over

B<Definition:> This attribute stores the string used to access the file (the file name).  
It will be L<tested|Spreadsheet::XLSX::Reader::LibXML::Types> as an XMLFile type.

B<Default> none

B<Range> any readable xml file (null not allowed)
		
=back

=head2 METHODS

These include methods to adjust attributes as well as providing methods to 
implement the functionality of the module.

=head3 get_position( $int )

=over

B<Definition:> This calls a routine that searches for the identified position of the 
defined _core_element in the xml file.  If no value is passed it will test for the 
current recorded position and pull the next one.

B<Accepts:> an integer representing the position in the array of elements indicated 
by the attribute _core_element.  It will calculate the next position if no value is 
passed.

B<Returns:> An L<XML::LibXML::Element> instance with the element data from the file 
contained in it.  If the requested position (or the next position) is passed the end 
of the file then this returns undef.

=back

=head3 error( $error_string )

=over

B<Definition:> This method is handled from the workbook link generally built by 
L<Spreadsheet::XLSX::Reader::LibXML>.  See L<Spreadsheet::XLSX::Reader::LibXML::Error> 
for details of the error_string attribute associated with this method.

B<Accepts:> a message string

B<Returns:> nothing

=back

=head3 encoding

=over

B<Definition:> This is the encoding of the file as recorded in the xml attribute.  
In general L<XML::LibXML> should be converting the data into unicode for perl..

B<Accepts:> nothing

B<Returns:> the value of the encoding attribute in the xml file

=back

=head3 where_am_i

=over

B<Definition:> The module tracks the last recorded _core_element position returned.  
This is the way to read that value.

B<Accepts:> nothing

B<Returns:> An integer counting from 0 of the last _core_element returned

=back

=head3 has_position

=over

B<Definition:> Either before the first position is returned or after the end of 
the _core_element list is reached the last recorded position will be undef.  
This is a way to test for that state.

B<Accepts:> nothing

B<Returns:> a boolean value indicating if there is a current _core_element 
position.

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

L<XML::LibXML>

L<XML::LibXML::Reader>

L<Spreadsheet::XLSX::Reader::LibXML::LogSpace>

L<Spreadsheet::XLSX::Reader::LibXML::Types>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::ParseExcel>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9