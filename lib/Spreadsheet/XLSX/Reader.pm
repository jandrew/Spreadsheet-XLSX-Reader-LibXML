package Spreadsheet::XLSX::Reader;
use version; our $VERSION = version->declare("v0.1_1");
use 5.010;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use	Archive::Zip;
use	OLE::Storage_Lite;
#~ use	Data::Dumper;
use	File::Temp;
#~ $File::Temp::DEBUG = 1;
use	DateTime::Format::Flexible;
#~ use	Carp qw( cluck );
#~ use	XML::LibXML;
#~ my	$parser = XML::LibXML->new;
#~ BEGIN{
	#~ $ENV{PERL_DESTRUCT_LEVEL} = 2;
#~ }
use Types::Standard qw(
        HashRef
		Str
		Bool
		ArrayRef
		Object
		Int
		CodeRef
    );
use lib	'../../../lib',;
with 	'Spreadsheet::XLSX::Reader::LogSpace',
		'Spreadsheet::XLSX::Reader::Error';
use Spreadsheet::XLSX::Reader::TempFilter;# Fix with release of Log::Shiras
###LogSD	use Log::Shiras::Telephone;# Fix with CPAN release of Log::Shiras
use Spreadsheet::XLSX::Reader::Types v0.1 qw(
		XLSXFile
        ParserType
		EpochYear
	);
use	Spreadsheet::XLSX::Reader::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::XMLReader::CalcChain;
use	Spreadsheet::XLSX::Reader::XMLDOM::Styles;
use	Spreadsheet::XLSX::Reader::XMLReader::Worksheet;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_name =>(
		isa			=> XLSXFile,
		writer		=> 'set_file_name',
		clearer		=> '_clear_file_name',
		predicate	=> 'has_file_name',
		trigger		=> \&_build_file,
	);
	
has creator =>(
		isa		=> Str,
		writer	=> '_set_creator',
		clearer	=> '_clear_creator',
	);
	
has modified_by =>(
		isa		=> Str,
		writer	=> '_set_modified_by',
		clearer	=> '_clear_modified_by',
	);
	
has date_created =>(
		isa		=> 'DateTime',
		writer	=> '_set_date_created',
		clearer	=> '_clear_date_created',
	);
	
has date_modified =>(
		isa		=> 'DateTime',
		writer	=> '_set_date_modified',
		clearer	=> '_clear_date_modified',
	);

has sheet_parser =>(
		isa		=> ParserType,
		writer	=> 'set_parser_type',
		reader	=> 'get_parser_type',
		default	=> 'reader',
		coerce	=> 1,
	);
	
has sheet_parser_modules =>(### Add Data::Walk::Extracted manipulators
		isa		=> HashRef,
		traits	=> ['Hash'],
		writer	=> 'set_sheet_parser_modules',
		reader	=> 'get_sheet_parser_modules',
		default	=> sub{
			{
				reader =>{
					sharedStrings =>{
						module		=> 'Spreadsheet::XLSX::Reader::XMLReader:SharedStrings',
					},
					calcChain		=>{
						module		=> 'Spreadsheet::XLSX::Reader::XMLReader::CalcChain',
					},
					styles =>{
						module		=> 'Spreadsheet::XLSX::Reader::DOM::Styles',
						attributes	=> [ 'excel_epoch', 'self' ],#### Does self cover excel_epoch?
					},
					worksheet =>{
						module		=> 'Spreadsheet::XLSX::Reader::XMLReader::Worksheet',
					},
				},
				dom	=>{
					sharedStrings =>{
						module		=> 'Spreadsheet::XLSX::Reader::DOM::SharedStrings',
					},
					calcChain =>{
						module		=> 'Spreadsheet::XLSX::Reader::DOM::CalcChain',
					},
					styles =>{
						module		=> 'Spreadsheet::XLSX::Reader::DOM::Styles',
						attributes	=> [ 'excel_epoch', 'self' ],#### Does self cover excel_epoch?
					},
					worksheet =>{
						module		=> 'Spreadsheet::XLSX::Reader::DOM::Worksheet',
					},
				},
			},
		},
		handles =>{
			_get_module_list => 'get',
			_has_parser_type => 'exists',
			
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


###############################################################################
#
# parse()
#
# A convenience method to duplicate Spreadsheet::ParseExcel processes

sub parse{

    my ( $self, $file_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::parse', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Arrived at parse for: ', $file_name ] );
	$self->set_file_name( $file_name );
	return ( $self->has_file_name ) ? $self : undef;
}

sub worksheets{

    my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::worksheets', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Attempting to build all worksheets: ', $self->get_worksheet_names ] );
	my	@worksheet_array;
	while( my $worksheet_object = $self->worksheet ){
	#~ for my $worksheet_name ( @{$self->get_worksheet_names} ){
		push @worksheet_array, $worksheet_object;#$self->worksheet( $worksheet_name );
	}
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		'sending worksheet array: ',@worksheet_array ] );
	return @worksheet_array;
}

sub worksheet{

    my ( $self, $worksheet_name ) = @_;
	my ( $next_position );
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::worksheet', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at worksheet with: ", $worksheet_name ] );
	
	# Handle an implied 'next sheet'
	if( !$worksheet_name ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No worksheet name passed" ] );
		$next_position = ( !$self->in_the_list ) ? 0 : ($self->_current_worksheet_position + 1);
		if( $next_position >= $self->number_of_sheets ){
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Reached the end of the worksheet list" ] );
			return undef;
		}
		$worksheet_name = $self->worksheet_name( $next_position );
	}
	
	# build the worksheet
	my	$worksheet_info = $self->_get_worksheet_info( $worksheet_name );
	$next_position //= $worksheet_info->{position};
	$worksheet_info->{_shared_strings_instance} = $self->_shared_strings_instance if
		$self->_shared_strings_instance;
	$worksheet_info->{_calc_chain_instance} = $self->_calc_chain_instance if
		$self->_calc_chain_instance;
	$worksheet_info->{_styles_instance} = $self->_styles_instance if
		$self->_styles_instance;
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"The next worksheet is: $worksheet_name",
	###LogSD		"Using Class: " . $self->_get_module_list( $self->get_parser_type )->{worksheet},
	###LogSD		"With worksheet ref:", 
	###LogSD		{
	###LogSD			name_space 	=> $self->get_log_space . "::Worksheet",
	###LogSD			name		=> $worksheet_name,
	###LogSD			%{$worksheet_info},
	###LogSD		}										] );
	my	$worksheet = ($self->_get_module_list( $self->get_parser_type )->{worksheet})->new(
						name_space 	=> $self->get_log_space . "::Worksheet",
						name		=> $worksheet_name,
						%{$worksheet_info},  );
	if( $worksheet ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Successfully loaded: $worksheet_name",
		###LogSD		"With worksheet ref:", $worksheet_info ] );
		$self->_set_current_worksheet_position( $next_position );
	}else{
		$self->error( "Failed to build the object for worksheet: $worksheet_name" );
		return undef;
	}
	my $cleaner_ref = $self->_link_cleaner_array;
	push @$cleaner_ref,
		sub{
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Clearing the worksheet connection: $worksheet_name",
			###LogSD		"For the file name: " . $worksheet_info->{file_name}, ] );
			if( $worksheet ){ $worksheet->DEMOLISH };
		};
	$self->_set_link_cleaner_array( $cleaner_ref );
	
	return $worksheet;
}
	

#~ ###############################################################################
#~ #
#~ # read_file()
#~ #
#~ # Unzip the XLSX file and read the [Content_Types].xml file to get the
#~ # structure of the contained XML files.
#~ #
#~ # Return a valid Workbook object if successful. If not return undef and set
#~ # the error status.
#~ #
#~ sub read_file {

    #~ my ( $self, $file_name ) = @_;
	#~ my	$phone = Log::Shiras::Telephone->new( fail_over => $fail_over,
					#~ name_space 	=> $self->get_log_space .  '::workbook::read_file', );
	#~ debug_line{ $phone->talk( level => 'info', message =>[
					#~ 'Arrived at read_file for: ', $file_name ] ) };
	#~ $self->_clear_shared_strings;
	#~ $self->_clear_calc_chain;
	#~ $self->_clear_temp_dir;
	#~ $self->_clear_worksheet_list;
	#~ $self->_clear_worksheet_lookup;
	#~ $self->_clear_creator;
	#~ $self->_clear_modified_by;
	#~ $self->_clear_date_created;
	#~ $self->_clear_date_modified;
	#~ my	$return_failure = 0;
    #~ if ( !-e $file_name ){
		#~ $phone->talk( level => 'warn', message =>[
			#~ "File doesn't exist: ", $file_name 		] );
		#~ $return_failure = 1;
	#~ }elsif( -z $file_name ) {
		#~ $phone->talk( level => 'warn', message =>[
			#~ "There is nothing in the file: ", $file_name ] );
		#~ $return_failure = 1;
	#~ }
	#~ if( $return_failure ){
		#~ $self->_clear_file_name;
		#~ return( undef );
	#~ }

    #~ # Check for xls or encrypted OLE files.
    #~ my $ole_file = $self->_check_if_ole_file( $file_name );
	
    #~ return( undef ) if $ole_file;

    #~ # Create a, locally scoped, temp dir to unzip the XLSX file into.
	#~ my	$temp_dir = File::Temp->newdir( DIR => undef );
    #~ $self->_set_temp_dir( $temp_dir );
	
    #~ # Create an Archive::Zip object to unzip the XLSX file.
    #~ my $zip_file = Archive::Zip->new();

    #~ # Read the XLSX zip file and catch any errors.
    #~ eval { $zip_file->read( $file_name ) };
    #~ if ( $@ ) {
		#~ $phone->talk( 	level	=> 'warn',
						#~ message =>[ "File has zip error(s):", $@ ] );
		#~ return undef;
	#~ }

    #~ # Extract the XML files from the XLSX zip.
    #~ $zip_file->extractTree( '', $self->_get_temp_dir . '/' );
	#~ debug_line{ $phone->talk( level => 'trace', message =>[ "Zip file: ", $zip_file ] ) };
	#~ my	$intermediate = $self->_get_temp_dir;
	#~ debug_line{ $phone->talk( level	=> 'trace', message =>[
					#~ "Temp dir: $intermediate", "Temp Dir contains: ", <$intermediate/*> ] ) };
	
	#~ # Load general workbook information to this instance
	#~ my	$workbook_file = $self->_get_temp_dir . '/xl/workbook.xml';
	#~ debug_line{ $phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] ) };
	#~ my ( $rel_lookup, $id_lookup ) = $self->_load_workbook_file( $workbook_file );
		#~ $workbook_file = $self->_get_temp_dir . '/xl/_rels/workbook.xml.rels';
	#~ debug_line{ $phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] ) };
	#~ my ( $pivot_lookup ) = $self->_load_rels_workbook_file( $rel_lookup, $workbook_file );
		#~ $workbook_file = $self->_get_temp_dir . '/docProps/core.xml';
	#~ debug_line{ $phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] ) };
	#~ $self->_load_doc_props_file( $workbook_file );
	
	#~ # Build the instances for all the shared worksheet files
	#~ if( $self->_has_parser_type( $self->get_parser_type ) ){
		#~ $self->_set_shared_worksheet_files( $self->_get_module_list( $self->get_parser_type) );
	#~ }else{
		#~ $phone->talk( level => 'fatal', message =>[
			#~ "No definitions for the sheet parser type: ",
			#~ $self->get_parser_type							] );
	#~ my	$wait = <>;
	#~ }
	
	#~ my	$shared_strings_file = $self->_get_temp_dir . '/xl/sharedStrings.xml';
	#~ $phone->talk( level => 'debug', message =>[ "Loading sharedStrings: $shared_strings_file",] );
	#~ my $shared_strings = 	Excel::Reader::XLSX::Shiras::SharedStrings->new( 
								#~ shared_string_file => $shared_strings_file,
							#~ );
	#~ $self->_set_shared_strings_instance( $shared_strings );
	#~ print "Loaded shared strings file\n";
	
	
	
	
	
	#~ my	$parser = XML::LibXML->new;
	#~ my	$content_doc = $parser->parse_file( $content_types_file );
	#~ my	$top = 0;
	#~ my	$has_workbook = 0;
	#~ for my $node ( $content_doc->getElementsByTagName( 'Override' ) ){
		#~ my	$part_name = $node->getAttribute( 'PartName' );
			#~ $part_name =~ s{^/}{};
		#~ $phone->talk(	level => 'trace',
						#~ message =>[	"Node: $top",
									#~ "Part Name: $part_name", ], );
		#~ if( -e $self->_get_temp_dir . $part_name ){
			#~ $phone->talk(	level => 'warn',
							#~ message =>[	
								#~ 'Listed file -' . $self->_get_temp_dir . $part_name .
								#~ '- does not exist', ], );
		#~ }elsif( -z $self->_get_temp_dir . $part_name ){
			#~ $phone->talk(	level => 'warn',
							#~ message =>[	
								#~ 'Listed file -' . $self->_get_temp_dir . $part_name .
								#~ '- has zero size', ], );
		#~ }
		#~ my	$failed_match = 1;
		#~ for my $match ( keys %$content_types ){
			#~ if( $part_name =~ /$match\d*\.xml$/ ){
				#~ $failed_match = 0;
				#~ my	$method = $content_types->{$match}->[0];
				#~ $self->$method(
					#~ @{$content_types->{$match}}[ 1 .. $#{$content_types->{$match}} ],
					#~ $part_name
				#~ );
				#~ last;
			#~ }
		#~ }
		#~ if( $failed_match ){
			#~ $phone->talk(
				#~ level 	=> 'warn',
				#~ message =>[	"No action found for Part Name: $part_name", ], );
		#~ }
		#~ $top++;
	#~ }
	#~ if( !$has_workbook ){
		#~ $phone->talk( level => 'warn', message =>[	"No workbook found!", ], );
		#~ return undef;
	#~ }
	
	#~ my	$shared_strings = $parser->parse_file( $self->_get_temp_dir . $self
    #~ # Create a reader object to read the sharedStrings.xml file.
    #~ my $shared_strings = Excel::Reader::XLSX::Package::SharedStrings->new();

    #~ # Read the sharedStrings if present. Only files with strings have one.
    #~ if ( $files{_shared_strings} ) {

        #~ $shared_strings->_parse_file( $tempdir . $files{_shared_strings} );
    #~ }

    #~ # Create a reader object for the workbook.xml file.
    #~ my $workbook = Excel::Reader::XLSX::Workbook->new(
        #~ $tempdir,
        #~ $shared_strings,
        #~ %files

    #~ );

    #~ # Read data from the workbook.xml file.
    #~ $workbook->_parse_file( $tempdir . $files{_workbook} );

    #~ # Store information in the reader object.
    #~ $self->{_files}          = \%files;
    #~ $self->{_shared_strings} = $shared_strings;
    #~ $self->{_package_dir}    = $tempdir;
    #~ $self->{_zipfile}        = $zipfile;

    #~ return $workbook;
	#~ return $self;
#~ }

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has epoch_year =>(
		isa		=> EpochYear,
		writer	=> '_set_epoch_year',
		reader	=> 'get_epoch_year',
		default	=> 1900,
	);

has _shared_strings_instance =>(
	isa		=> Object,
	writer	=>'_set_sharedStrings',
	handles	=>{
		get_shared_string_position => 'get_position',
	},
	clearer	=> '_clear_shared_strings',
);

has _calc_chain_instance =>(
	isa		=> Object,
	writer	=>'_set_calcChain',
	handles	=>{
		get_calc_chain_position => 'get_position',
	},
	clearer	=> '_clear_calc_chain',
);

has _styles_instance =>(
	isa		=> Object,
	writer	=>'_set_styles',
	#~ handles	=> [ qw( get_calc_chain_item ) ],
	clearer	=> '_clear_styles',
);

has '_temp_dir' =>(
	isa		=> 'File::Temp::Dir',
	writer	=> '_set_temp_dir',
	reader	=> '_get_temp_dir_object',
	clearer	=> '_clear_temp_dir',
	handles	=>{
		_get_temp_dir => 'dirname',
	},
);

has _worksheet_list =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_worksheet_list',
		clearer	=> '_clear_worksheet_list',
		reader	=> 'get_worksheet_names',
		handles	=>{
			worksheet_name => 'get',
		},
		default	=> sub{ [] },
	);

has _worksheet_lookup =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		writer	=> '_set_worksheet_lookup',
		clearer	=> '_clear_worksheet_lookup',
		reader	=> '_get_worksheet_lookup',
		handles	=>{
			_get_worksheet_info => 'get',
			number_of_sheets	=> 'count',
		},
		default	=> sub{ {} },
	);

has _current_worksheet_position =>(
		isa			=> Int,
		writer		=> '_set_current_worksheet_position',
		clearer		=> 'start_at_the_beginning',
		predicate	=> 'in_the_list',
	);
	
has _link_cleaner_array =>(
		isa		=> ArrayRef[CodeRef],
		clearer	=> '_clear_link_cleaner_array',
		writer	=> '_set_link_cleaner_array',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9
	
	

###########################################################################################
#
# _build_file()
#
# Unzip the .xlsx file and extract the generic information into perl data structures

sub _build_file{

    my ( $self, $file_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_build_file', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Arrived at _build_file for: ', $file_name ] );
	$self->_clear_shared_strings;
	$self->_clear_calc_chain;
	$self->_clear_styles;
	$self->_clear_temp_dir;
	$self->_clear_worksheet_list;
	$self->_clear_worksheet_lookup;
	$self->_clear_creator;
	$self->_clear_modified_by;
	$self->_clear_date_created;
	$self->_clear_date_modified;
	$self->_clear_error;
	$self->start_at_the_beginning;
	my	$message;
    if ( !-e $file_name ){
		$message = "Can't locate -$file_name-";
	}elsif( -z $file_name ) {
		$message = "There is nothing in the file -$file_name-";
	}
	if( $message ){
		$self->_set_error( $message );
		$self->_clear_file_name;
		return;
	}

    # Check for xls or encrypted OLE files.
    my $ole_file = $self->_check_if_ole_file( $file_name );
    return( undef ) if $ole_file;

    # Create a, locally scoped, temp dir to unzip the XLSX file into.
	my	$temp_dir = File::Temp->newdir( DIR => undef );
    $self->_set_temp_dir( $temp_dir );
	
    # Create an Archive::Zip object to unzip the XLSX file.
    my $zip_file = Archive::Zip->new();

    # Read the XLSX zip file and catch any errors.
    eval { $zip_file->read( $file_name ) };
    if ( $@ ) {
		$self->_set_error( "File has zip error(s): " . join( ' ~|~ ', $@ ) );
		return undef;
	}

    # Extract the XML files from the XLSX zip.
    $zip_file->extractTree( '', $self->_get_temp_dir . '/' );
	###LogSD	$phone->talk( level => 'trace', message =>[ "Zip file: ", $zip_file ] );
	my	$intermediate = $self->_get_temp_dir;
	###LogSD	$phone->talk( level	=> 'trace', message =>[
	###LogSD		"Temp dir: $intermediate", "Temp Dir contains: ", <$intermediate/*> ] );
	
	# Load general workbook information to this instance
	my	$workbook_file = $self->_get_temp_dir . '/xl/workbook.xml';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] );
	my ( $rel_lookup, $id_lookup ) = $self->_load_workbook_file( $workbook_file );
	return undef if !$rel_lookup;
	
	# Load the workbook rels file
	$workbook_file = $self->_get_temp_dir . '/xl/_rels/workbook.xml.rels';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] );
	my ( $pivot_lookup ) = $self->_load_rels_workbook_file( $rel_lookup, $workbook_file );
	return undef if !$pivot_lookup;
	
	# Load the docProps file
	$workbook_file = $self->_get_temp_dir . '/docProps/core.xml';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading $workbook_file"	] );
	$self->_load_doc_props_file( $workbook_file );
	#~ my $wait = <>;
	# Build the instances for all the shared files (data for sheets shared across worksheets)
	if( $self->_has_parser_type( $self->get_parser_type ) ){
		my	$result = 	$self->_set_shared_worksheet_files(
							$self->_get_module_list( $self->get_parser_type )
						);
		return undef if !$result;
	}else{
		$self->_set_error( "No definitions for the sheet parser type: " .
								$self->get_parser_type );
		return undef;
	}
	return $self;
}


###############################################################################
#
# _check_if_ole_file()
#
# Check if the file is an OLE compound doc. This can happen in a few cases.
# This first is when the file is xls and not xlsx. The second is when the
# file is an encrypted xlsx file. We also handle the case of unknown OLE
# file types.
#
# Porting note. As a lightweight test you can check for OLE files by looking
# for the magic number 0xD0CF11E0 (docfile0) at the start of the file.
#
sub _check_if_ole_file {

    my ( $self, $file_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_check_if_ole_file', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Arrived at _check_if_ole_file for: ', $file_name, ], );
    my	$ole = OLE::Storage_Lite->new( $file_name );
    my	$pps = $ole->getPpsTree();

    # If getPpsTree() failed then this isn't an OLE file.
    return if !$pps;

    # Loop through the PPS children below the root.
	my	$message;
    for my $child_pps ( @{ $pps->{Child} } ) {

        my 	$pps_name = OLE::Storage_Lite::Ucs2Asc( $child_pps->{Name} );
		
        # Match an Excel xls file.
        if ( $pps_name eq 'Workbook' || $pps_name eq 'Book' ) {
			$message = "File is xls not an xlsx file: $file_name";
			last;
        }elsif( $pps_name eq 'EncryptedPackage' ) {
			$message = "File is encrypted an encrypted xlsx file: $file_name";
			last;
        }
    }
	
	#Handle result
	if( $message ){
		$self->_set_error( $message );
	}else{
		###LogSD	$phone->talk( level => 'warn', message =>[
		###LogSD		'The OLE test passed (negative) for: ', $file_name, ], );
		return undef;
	}
	return 1;
}

sub _load_workbook_file{
	my( $self, $new_file ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_workbook_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the workbook file: $new_file" ] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	my ( $list, $sheet_ref, $rel_lookup, $id_lookup );
	my	$position = 0;
	my ( $setting_node ) = $dom->getElementsByTagName( 'workbookPr' );
	$self->_set_epoch_year( 1904 ) if $setting_node->getAttribute( 'date1904' );
	for my $sheet ( $dom->getElementsByTagName( 'sheet' ) ){
		my	$sheet_name = $sheet->getAttribute( 'name' );
		push @$list, $sheet_name;
		@{$sheet_ref->{$sheet_name}}{ 'sheet_id', 'rel_id', 'position' } = (
				$sheet->getAttribute( 'sheetId' ),
				$sheet->getAttribute( 'r:id' ),
				$position++,
		);
		$rel_lookup->{$sheet->getAttribute( 'r:id' )} = $sheet_name;
		$id_lookup->{$sheet->getAttribute( 'sheetId' )} = $sheet_name;
	}
	for my $sheet ( $dom->getElementsByTagName( 'pivotCache' ) ){
		my	$sheet_id = $sheet->getAttribute( 'cacheId' );
		my	$rel_id = $sheet->getAttribute( 'r:id' );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Sheet ID: $sheet_id", "Rel ID: $rel_id", ] );
		$rel_lookup->{$rel_id} = $sheet_id;
		$id_lookup->{$sheet_id} = $rel_id;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Sheet list: ", $list,
	###LogSD		"Worksheet lookup:", $sheet_ref,
	###LogSD		"rel lookup:", $rel_lookup,
	###LogSD		"id lookup:", $id_lookup,		] );
	if( !$list ){
		$self->_set_error( "No worksheets identified in this workbook" );
		return undef;
	}
	$self->_set_worksheet_list( $list );
	$self->_set_worksheet_lookup( $sheet_ref );
	return( $rel_lookup, $id_lookup );
}

sub _load_rels_workbook_file{
	my( $self, $rel_lookup, $new_file ) = @_;
	my ( $pivot_lookup, );
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_rels_workbook_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the rels/workbook file: $new_file",
	###LogSD			"With rel_lookup:", $rel_lookup								] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	my	$sheet_ref = $self->_get_worksheet_lookup;
	my	$found_file_names = 0;
	for my $sheet ( $dom->getElementsByTagName( 'Relationship' ) ){
		my	$rel_ID = $sheet->getAttribute( 'Id' );
		if( exists $rel_lookup->{$rel_ID} ){
			my	$file_name = $self->_get_temp_dir . '/xl/' . $sheet->getAttribute( 'Target' );
				$file_name =~ s/\\/\//g;
			if( $file_name =~ /worksheets/ ){
				$sheet_ref->{$rel_lookup->{$rel_ID}}->{file_name} = $file_name;
				$found_file_names = 1;
			}else{
				$pivot_lookup->{$rel_ID} = $file_name;
			}
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Worksheet lookup:", $sheet_ref,
	###LogSD		"Pivot lookup:", $pivot_lookup	] );
	if( !$found_file_names ){
		$self->_set_error( "Couldn't find any file names for the sheets" );
		return undef;
	}
	$self->_set_worksheet_lookup( $sheet_ref );
	return $pivot_lookup;
}

sub _load_doc_props_file{
	my( $self, $new_file ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_doc_props_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the doc props file: $new_file", ] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	$self->_set_creator( ($dom->getElementsByTagName( 'dc:creator' ))[0]->textContent() );
	$self->_set_modified_by( ($dom->getElementsByTagName( 'cp:lastModifiedBy' ))[0]->textContent() );
	$self->_set_date_created(
		DateTime::Format::Flexible->parse_datetime(
			($dom->getElementsByTagName( 'dcterms:created' ))[0]->textContent(),
			time_zone => 'floating'
		)
	);
	$self->_set_date_modified(
		DateTime::Format::Flexible->parse_datetime(
			($dom->getElementsByTagName( 'dcterms:modified' ))[0]->textContent(),
			time_zone => 'floating'
		)
	);
	###LogSD	$phone->talk( level => 'trace', message => [ "Current object:", $self ] );
	#~ my $wait = <>;
}

sub _set_shared_worksheet_files{
	my( $self, $object_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_set_shared_worksheet_files', );
	my	$temp_dir =$self->_get_temp_dir;
	my	@file_list = <$temp_dir/xl/*>;#/
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Building the shared worksheet files with the lookup ref:", $object_ref,
	###LogSD		"Reviewing files:", @file_list ] );
	my	$name_space = $self->get_log_space;
	for my $next ( @file_list ){
		if( $next =~ /[\/\\](([^\/\\]+)\.xml)$/ ){
			my	$name = $2;
			next if $name eq 'workbook';
			if( exists $object_ref->{$name} ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Loading the file: $next",
				###LogSD		"To the class: " . $object_ref->{$name}->{module},
				###LogSD		"With name: $name",
				###LogSD		"And log space: $name_space" ], );
				my	$sub_name;
				if( $object_ref->{$name}->{module} =~ /$name_space(.*)/ ){
					$sub_name = "$name_space$1";
				}else{
					$sub_name = $name;
				}
				###LogSD	$phone->talk( level => 'trace', message => [
				###LogSD		"Loading the base name space as: $sub_name", 
				###LogSD		"For package: " . __PACKAGE__ ], );
				my	$new_instance = $object_ref->{$name}->{module}->new(
										file_name 		=> $next,
										log_space 		=> $sub_name,
										_workbook_link	=> $self,
									);
				my	$attribute_setter = "_set_$name";
				$self->$attribute_setter( $new_instance );
			}else{
				$self->error( "No object available to load the -$name- xml file" );
				return undef;
			}
		}
	}
	return 1;
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Clearing the open files calcChain, sharedStrings, and styles files" ] );
	if( $self->_link_cleaner_array ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Attempting to clean up passed worksheets" ] );
		for my $sub ( @{$self->_link_cleaner_array} ){
			eval{ $sub->(); };
		}
	}
	$self->_clear_link_cleaner_array;
	$self->_clear_calc_chain;
	$self->_clear_shared_strings;
	$self->_clear_styles;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Clearing the Temporary Directory" ] );
	$self->_clear_temp_dir;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader - Read spreadsheet files with xlsx extentions


#~ ###############################################################################
#~ #
#~ # _check_files_exist()
#~ #
#~ # Verify that the subfiles read from the Content_Types actually exist;
#~ #
#~ sub _check_files_exist {

    #~ my $self    = shift;
    #~ my $tempdir = shift;
    #~ my %files   = @_;
    #~ my @filenames;

    #~ # Get the filenames for the files hash.
    #~ for my $key ( keys %files ) {
        #~ my $filename = $files{$key};

        #~ # Worksheets are stored in an aref.
        #~ if ( ref $filename ) {
            #~ push @filenames, @$filename;
        #~ }
        #~ else {
            #~ push @filenames, $filename;
        #~ }
    #~ }

    #~ # Verify that the files exist.
    #~ for my $filename ( @filenames ) {
        #~ if ( !-e $tempdir . $filename ) {
            #~ $self->{_error_extra_text} = $filename;
            #~ return;
        #~ }
    #~ }

    #~ return 1;
#~ }

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has '_temp_dir' =>(
	isa		=> 'File::Temp::Dir',
	writer	=> '_set_temp_dir',
	reader	=> '_get_temp_dir_object',
	clearer	=> '_clear_temp_dir',
	handles	=>{
		_get_temp_dir => 'dirname',
	},
);

has '_theme' =>(
	isa		=> Str,
	writer	=> '_set_theme',
	reader	=> '_get_theme',
	clearer	=> '_clear_theme',
);

has '_app' =>(
	isa		=> Str,
	writer	=> '_set_app',
	reader	=> '_get_app',
	clearer	=> '_clear_app',
);

has '_styles' =>(
	isa		=> Str,
	writer	=> '_set_styles',
	reader	=> '_get_styles',
	clearer	=> '_clear_styles',
);

has '_core' =>(
	isa		=> Str,
	writer	=> '_set_core',
	reader	=> '_get_core',
	clearer	=> '_clear_core',
);

has '_shared_strings' =>(
	isa		=> Str,
	writer	=> '_set_shared_strings',
	reader	=> '_get_shared_strings',
	clearer	=> '_clear_shared_strings',
);

has '_workbook' =>(
	isa		=> Str,
	writer	=> '_set_workbook',
	reader	=> '_get_workbook',
	clearer	=> '_clear_workbook',
);

has '_workbook_rels' =>(
	isa		=> Str,
	writer	=> '_set_workbook_rels',
	reader	=> '_get_workbook_rels',
	clearer	=> '_clear_workbook_rels',
);

has '_worksheets' =>(
	isa		=> ArrayRef,
	writer	=> '_set_worksheets',
	reader	=> '_get_worksheets',
	clearer	=> '_clear_worksheets',
	default => sub{ [] },
);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9


###############################################################################
#
# _check_if_ole_file()
#
# Check if the file in an OLE compound doc. This can happen in a few cases.
# This first is when the file is xls and not xlsx. The second is when the
# file is an encrypted xlsx file. We also handle the case of unknown OLE
# file types.
#
# Porting note. As a lightweight test you can check for OLE files by looking
# for the magic number 0xD0CF11E0 (docfile0) at the start of the file.
#
sub _check_if_ole_file {

    my ( $self, $file_name ) = @_;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk(
			level => 'info',
			message =>[ 'Arrived at _check_if_ole_file for: ', $file_name, ], );
    my	$ole = OLE::Storage_Lite->new( $file_name );
    my	$pps = $ole->getPpsTree();

    # If getPpsTree() failed then this isn't an OLE file.
    return if !$pps;

    # Loop through the PPS children below the root.
	my	$caught_failure = 0;
    for my $child_pps ( @{ $pps->{Child} } ) {

        my $pps_name = OLE::Storage_Lite::Ucs2Asc( $child_pps->{Name} );

        # Match an Excel xls file.
        if ( $pps_name eq 'Workbook' || $pps_name eq 'Book' ) {
			$phone->talk(
				level => 'warn',
				message =>[ 'File is xls not xlsx: ', $file_name, ], );
				$caught_failure = 1;
				last;
        }elsif( $pps_name eq 'EncryptedPackage' ) {
			$phone->talk(
				level => 'warn',
				message =>[ 'File is encrypted xlsx: ', $file_name, ], );
				$caught_failure = 1;
				last;
        }
    }
	
	$phone->talk(
		level => 'warn',
		message =>[ 'File is unknown OLE doc type: ', $file_name, ], );
	
	return 1;
}

sub _set_attribute{
	my ( $self, $attribute, $value ) = @_;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk(
			level => 'info',
			message =>[ "Setting attribute -$attribute- to value: $value" ], );
	my	$method = '_set' . $attribute;
	$self->$method( $value );
	$phone->talk( level => 'debug', message =>[ 'Action complete ...' ], );
}

sub _log_workbook{
	my ( $self, $attribute, $value ) = @_;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk(
			level => 'info',
			message =>[ "Setting attribute -$attribute- to value: $value" ], );
	my	$value_rels = $value;
		$value_rels =~ s{(workbook.xml)}{_rels/$1.rels};
	$self->_set_attribute( $attribute, $value );
	$self->_set_attribute( $attribute . '_rels', $value_rels );
	$phone->talk( level => 'debug', message =>[ 'Action complete ...' ], );
}

sub _push_attribute{
	my ( $self, $attribute, $value ) = @_;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk(
			level => 'info',
			message =>[ "Pushing value -$value- on to attribute: " . $attribute ], );
	my	$getter = '_get' . $attribute;
	my	$setter = '_set' . $attribute;
	my	$array_ref = $self->$getter;
	push @$array_ref, $value;
	$self->$setter( $array_ref );
	$phone->talk( level => 'debug', message =>[ 'Action complete for: ', $self->$getter ], );
}

my $xml_string = 
'<sst>
	<si>
		<r>
			<rPr>
				<sz val="11"/>
				<color rgb="FFFF0000"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t>Cell</t>
		</r>
		<r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t xml:space="preserve"> </t>
		</r>
		<r>
			<rPr>
				<b/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t>A2</t>
		</r>
	</si>
	<si>
		<r>
			<rPr>
				<sz val="11"/>
				<color rgb="FF00B0F0"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t>Cell</t>
		</r>
		<r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t xml:space="preserve"> </t>
		</r>
		<r>
			<rPr>
				<i/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t>B2</t>
		</r>
	</si>
</sst>';
sub _read_shared_strings{
	my ( $self, $attribute, $value ) = @_;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk(
			level => 'info',
			message =>[ "Putting the contents of -$value- in attribute: " . $attribute ], );
	#~ my $temp_dir = $self->_get_ temp_dir;
	#~ $temp_dir =~ s/\\/\//g;
	#~ print "$temp_dir\n";
	#~ print "$_\n" for <$temp_dir*>;
	my ( $string_list, );
	#~ my	$document 		= $parser->parse_file( ($self->_get_temp_dir . $value) );
						$parser->parse_chunk( $xml_string );
	my	$sst_node	= 	$parser->parse_chunk( "", 1 );
	#~ my ( $sst_node, ) 	= $document->getElementsByTagName( 'sst' );
	#~ my	$count			= $sst_node->getAttribute( 'count' );
	#~ my	$unique_count	= $sst_node->getAttribute( 'uniqueCount' );
	$phone->talk(
		level => 'info',
		message =>[	
			#~ "Opened xml sharedStrings with count -$count- and unique count: $unique_count",
			"Node: ", $sst_node->serialize() ], );#$xml_string, $sst_node,
	my	$one = 0;
	my ( $master_ref, $data_ref );
	for my $string_node ( $sst_node->getElementsByTagName( 'si' ) ){
		my	$string_data_ref;
		$phone->talk(
			level => 'trace',
			message =>[ "String node $one text : " . $string_node->serialize() ], );
		$data_ref = $self->_xml_to_data_ref( $string_node, $data_ref );
		push @$master_ref, $data_ref;
	#~ $self->$setter( $array_ref );
	#~ $phone->talk( level => 'debug', message =>[ 'Action complete for: ', $self->$getter ], );
	}
	$phone->talk(
			level => 'trace',
			message =>[ "Master Ref: ", $master_ref ], );
}

sub _xml_to_data_ref{
	my ( $self, $xml_node, $data_ref ) = @_;
	my	$message_ref = [ 	'Reached _xml_to_data_ref for node named -' .
							$xml_node->getName() . '- with string:',
							$xml_node->serialize() ];
	push( @$message_ref, ('Append to ref:', $data_ref) ) if $data_ref;
	my	$phone = Log::Shiras::Telephone->new;
		$phone->talk( level => 'info', message =>$message_ref );
	for my $child_node ( $xml_node->childNodes() ){
		$phone->talk( level => 'info', message =>[
			'Reached _xml_to_data_ref for node named -' . $child_node->getName() . 
			'- with string:' . $child_node->serialize() ], );
		if( $child_node->getName() eq 't' ){
			$data_ref->{t} = $child_node->textContent();
		}else{
			if( $child_node->textContent() ){
				$data_ref->{$child_node->getName()}->{value} = $child_node->textContent();
			}
			#~ my	$new_ref = $data_ref;
			$phone->talk( level => 'debug', message =>[
				'need to parse something besides a text node', $data_ref ], );
		}
		$phone->talk( level => 'fatal', message =>[
			'final parse:', $data_ref ], );
	}
	return $data_ref;
}
#~ ###############################################################################
#~ #
#~ # error().
#~ #
#~ # Return an error string for a failed read.
#~ #
#~ sub error {

    #~ my $self        = shift;
    #~ my $error_index = $self->{_error_status};
    #~ my $error       = $error_strings[$error_index];

    #~ if ( $self->{_error_extra_text} ) {
        #~ $error .= ': ' . $self->{_error_extra_text};
    #~ }

    #~ return $error;
#~ }


#~ ###############################################################################
#~ #
#~ # error_code().
#~ #
#~ # Return an error code for a failed read.
#~ #
#~ sub error_code {

    #~ my $self = shift;

    #~ return $self->{_error_status};
#~ }

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable(
	inline_constructor => 0,
);

1;
# The preceding line will help the module return a true value

#########1 main pod docs      3#########4#########5#########6#########7#########8#########9
__END__



=head1 NAME

Excel::Reader::XLSX - Efficient data reader for the Excel XLSX file format.

=head1 SYNOPSIS

The following is a simple Excel XLSX file reader using C<Excel::Reader::XLSX>:

    use strict;
    use warnings;
    use Excel::Reader::XLSX;

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }

    for my $worksheet ( $workbook->worksheets() ) {

        my $sheetname = $worksheet->name();

        print "Sheet = $sheetname\n";

        while ( my $row = $worksheet->next_row() ) {

            while ( my $cell = $row->next_cell() ) {

                my $row   = $cell->row();
                my $col   = $cell->col();
                my $value = $cell->value();

				print "  Cell ($row, $col) = ";
				print ((( $value ) ? $value : 'undef') . "\n");
            }
        }
    }

    __END__



=head1 DESCRIPTION

C<Excel::Reader::XLSX> is a fast and lightweight parser for Excel XLSX files. XLSX is the Office Open XML, OOXML, format used by Excel 2007 and later.

B<Note: This software is designated as alpha quality until this notice is removed.> The API shouldn't change but functionality is currently limited.

=head1 Reader

The C<Excel::Reader::XLSX> constructor returns a Reader object that is used to read an Excel XLSX file:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            while ( my $cell = $row->next_cell() ) {
                my $value = $cell->value();
                ...
            }
        }
    }

The C<Excel::Reader::XLSX> object is used to return sub-objects that represent the functional parts of an Excel spreadsheet, L</Workbook>, L</Worksheet>, L</Row> and L</Cell>:

     Reader
       +- Workbook
          +- Worksheet
             +- Row
                +- Cell

The C<Reader> object has the following methods:

    read_file()
    error()
    error_code()

=head2 read_file()

The C<read_file> Reader method is used to read an Excel XLSX file and return a C<Workbook> object:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    ...

It is recommended that the success of the C<read_file()> method is always checked using one of the error checking methods below.

=head2 error()

The C<error()> Reader method returns an error string if C<read_file()> fails:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }
    ...

The C<error()> strings and associated C<error_code()> numbers are:

    error()                              error_code()
    =======                              ============
    ''                                   0
    'File not found'                     1
    'File is xls not xlsx'               2
    'File is encrypted xlsx'             3
    'File is unknown OLE doc type'       4
    'File has zip error'                 5
    'File is missing subfile'            6
    'File has no [Content_Types].xml'    7
    'File has no workbook.xml'           8


=head2 error_code()

The C<error_code()> Reader method returns an error code if C<read_file()> fails:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die "Got error code ", $parser->error_code, "\n";
    }

This method is useful if you wish to use you own error strings or error handling methods.


=head1 Workbook

=head2 Workbook Methods

An C<Excel::Reader::XLSX> C<Workbook> object is returned by the Reader C<read_file()> method:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    ...

The C<Workbook> object has the following methods:

    worksheets()
    worksheet()

=head2 worksheets()

The Workbook C<worksheets()> method returns an array of
C<Worksheet> objects. This method is generally used to iterate through
all the worksheets in an Excel workbook and read the data:

    for my $worksheet ( $workbook->worksheets() ) {
      ...
    }


=head2 worksheet()

The Workbook C<worksheet()> method returns a single C<Worksheet>
object using the sheetname or the zero based index.

    my $worksheet = $workbook->worksheet( 'Sheet1' );

    # Or via the index.

    my $worksheet = $workbook->worksheet( 0 );


=head1 Worksheet

=head2 Worksheet Methods

The C<Worksheet> object is returned from a L</Workbook> object and is used to access row data.

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        ...
    }

The C<Worksheet> object has the following methods:

     next_row()
     name()
     index()

=head2 next_row()

The C<next_row()> method returns a L</Row> object representing the next
row in the worksheet.

        my $row = $worksheet->next_row();

It returns C<undef> if there are no more rows containing data or formatting in the worksheet. This allows you to iterate over all the rows in a worksheet as follows:

        while ( my $row = $worksheet->next_row() ) { ... }

Note, for efficiency the C<next_row()> method returns the next row in the file. This may not be the next sequential row. An option to read sequential rows, wheter they contain data or not will be added in a later release.

=head2 name()

The C<name()> method returns the name of the Worksheet object.

    my $sheetname = $worksheet->name();

=head2 index()

The C<index()> method returns the zero-based index of the Worksheet
object.

    my $sheet_index = $worksheet->index();


=head1 Row

=head2 Row Methods

The C<Row> object is returned from a L</Worksheet> object and is use to access cells in the worksheet.

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            ...
        }
    }

The C<Row> object has the following methods:

    values()
    next_cell()
    row_number()


=head2 values()

The C<values())> method returns an array of values for a row from the first column up to the last column containing data. Cells with no data value return an empty string C<''>.

    my @values = $row->values();

For example if we extracted data for the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    my $row = $worksheet->next_row();
    my @values = $row->values();
    ...

    # @values contains ( '', 'Foo', '', 'Bar' )


=head2 next_cell()

The C<next_cell> method returns the next, non-blank cell in the current row.

    my $cell = $row->next_cell();

It is usually used with a while loop. For example if we extracted data for the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    while ( my $cell = $row->next_cell() ) {
        my $value = $cell->value();
        print $value, "\n";
    }
    ...

    # Output:
    Foo
    Bar

Note, for efficiency the C<next_cell()> method returns the next cell in the row. This may not be the next sequential cell. An option to read sequential cells, wheter they contain data or not will be added in a later release.


=head2 row_number()

The C<row_number()> method returns the zero-indexed row number for the current row:

    my $row = $worksheet->next_row();
    print $row->row_number(), "\n";


=head1 Cell

=head2 Cell Methods

The C<Cell> object is used to extract data from Excel cells:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            while ( my $cell = $row->next_cell() ) {
                my $value = $cell->value();
               ...
            }
        }
    }

The C<Cell> object has the following methods:

    value()
    row()
    col()

For example if we extracted the data for the cells in the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    while ( my $row = $worksheet->next_row() ) {
        while ( my $cell = $row->next_cell() ) {
            my $row   = $cell->row();
            my $col   = $cell->col();
            my $value = $cell->value();

            print "Cell ($row, $col) = $value\n";
        }
    }
    ...

    # Output:
    Cell (0, 1) = Foo
    Cell (0, 2) = Bar


=head2 value()

The Cell C<value()> method returns the unformatted value from the cell.

    my $value = $cell->value();

The "value" of the cell can be a string or  a number. In the case of a formula it returns the result of the formula and not the formal string. For dates it returns the numeric serial date.


=head2 row()

The Cell C<row()> method returns the zero-indexed row number of the cell.

    my $row = $cell->row();


=head2 col()

The Cell C<col()> method returns the zero-indexed column number of the cell.

    my $col = $cell->col();


=head1 EXAMPLE

Simple example of iterating through all worksheets in a workbook and printing out values from cells that contain data.

    use strict;
    use warnings;
    use Excel::Reader::XLSX;

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }

    for my $worksheet ( $workbook->worksheets() ) {

        my $sheetname = $worksheet->name();

        print "Sheet = $sheetname\n";

        while ( my $row = $worksheet->next_row() ) {

            while ( my $cell = $row->next_cell() ) {

                my $row   = $cell->row();
                my $col   = $cell->col();
                my $value = $cell->value();

                print "  Cell ($row, $col) = $value\n";
            }
        }
    }

=head1 RATIONALE

The rationale for this module is to have a fast memory efficient module for reading XLSX files. This is based on my experience of user requirements as the maintainer of Spreadsheet::ParseExcel.


=head1 SEE ALSO

Spreadsheet::XLSX, an XLSX reader using the old Spreadsheet::ParseExcel hash based interface: L<http://search.cpan.org/dist/Spreadsheet-XLSX/>.

SimpleXlsx, a "rudimentary extension to allow parsing of information stored in Microsoft Excel XLSX spreadsheets": L<http://search.cpan.org/dist/SimpleXlsx/>.

Excel::Writer::XLSX, an XLSX file writer based on the Spreadsheet::WriteExcel interface: L<http://search.cpan.org/dist/Excel-Writer-XLSX/>.


=head1 TODO

There are a lot of features still to be added. This module is very much a work in progress.

=over

=item * Reading from filehandles.

=item * Option to read sequential rows via C<next_row()>.

=item * Option to read dates instead of raw serial style numbers. This is actually harder than it would seem due to the XLSX format.

=item * Option to read formulas, urls, comments, images.

=item * Spreadsheet::ParseExcel style interface.

=item * Direct cell access.

=item * Cell format data.

=back




=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org




=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.




=head1 DISCLAIMER OF WARRANTY

Because this software is licensed free of charge, there is no warranty for the software, to the extent permitted by applicable law. Except when otherwise stated in writing the copyright holders and/or other parties provide the software "as is" without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the software is with you. Should the software prove defective, you assume the cost of all necessary servicing, repair, or correction.

In no event unless required by applicable law or agreed to in writing will any copyright holder, or any other party who may modify and/or redistribute the software as permitted by the above licence, be liable to you for damages, including any general, special, incidental, or consequential damages arising out of the use or inability to use the software (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the software to operate with any other software), even if such holder or other party has been advised of the possibility of such damages.
