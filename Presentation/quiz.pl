use strict;
use warnings;
use DateTime::Format::Flexible;
print DateTime::Format::Flexible->parse_datetime( '12/31/2014' )->format_cldr( 'yyyy-MM' ) . "\n";