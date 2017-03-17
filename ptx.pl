#!/usr/bin/perl -w

use strict;
use Spreadsheet::ParseXLSX;

our $input;

if ( $#ARGV == 0 )
  { open( $input, "$ARGV[0]") || die "Could not open $ARGV[0]: $!\n"; }
elsif ( $#ARGV > 0 )
  { die "Too many arguments: ", $#ARGV+1, "\n"; }
else
  { open( $input, "-" ) || die "Could not open STDIN: $!\n"; }

my $parser   = Spreadsheet::ParseXLSX->new();
my $workbook = $parser->parse($input);

if ( !defined $workbook ) {
  die $parser->error(), ".\n";
  }

for my $worksheet ( $workbook->worksheets() ) {

  my ( $row_min, $row_max ) = $worksheet->row_range();
  my ( $col_min, $col_max ) = $worksheet->col_range();

  for my $row ( $row_min .. $row_max ) {
    for my $col ( $col_min .. $col_max ) {

      my $cell = $worksheet->get_cell( $row, $col );
      next unless $cell;

      print "Row, Col    = ($row, $col)\n";
      print "Type        = ", $cell->type(),        "\n";
      print "Value       = ", $cell->value(),       "\n";
      print "Unformatted = ", $cell->unformatted(), "\n";
      print "\n";
    }
  }
}
