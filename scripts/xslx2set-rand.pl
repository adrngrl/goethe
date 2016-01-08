#!/usr/bin/perl
use warnings;
use List::Util 'shuffle'; 
use Text::Iconv;
use File::Path qw(make_path);
use File::Spec;
use Spreadsheet::XLSX;
use File::Basename;


#my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
my $converter = Text::Iconv -> new ("utf-8", "utf-8");


# print help
if($ARGV[0] eq '-h' || $ARGV[0] eq '-help')
{
  help();
  exit;
}

# check for correct number of arguments
my($fnin, $sep, $setsize) = @ARGV;
if(@ARGV < 3)
{
  print "Wrong arguments. Type \"$0 -h\" for explanation.\n";
  exit;
}

my $excel = Spreadsheet::XLSX -> new ($fnin, $converter);

my @sheetarray = @{$excel -> {Worksheet}};
my $sheet = $sheetarray[0];

my @defs; # tablica ze wszystkimi definicjami

$sheet -> {MaxRow} ||= $sheet -> {MinRow}; 
foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
  $sheet -> {MaxCol} ||= $sheet -> {MinCol};
  my $line = "";
  foreach my $col ($sheet -> {MinCol} .. $sheet -> {MaxCol}) {        
    my $cell = $sheet -> {Cells} [$row] [$col];
    if ($cell) {
      my $celltext = $cell -> {Val};
      if($col == $sep)
      {
        $line = "${line}${celltext}\t";
      }
      else
      {
        $line = "${line}${celltext} ";
      }
    }
  }
  push @defs, $line;
}

@defs_rand = shuffle(@defs); # zmień losowo kolejność w tablicy

my $srcsize = scalar(@defs_rand);
my $remainder = $srcsize % $setsize;
my $fullsets = ($srcsize - $remainder) / $setsize;

my @numofdefs = ($setsize) x $fullsets;

printf("Total length of the main set is %d.\n", $srcsize);

if($remainder < 0.75*$setsize)
{
   printf("Generating %d shuffled subsets of: ", $fullsets);
   my $iterator = 0;
   while ($iterator < $remainder) 
   {
      $numofdefs[$iterator % $fullsets]++;
      $iterator++;
   }
}
else
{
   printf("Generating %d shuffled subsets of: ", $fullsets + 1);
   push @numofdefs, $remainder;
}

foreach (@numofdefs) {
  printf("%d ", $_);
}
print "elements.\n\n";

my $singlerow;

my $nmbase = fileparse($fnin, qr/\.[^.]*$/);

unlink "nmbase/*";

make_path($nmbase);
my $partindex = 0;
foreach my $defs (@numofdefs)
{
  print "Writing to $nmbase/${nmbase}-$partindex.txt\n";
  open(my $fh, '>', "$nmbase/${nmbase}-$partindex.txt")
    or die "Cant open file to write!\n";
  
  for(my $i = 0; $i < $defs; $i++)
  {
    $singlerow = pop @defs_rand;
    print $fh "${singlerow}\n";
  } 
  close $fh;
  $partindex++;
}


sub help
{
  my $usage = "  
Learning set generator for Quizlet;

Usage:
  $0 <src> <separator pos> <set size>

Parameters: 
  <src xlsx> source file with all entries to be processed
  <separator pos> index of the last column containing the definition
  <set size> desired size of the learning set\n\n";
  
  print $usage;
}