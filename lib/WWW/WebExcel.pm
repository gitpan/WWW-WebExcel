package WWW::WebExcel;

use 5.006;
use strict;
use warnings;
use Spreadsheet::WriteExcel;

require Exporter;

our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ();
our @EXPORT_OK   = ();
our @EXPORT      = qw();
our $VERSION     = '0.02';

sub new{
  my ($class,%opts) = @_;
  my $self = {};
  $self->{worksheets} = $opts{-worksheets} || [];
  $self->{type}       = 'application/vnd.ms-excel';
  bless($self,$class);
  return $self;
}# end new

sub add_worksheet{
  my ($self,@array) = @_;
  print "No Worksheet defined!" unless(defined $array[0]);
  push(@{$self->{worksheets}},[@array]);
}# end add_worksheet

sub del_worksheet{
  my ($self,$title) = @_;
  my @worksheets = grep{$_->[0] ne $title}@{$self->{worksheets}};
  $self->{worksheets} = [@worksheets];
}# end del_worksheet

sub add_row{
  my ($self,$title,$arref) = @_;
  foreach my $worksheet(@{$self->{worksheets}}){
    push(@{$worksheet->[1]->{'-data'}},$arref) if($worksheet->[0] eq $title);
  }
}# end add_data

sub set_headers{
  my ($self,$title,$arref) = @_;
  foreach my $worksheet(@{$self->{worksheets}}){
    $worksheet->[1]->{'-headers'} = $arref if($worksheet->[0] eq $title);
  }
}# end add_headers

sub output{
  my ($self) = @_;

  print "Content-type: ".$self->{type}."\n\n";
  my $EXCEL = new Spreadsheet::WriteExcel(\*STDOUT);

  foreach my $worksheet(@{$self->{worksheets}}){
    my $sheet = $EXCEL->addworksheet($worksheet->[0]);
    my $col = 0;
    my $row = 0;
    foreach(@{$worksheet->[1]->{-headers}}){
      $sheet->write($row,$col,$_);
      $col++;
    }
    $row++ if(scalar(@{$worksheet->[1]->{'-headers'}}) > 0);
    foreach my $data(@{$worksheet->[1]->{-data}}){
      $col = 0;
      foreach my $value(@$data){
        $sheet->write($row,$col,$value);
        $col++;
      }
      $row++;
    }
  }
  $EXCEL->close();
}# end output


# Preloaded methods go here.

1;
__END__
# Below is stub documentation for your module. You'd better edit it!

=head1 NAME

WWW::WebExcel - Perl extension for creating excel-files printed to STDOUT

=head1 SYNOPSIS

  use WWW::WebExcel;

  binmode(\*STDOUT);
  # data for spreadsheet
  my @header = qw(Header1 Header2);
  my @data   = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  # create a new instance
  my $excel = WWW::WebExcel->new();

  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_worksheet('Second Worksheet',{-data => \@data});
  $excel->add_worksheet('Test');

  # remove a worksheet
  $excel->del_worksheet('Test');

  # create the spreadsheet
  $excel->output();

  ## or

  # data
  my @data2  = (['Row1Col1', 'Row1Col2'],
                ['Row2Col1', 'Row2Col2']);

  my $worksheet = ['NAME',{-data => \@data2}];
  # create a new instance
  my $excel2    = WWW::WebExcel->new(-worksheets => [$worksheet]);

  # add headers to 'NAME'
  $excel2->set_headers('NAME',[qw/this is a test/]);
  # append data to 'NAME'
  $excel2->add_row('NAME',[qw/new row/]);

  $excel2->output();

=head1 DESCRIPTION

WWW::WebExcel simplifies the creation of excel-files in the web. It does not
provide any access to cell-formats yet. This is just a raw version that will be
extended within the next few weeks.

=head1 METHODS

=head2 new

  # create a new instance
  my $excel = WWW::WebExcel->new();

  # or

  my $worksheet = ['NAME',{-data => ['This','is','an','Test']}];
  my $excel2    = WWW::WebExcel->new(-worksheets => [$worksheet]);

=head2 add_worksheet

  # add worksheets
  $excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
  $excel->add_worksheet('Second Worksheet',{-data => \@data});
  $excel->add_worksheet('Test');

The first parameter of this method is the name of the worksheet and the second one is
a hash with (optional) information about the headlines and the data.

=head2 del_worksheet

  # remove a worksheet
  $excel->del_worksheet('Test');

Deletes all worksheets named like the first parameter

=head2 add_row

  # append data to 'NAME'
  $excel->add_row('NAME',[qw/new row/]);

Adds a new row to the worksheet named 'NAME'

=head2 set_headers

  # add headers to 'NAME'
  $excel->set_headers('NAME',[qw/this is a test/]);

set the headers for the worksheet named 'NAME'

=head2 output

  $excel2->output();

prints the worksheet to the STDOUT.

=head1 DEPENDENCIES

This module requires Spreadsheet::WriteExcel

=head1 BUGS

I'm sure there are some bugs in this module. Feel free to contact me if you
experienced any problem.

=head1 ToDo

* add formats to cell
* write spreadsheet into file
* add data (rows) to worksheet
* add headers to worksheet (replace the existing list)
* widen range of headers (add headers to the existing ones)
* widen range of data (add cols to data)

=head1 SEE ALSO

Spreadsheet::WriteExcel

=head1 AUTHOR

Renee Baecker, E<lt>module@renee-baecker.deE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2004 by Renee Baecker

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.1 or,
at your option, any later version of Perl 5 you may have available.


=cut
