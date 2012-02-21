package SimpleXlsx;
    
use strict;
use warnings;
use Archive::Zip qw( :ERROR_CODES );
use XML::Simple;
use File::Basename;
use Data::Dumper;

require Exporter;

our @ISA = qw(Exporter);

# Items to export into callers namespace by default. Note: do not export
# names by default without a very good reason. Use EXPORT_OK instead.
# Do not simply export all your public functions/methods/constants.

our @EXPORT_OK = ( 'parse', 'as_HTML' );

our $VERSION = '0.05';

our %indexedColors2ARGB = (0=>'00000000', 1=>'00FFFFFF', 2=>'00FF0000', 3=>'0000FF00', 4=>'000000FF', 5=>'00FFFF00', 6=>'00FF00FF', 7=>'0000FFFF', 8=>'00000000', 9=>'00FFFFFF', 10=>'00FF0000', 11=>'0000FF00', 12=>'000000FF', 13=>'00FFFF00', 14=>'00FF00FF', 15=>'0000FFFF', 16=>'00800000', 17=>'00008000', 18=>'00000080', 19=>'00808000', 20=>'00800080', 21=>'00008080', 22=>'00C0C0C0', 23=>'00808080', 24=>'009999FF', 25=>'00993366', 26=>'00FFFFCC', 27=>'00CCFFFF', 28=>'00660066', 29=>'00FF8080', 30=>'000066CC', 31=>'00CCCCFF', 32=>'00000080', 33=>'00FF00FF', 34=>'00FFFF00', 35=>'0000FFFF', 36=>'00800080', 37=>'00800000', 38=>'00008080', 39=>'000000FF', 40=>'0000CCFF', 41=>'00CCFFFF', 42=>'00CCFFCC', 43=>'00FFFF99', 44=>'0099CCFF', 45=>'00FF99CC', 46=>'00CC99FF', 47=>'00FFCC99', 48=>'003366FF', 49=>'0033CCCC', 50=>'0099CC00', 51=>'00FFCC00', 52=>'00FF9900', 53=>'00FF6600', 54=>'00666699', 55=>'00969696', 56=>'00003366', 57=>'00339966', 58=>'00003300', 59=>'00333300', 60=>'00993300', 61=>'00993366', 62=>'00333399', 63=>'00333333');
our %CellDataTypes      = ('b'=>'Boolean', 'd'=>'Date', 'e'=>'Error', 'inlineStr'=>'Inline String', 'n'=>'Number', 's'=>'Shared String', 'str'=>'String');
our %HtmlBorderStyles   = (map {$_=>1} qw(none hidden dotted dashed solid double groove ridge inset outset));

sub new
{
  my $package = shift;

  my $self = {};
  $self->{zip} = Archive::Zip->new();
  
  return bless $self, $package;
}
sub d { $Data::Dumper::Deepcopy = 0; print Data::Dumper::Dumper @_ }

sub zip
{
  return $_[0]->{zip};
}

sub parse
{
  my($self, $xfile) = @_;
  
  my($ret) = $self->zip->read($xfile);
  unless ($ret == AZ_OK)
  {
    warn "Unable to read file \"$xfile\" ($!)\n";
    return undef;
  }

  my $workbook = $self->getWorkbook($xfile);
  my $workbookRelations = $self->getWorkbookRelations($xfile);
  my @zWorksheets = $self->getWorksheets($workbookRelations, $workbook);
  my @strings     = $self->getValues($workbookRelations);
  my($styles)     = $self->getStyles($workbookRelations);

  my(%worksheets);
  my(@sheetNames);
  
  $worksheets{'Worksheets'} = [];
  
  $worksheets{'Total Worksheets'} = ($#zWorksheets + 1);
  for my $zWorksheet (@zWorksheets)
  {
    my $file = $zWorksheet->{z};

    my $sRelations = $self->getSheetRelations($file); $sRelations = [$sRelations] if ref $sRelations ne 'ARRAY';
    my $sComments  = $self->getSheetComments($sRelations);

    my(%worksheet);
    $worksheet{'Rows'}  = [];
    $worksheet{'Data'}  = {};
    $worksheet{'Merge'} = {};
        
    my($name) = basename($file->fileName());
    $name =~ s/\.xml$//;
    
    my($contents) = $file->contents();
    my($xml) = new XML::Simple;
    my($data) = $xml->XMLin($contents);
    
    my $sData = $data->{'sheetData'}->{'row'} || [];
       $sData = [$sData] if ref $sData ne 'ARRAY';
    if (@$sData) {
      my %tcols;
      for my $col (0 .. $#{$sData->[0]->{'c'}})
      {
        (my $key = $sData->[0]->{'c'}->[$col]->{'r'}) =~ s|\d+$||o;
        $tcols{$key} = $col;
      }
      $worksheet{'Columns'} = \%tcols;
      
      my(@trow);
      my(%tdata);
      for my $row (0 .. $#{$sData})
      {
        my($cols) = $sData->[$row]->{'c'};
        $cols = [$cols] if $cols and ref $cols ne 'ARRAY';
        
        my(@rdata);
        for my $col (0 .. $#{$cols})
        {
          my $xCell     = $cols->[$col] || {};
          my $xCellData = {};
          $xCellData->{Data}     = !defined $xCell->{'v'} ? undef :
                                    defined $xCell->{'t'} && $xCell->{'t'} eq 's' && defined $xCell->{'v'} ? $strings[$xCell->{'v'}] :
                                   {Text=>$xCell->{'v'}};
          $xCellData->{Style}    = exists $xCell->{'s'} ? $styles->[$xCell->{'s'}] : undef;
          $xCellData->{Comments} = $sComments->{$xCell->{'r'}};
          $xCellData->{DataType} = exists $xCell->{'t'} ? $CellDataTypes{$xCell->{'t'}} : undef;        

          push @rdata, $xCellData;
        }


        if (exists $sData->[$row]->{'r'}) # Row Index, §18.3.1.73
        {
          push @trow, $sData->[$row]->{'r'};
          $tdata{$sData->[$row]->{'r'}}{'Data'} = \@rdata;
          
          if (exists $sData->[$row]->{'s'})
          {
            $tdata{$sData->[$row]->{'r'}}{'Style'} = $styles->[$sData->[$row]->{'s'}];
          }
        }
      }
      
      $worksheet{'Rows'} = \@trow;
      $worksheet{'Data'} = \%tdata;
    }


    my $sMerge = $data->{'mergeCells'}->{'mergeCell'} || [];
       $sMerge = [$sMerge] if ref $sMerge ne 'ARRAY';
    if (@$sMerge)
    {
      my(%merge);
      for my $mc (@{$sMerge})
      {
        my($from, $to) = split(':', $mc->{'ref'});
        
        $from =~ /([a-zA-Z]+)([0-9]+)/;
        my($col1, $row1) = ($1, $2);
        
        $to =~ /([a-zA-Z]+)([0-9]+)/;
        my($col2, $row2) = ($1, $2);

        push @{$merge{$row1}},
        {
          'From' => { 'Row' => $row1, 'Column' => $col1 },
          'To' => { 'Row' => $row2, 'Column' => $col2 }
        };

        my @rowNumbers  = ($row1 .. $row2);
        my @cellLetters = ($col1 .. $col2);
        my $dominateCell;
        for my $rowNumber (@rowNumbers)
        {
          for my $colLetter (@cellLetters)
          {
            my $colNumber = $worksheet{'Columns'}{$colLetter};
            my $cell = $worksheet{'Data'}{$rowNumber}{'Data'}[$colNumber];
            $cell->{'Merged'} = 1;

            $dominateCell ||= $cell;
          }
        }
        $dominateCell->{'Merged'} = {};
        $dominateCell->{'Merged'}{'Cols'} = scalar @cellLetters;
        $dominateCell->{'Merged'}{'Rows'} = scalar @rowNumbers;
        $dominateCell->{'Merged'}{'_cols'} = \@cellLetters;
        $dominateCell->{'Merged'}{'_rows'} = \@rowNumbers;
      }

      $worksheet{'Merge'} = \%merge;
    }

    $worksheet{'Name'} = $zWorksheet->{'name'};
    
    $worksheets{$name} = \%worksheet;

    push @sheetNames, $name;
  }
  
  $worksheets{'Worksheets'} = \@sheetNames;

  return \%worksheets;;
}

sub getValues
{
  my ($self, $relations) = @_;
  
  my(@zStrings) = map {$self->zip->membersMatching('^xl/'.$_->{'Target'})}
                  grep {$_->{'Type'} =~ m|relationships/sharedStrings$|} @$relations;
  if ($#zStrings > 0)
  {
    warn "Error: Multiple shared strings are not [yet] supported\n";
  }
  
  my($xml) = new XML::Simple;
  my($sstrings) = $zStrings[0];
  $sstrings = $sstrings->contents();
  my($tstrings) = $xml->XMLin($sstrings);
  
  my(@strings);
  for my $idx (0 .. $#{$tstrings->{'si'}})
  { # §18.4.8 string Item
    my @string;

    my $si = $tstrings->{'si'}->[$idx]; $si = [$si] if $si && ref $si ne 'ARRAY';
    foreach my $child (@$si)
    {
      if (exists $child->{'t'})
      {
        push @string, {'Text' => ref $child->{'t'} ? $child->{'t'}{'content'} || ' ' : $child->{'t'}};
      }
      if (exists $child->{'r'})
      {
        my $childR = $child->{'r'}; $childR = [$childR] if ref $childR ne 'ARRAY';
        foreach my $r (@$childR)
        {
          push @string, $self->_parseRichTextRun($r);
        }
        push @string, {};
      }
    }

    push @strings, \@string;
  }
  
  return @strings;
}

sub getWorkbook
{
  my ($self) = @_;

  my @zWorkbooks = $self->zip->membersMatching('^xl/workbook.xml');
  my $data = [];

  if ($zWorkbooks[0])
  {
    ($data) = $zWorkbooks[0]->contents();

    my($xml) = new XML::Simple;
    $data = $xml->XMLin($data);
  }

  return $data;
}

sub getWorkbookRelations
{
  my ($self) = @_;

  my $data = $self->_parseRelations('^xl/_rels/workbook.xml.rels');

  return $data;
}

sub getWorksheets
{
  my ($self, $relations, $workbook) = @_;

  my $wbSheets = $workbook->{'sheets'}{'sheet'};
  my @sheets;
  foreach my $rl (grep {$_->{'Type'} =~ m|relationships/worksheet$|} @$relations)
  {
    my @zSheet    = $self->zip->membersMatching('^xl/'.$rl->{'Target'});
    my $sheetName = undef;
    if (exists $wbSheets->{'name'})
    {
      $sheetName = $wbSheets->{'name'};
    }
    else
    {
      foreach my $wbSheetName (keys %$wbSheets)
      {
        if ($wbSheets->{$wbSheetName}{'r:id'} eq $rl->{'Id'})
        {
          $sheetName = $wbSheetName;
        }
      }
    }
    push @sheets, {z=>$zSheet[0], name=>$sheetName};
  }

  return @sheets;
}

sub getSheetRelations
{
  my ($self, $file) = @_;
  
  my($name) = basename($file->fileName());
  my $data = $self->_parseRelations('^xl/worksheets/_rels/'.$name.'.rels');
      
  return $data;
}

sub getSheetComments
{
  my ($self, $sRelations) = @_;
  
  my %comments;

  my($rCommentsFile) = grep {$_->{'Type'} =~ m|relationships/comments$|} @$sRelations;
  if ($rCommentsFile)
  {
    my $zCommentsPath = '^xl/worksheets/'.$rCommentsFile->{'Target'};
    $zCommentsPath =~ s|/.+?/\.\./|/|go;
    
    my(@zComments) = $self->zip->membersMatching($zCommentsPath);
    
    if ($zComments[0])
    {
      my($data) = $zComments[0]->contents();
      my($xml) = new XML::Simple;
      $data = $xml->XMLin($data);

      my($authors)  = $data->{'authors'}{'author'};      $authors  = [$authors] if ref $authors ne 'ARRAY';
      my($comments) = $data->{'commentList'}{'comment'}; $comments = [$comments] if ref $comments ne 'ARRAY';
      
      foreach my $c (@$comments)
      {
        # §18.7.*
        $comments{$c->{'ref'}} = {};
        $comments{$c->{'ref'}}{'Author'} = $authors->[$c->{'authorId'}];
        $comments{$c->{'ref'}}{'Text'} = [];
        
        # §18.7.7
        if (exists $c->{'text'}{'t'})
        {  # (Text) 
          push @{$comments{$c->{'ref'}}{'Text'}}, $_ for ref $c->{'text'}{'t'} eq 'ARRAY' ? @{$c->{'text'}{'t'}} : $c->{'text'}{'t'};
        }
        elsif (exists $c->{'text'}{'r'})
        {  # (Rich Text Run)
          foreach my $r (ref $c->{'text'}{'r'} eq 'ARRAY' ? @{$c->{'text'}{'r'}} : $c->{'text'}{'r'})
          {
            push @{$comments{$c->{'ref'}}{'Text'}}, $self->_parseRichTextRun($r)->{'Text'};
          }
        }
      }
    }    
  }
  
  return \%comments;
}

sub getStyles
{
  my ($self, $relations) = @_;

  my(@zStyles) = map {$self->zip->membersMatching('^xl/'.$_->{'Target'})}
                 grep {$_->{'Type'} =~ m|relationships/styles$|} @$relations;
  if ($#zStyles > 0)
  {
    warn "Error: Multiple shared strings are not [yet] supported\n";
  }     

  my($data) = $zStyles[0]->contents();
  
  my($xml) = new XML::Simple;
  $data = $xml->XMLin($data);
  
  my(@cellFormats);
  my(%fonts);
  my(%fills);
  my(%borders);
  
  my($xcellFormats) = $data->{'cellXfs'}{'xf'}; $xcellFormats = [$xcellFormats] if ref $xcellFormats ne 'ARRAY';
  my($xfonts) = $data->{'fonts'}{'font'};       $xfonts   = [$xfonts] if ref $xfonts ne 'ARRAY';
  my($xfills) = $data->{'fills'}{'fill'};       $xfills   = [$xfills] if ref $xfills ne 'ARRAY';
  my($xborders) = $data->{'borders'}{'border'}; $xborders = [$xborders] if ref $xborders ne 'ARRAY';

  my($idx) = 0;
  
  # §18.8.22
  for my $ind (0 .. $#{$xfonts})
  {
    $fonts{$idx} = $self->_parseRunProperties($xfonts->[$ind]);
    if ($xfonts->[$ind]{'name'})
    {
      $fonts{$idx}{'font-name'} = $xfonts->[$ind]{'name'}{'val'};
    }
    
    $idx++;
  }
  
  # §18.8.5
  $idx = 0;
  for my $ind (0 .. $#{$xborders})
  {
    $borders{$idx} = {};
    foreach my $side (qw(left right top bottom diagonal))
    {
      if (exists $xborders->[$ind]{$side}{'color'})
      {
        $borders{$idx}{$side}{'color'} = $self->_color2RGB($xborders->[$ind]{$side}{'color'});
      }
      if (exists $xborders->[$ind]{$side}{'style'})
      {
        $borders{$idx}{$side}{'style'} = $xborders->[$ind]{$side}{'style'};
      }
    }

    $idx++;
  }
  
  # §18.8.20
  $idx = 0;
  for my $ind (0 .. $#{$xfills})
  {
    $fills{$idx} = undef;
    if (exists $xfills->[$ind]{'patternFill'})
    {
      $fills{$idx}{'bgColor'} = $self->_color2RGB($xfills->[$ind]{'patternFill'}{'bgColor'});
      $fills{$idx}{'fgColor'} = $self->_color2RGB($xfills->[$ind]{'patternFill'}{'fgColor'});
    }
    
    $idx++;
  }

  # §18.8.45
  for my $ind (0 .. $#{$xcellFormats})
  {
    my $cFormat = {%{$fonts{$xcellFormats->[$ind]->{'fontId'}} || {}}}; # copy hash
    if (defined $xcellFormats->[$ind]->{'borderId'} && $borders{$xcellFormats->[$ind]->{'borderId'}} && %{$borders{$xcellFormats->[$ind]->{'borderId'}}})
    {
      $cFormat->{'border'} = $borders{$xcellFormats->[$ind]->{'borderId'}};
    }
    if ($xcellFormats->[$ind]{'alignment'})
    {
      if ($xcellFormats->[$ind]{'alignment'}{'horizontal'})
      {
         $cFormat->{'text-align'} = $xcellFormats->[$ind]{'alignment'}{'horizontal'};
      }
      if ($xcellFormats->[$ind]{'alignment'}{'vertical'})
      {
         $cFormat->{'vertical-align'} = $xcellFormats->[$ind]{'alignment'}{'vertical'};
      }
      if (exists $xcellFormats->[$ind]{'alignment'}{'wrapText'} && $xcellFormats->[$ind]{'alignment'}{'wrapText'} == 0)
      {
         $cFormat->{'vertical-align'} = 'nowrap';
      }
    }
    if ($fills{$xcellFormats->[$ind]->{'fillId'}}{'bgColor'}) {
      $cFormat->{'background-color'} = $fills{$xcellFormats->[$ind]->{'fillId'}}{'bgColor'};
    }
    elsif ($fills{$xcellFormats->[$ind]->{'fillId'}}{'fgColor'}) {
      $cFormat->{'background-color'} = $fills{$xcellFormats->[$ind]->{'fillId'}}{'fgColor'};
    }

    push @cellFormats, $cFormat;

  }

  return \@cellFormats;
}

sub _parseRelations
{
  my ($self, $filename) = @_;

  my @zRelations = $self->zip->membersMatching($filename);
  my $data = [];

  if ($zRelations[0])
  {
    ($data) = $zRelations[0]->contents();

    my($xml) = new XML::Simple;
    $data = $xml->XMLin($data);
    $data = $data->{'Relationship'} || [];
  }

  return $data;
}

sub _parseRichTextRun
{
  my ($self, $r) = @_;
  
  my %t;
  if (exists $r->{'t'})
  {
    $t{Text} = ref $r->{'t'} ? $r->{'t'}{'content'} || ' ' : $r->{'t'};
  }
  if (exists $r->{'rPr'})
  {
    $t{'Style'} = $self->_parseRunProperties($r->{'rPr'});
  }

  return \%t;
}

sub _parseRunProperties
{
  # §18.4.7
  my ($self, $rPr) = @_;

  my $style = {};

  $style->{'bold'}++        if exists $rPr->{'b'};
  $style->{'italic'}++      if exists $rPr->{'i'};
  $style->{'strike'}++      if exists $rPr->{'strike'};
  $style->{'text-shadow'}++ if exists $rPr->{'shadow'};
  $style->{'outline'}++     if exists $rPr->{'outline'};
  $style->{'vertical-align'}++ if exists $rPr->{'vertAlign'};
  $style->{'font-size'} = $rPr->{'sz'}{'val'}    if exists $rPr->{'sz'};
  $style->{'font-name'} = $rPr->{'rFont'}{'val'} if exists $rPr->{'rFont'};
  if (exists $rPr->{'color'} && $rPr->{'color'}{'indexed'})
  {
    $style->{'color'} = $self->_color2RGB($rPr->{'color'});
  }
  unless ($style->{'color'})
  {
    #
    # toDo: for know correct default text color
    #       we must use xfId->cellStyleXfs->fontId->(color=>theme)->xl/theme/theme1.xml->sysClr val="windowText"
    #
    # sorry, now 000000 - default text color
    #
    $style->{'color'} = '000000';
  }
  return $style;
}

sub _color2RGB
{
  my ($self, $color) = @_;

  my $c = $color->{'indexed'} ? $self->_indexedColors2RGB($color->{'indexed'}) :
          $color->{'rgb'}     ? $self->_argbColors2RGB($color->{'rgb'}) :
          '';

  return $c;
}

sub _indexedColors2RGB
{
  my $c = $indexedColors2ARGB{$_[1]} || '';
     $c =~ s|^00||o;
  return $c;
}

sub _argbColors2RGB
{
  my $c = $_[1];
     $c =~ s|^\w{2}||o;
  return $c;
}

sub as_HTML
{
  #
  # opts: tagName, containerTagName, notEmbedComment
  #
  my($self, $xfile, %opts) = @_;

  my $xlsx = $self->parse($xfile);

  my @tables;
  foreach my $sheetName (@{$xlsx->{'Worksheets'}})
  {
    my $sheet = $xlsx->{$sheetName};
    my @rows;

    foreach my $rNum (@{$sheet->{'Rows'}})
    {
      my $rowData = $sheet->{'Data'}{$rNum} || {};
      my @cells;

      foreach my $cell (@{$rowData->{'Data'} || []})
      {
        my $cellHTML = $self->cell_as_HTML($cell, %opts);
        if ($cellHTML ne '___merged___')
        {
          push @cells, $cellHTML;
        }
      }

      push @rows, \@cells;
    }
    
    if (@rows)
    {
      push @tables, \@rows;
    }
  }


  my @html;
  $opts{'containerTagName'} ||= !$opts{'tagName'} || $opts{'tagName'} eq 'td' ? 'tr' : 'div';
  foreach my $table (@tables)
  {
    my @str;
    for my $row (@$table)
    {
      push @str, '<'.$opts{'containerTagName'}.'>'.join('', @$row).'</'.$opts{'containerTagName'}.'>';
    }
    
    my $html = join("\n", @str);
    if ($opts{'containerTagName'} eq 'tr')
    {
      $html = '<table><tbody>'.$html.'</tbody></table>';
    }
    push @html, $html;
  }

  return wantarray ? @html : join("\n", @html);
}

sub cell_as_HTML
{
  #
  # opts: tagName, notEmbedComment
  #
  my ($self, $cell, %opts) = @_;

  $cell->{'Data'} ||= [];
  $cell->{'Data'}   = [$cell->{'Data'}] if ref $cell->{'Data'} ne 'ARRAY';

  my %tagAttrs;

  if ($cell->{'Merged'})
  {
    if ($cell->{'Merged'} == 1)
    {
      return '___merged___';
    }
    else
    {
      $tagAttrs{'colspan'} = $cell->{'Merged'}{'Cols'} if $cell->{'Merged'}{'Cols'} > 1;
      $tagAttrs{'rowspan'} = $cell->{'Merged'}{'Rows'} if $cell->{'Merged'}{'Rows'} > 1;
    }
  }

  my @text;
  foreach my $t (@{$cell->{'Data'} || []})
  {
    my %tStyle;

    if ($t->{'Style'})
    {
      %tStyle = %{$t->{'Style'}};

      $tStyle{'color'}            = '#'.$tStyle{'color'} if $tStyle{'color'} && $tStyle{'color'} !~ /^#/;
      $tStyle{'background-color'} = '#'.$tStyle{'background-color'} if $tStyle{'background-color'} && $tStyle{'background-color'} !~ /^#/;
      $tStyle{'font-weight'}      = 'bold'         if delete $tStyle{'bold'};
      $tStyle{'font-style'}       = 'italic'       if delete $tStyle{'italic'};
      $tStyle{'text-decoration'}  = 'line-through' if delete $tStyle{'strike'};
      $tStyle{'font-size'}       .= 'pt'           if $tStyle{'font-size'} !~ /pt$/;

      if (my $font = delete $tStyle{'font-name'})
      {
        $tStyle{'font-family'} = $font;
      }

      if ($cell->{Style})
      {
        if (($cell->{'Style'}{'font-weight'} || $cell->{'Style'}{'bold'}) && !$tStyle{'font-weight'})
        {
          # psevdo-Default font-weight
          $tStyle{'font-weight'} = 'normal';
        }
        if (($cell->{'Style'}{'font-style'} || $cell->{'Style'}{'italic'}) && !$tStyle{'font-style'})
        {
          # psevdo-Default font-style
          $tStyle{'font-style'} = 'normal';
        }
        if (($cell->{'Style'}{'text-decoration'} || $cell->{'Style'}{'strike'}) && !$tStyle{'text-decoration'})
        {
          # psevdo-Default text-decoration
          $tStyle{'text-decoration'} = 'none';
        }
      }
    }

    my $tVal = $t->{'Text'} || '';
    if (%tStyle)
    {
      $tVal = '<span style="'.join(';', map {$_.':'.$tStyle{$_}} keys %tStyle).'">'.$tVal.'</span>';
    }

    push @text, $tVal;
  }
  
  my $cellText = @text ? join("\n", @text) : '';
  
  if ($cell->{'Style'} && $cellText =~ /\S/o)
  {
    #
    # no sense to write node style if content is empty
    #

    $cell->{'Style'}{'color'}            = '#'.$cell->{'Style'}{'color'} if $cell->{'Style'}{'color'} && $cell->{'Style'}{'color'} !~ /^#/;
    $cell->{'Style'}{'background-color'} = '#'.$cell->{'Style'}{'background-color'} if $cell->{'Style'}{'background-color'} && $cell->{'Style'}{'background-color'} !~ /^#/;
    $cell->{'Style'}{'font-style'}       = 'italic'       if delete $cell->{'Style'}{'italic'};
    $cell->{'Style'}{'font-weight'}      = 'bold'         if delete $cell->{'Style'}{'bold'};
    $cell->{'Style'}{'text-decoration'}  = 'line-through' if delete $cell->{'Style'}{'strike'};
    $cell->{'Style'}{'font-size'}       .= 'pt'           if $cell->{'Style'}{'font-size'} !~ /pt$/;

    if (my $font = delete $cell->{'Style'}{'font-name'})
    {
      $cell->{'Style'}{'font-family'} = $font;
    }
    
    if ($cell->{'Style'}{'border'} && ref $cell->{'Style'}{'border'})
    {
      my $border = delete $cell->{'Style'}{'border'};
      for my $b (keys %$border)
      {
        $cell->{'Style'}{"border-$b"} = join ' ', '1px', 
                                                  ($border->{$b}{'style'} && exists $HtmlBorderStyles{$border->{$b}{'style'}} ? $HtmlBorderStyles{$border->{$b}{'style'}} : 'solid'),
                                                  ($border->{$b}{'color'} ? '#'.$border->{$b}{'color'} : 'black');
      }
      
      if ($cell->{'Style'}{'border-top'} eq $cell->{'Style'}{'border-right'} && $cell->{'Style'}{'border-top'} eq $cell->{'Style'}{'border-bottom'} && $cell->{'Style'}{'border-top'} eq $cell->{'Style'}{'border-left'})
      {
        #
        # merge border if it possible
        #
        $cell->{'Style'}{'border'} = $cell->{'Style'}{'border-top'};
        for my $b (keys %$border)
        {
          delete $cell->{'Style'}{"border-$b"};
        }
      }
    }
    
    $tagAttrs{style} = join(';', map {$_.':'.$cell->{'Style'}{$_}} keys %{$cell->{'Style'}});
  }

  my $comment = join "\n", @{$cell->{'Comments'}{'Text'} || []};
  $comment =~ s|\s+$||gmso;
  if ($comment =~ /\S/o && !$opts{'notEmbedComment'})
  {
    $comment =~ s|"|\"|go;
    $comment =~ s|>|&gt;|go;
    $comment =~ s|<|&lt;|go;

    $tagAttrs{'title'} = $comment;
    $tagAttrs{'class'} = 'commented';
  }


  my $tagName = $opts{'tagName'} || 'td';
  
  my $cellHTML = '<'.$tagName.' '.join(' ', map {qq|$_="$tagAttrs{$_}"|} grep {$tagAttrs{$_} =~ /\S/o} keys %tagAttrs).'>'.$cellText.'</'.$tagName.'>';
  
  return wantarray ? ($cellHTML, $comment) : $cellHTML;
}


1;

__END__

=head1 NAME

SimpleXlsx - Perl extension to read data from a Microsoft Excel 2007 XLSX file

=head1 SYNOPSIS

  use SimpleXlsx;
  
  my($xlsx) = SimpleXlsx->new();
  my($worksheets) = $xlsx->parse('/path/to/workbook.xlsx');

=head1 DESCRIPTION

SimpleXlsx is a rudamentary extension to allow parsing information stored in
Microsoft Excel XLSX spreadsheets.

=head2 EXPORT

None by default.

=head1 SEE ALSO

This module is intended as a quick method of extracting the raw data from
the XLSX file format. This module uses Archive::Zip to extract the contents
of the XLSX file and XML::Simple for parsing the contents.

=head1 AUTHOR

Joe Estock, E<lt>jestock@nutextonline.comE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2010 by Joe Estock

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.8 or,
at your option, any later version of Perl 5 you may have available.

=cut
