#!/usr/bin/perl

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use utf8;
use Tk;
use Tk::GridColumns;

# файл с базой
my $DATAFILE = 'db.xls';

# создаём окно
my $mw = tkinit( -title => 'DB simple crud' );
$mw->geometry( '=960x700+100+100' );

# @columns - список имён столбцов
# @data - данные
my (@data, @columns);

# создаём новый грид, делаем бинд на правую кнопку мыши для изменения значений ячейки
my $gc = $mw->Scrolled(
    'GridColumns' =>
    -scrollbars => 'ose',
    -data => \@data,
    -columns => \@columns,
    -bg => 'white',
    -itemattr => { -anchor => 'w', -bg => Tk::NORMAL_BG },
    -itemgrid => { -padx => 1, -pady => 1 },
    -item_bindings => { '<ButtonPress-3>' => \&edit_item },
)->pack(
    -fill => 'both',
    -expand => 1,
)->Subwidget( 'scrolled' );

# загружаем данные
open_adrbook( $gc, $DATAFILE )->refresh;

# рисуем кнопки
my $frm_bottom = $mw->Frame->pack(
    -side => 'bottom',
    -fill => 'x',
);

$frm_bottom->Button(
    -text => 'Добавить',
    -command => sub { $gc->add_row( '','' )->refresh_items },
)->pack(
    -side => 'left',
);

$frm_bottom->Button(
    -text => 'Удалить',
    -command => sub {
        my @sel = @{ $gc->curselection };
        if ( @sel ) {
            $gc->deselect( $sel[0][0], $_ ) for 0 .. 2;
            splice @data, $sel[0][0], 1;
            $gc->refresh_items;
        } # if
    },
)->pack(
    -side => 'left',
);

$frm_bottom->Button(
    -text => 'Сохранить',
    -command => sub { save_adrbook( $gc, $DATAFILE ) },
)->pack(
    -side => 'right',
);

MainLoop;


sub open_adrbook {
    my( $gc, $file ) = @_;

    my $parser   = Spreadsheet::ParseExcel->new(); # создаём новый парсер, загружаем файл
    my $workbook = $parser->parse( $file );

    # @table_data - данные, @table_head - имена столбцов
    my @table_data; 
    my @table_head;
    my @col_data;

    if ( !defined $workbook ) {
      die $parser->error(), ".\n";
    }

    for my $worksheet ( $workbook->worksheets() ) {
      
      # определяем сколько столбцов и сторк
      my ( $row_min, $row_max ) = $worksheet->row_range();
      my ( $col_min, $col_max ) = $worksheet->col_range();
   
      for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {

          my $cell = $worksheet->get_cell( $row, $col );
          next unless $cell;
          
          # если 1 строка - кладём значения в @table_head ( имена стобцов )
          if ($row == 0){
            push @table_head, { -text => $cell->value, -command => $gc->sort_cmd( $col, 'abc' )}
          }
          # в противном случаем пишем в @table_data ( данные )
          else {
            push @col_data, $cell->value;
          }

        }
        @col_data ? push @table_data, [@col_data] : ();
        @col_data = ();
      }
    }

    @data = @table_data;
    @columns = @table_head;
    return $gc;
}

sub save_adrbook {
  my( $gc, $file ) = @_;

  # создаем новый файл
  my $workbook = Spreadsheet::WriteExcel->new('new.xls');
  my $worksheet = $workbook->add_worksheet();
  
  # для имен столбцов добавим форматирование ( жирный шрифт )
  my $format = $workbook->add_format();
  $format->set_bold();

  # пишем данные в файл
  for my $row ( 0 .. $#{ $gc->data }+1 ) {
    for my $col ( 0 .. $#{ $gc->data->[$row] } ){
      if ($row == 0){
        $worksheet->write($row, $col, $gc->columns->[$col]{'-text'}, $format );
      }
      else {
        $worksheet->write($row, $col, $gc->data->[$row][$col] );
      }
    }
  }
  return $gc;
}

sub edit_item {
    my( $self, $w, $row, $col ) = @_;
    $w->destroy;
    
    my $entry = $self->Entry(
        -textvariable => \$data[$row][$col],
        -width => 0,
    )->grid(
        -row => $row+1,
        -column => $col,
        -sticky => 'nsew',
    );
    
    $entry->selectionRange( 0, 'end' );
    $entry->focus;

    $entry->bind( '<Return>' => sub { $self->refresh_items } );
    $entry->bind( '<FocusOut>' => sub { $self->refresh_items } );
}
__END__