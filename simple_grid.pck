create or replace package simple_grid is

  -- Author  : EYDENZONBA
  -- Created : 29.03.2017 14:44:33
  -- Purpose : creation simple report

  -- settings style default
  def_font varchar2(50) := 'Arial';
  def_font_size pls_integer := 8;
  -- width column default
  def_width number := 42;

  -- columns description
  type t_column_type_rec is record(
    title varchar2(250),
    datatype t_excel_format_data,
    width number);
  type t_column_type_tbl is table of t_column_type_rec index by pls_integer;
  -- custom rows
  type t_custom_rows_rec is table of varchar2(250) index by pls_integer;

  -- creation book
  function book(
    sheet_name varchar2,                           -- sheet title
    report_name varchar2,                          -- report title
    columns_set t_column_type_tbl,                 -- columns dimension
    custom_rows t_custom_rows_rec,                 -- custom rows dimension
    sql_text varchar2)                             -- cursor text
  return blob;

end simple_grid;
/
create or replace package body simple_grid is

  -- line feed
  lf constant varchar2(2) := chr(13);
  -- rows number maximum
  row_count constant pls_integer := 100000;
  -- xml file report
  report_clob clob;
  -- report description
  type t_report_rec is record(
    sheet_title varchar2(250),
    report_title varchar2(250),
    col_set t_column_type_tbl,
    custom_str t_custom_rows_rec,
    row_offset number);
  report t_report_rec;

  -- book title
  procedure header_book is
  begin
    dbms_lob.append(report_clob,
                    '<?xml version="1.0" encoding="utf-8"?><?mso-application progid="Excel.Sheet"?>' || lf ||
                    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">' || lf ||
                    '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' || lf ||
                    '<WindowHeight>12075</WindowHeight>' || lf ||
                    '<WindowWidth>24915</WindowWidth>' || lf ||
                    '<WindowTopX>120</WindowTopX>' || lf ||
                    '<WindowTopY>150</WindowTopY>' || lf ||
                    '<ProtectStructure>False</ProtectStructure>' || lf ||
                    '<ProtectWindows>False</ProtectWindows>' || lf ||
                    '</ExcelWorkbook>' || lf ||
                    -- default style
                    '<Styles>' || lf ||
                    '<Style ss:ID="Default" ss:Name="Normal">' || lf ||
                    '<Alignment ss:Vertical="Top"/>' || lf ||
                    '<Borders/>' || lf ||
                    '<Font ss:FontName="' || def_font || '" x:CharSet="204" x:Family="Swiss" ss:Size="' || def_font_size || '" ss:Color="#000000"/>' || lf ||
                    '<Interior/>' || lf ||
                    '<NumberFormat/>' || lf ||
                    '<Protection/>' || lf ||
                    '</Style>' || lf ||
                    -- title
                    '<Style ss:ID="s1">' || lf ||
                    '<Alignment ss:Vertical="Center"/>' || lf ||
                    '<Font ss:FontName="' || def_font || '" x:CharSet="204" x:Family="Swiss" ss:Size="' || (def_font_size + 2) || '" ss:Color="#000000" ss:Bold="1"/>' || lf ||
                    '<NumberFormat ss:Format="@"/>' || lf ||
                    '</Style>' || lf ||
                    -- custom row
                    '<Style ss:ID="s2">' || lf ||
                    '<Alignment ss:Vertical="Center"/>' || lf ||
                    '<Font ss:FontName="' || def_font || '" x:CharSet="204" x:Family="Swiss" ss:Size="' || (def_font_size + 2) || '" ss:Color="#000000"/>' || lf ||
                    '<NumberFormat ss:Format="@"/>' || lf ||
                    '</Style>' || lf ||
                    -- title
                    '<Style ss:ID="s3">' || lf ||
                    '<Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' || lf ||
                    '<Borders>' || lf ||
                    '<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>' || lf ||
                    '<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>' || lf ||
                    '</Borders>' || lf ||
                    '<Font ss:FontName="' || def_font || '" x:CharSet="204" x:Family="Swiss" ss:Size="' || def_font_size || '" ss:Color="#000000" ss:Bold="1"/>' || lf ||
                    '<NumberFormat ss:Format="@"/>' || lf ||
                    '</Style>' || lf ||
                    -- string
                    '<Style ss:ID="s4">' || lf ||
                    '<NumberFormat ss:Format="@"/>' || lf ||
                    '</Style>' || lf ||
                    -- string with line feed
                    '<Style ss:ID="s5">' || lf ||
                    '<Alignment ss:Vertical="Top" ss:WrapText="1"/>' || lf ||
                    '<NumberFormat ss:Format="@"/>' || lf ||
                    '</Style>' || lf ||
                    -- date
                    '<Style ss:ID="s6">' || lf ||
                    '<NumberFormat ss:Format="dd/mm/yyyy"/>' || lf ||
                    '</Style>' || lf ||
                    -- date/time
                    '<Style ss:ID="s7">' || lf ||
                    '<NumberFormat ss:Format="dd/mm/yyyy\ h:mm:ss;@"/>' || lf ||
                    '</Style>' || lf ||
                    -- integer
                    '<Style ss:ID="s8">' || lf ||
                    '<NumberFormat ss:Format="0"/>' || lf ||
                    '</Style>' || lf ||
                    -- float (2 signs)
                    '<Style ss:ID="s9">' || lf ||
                    '<NumberFormat ss:Format="0.00"/>' || lf ||
                    '</Style>' || lf ||
                    -- float (4 signs)
                    '<Style ss:ID="s10">' || lf ||
                    '<NumberFormat ss:Format="0.0000"/>' || lf ||
                    '</Style>' || lf ||
                    -- float (2 signs) with digit grouping
                    '<Style ss:ID="s11">' || lf ||
                    '<NumberFormat ss:Format="#,##0.00"/>' || lf ||
                    '</Style>' || lf ||
                    -- float (4 signs) with digit grouping
                    '<Style ss:ID="s12">' || lf ||
                    '<NumberFormat ss:Format="#,##0.0000"/>' || lf ||
                    '</Style>' || lf ||
                    '</Styles>' || lf);
  end header_book;

  -- sheet title
  procedure header_worksheet is
  begin
    dbms_lob.append(report_clob,
                    '<Worksheet ss:Name="' || report.sheet_title || '">' || lf ||
                    '<Table ss:ExpandedColumnCount="' || report.col_set.count || '" ss:ExpandedRowCount="' || row_count || '" x:FullColumns="1" x:FullRows="1">' || lf);
  end header_worksheet;

  -- add columns
  procedure add_columns is
  begin
    for i in 1 .. report.col_set.count loop
      dbms_lob.append(report_clob,
                      '<Column ss:StyleID="s' || report.col_set(i).datatype.get_index ||
                      '" ss:Index="' || i ||
                      '" ss:AutoFitWidth="0" ss:Width="' || report.col_set(i).width || '"/>' || lf);
    end loop;
  end add_columns;

  -- add title report
  procedure add_report_title is
  begin
    dbms_lob.append(report_clob,
                    '<Row><Cell ss:StyleID="s1"><Data ss:Type="String">' || report.report_title || '</Data></Cell></Row>' || lf);
  end add_report_title;

  -- add custom rows
  procedure add_custom_str is
  begin
    for i in 1 .. report.custom_str.count loop
      dbms_lob.append(report_clob,
                      '<Row><Cell ss:StyleID="s2"><Data ss:Type="String">' || report.custom_str(i) || '</Data></Cell></Row>' || lf);
    end loop;
  end add_custom_str;

  -- add columns descriptions
  procedure add_columns_title is
  begin
    dbms_lob.append(report_clob, '<Row ss:AutoFitHeight="1">' || lf);
    for i in 1 .. report.col_set.count loop
      dbms_lob.append(report_clob,
                      '<Cell ss:StyleID="s3"><Data ss:Type="String">' || report.col_set(i).title || '</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>' || lf);
    end loop;
    dbms_lob.append(report_clob, '</Row>' || lf);
  end add_columns_title;

  -- body report
  procedure body_book(
    sql_text varchar2)                             -- cursor text
  is
    cr pls_integer;
    cr_count pls_integer := 0;
    rec_tbl dbms_sql.desc_tab;
    row_data clob;
    exec_sql clob;
  begin
    -- get descriptions fields cursor
    cr := dbms_sql.open_cursor;
    dbms_sql.parse(cr, sql_text, dbms_sql.native);
    dbms_sql.describe_columns(cr, cr_count, rec_tbl);
    dbms_sql.close_cursor(cr);
    row_data := 'str := ''<Row>'' || lf;' || lf;
    for i in 1 .. report.col_set.count loop
      case
        when report.col_set(i).datatype.get_type = 'Number' then
             row_data := row_data || 'if cr.' || rec_tbl(i).col_name || ' is null then str := str || ''<Cell/>'' || lf; else str := str || ''<Cell><Data ss:Type="' ||
                         report.col_set(i).datatype.get_type || '">'' || cr.' || rec_tbl(i).col_name || ' || ''</Data></Cell>'' || lf; end if;' || lf;
        when report.col_set(i).datatype.get_type = 'String' then
             row_data := row_data || 'str := str || ''<Cell><Data ss:Type="' || report.col_set(i).datatype.get_type ||
                         '"><![CDATA['' || convert(cr.' || rec_tbl(i).col_name || ', ''UTF8'') || '']]></Data></Cell>'' || lf;' || lf;
        when report.col_set(i).datatype.get_type = 'DateTime' then
             row_data := row_data || 'str := str || ''<Cell><Data ss:Type="' || report.col_set(i).datatype.get_type ||
                         '">'' || to_char(cr.' || rec_tbl(i).col_name || ', ''yyyy-mm-dd'') || ''T'' || to_char(cr.'
                                               || rec_tbl(i).col_name || ', ''hh24:mi:ss'') || ''</Data></Cell>'' || lf;' || lf;
      end case;
    end loop;
    row_data := row_data || 'str := str || ''</Row>'' || lf;' || lf;
    exec_sql :=
     'declare
        lf constant varchar2(2) := chr(13);
        str clob;
      begin
        for cr in (' || sql_text || ')
        loop ' ||
          row_data || '
          dbms_lob.append(:rcod, str);
        end loop;
      end;';
    execute immediate exec_sql using in out report_clob;
  exception
    when others then
      if dbms_sql.is_open(cr) then
        dbms_sql.close_cursor(cr);
      end if;
      raise_application_error(-20001, $$plsql_unit || ': ' || sqlerrm);
  end body_book;

  -- sheet footer
  procedure footer_worksheet is
  begin
    dbms_lob.append(report_clob,
                    '</Table>' || lf ||
                    '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">' || lf ||
                    '<PageSetup>' || lf ||
                    '<Layout x:CenterHorizontal="1"/>' || lf ||
                    '<Header x:Margin="0.2" x:Data="' || convert('&amp;RСтраница &amp;P из &amp;N', 'UTF8') || '"/>' || lf ||
                    '<Footer x:Margin="0.2"/>' || lf ||
                    '<PageMargins x:Bottom="0.4" x:Left="0.4" x:Right="0.4" x:Top="0.4"/>' || lf ||
                    '</PageSetup>' || lf ||
                    '<Selected/>' || lf ||
                    '<FreezePanes/>' || lf ||
                    '<FrozenNoSplit/>' || lf ||
                    '<SplitHorizontal>' || (report.row_offset + 2) || '</SplitHorizontal>' || lf ||
                    '<TopRowBottomPane>' || (report.row_offset + 2) || '</TopRowBottomPane>' || lf ||
                    '<ActivePane>2</ActivePane>' || lf ||
                    '<ProtectObjects>False</ProtectObjects>' || lf ||
                    '<ProtectScenarios>False</ProtectScenarios>' || lf ||
                    '</WorksheetOptions>' || lf ||
                    '<AutoFilter x:Range="R' || (report.row_offset + 2) || 'C1:R' || (report.row_offset + 2) || 'C' || report.col_set.count || '" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>' || lf ||
                    '</Worksheet>' || lf);
  end footer_worksheet;

  -- book footer
  procedure footer_book is
  begin
    dbms_lob.append(report_clob, '</Workbook>' || lf);
  end footer_book;

  -- creation book
  function book(
    sheet_name varchar2,                           -- sheet title
    report_name varchar2,                          -- report title
    columns_set t_column_type_tbl,                 -- columns dimension
    custom_rows t_custom_rows_rec,                 -- custom rows dimension
    sql_text varchar2)                             -- cursor text
  return blob is
    dst_offset number := 1;
    src_offset number := 1;
    lng_context number := dbms_lob.default_lang_ctx;
    warn number := dbms_lob.warn_inconvertible_char;
    rcod blob;
  begin
    -- initialization objects
    dbms_lob.createtemporary(rcod, false, dbms_lob.call);
    dbms_lob.createtemporary(report_clob, false, dbms_lob.call);
    report.sheet_title := convert(sheet_name, 'UTF8');
    report.report_title := '<![CDATA[' || convert(report_name, 'UTF8') || ']]>';
    for i in 1 .. columns_set.count loop
      report.col_set(i).title := '<![CDATA[' || convert(columns_set(i).title, 'UTF8') || ']]>';
      report.col_set(i).datatype := coalesce(columns_set(i).datatype, t_excel_format_data('dtString'));
      report.col_set(i).width := coalesce(columns_set(i).width, def_width);
    end loop;
    for i in 1 .. custom_rows.count loop
      report.custom_str(i) := '<![CDATA[' || convert(custom_rows(i), 'UTF8') || ']]>';
    end loop;
    report.row_offset := report.custom_str.count;
    -- book title
    header_book;
    -- sheet title
    header_worksheet;
    -- columns
    add_columns;
    -- report title
    add_report_title;
    -- custom rows
    add_custom_str;
    -- columns descriptions
    add_columns_title;
    -- data
    body_book(sql_text);
    -- sheet footer
    footer_worksheet;
    -- book footer
    footer_book;
    -- conversion
    dbms_lob.converttoblob(rcod,
                           report_clob,
                           dbms_lob.lobmaxsize,
                           dst_offset,
                           src_offset,
                           dbms_lob.default_csid,
                           lng_context,
                           warn);
    -- compression
    return utl_compress.lz_compress(rcod);
  end book;

end simple_grid;
/
