declare
  f simple_grid.t_custom_rows_rec;
  m simple_grid.t_column_type_tbl;
  sql_text varchar2(32000) := 'select * from all_users';
  rcod blob;
begin
  f(1) := 'Date/time create: ' || to_char(sysdate, 'dd.mm.rrrr hh24:mi:ss');
  f(2) := user;
  f(3) := 'REPORT_DATE: ' || trunc(sysdate);
  for cr in (select a.table_name, a.column_name, a.data_type, a.data_length, b.comments, a.column_id
               from all_tab_cols a
               join all_col_comments b
                 on 1 = 1
                and b.owner = a.owner
                and b.table_name = a.table_name
                and b.column_name = a.column_name
              where 1 = 1
                and a.table_name = 'ALL_USERS'
              order by a.column_id)
  loop
    m(cr.column_id).title := cr.comments;
    case
      when cr.data_type = 'VARCHAR2' then m(cr.column_id).datatype := t_excel_format_data('dtString');
      when cr.data_type = 'NUMBER' then m(cr.column_id).datatype := t_excel_format_data('dtNumber2');
      else m(cr.column_id).datatype := t_excel_format_data('dt' || initcap(cr.data_type));
    end case;
    if cr.column_id = 1 then
      m(cr.column_id).datatype := t_excel_format_data('dtStringWrap');
    end if;
  end loop;
  simple_grid.def_width := 120;
  rcod := simple_grid.book(sheet_name  => 'ALL_USERS',
                           report_name => 'Example report',
                           columns_set => m,
                           custom_rows => f,
                           sql_text    => sql_text);
end;
