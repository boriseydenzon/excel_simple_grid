create or replace type t_excel_format_data as object(

  -- Author  : EYDENZONBA
  -- Created : 03.04.2017 10:20:42
  -- Purpose : data types for excel

  data_type varchar2(30),

  -- constructor
  constructor function t_excel_format_data(
    data_type varchar2)                            -- data type
  return self as result,

  -- data type index
  member function get_index return number,

  -- data type description
  member function get_type return varchar2

)
/
create or replace type body t_excel_format_data is

  -- contructor
  constructor function t_excel_format_data(
    data_type varchar2)                            -- data type
  return self as result
  is
  begin
    if data_type in ('dtString', 'dtStringWrap', 'dtDate', 'dtDateTime', 'dtInteger',
                     'dtNumber2', 'dtNumber4', 'dtNumberSeparator2', 'dtNumberSeparator4') then
      self.data_type := data_type;
      return;
    end if;
    raise_application_error(-20001,
                            'Invalid datatype "' || data_type || '". Only dtString, dtStringWrap, dtDate, dtDateTime, ' ||
                            'dtInteger, dtNumber2, dtNumber4, dtNumberSeparator2 or dtNumberSeparator4 are allowed.');
  end;

  -- data type index
  member function get_index return number is
    rcod number;
  begin
    case
      when self.data_type = 'dtString' then rcod := 4;
      when self.data_type = 'dtStringWrap' then rcod := 5;
      when self.data_type = 'dtDate' then rcod := 6;
      when self.data_type = 'dtDateTime' then rcod := 7;
      when self.data_type = 'dtInteger' then rcod := 8;
      when self.data_type = 'dtNumber2' then rcod := 9;
      when self.data_type = 'dtNumber4' then rcod := 10;
      when self.data_type = 'dtNumberSeparator2' then rcod := 11;
      when self.data_type = 'dtNumberSeparator4' then rcod := 12;
      else rcod := 0;
    end case;
    return rcod;
  end;

  -- data type decription
  member function get_type return varchar2 is
    rcod varchar2(10);
  begin
    case
      when self.data_type in ('dtString', 'dtStringWrap') then rcod := 'String';
      when self.data_type in ('dtDate', 'dtDateTime') then rcod := 'DateTime';
      else rcod := 'Number';
    end case;
    return rcod;
  end;

end;
/
