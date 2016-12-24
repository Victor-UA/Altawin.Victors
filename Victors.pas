procedure FillTListView(LV: TListView; Var DL: icmDictionaryList; SQL: String=''; Fields: icmDictionary = CreateDictionary; CheckBoxes: Boolean=False; Filter: icmDictionaryList = CreateDictionaryList; CheckField: String=''; _Checked: Variant=1);
{
    Fields:=MakeDictionasry(['Caption0', 'FieldName0','Caption1', 'FieldName1','Caption2', 'FieldName2',..]);
      ['Артикул','marking','Найменування','ggname'];

}
{
    Filter:=CreateDictionaryList; Запис - фільтри із загальним правилом
      ['marking', '560510', 'orderno', 'замовлення №', '.Rule', 'partialy']
      ['marking', '560510', '.Rule', 'fully']
    //.Rule - за умовчанням = 'partialy'
}
var
  i: Integer;
  n: Integer;
  k: Integer;
  fi: Integer;
  LC: TListColumn;
  Checked: Boolean;
  GridLines: Boolean;
begin
  try
    GridLines:=LV.GridLines;
    LV.GridLines:=False;
    Application.ProcessMessages;
    if SQL<>'' then begin
      DL:=QueryRecordList(SQL, CreateDictionary);
    end;
    LV.Items.Clear;
    LV.Items.BeginUpdate;

    //Заповнюємо заголовки
    i:=0;
    while i<Fields.Count do begin
      try
        LC:=LV.Columns[i];
      except
        LC:=LV.Columns.Add;
      end;
      with LC do begin
        Aligment:=taLeftJustify;
        AutoSize:=False;
        Caption:=Fields.Name[i];
        MinWidth:=TextWidth(Caption, LV.Font)+12;
        Width:=MinWidth;
      end;
      inc(i);
    end;
    LV.Items.BeginUpdate;
    k:=0;
    i:=0;
    while i<DL.Count do begin
      if Filter.Count=0 then begin
        Checked:=True;
      end
      else begin
        Checked:=True;
        n:=0;
        while (n<Filter.Count) and Checked do begin
          try
            fi:=0;
            while (fi<Filter[n].Count) and Checked do begin
              if (Filter[n].Count>1) and (Filter[n].Exists('.Rule')) and (LowerCase(Filter[n]['.Rule'])='fully') then begin
                if lowercase(Filter[n].Name[fi])<>'.rule' then begin
                  Checked:=(VarToStr(Filter[n][Filter[n].Name[fi]])=VarToStr(DL[i][Filter[n].Name[fi]])) or (Filter[n][Filter[n].Name[fi]]='');
                end;
              end
              else begin
                if lowercase(Filter[n].Name[fi])<>'.rule' then begin
                  Checked:=(Pos(LowerCase(Filter[n][Filter[n].Name[fi]]), LowerCase(VarToStr(DL[i][Filter[n].Name[fi]])))>0) or (Filter[n][Filter[n].Name[fi]]='');
                end;
              end;
              inc(fi);
            end;
          except
            Checked:=True;
          end;
          inc(n);
        end;
      end;

      //Заповнюємо поля
      if Checked then begin
        LV.Items.Add();
        LV.Items[k].Data:=i;
        LV.Items[k].Caption:=DL[i][Fields.Value[Fields.Name[0]]];
        if CheckBoxes then begin
          LV.Items[k].Checked:=DL[i][CheckField]=_Checked;
        end;
        n:=1;
        while n<Fields.Count do begin
          LV.Items[k].SubItems.Add(DL[i][Fields.Value[Fields.Name[n]]]);
          inc(n);
        end;
        inc(k);
      end;

      inc(i);
    end;
    LV.Items.EndUpdate;
    LV.GridLines:=GridLines;
    i:=0;
    while i<Fields.Count do begin
      LV.Columns[i].Tag:=i+1;
      LV.Columns[i].Width:=-1;
      inc(i);
    end;
    if LV.Items.Count>0 then begin
      LV.Items[0].Selected:=True;
    end;
    LV.Items.EndUpdate;
  except
    showmessage('Помилка заповнення ListView'+#13+ExceptionMessage+#13#13+SQL);
  end;
end;

function TextWidth(Text: String; _Font: TFont = TFont.Create): Integer;
var
  BM: TBitmap = TBitmap.Create;
  Canv: TCanvas;
begin
  try
    try
      Canv:=BM.Canvas;
      with Canv do begin
        Control:=TControl.Create(null);
        Font:=_Font;
        Result:=TextWidth(Text);
      end;
    except
      RaiseException('Помилка визначення ширини тексту'+#13+Text+#13+ExceptionMessage);
    end;
  finally
    BM.Free;
  end;
end;

function TextHeight(Text: String; _Font: TFont = TFont.Create): Integer;
var
  BM: TBitmap = TBitmap.Create;
  Canv: TCanvas;
begin
  try
    try
      Canv:=BM.Canvas;
      with Canv do begin
        Control:=TControl.Create(null);
        Font:=_Font;
        Result:=TextHeight(Text);
      end;
    except
      RaiseException('Помилка визначення ширини тексту'+#13+Text+#13+ExceptionMessage);
    end;
  finally
    BM.Free;
  end;
end;

function VarToSQL(Value: Variant): String;
begin
  case VarType(Value) and VarTypeMask of
    varString, varOleStr: begin
      if LowerCase(Value)='null' then begin
        Result:=Value;
      end
      else begin
        Result:=ReplaceText(Value, '''', '''''');
        Result:=''''+Result+'''';
      end;
    end;
    varDate: Result:=''''+DateTimeToStr(Value)+'''';
    varEmpty: Result:='''''';
    varNull: Result:='null';
  else
    Result:=VarToStr(Value);
  end;
end;

procedure OpenReport(DocClass, ReportName: String; TransData: Variant);
var
  SQL: String;
begin
  SQL:='select'+#13+
       '  rtl.reportblob'+#13+
       'from reporttemplateslink rtl'+#13+
       'where lower(rtl.docclass)=:docclass'+#13+
       '  and lower(rtl.reportname)=:reportname';
  SQL:=ReplaceText(SQL, ':docclass', VarToSQL(LowerCase(DocClass)));
  SQL:=ReplaceText(SQL, ':reportname', VarToSQL(LowerCase(ReportName)));
  try
    FastReportExecuteStream(
      TStringStream.Create(QueryValue(SQL, CreateDictionary)),
      ['TransData', TransData]
    );
  except
    Showmessage('Помилка відкриття '+'['+ReportName+']'+#13+ExceptionMessage);
  end;
end;

function GetXMLParamValue(XML,ParamName: String; Var PosBegin,PosEnd: Integer): String;
var
  XMLModify: String;
begin
  Result:='';
  try
    if PosBegin<=0 then begin
      PosBegin:=Pos('<'+ParamName+'>',XML);
    end;
    PosEnd:=Pos('</'+ParamName+'>',XML)+Length('</'+ParamName+'>');
    Result:=UTF8ToAnsi(Copy(XML, PosBegin+Length('<'+ParamName+'>'),Pos('</'+ParamName+'>',XML)-(PosBegin+Length('<'+ParamName+'>'))));
  except
    Result:='';
  end;
end;

function SetXMLParamValue(XML1,ParamName,OldValue1,NewValue: String; PosBegin: Integer=-1; Count: Integer=-1): String;
var
  XMLStart: String;
  XMLModify: String;
  XMLEnd: String;
  OldValue: String;
  pBegin: Integer;
  pEnd: Integer;
  XML: String;
  i: Integer;
  qty: Integer;
begin
  Result:=XML1;
  XML:=XML1;
  OldValue:=OldValue1;
  try
    if lowercase(OldValue)='any value' then begin
      qty:=0;
      i:=0;
      while (i<length(XML)) and (Pos('<'+ParamName+'>',XML)>0) and ((Count<0) or (qty<Count)) do begin
        pBegin:=Pos('<'+ParamName+'>',XML);
        XML:=SetXMLParamValue(XML, ParamName, GetXMLParamValue(XML,ParamName,pBegin,pEnd), NewValue, pBegin);
        if pEnd=0 then begin
          i:=length(XML);
        end
        else begin
          inc(qty);
        end;
        i:=pEnd-1;
        inc(i);
      end;
      Result:=XML;
    end
    else begin
      NewValue:=AnsiToUTF8(NewValue);
      OldValue:=AnsiToUTF8(OldValue);
      if PosBegin<=0 then begin
        Result:=ReplaceText(XML, '<'+ParamName+'>'+OldValue+'</'+ParamName+'>', '<'+ParamName+'>'+NewValue+'</'+ParamName+'>');
      end
      else begin
        XMLStart:=copy(XML, 1, PosBegin-1);
        XMLModify:=copy(XML, PosBegin, Length('<'+ParamName+'>'+OldValue+'</'+ParamName+'>'));
        XMLEnd:=copy(XML, Length(XMLStart)+Length(XMLModify)+1, Length(XML)-(Length(XMLStart)+Length(XMLModify)));
        XMLModify:=ReplaceText(XMLModify, '<'+ParamName+'>'+OldValue+'</'+ParamName+'>', '<'+ParamName+'>'+NewValue+'</'+ParamName+'>');
        XML:=XMLStart+XMLModify+XMLEnd;
        Result:=XMLStart+XMLModify+XMLEnd;
      end;
    end;
  except
    Showmessage('Помилка зміни значення параметру ['+ParamName+'] XML'+#13+ExceptinoMessage+#13+XML);
    Result:=XML1;
  end;
end;

function OrderExist(OrderId: Integer): Boolean;
var
  SQL: String;
  QTY: Integer;
begin
  try
    SQL:='select'+#13+
         '  count(o.orderid)'+#13+
         'from orders o'+#13+
         'where o.orderid=:orderid';
    SQL:=ReplaceText(SQL, ':orderid', VarToStr(OrderId));
    QTY:=QueryValue(SQL, MakeDictionary([]));
    Result:=QTY>0;
  except
    Showmessage('Помилка визначення кількості замовлень'+#13+ExceptionMessage+#13+SQL);
    Result:=False;
  end;
end;

procedure SendEmail(RecipientsField,CopyField,Subject,Body: String);
var
  Mailer: IpubMailer;
  Email:  IpubEmail;
  A:  IpubEmailAttachment;
  i: Integer;
  Response: Boolean=False;
begin
  try
    Mailer:=CreateMailer;
    Mailer.Account.Address:='altawin@gazda.ua';
  //  Mailer.Account.SmtpHost:='mail.gazda.ua';
    Mailer.Account.SmtpHost:='94.154.220.129';
    Mailer.Account.UserName:='[Altawin]';
    Mailer.Account.SmtpUser:='altawin@gazda.ua';
    Mailer.Account.SmtpPassword:='UybCPZYbJS';
//    Mailer.Account.SmtpPassword:='ejBJ4ZZp155w3ITT';
    Email:=Mailer.NewEmail;
    Email.RecipientsField:=RecipientsField;
    Email.CopyField:=CopyField;
    Email.Subject:=Subject;
    Email.Body:=Body;
    i:=0;
    while (i<15) and not Response do begin
      Response:=Mailer.SendEmail(Email, False);
      inc(i);
    end;
    Mailer:=Empty;
    Email:=Empty;
  except
    Showmessage('Помилка надсилання листа'+#13+ExceptionMessage);
  end;
end;

function DictListIndex(DictList: icmDictionaryList; Name: String; Value: Variant; Lower: Boolean=False): Integer;
var
  i: Integer;
begin
  Result:=-100500;
  i:=0;
  try
    while (i<DictList.Count) and (Result=-100500) do begin
      if Lower then begin
        if LowerCase(DictList[i][Name])=LowerCase(VarToStr(Value)) then
          Result:=i;
      end
      else begin
        if DictList[i][Name]=Value then
          Result:=i;
      end;
      inc(i);
    end;
  except
    Result:=-100500;
  end;
end;

function IsNumeric(S: String; Comma: String=''): Boolean;
var
  i: Integer;
  flt: Variant;
begin
  Result:=True;
  if (Comma='') then begin
    try
      flt:=StrToFloat(S);
    except
      Result:=False;
    end;
  end
  else begin
    i:=1;
    while (i<=length(S)) and Result do begin
      Result:=S[i] in ['0'..'9',Comma];
      inc(i);
    end;
  end;

end;

function TimeOf(vDateTime: TDateTime): TDateTime;
var
  S: String='';
begin
  S:=Copy(VarToStr(vDateTime), 12, 8);
  Result:=StrToTime(S);
end;

function DateOf(vDateTime: TDateTime): TDateTime;
var
  S: String='';
begin
  S:=Copy(VarToStr(vDateTime), 1, 10);
  S:=Copy(S, 9, 2)+'.'+Copy(S, 6, 2)+'.'+Copy(S, 1, 4);
  Result:=StrToDate(S);
end;

function QueryFromFile(SQL: String=''; Params: icmDictionaryList=CreateDictionaryList; Path, FileName: String): String;
//Params.Add(MakeDictionary(['Name', {value}]))
var
  Names, Strings: TStringList;
  Dict: icmDictionary;
  DL: icmDictionaryList;
  Sep: String;
  Str: String;
  tmp: String;
  Err: String;
  sepPos: Integer;
  i,k: Integer;
  FormPB: TForm;
  fPB: TProgressBar;
  fLabel: TLabel;
  NewP, OldP: Integer;
begin
  Dict:=CreateDictionary;
  DL:=CreateDictionaryList;
  Strings:=TStringList.Create;
  Names:=TStringList.Create;

  FormPB:=TForm.Create(nil);
  FormPB.Height:=100;
  FormPB.Width:=330;
  FormPB.BorderStyle:=bsToolWindow;
  FormPB.Position:=poScreenCenter;
  fPB:=TProgressBar.Create(nil);
  fPB.Parent:=FormPB;
  fPB.Height:=20;
  fPB.Width:=300;
  fPB.Top:=30;
  fPB.Left:=10;
  fPB.Min:=0;
  fPB.Position:=0;
  fPB.Step:=1;
  fPB.Visible:=True;
  fLabel:=TLabel.Create(nil);
  fLabel.Parent:=FormPB;
  fPB.Height:=20;
  fLabel.Top:=10;
  fLabel.Left:=10;
  fLabel.Caption:='Підготовка...';
  fLabel.Visible:=True;
  FormPB.Show;
  Application.ProcessMessages;

  try
    Strings.LoadFromFile(Path+FileName);
  except
   showmessage('Помилка завантаження даних з файлу ['+Patch+FileName+']'+#13+ExceptionMessage);
   Exit;
  end;
  if Strings.Count>2 then begin
    try
      Sep:=copy(Strings[0], 6, 1);
      Str:=Strings[1];
      sepPos:=Pos(Sep, Str);
      while sepPos>0 do begin
        tmp:=copy(Str, 1, SepPos-1);
        if length(tmp)>0 then
          Names.add(tmp);
        Str:=copy(Str, SepPos+1, Length(Str)-SepPos);
        sepPos:=Pos(Sep, Str);
      end;
      if length(Str)>0 then                     //Останній запис після коми
        Names.add(Str);
      fPB.Max:=Strings.Count-1;
      OldP:=0;
      for k:=2 to Strings.Count-1 do begin               //Values
        Str:=Strings[k];
        i:=0;
        sepPos:=Pos(Sep, Str);
        while sepPos>0 do begin
          tmp:=copy(Str, 1, SepPos-1);
          if k=2 then
            Dict.add(Names[i],tmp)
          else
            Dict.Value[Names[i]]:=tmp;
          inc(i);
          Str:=copy(Str, SepPos+1, Length(Str)-SepPos);
          sepPos:=Pos(Sep, Str);
        end;
        if length(Str)>0 then begin                    //Останній запис після коми
          if k=2 then
            Dict.add(Names[i],Str)
          else
            Dict.Value[Names[i]]:=Str;
        end;
        DL.Add(Dict);
        if length(SQL)>0 then begin
          tmp:=SQL;
          i:=0;
          while i<Params.Count do begin
            tmp:=ReplaceText(tmp, ':'+Params[i]['Name'], VarToSQL(Dict.Value[Params[i]['Name']]));
            inc(i);
          end;
          Err:=ExecuteSQLCommit(tmp);
          if Length(Err)>0 then begin
            showmessage('Помилка виконання запиту'+#13+Err);
            exit;
          end;
        end;
        NewP:=Trunc(k/fPB.Max*100);
        if NewP<>OldP then begin
          fLabel.Caption:=FileName+': '+VarToStr(NewP)+'%';
          fPB.Position:=k;
          Application.ProcessMessages;
        end;
        OldP:=P;
      end;
    except
      showmessage('Помилка застосування даних'+#13+tmp+#13+ExceptionMessage);
    end;
  end;
//  showmessage(tmp);//!
  Result:=JSONEncode(DL);
  DL.Clear;
  Dict.Clear;
  Strings.Clear;
  Names.Clear;
  FormPB.Close;
  FormPB.Free;
end;

procedure QueryToFile(SQL, Path, FileName: String);
var
  Strings: TStringList;
  Str: String;
  DictList: icmDictionaryList;
  i,k: Integer;
  FormPB: TForm;
  fPB: TProgressBar;
  fLabel: TLabel;
  NewP, OldP: Integer;
begin
  try
    ForceDirectories(Path);
  except
    showmessage('Помилка створення каталогу ['+Path+']'+#13+ExceptionMessage);
    exit;
  end;
  Strings:=TStringList.Create;

  FormPB:=TForm.Create(nil);
  FormPB.Height:=100;
  FormPB.Width:=330;
  FormPB.BorderStyle:=bsToolWindow;
  FormPB.Position:=poScreenCenter;
  fPB:=TProgressBar.Create(nil);
  fPB.Parent:=FormPB;
  fPB.Height:=20;
  fPB.Width:=300;
  fPB.Top:=30;
  fPB.Left:=10;
  fPB.Min:=0;
  fPB.Position:=0;
  fPB.Step:=1;
  fPB.Visible:=True;
  fLabel:=TLabel.Create(nil);
  fLabel.Parent:=FormPB;
  fPB.Height:=20;
  fLabel.Top:=10;
  fLabel.Left:=10;
  fLabel.Caption:='Підготовка...';
  fLabel.Visible:=True;
  FormPB.Show;
  Application.ProcessMessages;

  try
    DictList:=QueryRecordList(SQL, MakeDictionary([]));
    i:=0;
    Str:='sep =,';
    Strings.add(Str);
    fPB.Max:=DictList.Count;
    OldP:=0;
    while i<DictList.Count do begin
      Str:='';
      k:=0;
      while k<DictList[i].Count do begin
        if i=0 then
          Str:=Str+DictList[i].Name[k]
        else
          Str:=Str+VarToStr(DictList[i].Value[DictList[i].Name[k]]);
        inc(k);
        if k<DictList[i].Count then
          Str:=Str+',';
      end;
      Strings.add(Str);
      inc(i);
      NewP:=Trunc(i/fPB.Max*100);
      if NewP<>OldP then begin
        fLabel.Caption:=FileName+': '+VarToStr(NewP)+'%';
        fPB.Position:=i;
        Application.ProcessMessages;
      end;
      OldP:=P;
    end;
    Strings.SaveToFile(Path+FileName);
  except
    showmessage('Помилка збереження'+#13+ExceptionMessage);
  end;
  DictList.Clear;
  Strings.Clear;
  FormPB.Close;
  FormPB.Free;
end;

function DictList2Dict(DictList: icmDictionaryList): icmDictionary;
var
  i: Integer;
begin
  try
    Result:=CreateDictionary;
    i:=0;
    while i<DictList.Count do begin
      Result.Add(DictList[i]['intl_varname'], DictList[i]['value1']);
      inc(i);
    end;
  except
    Showmessage('Помилка перетворення DcitionaryList на Dictionary'+#13#13+ExceptionMessage);
  end;
end;

function Masking(enCodeStr: String; MaxInterval: Integer; NullOff: Boolean=False): String;
const
  Letters=['a','c','e','g','m','n','o','p','q','r','s','t','u','v','w','x','y','z'];
var
  Index:Byte;
  IntervalSum, Interval: Byte;
  MaxIntervalSum: Byte;
  Mask, Code: String;
  i: Integer;
begin
  Mask:='#';
  for i:=1 to 12 do begin
    Index:=Rand(Length(Letters));
    Mask:=Mask+Letters[Index];
  end;
  Code:=Mask;
  try
    Index:=1;
    IntervalSum:=0;
    MaxIntervalSum:=Length(Code)-Length(enCodeStr)-1;
    if (StrToInt(enCodeStr)<>0) or not NullOff then begin
      for i:=1 to Length(enCodeStr) do begin
        Interval:=Rand(MaxInterval)+1;
        if IntervalSum+Interval>MaxIntervalSum then
          Interval:=0;
        IntervalSum:=IntervalSum+Interval;
        Index:=Index+Interval;
        Code:=Copy(Code, 1, Index)+Copy(enCodeStr, i, 1)+Copy(Code, Index+2, Length(Code)-Index-1);
        Index:=Index+1;
      end;
    end;
  except
    Code:=Mask;
  end;
  Result:=Code;
end;

function ExecuteSQLCommit(pSQL: String; Connection: String=''; Params: icmDictionary=MakeDictionary([])): String;
var
  Session: TAltecSession;
  Query: TAltecQuery;
  i: Integer;
begin
  Result:='';
  Session:=TAltecSession.Create(Application);
  Session.Connection:=Connection;
  try
    Session.Start;
    try
      Query:=TAltecQuery.Create(Session);
      try
        Query.Session:=Session;
        Query.SQL.Text:=pSQL;
        i:=0;
        while i<Params.Count do begin
          Query.Params[i].Value:=Params.Value[Query.Params[i].Name];
          inc(i);
        end;
        Query.ExecQuery;
        Session.Commit;
      finally
        Query.Free;
      end;
    except
      Result:=ExceptionMessage+#13+pSQL;
      Session.Rollback;
      RaiseException(ExceptionMessage);
    end;
  finally
    Session.Free;
  end;
end;

function Rand(Count: Integer; Self: Boolean=False): Integer;
var
  vRand: String;
  Index: Byte;
begin
  Randomize;
  vRand:=VarToStr(Random);
  if Self then begin
    Index:=Rand(Length(vRand)-2);
    if Index<0 then
      Index:=0;
  end
  else
    Index:=0;
  Result:=StrToInt(Copy(vRand, Length(vRand)-1-Index, 1));
  Result:=Result-Count*Int(Result/Count);
end;

{ **** UBPFD *********** by delphibase.endimus.com ****
>> Дополнение строки пробелами слева

Дополненяет строку слева пробелами до указанной длины

Зависимости: нет
Автор:       Anatoly Podgoretsky, anatoly@podgoretsky.com, Johvi
Copyright:
Дата:        26 апреля 2002 г.
***************************************************** }

function PADL(Src: string; Lg: Integer; Symb: String=' '): string;
begin
  Result := Src;
  while Length(Result) < Lg do
    Result := Symb + Result;
end;
{ **** UBPFD *********** by delphibase.endimus.com ****
>> Дополнение строки пробелами справа

Дополняет строку пробелами справа до указанной длины.

Зависимости: нет
Автор:       Anatoly Podgoretsky, anatoly@podgoretsky.com, Johvi
Copyright:   Anatoly Podgoretsky
Дата:        26 апреля 2002 г.
***************************************************** }

function PADR(Src: string; Lg: Integer; Symb: String=' '): string;
begin
  Result := Src;
  while Length(Result) < Lg do
    Result := Result + Symb;
end;
{ **** UBPFD *********** by delphibase.endimus.com ****
>> Дополнение строки пробелами с обоих сторон

Дополнение строки пробелами с обоих сторон до указанной длины

Зависимости: нет
Автор:       Anatoly Podgoretsky, anatoly@podgoretsky.com, Johvi
Copyright:
Дата:        26 апреля 2002 г.
***************************************************** }

function PADC(Src: string; Lg: Integer): string;
begin
  Result := Src;
  while Length(Result) < Lg do
  begin
    Result := Result + ' ';
    if Length(Result) < Lg then
    begin
      Result := ' ' + Result;
    end;
  end;
end;
