unit Unit1;

{
Допустимые символы: 0-9 и A-z, А-я, остальные считаются разделителями.
GetWord(n:integer):string; - пропускаются все не значащие символы, потом слово до первого
значащего символа,
n - номер по счету слова (начиная с 0 ! )
GetWord(s,-1) возвращает весь текст после первого слова
Возвращает '' если нет слова под номером n (например, конец строки) 

Первое слово - комманда, остальные - параметры.
Команда состоит из 1 слова. Правило синтаксиса: 2 и более слов разделяются знаком "_" , например, end_repeat
если первое слово не команда, то эта строка считается комментарием

FIX:
в диапазоне цвета проверка не по байтам (по исний, зеленый, красный), а сразу по всем как по одному числу, это может глючить
добавить if str
// возможны траблы с нажатиями клавиш и щелчками мышью. проследить WM_CHAR, WM_SETFOCUS и паузы между ними. А еще при дв. щелчке происходит реально 3 (!)
//Возможны траблы в будущем с lowercase(s) для рус. алфавита в GetWord
//при if коорд цвет и if_not коорд цвет проверка реально происхродит только когда окно УО на виду
//скрипт пишется для первого окна (ctrl+A)
//после команды msg скрипт останавливается и ждет пока нажмем OK. это важно
// на проигрывание файла msg.wave дается 400 мсек

}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, HotKeyMgr, ComCtrls, Menus, ExtCtrls,jpeg,
  mmsystem, Recorder;

type
  TfmMain = class(TForm)
    pcAll: TPageControl;
    tsGeneral: TTabSheet;
    tsScript: TTabSheet;
    btAddM: TSpeedButton;
    mmScript: TMemo;
    gbC: TGroupBox;
    ed0: TEdit;
    ed1: TEdit;
    ed2: TEdit;
    ed3: TEdit;
    cb1: TComboBox;
    cb2: TComboBox;
    cb3: TComboBox;
    tsHelp: TTabSheet;
    mmHelp: TMemo;
    btS1: TSpeedButton;
    btS2: TSpeedButton;
    btS3: TSpeedButton;
    btS0: TSpeedButton;
    btStart: TSpeedButton;
    mnHotKey: TMainMenu;
    miCtrlA: TMenuItem;
    odLoad: TOpenDialog;
    tm0: TTimer;
    tm2: TTimer;
    tm3: TTimer;
    tm1: TTimer;
    HotKeyManager1: THotKeyManager;
    sdSave: TSaveDialog;
    cbM: TComboBox;
    mnCom: TPopupMenu;
    miM: TMenuItem;
    miW: TMenuItem;
    miRt: TMenuItem;
    mieRt: TMenuItem;
    miDL: TMenuItem;
    miDR: TMenuItem;
    miStop: TMenuItem;
    miD: TMenuItem;
    miS: TMenuItem;
    miMsg: TMenuItem;
    miIF: TMenuItem;
    mieIF: TMenuItem;
    Label14: TLabel;
    miIFNot: TMenuItem;
    gbOtherWindow: TGroupBox;
    btS4: TSpeedButton;
    btS5: TSpeedButton;
    ed4: TEdit;
    ed5: TEdit;
    cb4: TComboBox;
    cb5: TComboBox;
    miOpen: TMenuItem;
    miSave: TMenuItem;
    miSaveAs: TMenuItem;
    N12: TMenuItem;
    miExit: TMenuItem;
    miNew: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    miRec: TMenuItem;
    miStopRec: TMenuItem;
    miPlay: TMenuItem;
    miSpeed: TMenuItem;
    mi25: TMenuItem;
    mi50: TMenuItem;
    miRepeat: TMenuItem;
    N19: TMenuItem;
    N52: TMenuItem;
    N152: TMenuItem;
    N9991: TMenuItem;
    N20: TMenuItem;
    mi75: TMenuItem;
    mi100: TMenuItem;
    N2001: TMenuItem;
    N3001: TMenuItem;
    edM: TEdit;
    miCtrlB: TMenuItem;
    tm4: TTimer;
    tm5: TTimer;
    miAlarm: TMenuItem;
    N1: TMenuItem;
    gbScreenShot: TGroupBox;
    cbDate: TCheckBox;
    rbBmp: TRadioButton;
    rbJpg: TRadioButton;
    edScr: TEdit;
    Label4: TLabel;
    cb0: TComboBox;
    miContinue: TMenuItem;
    miBreak: TMenuItem;
    miCut: TMenuItem;
    miPaste: TMenuItem;
    miCopy: TMenuItem;
    N5: TMenuItem;
    miUndo: TMenuItem;
    btXY: TSpeedButton;
    btColor: TSpeedButton;
    cbInsertXY: TCheckBox;
    Label3: TLabel;
    Label1: TLabel;
    edPause: TEdit;
    Label2: TLabel;
    N2: TMenuItem;
    N3: TMenuItem;
    left1: TMenuItem;
    right1: TMenuItem;
    N4: TMenuItem;
    miElse: TMenuItem;
    move1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure SetCoord(Sender: TObject);
    procedure btStartClick(Sender: TObject);
    procedure btCStartClick(Sender: TObject);
    procedure tm0Timer(Sender: TObject);
    procedure HotKeyScr(Sender: TObject);
    procedure HotKeyStartScript(
      Sender: TObject);
    procedure HotKeyRec(
      Sender: TObject);
    procedure btLoadClick(Sender: TObject);
    procedure btSaveClick(Sender: TObject);
    procedure miComClick(Sender: TObject);
    procedure edScrExit(Sender: TObject);
    procedure miRadioClick(Sender: TObject);
    procedure miNewClick(Sender: TObject);
    procedure miExitClick(Sender: TObject);
    procedure miCtrlBClick(Sender: TObject);
    procedure HotKeyRecStop(
      Sender: TObject);
    procedure HotKeyPlay(Sender: TObject);
    procedure miSaveClick(Sender: TObject);
    procedure HotKeyManager1HotKeyshk1HotKeyActivation(Sender: TObject);
    procedure ed1Enter(Sender: TObject);
    procedure btAddColClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function GetWord(s:string;n:integer):string;   // получение n-нного слова, состоящего из значащих символов (0-1, A-z, А-я)
    procedure Scan(ind:integer);
    procedure MouseClick(button:byte; XY:integer);
  end;

TLegalSymbols = set of char;

var
  fmMain                  :TfmMain;
  wnd1,wnd2               :HWND; //окно ультимы
  doScript,isRec          :boolean;
  ptC                     :integer ;   // куда щелкать через интервал
  doBreak                 :integer;    // уровень прерывания цикла. если 1 или ничего - то прерывается текущий цикл, если 2 и более - то соотв. число родительских циклов
  fileNameSave            :string;

  
const

  KeyCodes: array[0..27] of Integer =
    (
     VK_F1, VK_F2, VK_F3, VK_F4, VK_F5, VK_F6, VK_F7, VK_F8, VK_F9, VK_F10, VK_F11, VK_F12,
     VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT, VK_ESCAPE, VK_TAB, VK_INSERT, VK_DELETE,
     VK_HOME, VK_END, VK_PRIOR, VK_NEXT, VK_BACK, VK_RETURN, VK_PAUSE, VK_SCROLL );

  KeyLabels: array[0..27] of string =
    (
     'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12',
     'Up', 'Down', 'Left', 'Right', 'Escape', 'Tab', 'Insert', 'Delete',
     'Home', 'End', 'PageUp', 'PageDown','Backspace','Enter','Pause', 'ScrolLsock');

LegalSymbols:TLegalSymbols = ['0'..'9','A'..'z','А'..'я']; // допустимые символы в слове (включая рус. алфавит)

implementation

{$R *.DFM}


function TfmMain.GetWord(s:string; n:integer):string;  // n=0 - первое слово в строке
var ts:string; i,tn:integer;  // ts - текущее слово, tn - текущий номер слова в строке
begin
// ИДЕЯ: перебираем по очереди все символы, Если легальный, то добавляем в ts. Если нелегальный, то если до него тоже был нелегальный (т.е. ts='' , еще не успели заполнить), то далее. Если нелегальный, но ts<>'' (т.е. заполняли), то возможно это наше искомое слово, проверим его номер.
tn:=0; ts:='';             // обнуляем на всякий случай (все равно они по нулям :-))
// ------ сам алгоритм -------
for i:=1 to length(s) do    // перебираем каждый символ
begin
// если n=-1, то получим всю строку после первого слова, например: s='say vendor buy', получим 'vendor buy'
if (s[i] in LegalSymbols)or((n=-1) and (tn=1)) then begin ts:=ts+s[i]; continue; end;  // увеличиваем допустимую строку
//--- встретился нелегальный символ
if ts='' then continue; // это значит, перед ним тоже был нелегальный символ. Просто пропускаем последовательность нелегальных символов
if tn=n then break; // ок, получили искомое слово (в нижнем регистре, с рус. алфавитом могут быть траблы!)
// слово кончилось, но это не наш номер слова! увеличим счетчик и очистим ts
inc(tn);
ts:='';
end; //for
//----------------------------
if (n=-1) or (n=tn) then result:=LowerCase(ts) else result:=''; // если '', то ничего не нашли
end;



procedure TfmMain.Scan(ind:integer);
var i,k,sel,key,nif,c,nrepeat,nElse:integer; s,str,command:string; FirstTick:Longint; DC:HDC; pt,pt1:integer;   //str - временная строка

             procedure Delay(ms:string);      // пауза в миллисекундах
             begin
                    if (ms='') or (ms='0') then exit;
                    FirstTick := GetTickCount;
                     repeat
                         if doScript=false then begin btStart.Down:=false; mmScript.Enabled:=true; exit; end;  // если отжали кнопку старта скрипта
                         Application.ProcessMessages;
                     until GetTickCount - FirstTick >= strtoint(ms);
             end;

begin
try
{сканирование и выполнение скрипта}
if mmScript.Text='' then begin btStart.Down:=false; mmScript.Enabled:=true; exit; end; // пустой скрипт, выход
i:=ind; // это если мы начинаем не с начала скрипта, например, цикл или условие if
while i<mmScript.Lines.Count do begin     // для каждой строки скрипта...
Delay(edPause.Text); // пауза между строками скрипта
if mmScript.Lines.Strings[i]='' then begin inc(i); continue; end; // текущая строчка пустая
if doScript=false then begin btStart.Down:=false; mmScript.Enabled:=true; exit; end;  // если отжали кнопку старта скрипта
s:=mmScript.Lines.Strings[i]; //текущая строка

//--- сейчас выделим первый символ в текущей строке (блин, неужто это я сам написал? :-))
sel:=0;
for k:=0 to i do sel:=sel+(length(mmScript.Lines.Strings[k])+2);
sel:=sel-(length(mmScript.Lines.Strings[i])+2);
mmScript.SelStart:=sel;mmScript.SelLength:=1;
//---

// ----- далее обрабатываем комманды --------------

command:=GetWord(s,0);  // чтобы для одной строчки несколько раз не рассчитывать команду (первое слово)

if command='wait' then delay(GetWord(s,1)); // -wait
//----------------------------------------
if command='msg' then showmessage(GetWord(s,-1));
//----------------------------------------
if command='say' then begin
str:=GetWord(s,-1);
for k:=1 to length(str) do begin
PostMessage(wnd1,WM_KEYDOWN,ord(str[k]),0);
PostMessage(wnd1,WM_CHAR,ord(str[k]),0);
PostMessage(wnd1,WM_KEYUP,ord(str[k]),0);
                             end;// - for
PostMessage(wnd1,WM_CHAR,VK_RETURN,0);
                   end; // -say

//----------------------------------------
if command='send' then begin
str:=GetWord(s,1);  // клавиша типа backspace  (учтите, что LowerCase(keylabels[k]), т.к. сравниваем строки, а при GetWord получаем в нижнем регистре)
for k:=0 to high(keylabels) do if comparetext(str,LowerCase(keylabels[k]))=0 then key:=keycodes[k];
PostMessage(wnd1,WM_KEYDOWN ,key,0);
PostMessage(wnd1,WM_KEYUP,key,0);
if GetWord(s,2)<>'' then Delay(GetWord(s,2));  // задержка после нажатия на клавишу (если есть)
                    end;
//----------------------------------------
if (command='if')or(command='if_not') then begin
//ИДЕЯ: если соотв. условие выполняется, то вызов Scan(i+1), если нет то ищем наш else и Scan(numElse). После этого ищем наш end_if и переходим на него (а т.к. в конце inc(i), то попадем на следующую строчку после него)

//--- Ищем наш end_if и nElse - номер строки с нашим else:
k:=i+1;     // номер строки с нашим end_if , начинаем считать кол-во вложенных условий if и if_not со след. от текущей
nif:=0;     // текущее кол-во вложенных условий if и if-not. если <0 то мы нашли наш end_if
nElse:=0;   // номер строки с нашим else, если есть. если нету, то 0
while k<mmScript.Lines.Count do begin
if (GetWord(mmScript.Lines.Strings[k],0)='if') or (GetWord(mmScript.Lines.Strings[k],0)='if_not') then inc(nif);
if GetWord(mmScript.Lines.Strings[k],0)='end_if' then dec(nif);
if GetWord(mmScript.Lines.Strings[k],0)='else' then if nif=0 then nElse:=k; // если кол-во вложен. условий=0 (т.е. наше условие)
if nif<0 then begin break; end;   // k=номеру строки с нашим end_if
inc(k);
                                end;
if k=mmScript.Lines.Count then begin doScript:=false; btStart.Down:=false; mmScript.Enabled:=true; showmessage('Не могу найти конец условия: "End_IF", проверьте скрипт'); exit; end; // цикл while дошел до конца скрипта и не нашел end_if

//---  теперь определим, равен ли цвет точки или нет.
DC:=GetDC(wnd1);
c:=GetPixel(DC,strtoint(GetWord(s,1)),strtoint(GetWord(s,2))); // 1 и 2 параметры -координаты точки, 3 - цвет точки, с кот. надо сравнить
ReleaseDC(wnd1,DC);
//--- переход согласно условию (а если есть else, то на него, если условие не выполняется)
if GetWord(s,4)='' then begin   // это если четко указан цвет
if command='if' then if c=strtoint(GetWord(s,3)) then Scan(i+1) else if nElse<>0 then Scan(nElse+1);
if command='if_not' then if c<>strtoint(GetWord(s,3)) then Scan(i+1) else if nElse<>0 then  Scan(nElse+1);
                        end 
                        else begin // это если указан диапазон цвета
if command='if' then if ((c>=strtoint(GetWord(s,3)))and(c<=strtoint(GetWord(s,4)))) then Scan(i+1) else if nElse<>0 then Scan(nElse+1);
if command='if_not' then if not ((c>=strtoint(GetWord(s,3))) and(c<=strtoint(GetWord(s,4)))) then Scan(i+1) else if nElse<>0 then  Scan(nElse+1);

                             end;
//---
i:=k; // k=end_repeat, а в самом конце inc(i), поэтому попадем на следующую строчку после end_if
                 end; // -"if"


//------------------------------------------
if command='left' then begin
MouseClick(1,MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))));
                   end else // -left
if command='right' then begin
MouseClick(2,MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))));
                   end; // -left
if command='double_left' then begin
MouseClick(11,MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))));
                   end; // -double left & right
if command='double_right' then begin
MouseClick(22,MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))));
                   end; // -double left & right
//-------------------------------------------------
if command='alarm' then begin
PlaySound(pChar('Msg.wav'), 0, SND_FILENAME or SND_ASYNC or SND_NOWAIT);
delay('400');  // время на проигрывание файла.
                             end; //alarm
//--------------------------------------------------
if command='drag' then begin
pt:=MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))); // (x1,y1) (x2,y2) (кoличество, если не указано или all, то все)
pt1:=MakeLong(strtoint(GetWord(s,3)),strtoint(GetWord(s,4)));
// --- собственно перетаскивание:
PostMessage(wnd1, WM_LBUTTONDOWN,0,pt);
PostMessage(wnd1, WM_SETCURSOR, wnd1, MakeLong(HTCLIENT,WM_LBUTTONDOWN));
PostMessage(wnd1, WM_MOUSEMOVE,0,pt1);
PostMessage(wnd1, WM_SETCURSOR, wnd1, MakeLong(HTCLIENT,WM_MOUSEMOVE));
delay('400');   // иначе не просекает...
PostMessage(wnd1, WM_LBUTTONUP,0,pt1);
PostMessage(wnd1, WM_SETCURSOR, wnd1, MakeLong(HTCLIENT,WM_LBUTTONUP));
str:=GetWord(s,5);        // Это строковое представление кол-ва перетаскиваемых объектов
//введем сколько  (достаточно WM_CHAR, а WM_KEYDOWN/UP и не надо???)
if (str<>'')and(str<>'all') then for k:=1 to length(str) do begin PostMessage(wnd1,WM_CHAR,ord(str[k]),0);end;
//в любом случае нажмем Enter, а если перед этим с клавиатуры не ввели число, то значит все
PostMessage(wnd1,WM_CHAR,VK_RETURN,0);
// щелкнем, чтобы бросить
PostMessage(wnd1, WM_LBUTTONDOWN,0,pt1);
PostMessage(wnd1, WM_SETCURSOR, wnd1, MakeLong(HTCLIENT,WM_LBUTTONDOWN));
PostMessage(wnd1, WM_LBUTTONUP,0,pt1);
PostMessage(wnd1, WM_SETCURSOR, wnd1, MakeLong(HTCLIENT,WM_LBUTTONUP));
                   end; // -drag
//--------------------------------------------------
if command='end_script' then doScript:=false;     // потому что эту функцию (Scan) мы можем вызывать рекурсивно, так что просто выход не поможет
//--------------------------------------------------
if (command='end_repeat') or (command='continue') or (command='else') or (command='end_if') then exit;  // в родительскую процедуру, возможно, сами себе (рекурсия)
//--------------------------------------------------
if command='break' then begin if (GetWord(s,1)='')or (GetWord(s,1)='0') then doBreak:=1 else doBreak:=strtoint(GetWord(s,1)); exit; end; // в родительскую процедуру, возможно, сами себе (рекурсия)
//--------------------------------------------------
if command='repeat' then begin
//ИДЕЯ: вначале вызовем сколько надо раз scan(i+1) (тогда при встрече end_repeat вернемся в эту точку), а потом перескочим на строчку с нашим end_repeat.
for k:=1 to strtoint(GetWord(s,1)) do begin
                                      if doBreak>0 then begin doBreak:=doBreak-1; break; end; // если в теле цикла встретим break , то прекратим выполнение текущего цикла.
                                      Scan(i+1);
                                      end;
//---  перескочим на наш end_repeat
k:=i+1;     // Номер след. строки после repeat
nrepeat:=0; // число вложенных циклов

while k<mmScript.Lines.Count do begin
if GetWord(mmScript.Lines.Strings[k],0)='repeat' then inc(nrepeat);
if GetWord(mmScript.Lines.Strings[k],0)='end_repeat' then dec(nrepeat);
if nrepeat<0 then begin i:=k; break; end;   // сейчас i=k=номеру с нашим end_repeat, увеличится этот номер строки в конце общего цикла, в итоге перескочим на строку после endrepeat
inc(k);
                                end;
if k=mmScript.Lines.Count then begin doScript:=false; btStart.Down:=false; mmScript.Enabled:=true; showmessage('Не могу найти конец цикла: "End_Repeat", проверьте скрипт'); exit; end; // не нашли соотв. конец цикла end_repeat
//---
                          end; // -repeat
//--------------------------------------------------
if command='move' then begin // сдвигаем курсор, аналог SetCursorPos(x,y);
PostMessage(wnd1,WM_MOUSEMOVE,0,MakeLong(strtoint(GetWord(s,1)),strtoint(GetWord(s,2))));
                       end;
//--------------------------------------------------
                          
//------------------- конец проверки операторов (команд) -----------------------

inc(i); // увеличиваем номер строки в мактросе
                                end;  //-while обработка строки
Application.ProcessMessages;
if doScript then Scan(0) else begin btStart.Down:=false; mmScript.Enabled:=true; exit; end;
except
doScript:=false; btStart.Down:=false; mmScript.Enabled:=true; showmessage('Ошибка! Проверьте правильность скрипта!');
end;

end;  //процедуры Scan



procedure TfmMain.MouseClick(button:byte; XY:integer);
begin
if button=11 then begin {дв. клик}
PostMessage(wnd1,WM_LBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONDOWN));
PostMessage(wnd1,WM_LBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONUP));
//----- хз! может этого раздела не надо, но УО тогда воспринимает как одинарный клик
PostMessage(wnd1,WM_LBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONDOWN));
PostMessage(wnd1,WM_LBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONUP));
//-------
PostMessage(wnd1,WM_LBUTTONDBLCLK,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONDBLCLK));
PostMessage(wnd1,WM_LBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONUP));
                        end;

if button=22 then begin {дв. клик правой}
PostMessage(wnd1,WM_RBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONDOWN));
PostMessage(wnd1,WM_RBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONUP));
//----- хз! может этого раздела не надо, но УО тогда воспринимает как одинарный клик
PostMessage(wnd1,WM_RBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONDOWN));
PostMessage(wnd1,WM_RBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONUP));
//-------
PostMessage(wnd1,WM_RBUTTONDBLCLK,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONDBLCLK));
PostMessage(wnd1,WM_RBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONUP));
                        end;

if button=1 then begin {левой клик}
PostMessage(wnd1,WM_LBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONDOWN));
PostMessage(wnd1,WM_LBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_LBUTTONUP));
                        end;

if button=2 then begin {правой клик}
PostMessage(wnd1,WM_RBUTTONDOWN,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONDOWN));
PostMessage(wnd1,WM_RBUTTONUP,0,XY);
PostMessage( wnd1,WM_SETCURSOR,wnd1,MakeLong(HTCLIENT,WM_RBUTTONUP));
                        end;
end;


procedure TfmMain.FormCreate(Sender: TObject);
var i:integer;
begin
TheRecorder.SpeedFactor:=miSpeed.Tag;

// заполним ниспадающие списки. тут по хитрому: заполняем всего 1 список, а остальные просто ссылаются на заполненный тот единственный список у cb1
cb1.Items.Clear;
for i:=0 to High(keylabels) do begin
cb1.Items.Add(keylabels[i]);
                            end;
cb2.Items:=cb1.Items;
cb3.Items:=cb1.Items;
cb4.Items:=cb1.Items;
cb5.Items:=cb1.Items;
cbM.Items:=cb1.Items;
//-заполнение

end;

procedure TfmMain.SetCoord(Sender: TObject);
var c: Integer; p:tpoint; DC:HDC;
begin
GetCursorPos(p);
wnd1:=WindowFromPoint(p);
if wnd1=0 then begin showmessage('Не могу найти окно UO'); exit; end;
windows.ScreenToClient(wnd1,p);
if pcAll.ActivePage=tsGeneral then begin   // только для щелкалки
ptC:=MakeLong(p.x,p.y);   // это коорд. куда щелкать
btS0.Caption:=inttostr(p.x)+', '+inttostr(p.y);
                                   end;

if pcAll.ActivePage=tsScript then begin   // только для скрипта
DC:=GetDC(wnd1);
c:=GetPixel(DC,p.x,p.y);
btColor.Caption:=inttostr(c);
btXY.Caption:=inttostr(p.x)+', '+inttostr(p.y);
if cbInsertXY.Checked then mmScript.SelText:=btXY.Caption+' ';
ReleaseDC(wnd1,DC);
                                   end;
end;

procedure TfmMain.btStartClick(Sender: TObject);        // старт/стоп скрипта
begin
if btStart.Down then begin            // только что нажали (пуск)
if wnd1=0 then wnd1:=FindWindow(Pchar('Ultima Online'),nil);
if wnd1=0 then begin btStart.Down:=false; showmessage('Не могу найти UO, попробуйте указать его принудительно:'+#13+#10+'Установите курсор мыши над окном UO и нажмите "Ctrl+A"');exit; end;

mmScript.Enabled:=false;
doScript:=true;                       // это значит что запустили скрипт
scan(0);                              // и начинаем сканировать текст с первой строки
                                end else begin          // только что отжали (стоп)
mmScript.Enabled:=true;
doScript:=false;
                                         end;
end;


procedure TfmMain.btCStartClick(Sender: TObject);
var interval:integer;  c:TComponent;
begin
//если щелкнули по кнопкам для 2-го окна, проверим, есть ли окна uo
if ((sender as TSpeedButton).Name='btS4')or((sender as TSpeedButton).Name='btS5') then
begin
if wnd2=0 then wnd2:=FindWindow(Pchar('Ultima Online'),nil);
if wnd2=0 then begin (Sender as TSpeedButton).Down:=false; showmessage('Не могу найти UO, попробуйте указать его принудительно:'+#13+#10+'Установите курсор мыши над окном UO и нажмите "Ctrl+B"');exit; end;
end else begin  // иначе для первого окна
if wnd1=0 then wnd1:=FindWindow(Pchar('Ultima Online'),nil);
if wnd1=0 then begin (Sender as TSpeedButton).Down:=false; showmessage('Не могу найти UO, попробуйте указать его принудительно:'+#13+#10+'Установите курсор мыши над окном UO и нажмите "Ctrl+A"');exit; end;
         end;
// --- теперь запустим/остановим таймеры ---
if (Sender as TSpeedButton).Down then begin  // только что нажали кнопку
// проверим, выбрана ли клавиша, используя Tag кнопки для определения cb и ed
if (Sender as TSpeedButton).Name='btS0' then begin if cb0.Text='' then begin showmessage('Выберите тип кликания'); btS0.Down:=false; exit; end; end // 2 end из-за else
else if (fmMain.FindComponent('cb'+inttostr((Sender as TComponent).Tag)) as TComboBox).Text='' then begin showmessage('Задайте клавишу'); (Sender as TSpeedButton).Down:=false; exit; end;
  try
interval:=strtoint((fmMain.FindComponent('ed'+inttostr((Sender as TComponent).Tag)) as TEdit).Text);
(fmMain.FindComponent('tm'+inttostr((Sender as TComponent).Tag)) as TTimer).Interval:=interval;
(fmMain.FindComponent('tm'+inttostr((Sender as TComponent).Tag)) as TTimer).Enabled:=true;
//--- отключим enable для cd и ed
(fmMain.FindComponent('cb'+inttostr((Sender as TComponent).Tag)) as TComboBox).Enabled:=false;
(fmMain.FindComponent('ed'+inttostr((Sender as TComponent).Tag)) as TEdit).Enabled:=false;
//---
  except showmessage('Задайте правильный интервал (в миллисекундах)'); (Sender as TSpeedButton).Down:=false; exit; end;
                                      end else begin  // только что отжали кнопку
(fmMain.FindComponent('tm'+inttostr((Sender as TComponent).Tag)) as TTimer).Enabled:=false;
//--- включим enable для cd и ed
(fmMain.FindComponent('cb'+inttostr((Sender as TComponent).Tag)) as TComboBox).Enabled:=true;
(fmMain.FindComponent('ed'+inttostr((Sender as TComponent).Tag)) as TEdit).Enabled:=true;
//---
                                               end;
end;


procedure TfmMain.tm0Timer(Sender: TObject);
var s1:string; i:integer;         // PostMessage(wnd,WM_MOUSEMOVE,0{MK_LBUTTON},ptC.l);
begin

if (Sender as TTimer).Name='tm0' then begin
if cb0.ItemIndex=0 then {дв. клик}        MouseClick(11,ptC);
if cb0.ItemIndex=1 then {дв. клик правой} MouseClick(22,ptC);
if cb0.ItemIndex=2 then {левой клик}      MouseClick(1,ptC);
if cb0.ItemIndex=3 then {правой клик}     MouseClick(2,ptC);
                                         end
   else begin
//--- определим клавишу, используя Tag у таймера
s1:=(fmMain.FindComponent('cb'+inttostr((Sender as TComponent).Tag)) as TComboBox).Text;
//---
try
i:= 0;
  while i < High(KeyLabels) do
  begin
    Inc(i);
    if CompareText(s1, KeyLabels[i]) = 0 then
      Break;
  end;
if ((sender as TTimer).Name='tm4')or((sender as TTimer).Name='tm5') then
begin
PostMessage(wnd2,WM_KEYDOWN ,keycodes[i],0 );
PostMessage(wnd2,WM_KEYUP,keycodes[i],0);
end else begin
PostMessage(wnd1,WM_KEYDOWN ,keycodes[i],0 );
PostMessage(wnd1,WM_KEYUP,keycodes[i],0);
         end;
except
showmessage('Ошибка при попытке определить клавишу');
end;
        end;
end;

procedure TfmMain.HotKeyScr(Sender: TObject);    // скринсэйвер
var fs:TSearchRec; attr,n, Found:integer; bmp: TBitmap; jpg:TJPEGImage; DC: HDC; s:string;
begin
try
{}
// --------  если сохраняем по номеру, то -----------
if not cbDate.Checked then begin
n:=0;
Found := FindFirst(ExtractFilePath(edScr.text)+'*.*',attr,fs);
while Found = 0 do
begin
   try
// учитывая, что имена файлов выдаются по алфавиту, а не по порядковому номеру (10 идет раньше 9 !)
if copy(fs.Name,1,3)='Pic' then if n<strtoint(copy(fs.Name,4,Length(fs.Name)-7)) then n:=strtoint(copy(fs.Name,4,Pos('.',fs.Name)-4));
   except //файл начинается на Pic , но не удалось распознать номер
   end; // except
Found := FindNext(fs);
end;  // while
FindClose(fs);
inc(n);  // теперь минимум 1, а если есть файлы, то на 1 больше последнего номера
s:=ExtractFilePath(edScr.text)+'Pic'+inttostr(n);   // имя файла без расширения (!), т.к. потом определим, jpg или bmp
                      end;// if miNum
// --------  если сохраняем по дате, то -----------
if cbDate.Checked then begin
s:=ExtractFilePath(edScr.text)+'Pic'+FormatDateTime('dd.mm_hh.nn.ss',Now);
// если такой файл существует, то создадим файл как бы на 1 секунду позже
if FileExists(s+'.bmp') or FileExists(s+'.jpg') then s:=ExtractFilePath(edScr.text)+'Pic'+FormatDateTime('dd.mm_hh.nn.ss',Now+strtotime('00:00:01'));
                       end; // if midate

// --- сам скрин сэйвер ------
DC:=GetDC(0);

bmp:=TBitmap.Create;
bmp.Height:=Screen.Height;
bmp.Width:=Screen.Width;
bitblt(bmp.Canvas.Handle, 0, 0, Screen.Width, Screen.Height,DC, 0, 0, SRCCOPY);

if rbJpg.Checked then begin
jpg:=TJPEGImage.Create;
jpg.Assign(bmp);
jpg.CompressionQuality:=70;
jpg.Compress;
jpg.SaveToFile(s+'.jpg');
jpg.Free;
                      end;

if rbBmp.Checked then begin
bmp.SaveToFile(s+'.bmp');
                      end; // if bmp

bmp.FreeImage;
bmp.Free;

ReleaseDC(0, DC);
//----------------------------
except
showmessage('Ошибка! Убедитесь, что каталог, куда сохранять экран, существует');
end;
end;

procedure TfmMain.HotKeyStartScript(
  Sender: TObject);
begin
{}
pcAll.ActivePage:=tsScript;
if not btStart.Down then begin
if wnd1=0 then wnd1:=FindWindow(Pchar('Ultima Online'),nil);
if wnd1=0 then begin btStart.Down:=false; showmessage('Не могу найти UO, попробуйте указать его принудительно:'+#13+#10+'Установите курсор мыши над окном UO и нажмите "Ctrl+A"');exit; end;
btStart.Down:=true;
mmScript.Enabled:=false;
doScript:=true;
scan(0);
                                end else begin
btStart.Down:=false;
mmScript.Enabled:=true;
doScript:=false;
                                         end;
end;

procedure TfmMain.HotKeyRec(Sender: TObject);  // начало записи макроса
begin
    {}
TheRecorder.FStream.free;
TheRecorder.FStream := TMemoryStream.Create;

  TheRecorder.SpeedFactor:=miSpeed.Tag;
  TheRecorder.DoStop;
  TheRecorder.DoRecord(true);
end;

procedure TfmMain.btLoadClick(Sender: TObject);  // загрузка скрипта
var s:string; f:textfile;
begin
odLoad.InitialDir:=ExtractFilePath(Application.ExeName);
odLoad.FileName:='';
odLoad.Title:='Загрузить скрипт...';
if odLoad.Execute then assignfile(f,odLoad.FileName) else exit;
reset(f);
mmScript.Clear;
while not eof(f) do begin
ReadLn(f,s);
mmScript.Lines.Add(s);
                    end;
CloseFile(f);
end;

procedure TfmMain.btSaveClick(Sender: TObject);   // сохранение скрипта
var s:string; i:integer; f:textfile;
begin
sdSave.InitialDir:=ExtractFilePath(Application.ExeName);
sdSave.FileName:='';
sdSave.Title:='Сохранить скрипт как...';
if sdSave.Execute then assignfile(f,sdSave.FileName) else exit;
fileNameSave:=sdSave.FileName;
rewrite(f);
i:=0;
while i<mmScript.Lines.Count do begin
s:=mmScript.Lines.Strings[i];
WriteLn(f,s);
inc(i);
                    end;
CloseFile(f);
end;

procedure TfmMain.miComClick(Sender: TObject);  // щелкнули по контекстному меню с командами
var s:string;
begin
// команды Copy, Paste, Cut, Undo
if (Sender as TMenuItem).Name='miPaste' then begin mmScript.PasteFromClipboard; exit; end;
if (Sender as TMenuItem).Name='miCopy' then begin mmScript.CopyToClipboard; exit; end;
if (Sender as TMenuItem).Name='miCut' then begin mmScript.CutToClipboard; exit; end;
if (Sender as TMenuItem).Name='miUndo' then begin mmScript.Perform(WM_UNDO,0,0); exit; end;

s:=GetWord((Sender as TMenuItem).Caption,0)+' ';
mmScript.SelText:=s;
end;


procedure TfmMain.edScrExit(Sender: TObject);  // добавим слэш в конце пути. А надо?
begin
if copy(edScr.text,length(edScr.text),1)<>'\' then edScr.text:=edScr.text+'\';
end;


procedure TfmMain.miRadioClick(Sender: TObject);  // выбираем меню checked + определение скорости и кол-ва повтора
begin
(Sender as TMenuItem).Checked:=true;
if (Sender as TMenuItem).Parent.Name='miSpeed' then TheRecorder.SpeedFactor:=(Sender as TMenuItem).Tag;
if (Sender as TMenuItem).Parent.Name='miRepeat' then TheRecorder.Tag:=(Sender as TMenuItem).Tag;

end;

procedure TfmMain.miNewClick(Sender: TObject);  // подтверждение на очистку скрипта TODO: надо только если не сохранен
begin
if mmScript.Text<>'' then if MessageDlg('Вы уверены, что хотите'+#13+'очистить существующий скрипт?',mtConfirmation,[mbYes,mbNo],0)=mrYes then mmScript.Clear;
end;


procedure TfmMain.miExitClick(Sender: TObject);    // выход
begin
Application.Terminate;
end;

procedure TfmMain.miCtrlBClick(Sender: TObject);    // выбор 2-го окна
var p:tpoint;
begin
GetCursorPos(p);
wnd2:=WindowFromPoint(p);
if wnd2=0 then begin showmessage('Не могу найти окно UO'); exit; end
else gbOtherWindow.Caption:='Выбрано! Для выбора другого: (Ctrl+B)';
end;


procedure TfmMain.HotKeyRecStop(Sender: TObject);   // стоп записи и проигрывания  макро
begin
  TheRecorder.DoStop;
end;

procedure TfmMain.HotKeyPlay(Sender: TObject);      // воспроизведение макро
var i:integer;
begin
  TheRecorder.DoStop;
  TheRecorder.DoPlay;
end;

procedure TfmMain.miSaveClick(Sender: TObject);      // сохранение скрипта
var s:string;i:integer; f:textfile;
begin
if fileNameSave='' then begin miSaveAs.Click; exit; end;
assignfile(f,fileNameSave);
rewrite(f);
i:=0;
while i<mmScript.Lines.Count do begin
s:=mmScript.Lines.Strings[i];
WriteLn(f,s);
inc(i);
                    end;
CloseFile(f);
end;

procedure TfmMain.HotKeyManager1HotKeyshk1HotKeyActivation(Sender: TObject);
begin
//showmessage(Copy((Sender as THotKeyItem).Name,3,1));
(fmMain.FindComponent('btS'+Copy((Sender as THotKeyItem).Name,3,1)) as TSpeedButton).Down:= not (fmMain.FindComponent('btS'+Copy((Sender as THotKeyItem).Name,3,1)) as TSpeedButton).Down;
(fmMain.FindComponent('btS'+Copy((Sender as THotKeyItem).Name,3,1)) as TSpeedButton).click;
end;

procedure TfmMain.ed1Enter(Sender: TObject);
begin
(sender as TEdit).SelectAll;
end;




procedure TfmMain.btAddColClick(Sender: TObject);
var s:string;
begin
// добавим координаты и цвет, либо клавишу и паузу
if (Sender as TSpeedButton).Name='btColor' then s:=btColor.Caption+' ';
if (Sender as TSpeedButton).Name='btXY' then s:=btXY.Caption+' ';
if (Sender as TSpeedButton).Name='btAddM' then if cbM.Text='' then begin showmessage('Задайте клавишу'); exit; end else s:=cbM.text+' '+edM.text;

mmScript.SelText:=s;
mmScript.SetFocus;
end;

end.
