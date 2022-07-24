function varargout = Projek_Akhir_PraktikumSCPK(varargin)
% PROJEK_AKHIR_PRAKTIKUMSCPK MATLAB code for Projek_Akhir_PraktikumSCPK.fig
%      PROJEK_AKHIR_PRAKTIKUMSCPK, by itself, creates a new PROJEK_AKHIR_PRAKTIKUMSCPK or raises the existing
%      singleton*.
%
%      H = PROJEK_AKHIR_PRAKTIKUMSCPK returns the handle to a new PROJEK_AKHIR_PRAKTIKUMSCPK or the handle to
%      the existing singleton*.
%
%      PROJEK_AKHIR_PRAKTIKUMSCPK('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PROJEK_AKHIR_PRAKTIKUMSCPK.M with the given input arguments.
%
%      PROJEK_AKHIR_PRAKTIKUMSCPK('Property','Value',...) creates a new PROJEK_AKHIR_PRAKTIKUMSCPK or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Projek_Akhir_PraktikumSCPK_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Projek_Akhir_PraktikumSCPK_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Projek_Akhir_PraktikumSCPK

% Last Modified by GUIDE v2.5 13-May-2022 00:19:08

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Projek_Akhir_PraktikumSCPK_OpeningFcn, ...
                   'gui_OutputFcn',  @Projek_Akhir_PraktikumSCPK_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Projek_Akhir_PraktikumSCPK is made visible.
function Projek_Akhir_PraktikumSCPK_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Projek_Akhir_PraktikumSCPK (see VARARGIN)

% Choose default command line output for Projek_Akhir_PraktikumSCPK
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Projek_Akhir_PraktikumSCPK wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Projek_Akhir_PraktikumSCPK_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function nama_Callback(hObject, eventdata, handles)
% hObject    handle to nama (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of nama as text
%        str2double(get(hObject,'String')) returns contents of nama as a double


% --- Executes during object creation, after setting all properties.
function nama_CreateFcn(hObject, eventdata, handles)
% hObject    handle to nama (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c1_Callback(hObject, eventdata, handles)
% hObject    handle to c1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c1 as text
%        str2double(get(hObject,'String')) returns contents of c1 as a double


% --- Executes during object creation, after setting all properties.
function c1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c2_Callback(hObject, eventdata, handles)
% hObject    handle to c2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c2 as text
%        str2double(get(hObject,'String')) returns contents of c2 as a double


% --- Executes during object creation, after setting all properties.
function c2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c3_Callback(hObject, eventdata, handles)
% hObject    handle to c3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c3 as text
%        str2double(get(hObject,'String')) returns contents of c3 as a double


% --- Executes during object creation, after setting all properties.
function c3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c4_Callback(hObject, eventdata, handles)
% hObject    handle to c4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c4 as text
%        str2double(get(hObject,'String')) returns contents of c4 as a double


% --- Executes during object creation, after setting all properties.
function c4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c5_Callback(hObject, eventdata, handles)
% hObject    handle to c5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c5 as text
%        str2double(get(hObject,'String')) returns contents of c5 as a double


% --- Executes during object creation, after setting all properties.
function c5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c6_Callback(hObject, eventdata, handles)
% hObject    handle to c6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c6 as text
%        str2double(get(hObject,'String')) returns contents of c6 as a double


% --- Executes during object creation, after setting all properties.
function c6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c7_Callback(hObject, eventdata, handles)
% hObject    handle to c7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c7 as text
%        str2double(get(hObject,'String')) returns contents of c7 as a double


% --- Executes during object creation, after setting all properties.
function c7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c8_Callback(hObject, eventdata, handles)
% hObject    handle to c8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c8 as text
%        str2double(get(hObject,'String')) returns contents of c8 as a double


% --- Executes during object creation, after setting all properties.
function c8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c9_Callback(hObject, eventdata, handles)
% hObject    handle to c9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c9 as text
%        str2double(get(hObject,'String')) returns contents of c9 as a double


% --- Executes during object creation, after setting all properties.
function c9_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function c10_Callback(hObject, eventdata, handles)
% hObject    handle to c10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of c10 as text
%        str2double(get(hObject,'String')) returns contents of c10 as a double


% --- Executes during object creation, after setting all properties.
function c10_CreateFcn(hObject, eventdata, handles)
% hObject    handle to c10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in simpan.
function simpan_Callback(hObject, eventdata, handles)
% hObject    handle to simpan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

nama = get(handles.nama, 'String');
dc1 = str2num(get(handles.c1, 'String'));
dc2 = str2num(get(handles.c2, 'String'));
dc3 = str2num(get(handles.c3, 'String'));
dc4 = str2num(get(handles.c4, 'String'));
dc5 = str2num(get(handles.c5, 'String'));
dc6 = str2num(get(handles.c6, 'String'));
dc7 = str2num(get(handles.c7, 'String'));
dc8 = str2num(get(handles.c8, 'String'));
dc9 = str2num(get(handles.c9, 'String'));
dc10 = str2num(get(handles.c10, 'String'));

Data = [{nama} dc1 dc2 dc3 dc4 dc5 dc6 dc7 dc8 dc9 dc10];

tabeldata = readcell('Data.xlsx');
tabeldata = [tabeldata; Data];
writecell(tabeldata,'Data.xlsx');
 
% tampil data
%p.Data = [p.Data; [{nama} {c1} {c2} {c3} {c4} {c5} {c6} {c7} {c8} {c9} {c10}] ];
%set(handles.DataTabel, 'Data', p.Data);

% setelah data masuk tabel, inputan direset
reset = '';
set(handles.nama, 'String', reset);
set(handles.c1, 'String', reset);
set(handles.c2, 'String', reset);
set(handles.c3, 'String', reset);
set(handles.c4, 'String', reset);
set(handles.c5, 'String', reset);
set(handles.c6, 'String', reset);
set(handles.c7, 'String', reset);
set(handles.c8, 'String', reset);
set(handles.c9, 'String', reset);
set(handles.c10, 'String', reset);

%Excel.ActiveWorkbook.Save;
%Excel.Quit;
%fclose('all');

% --- Executes on button press in preverensi.
function preverensi_Callback(hObject, eventdata, handles)
% hObject    handle to preverensi (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%ambil data dari tabel
DataIsi = readtable('Data.xlsx','ReadVariableNames',false);
DataNama = DataIsi(:,1);
DataIsi = DataIsi(:,2:11);
DataIsi = table2array(DataIsi);
DataNama = table2array(DataNama);

% Kriteria => benefit semua
k = [1 1 1 1 1 1 1 1 1 1];

% bobot
w = [4 3 3 4 5 4 2 3 1 5];

%tahapan pertama, perbaikan bobot
[m, n]=size (DataIsi); %inisialisasi ukuran Data
w = w./sum(w); %membagi bobot per kriteria dengan jumlah total seluruh bobot

%tahapan kedua, melakukan perhitungan vektor(S) per baris (alternatif)
% 1) perkalian dengan kriteria
for j=1:n
    if k(j)==0 
        w(j)= -1*w(j);
    end
end
% 2) Perpangkatan Rating dengan bobot
for i=1:m
    S(i)=prod(DataIsi(i,:).^w);
end

%tahapan ketiga, proses perangkingan
V= S/sum(S);

rank = sort(V, 'descend');

for i=1:m
    for j=1:m
        if rank(i) == V(j)
            NamaUrut(i) = DataNama(j);
        end
    end
end

Preverensi = num2cell(transpose(rank));
hasil = [Preverensi(:) NamaUrut(:)];
            
%xlswrite('Preverensi.xlsx',hasil);
set(handles.uitable2, 'Data', hasil);

[row, rowCons] = max(V);
set(handles.kesimpulan, 'String', DataNama(rowCons));

% --- Executes on button press in Tampil.
function Tampil_Callback(hObject, eventdata, handles)
% hObject    handle to Tampil (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
ReadData = readcell('Data.xlsx');
set(handles.DataTabel, 'Data', ReadData);
