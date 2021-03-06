% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Dependencies include amp.m, Cell2char.m, Cell2JavaString.m, dhpg.m, 
% excelcols.m, excelsave.m, freq.m, iocurve.m, jxl.jar, MXL.jar, ltd.m,
% ltp.m, nmda.m, percent_ltd.m, percent_ltp.m, ppr.m, ppr2.m,
% WriteXL.java, and xlwrite.m.


function varargout = Ephys_Analysis_OSX(varargin)
%EPHYS_ANALYSIS_OSX M-file for Ephys_Analysis_OSX.fig

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Ephys_Analysis_OSX_OpeningFcn, ...
                   'gui_OutputFcn',  @Ephys_Analysis_OSX_OutputFcn, ...
                   'gui_LayoutFcn',  [], ...
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


% --- Executes just before Ephys_Analysis_OSX is made visible.
function Ephys_Analysis_OSX_OpeningFcn(hObject, eventdata, handles, varargin)
% Choose default command line output for Ephys_Analysis_OSX
handles.output = hObject;
handles.home_dir = pwd;
handles.filename = 'nofile';
handles.pathname = 'nopath';
handles.worksheet = 'noworksheet';
handles.column_number = '1';
handles.column_letter = 'A';
handles.dhpg_sweep = 0;
handles.freq = 0;
handles.wash = 0;
handles.dhpg_on = 0;

% Update handles structure
guidata(hObject, handles)


% --- Outputs from this function are returned to the command line.
function varargout = Ephys_Analysis_OSX_OutputFcn(hObject, eventdata, handles)

varargout{1} = handles.output;



% --- Executes on button press in push_openfile.
function push_openfile_Callback(hObject, eventdata, handles)    
    
    %Choose the excel file
    
    [filename,pathname] = uigetfile('*.xlsx','Select the .xlxs workbook you want to analyze');
    
    if filename ~= 0
        
        % Change directory to the one with the excel file
        home_dir = pwd;
        cd (pathname);
        file_path = strcat (pathname, filename);
        set(handles.txt_file, 'String', file_path);
        
        % Now get the list of sheet names
        [status,sheets] = xlsfinfo(filename);
   
        % Now put the sheet names into the list box
        set(handles.listbox1,'String',sheets);
        
        % save file data
        handles.filename = filename;
        handles.pathname = pathname;
        guidata(hObject,handles);
        cd (home_dir);
    else
        clearvars;
    end
    
    
% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
% Pick a sheet. Any sheet.
    items = get(hObject,'String');
    numeric_index = get(hObject,'Value');
    item_selected = items{numeric_index};
    handles.worksheet = item_selected;
    guidata(hObject,handles);
    txt_status = strcat(handles.worksheet, ' is selected. Now choose an experiment type.')
    set(handles.txt_status,'String',txt_status);
    

% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)


% --- Executes on button press in push_clear.
function push_clear_Callback(hObject, eventdata, handles)
    set(handles.listbox1, 'String', '');
    set(handles.txt_status,'String','Please choose an experiment type');
    clearvars ;
    
% --- Executes on button press in push_help.
function push_help_Callback(hObject, eventdata, handles)
    web('http://haleycodes.blogspot.com/p/ephysanalysishelp.html');


% --- Executes when selected object is changed in rad_exp.
function rad_exp_SelectionChangedFcn(hObject, eventdata, handles);
       current_rad = get(eventdata.NewValue,'Tag');
    

% --- Executes on button press in rad_curve.
function rad_curve_Callback(hObject, eventdata, handles)
    curve_text = ['For IO Curve: Data should be arranged in a single column'...
        ' with Row 1 reserved for the Experiment Title. The program assumes '...
        'that Stimulus Intensity changes every 5 rows. Output'...   
        ' will be saved in a new sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',curve_text);


% --- Executes on button press in rad_ppr.
function rad_ppr_Callback(hObject, eventdata, handles)

    ppr_text = ['For Paired pulse: Data should be arranged '...
        'in sets of 3 columns: Column 1 = Interstimulus Interval, '...                                                 '...
        'Column 2 = Pulse 1 slope or amplitude, Column 3 = Pulse 2 slope '...
        'or amplitude. The header should be formatted as: Row 1, Column 1 '...
        '= Experiment Title. Output will be saved in a new sheet as'...
        '"Sheet_name Analyzed"'];
    set(handles.txt_status,'String',ppr_text);


% --- Executes on button press in rad_LTP.
function rad_LTP_Callback(hObject, eventdata, handles)
    ltp_text = ['For LTP: Data should be arranged in a single column'...
        ' with Row 1 reserved for the Experiment Title. The program assumes '...
        ' that Rows 1-60 are baseline slopes. Output will be saved in a new'... 
        ' sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',ltp_text);


% --- Executes on button press in rad_ltd.
function rad_ltd_Callback(hObject, eventdata, handles)
    ltd_text = ['For LTD: Data should be arranged in a single column with'...
        ' Row 1 reserved for the Experiment Title. The program assumes '...
        'that Rows 1-60 are baseline slopes. Output will be saved in a new'...
        ' sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',ltd_text);


% --- Executes on button press in rad_dhpg.
function rad_dhpg_Callback(hObject, eventdata, handles)
    dhpg_text = ['For DHPG LTD: Data should be arranged in a single column'...
        ' with Row 1 reserved for the Experiment Title. Enter the first sweep # of' ... 
        ' DHPG acute depression. Output will be saved in a new sheet as'...
        ' "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',dhpg_text);


% --- Executes on button press in rad_freq.
function rad_freq_Callback(hObject, eventdata, handles)
    freq_text = ['For mEPSC frequency: Data (interevent interval) should be' ...
        ' arranged in a single column with Row 1 reserved for the'...
        ' Experiment Title. Enter the length of the analyzed trace (in minutes).'... 
        ' Output will be saved in a new sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',freq_text);


% --- Executes on button press in rad_amp.
function rad_amp_Callback(hObject, eventdata, handles)
    amp_text = ['For mEPSC amplitude: Data should be arranged in a single'... 
        ' column with Row 1 reserved for the Experiment Title. Enter the' ...
        ' length of the analyzed trace (in minutes). Output will'...
        ' be saved in a new sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',amp_text);


% --- Executes on button press in rad_nmda.
function rad_nmda_Callback(hObject, eventdata, handles)
    nmda_text = ['For NMDA/AMPA: Data should be arranged in two columns:'...                           '...
        ' Column 1 = AMPA, Column 2 = NMDA. Row 1, Column 1 = Experiment Title.'...                   '...
        ' Output will be saved in a new sheet as "Sheet_Name Analyzed"'];
    set(handles.txt_status,'String',nmda_text);

% --- Executes on button press in rad_ppr2.
function rad_ppr2_Callback(hObject, eventdata, handles)

    ppr_text = ['For Raw Paired Pulse, see "ppr_raw.xls"/help/readme/github description '...
        'for layout guidelines. More flexibility in data layout will be included '...
        'in future versions. Output will be saved in a new sheet as '...
        '"Sheet_name Analyzed"'];
    set(handles.txt_status,'String',ppr_text);

    % --- Input time of DHPG wash on  
function edit_dhpg_Callback(hObject, eventdata, handles)
handles.dhpg_on = str2double (get (hObject, 'String'));
guidata(hObject,handles);

% --- Executes during object creation, after setting all properties.
function edit_dhpg_CreateFcn(hObject, eventdata, handles)
%if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
 %   set(hObject,'BackgroundColor','white');
%end

function edit_freq_CreateFcn(hObject, eventdata, handles)


function edit_freq_Callback(hObject, eventdata, handles)
handles.freq = get (hObject, 'String');
guidata(hObject,handles);

% --- Executes on button press in push_analyze.
function push_analyze_Callback(hObject, eventdata, handles)

% The selected sheet is passed to a subroutine (object) based on the 
% experiment type and triggers the appropriate analysis routine.

    selected_list=get(handles.rad_exp, 'SelectedObject');
    rad_exp = get(selected_list,'Tag');
    filename = handles.filename;
    pathname = handles.pathname;
    worksheet = handles.worksheet;
    freq_length = handles.freq;
        
    set(handles.txt_status,'String','working');
    
    switch rad_exp     
        case 'rad_curve' 
            iocurve;     
        case 'rad_ppr'     
            ppr;
        case 'rad_ppr2'
            ppr2;
        case 'rad_LTP'      
            ltp;
        case 'rad_ltd'      
            ltd;  
        case 'rad_dhpg'
            dhpg_on = handles.dhpg_on;
            wash = handles.wash;
            if dhpg_on == 0
                txt_status = ['Enter the starting Sweep of DHPG acute depression'];
            elseif dhpg_on < 31
                txt_status = ['Sweep number for onset of DHPG acute depression' ...
                    'must be > 30 to allow for a 10 minute baseline.'];
            elseif wash == 0 | wash == NaN
                txt_status = ['Enter the lenght of DHPG wash in (mins)'];
            else     
                dhpg;
            end    
        case 'rad_freq'     
            freq_length = handles.freq;
            if freq_length ~= 0 & freq_length ~= NaN
                freq; 
            else
                txt_status = ['Enter the length of the analyzed mPSC trace (in minutes)'];
            end
        case 'rad_amp' 
             amp;   
        case 'rad_nmda'     
            nmda;              
        otherwise, disp ('Nothing was selected');
    end


function edit_wash_Callback(hObject, eventdata, handles)
handles.wash =str2double (get (hObject, 'String'));
guidata(hObject,handles);


% --- Executes during object creation, after setting all properties.
function edit_wash_CreateFcn(hObject, eventdata, handles)



