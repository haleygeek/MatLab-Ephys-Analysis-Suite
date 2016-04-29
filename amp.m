% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Calculates the average and cumualitve amplitude and frequency of 
% spontaneous events from raw event detection data (i.e from Clampfit)

% set text box status to 'Working'
analyze_status = 'Working';

%Read in the directory, filename, and worksheet selected from the GUI
pathname = handles.pathname;
filename = handles.filename;
worksheet = handles.worksheet;
home_dir = handles.home_dir;
summary_data = 1;
save_cumulative = {};

% Set bin edges (right side of bin is included in analysis. i.e. 25 is in
% bin #1) These are set to mPSCs, increase the edge size if looking at APs
% or in the absence of TTX
edges = [0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 ...
    26 27 28 29 30 31 32 33 34 35 36 37 38 39 40];
trans_edges = num2cell(transpose (edges));

%Change to the directory where the xlsx file is located
cd (pathname);

% Read the data from the excel worksheet into a variable 
[data,header] = xlsread (filename,worksheet)

% Initialize variables
rows_cols = size(data);          % Gets array size (rows,columns)
max_row = rows_cols (1,1);       % Separates into row and col variables
max_col = rows_cols (1,2);
current_col = 1;
current_row = 1;

%worksheet2 = strcat(worksheet,' Analyzed');
worksheet2 = strcat(worksheet, ' Summary');
worksheet3 = strcat(worksheet, ' Cumulative');

%save_data = {}; Shouldn't need Save_Data
save_summary = {};
save_col = 0;

% Gets header data
while current_col < max_col + 1
    
     % Save first column with exp_name
     % Saves inter-event data in the same format as raw data
     % Transposes experiment names to rows in Column 1 for summary
     % Column 2 is the average amplitude
     % %%%Column 3 WAS the average interstimulus interval
     
     exp_name = header (1, current_col);
        
     if current_col == 1
         save_cumulative (1,1) = cellstr ('Bin');
         save_cumulative (1,2) = exp_name;
         
         x = 1;
         for x = 1:41;
            save_cumulative (x + 1, 1)= trans_edges (x, 1);
            x = x + 1;
         end
         
         save_summary (1, 1) = cellstr ('Experiment');
         save_summary (1, 2) = cellstr('Average Amplitude (pA)');
         save_summary (current_col + 1, 1) = exp_name;
         
    
     else 
         save_summary (current_col + 1, 1) = exp_name;
         save_cumulative (1, current_col + 1) = exp_name;
    
     end
     current_col = current_col + 1;
end
    
% Reset row and column to 1 for I/O curve analysis    
current_col = 1;
event_col = 1;
event_row = 2;


% Assumes every column contains the "start" time of each mPSC event
while current_col < max_col + 1 
    current_row = 1;
    current_mPSC = 0;
    sum = 0;
    avg_amp = 0;
    events = 1;
    
   
    % Calculates the average amp and cumulative prob for all events
   
    while current_row < max_row  
        
        if ~isnan(data(current_row + 1, current_col));
            current_mPSC = data (current_row, current_col);
            
            % Save variable for histogram data
            save_amp (current_row, current_col) = current_mPSC
            
            sum = sum + current_mPSC;
            events = events + 1
        end
        % Histogram function
        get_hist = histcounts(abs (save_amp), edges, 'Normalization','cdf')
   
        x = 1;
        hist_row = 2;
        for x = 1:40;
            save_cumulative (hist_row, current_col + 1) = num2cell (get_hist (1, hist_row - 1));
            hist_row = hist_row + 1;
            x = x + 1;
        end
        
        current_row = current_row + 1;
    end
    
    avg_amp = sum/events
    save_summary (event_row, 2) = num2cell (avg_amp)
    event_row = event_row + 1
       
    current_col = current_col + 1
   
end

% Save the file (does not use regular excelsave.m)
cd (home_dir);

if ispc == 1
    cd (pathname)
    xlswrite (filename, save_summary, worksheet2);
    xlswrite (filename, save_cumulative, worksheet3);
    
    analyze_status = strcat ('Analysis complete. ', worksheet, ' and ', worksheet2, ' have been added to: ',pathname, filename);
    cd (home_dir);
else
    % Importing java files for saving xlsx on Macs (Not supported beyond 
    % OS Yosimite
    javaaddpath('/library/java/extensions/jxl.jar');
    javaaddpath('/library/java/extensions/MXL.jar');

    import mymxl.*;
    import jxl.*; 
   
    % Have to save to a different file or else get java error 
    pathname2 = strcat (pathname, 'Analyzed/')
    filename2 = strcat(pathname2, filename, ' Analyzed')
    
    if ~exist(pathname2, 'dir')
        mkdir(pathname2)
    end
    
    % Save to the new worksheet with a java workaround
    xlwrite (filename2, save_summary, worksheet2);
    xlwrite (filename2, save_cumulative, worksheet3);
   
    % Display save path in text_box
    analyze_status = strcat ('Analysis complete. Data saved as ', filename2)
    
end

% Update text box with path name
set(handles.txt_status,'String', analyze_status)

% Reset Folder
cd(home_dir)