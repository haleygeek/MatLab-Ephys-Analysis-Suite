% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Analyzes raw event detection data from Clampfit or something similar and
% calculates the average and cumulative probability of event amplitude and
% frequency.

% set text box status to 'Working'
analyze_status = 'Working'

%Read in the directory, filename, and worksheet selected from the GUI
pathname = handles.pathname;
filename = handles.filename;
worksheet = handles.worksheet;
home_dir = handles.home_dir;
freq_length = handles.freq
summary_data = 1;
save_cumulative = {};

% Set bin edges (right side of bin is included in analysis. i.e. 25 is in
% bin #1)
edges = [0 25 50 75 100 125 150 175 200 225 250 275 300 325 350 375 400 ...
    425 450 475 500 525 550 575 600 625 650 675 700 725 750 775 800 825 ...
    850 875 900 925 950 975 1000 1025 1050 1075 1100 1125 1150 1175 1200 ...
    1225 1250 1275 1300 1325 1350 1375 1400 1425 1450 1475 1500 1525 1550 ...
    1575 1600 1625 1650 1675 1700 1725 1750 1775 1800 1825 1850 1875 1900 ...
    1925 1950 1975 2000 2025 2050 2075 2100 2125 2150 2175 2200 2225 2250 ...
    2275 2300 2325 2350 2375 2400 2425 2450 2475 2500 2525 2550 2575 2600 ...
    2625 2650 2675 2700 2725 2750 2775 2800 2825 2850 2875 2900 2925 2950 ...
    2975 3000];
trans_edges = num2cell(transpose (edges))

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

worksheet2 = strcat(worksheet,' Analyzed');
worksheet3 = strcat(worksheet, ' Summary');
worksheet4 = strcat(worksheet, ' Cumulative');

save_data = {};
save_summary = {};
save_col = 0;

% Gets header data
while current_col < max_col + 1
    
     % Save first column with exp_name
     % Saves inter-event data in the same format as raw data
     % Transposes experiment names to rows in Column 1 for summary
     % Column 2 is the average Frequency
     % Column 3 is the average interevent interval
     
     exp_name = header (1, current_col);
        
     if current_col == 1
         save_data(1, current_col) = exp_name;
         save_cumulative (1,1) = cellstr ('Bin');
         save_cumulative (1,2) = exp_name;
         
         x = 1;
         for x = 1:120;
            save_cumulative (x + 1, 1)= trans_edges (x, 1);
            x = x + 1;
         end
         
         save_summary (1, 1) = cellstr ('Experiment');
         save_summary (1, 2) = cellstr('Frequency (Hz)');
         save_summary (1, 3) = cellstr('Average Interval (ms)');
         save_summary (current_col + 1, 1) = exp_name;
         
    
     else
         save_data(end + 1) = exp_name ;  
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
    next_mPSC = 0;
    interval = 0;
    save_row = 2;
    sum = 0;
    avg_interval = 0;
    events = 1;
    hz = 0;
   
    % Calculates the inter-event interval for all events
   
    while current_row < max_row  
        
        if ~isnan(data(current_row + 1, current_col)) & ~isnan(interval);
            current_mPSC = data (current_row, current_col);
            next_mPSC = data (current_row + 1, current_col);
            interval = next_mPSC - current_mPSC;
            
            % Save variable for histogram data
            save_interval (save_row, current_col) = interval;
            
            % Save variabe for inter-event data
            save_data (save_row, current_col) = num2cell (interval);
            
           
            sum = sum + interval;
            save_row = save_row + 1;
            events = events + 1
        end
    % Histogram function
    get_hist = histcounts(save_interval, edges, 'Normalization','cdf');
   
    x = 1;
    hist_row = 2;
    for x = 1:120;
        save_cumulative (hist_row, current_col + 1) = num2cell (get_hist (1, hist_row - 1))
        hist_row = hist_row + 1;
        x = x + 1;
    end    
    current_row = current_row + 1;
    
    end
    
    hz = events / freq_length
    avg_interval = sum/events
    save_summary (event_row, 2) = num2cell (hz)
    save_summary (event_row, 3) = num2cell (avg_interval)
    event_row = event_row + 1
       
    current_col = current_col + 1
   
end

% Save the file (does not use regular excelsave.m)
cd (home_dir);

if ispc == 1
    cd (pathname)
    xlswrite (filename, save_data, worksheet2);
    xlswrite (filename, save_summary, worksheet3);
    xlswrite (filename, save_cumulative, worksheet4);
    
    analyze_status = strcat ('Analysis complete. ', worksheet, ' and ', worksheet2, ' have been added to: ',pathname, filename);
    cd (home_dir);
else
    % Importing java files for saving xlsx on Macs
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
    xlwrite(filename2, save_data, worksheet2);
    xlwrite (filename2, save_summary, worksheet3);
    xlwrite (filename2, save_cumulative, worksheet4);
   
    % Display save path in text_box
    analyze_status = strcat ('Analysis complete. Data saved as ', filename2)
    
end

% Update text box with path name
set(handles.txt_status,'String', analyze_status)

% Reset Folder
cd(home_dir)






