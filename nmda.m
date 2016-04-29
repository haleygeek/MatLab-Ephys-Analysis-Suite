% setting up files for reading and writing to excel
pathname = handles.pathname;
filename = handles.filename;
home_dir = handles.home_dir;
worksheet = handles.worksheet;

cd (pathname)

% read in data stored in worksheet and store it in an array (variable)
[data,header] = xlsread (filename,worksheet)

% Initialize variables
rows_cols = size(data);          % Gets array size (rows,columns)
max_row = rows_cols (1,1);       % Separates into row and col variables
max_col = rows_cols (1,2);
current_col = 1;
save_row = 2;
worksheet2 = strcat(worksheet,' Analyzed');
save_data = {};

save_data (1, 1) = cellstr('Experiment');
save_data (1, 2) = cellstr('NMDA/AMPA');

while current_col <= max_col 
    
    if mod(current_col, 2) ~= 0 
        row = 1;
        i = 1;
        j = 1;
        ampa_col = 0;
        nmda_col = 0;
    
        while row <= max_row  
            % Calculate NMDA/AMPA for each column
        
            % Average -70 mV Col 1 if odd
            ampa_col = ampa_col + data (row, current_col);
        
            % Average +40 mv Col 2 if even
            nmda_col = nmda_col + data (row, current_col + 1);       
            i = i + 1;
            j = j + 1;
            
            % Move header to new cell variable
            if row == 1 
                exp_name = header (1, current_col);
                save_data (save_row, 1) = exp_name
            end
            row = row + 1;
        end
    
        % Calculate average NMDA/AMPA
        ampa_avg = abs (ampa_col / i);
        nmda_avg = abs (nmda_col / j);
        nmda_ampa = nmda_avg / ampa_avg;
        
        % Save to save_data
        save_data (save_row, 2) = num2cell (nmda_ampa)
        save_row = save_row + 1;
    end
    
    current_col = current_col + 1;
end

cd (home_dir);

% Save the worksheets
if ispc == 1
    cd (pathname);
    xlswrite (filename, save_data, worksheet2);
    
    analyze_status = strcat ('Analysis complete. ', worksheet, ' and ', worksheet2, ' have been added to: ',pathname, filename);
    
    cd (home_dir);
else
    % Importing java files for saving xlsx on Macs
    javaaddpath('/library/java/extensions/jxl.jar');
    javaaddpath('/library/java/extensions/MXL.jar');

    import mymxl.*;
    import jxl.*; 
   
    % Have to save to a different file or else get java error 
    pathname2 = strcat (pathname, 'Analyzed/');
    filename2 = strcat(pathname2, filename, ' Analyzed');
    
    if ~exist(pathname2, 'dir')
        mkdir(pathname2)
    end
    
    % Save to the new worksheet with a java workaround
    xlwrite(filename2, save_data, worksheet2);
    
    % Display save path in text_box
    analyze_status = strcat ('Analysis complete. Data saved as ', filename2);
    
end

% Update text box with path name
set(handles.txt_status,'String', analyze_status);

% Reset Folder
cd(home_dir);


