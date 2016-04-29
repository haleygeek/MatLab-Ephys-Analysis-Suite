% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Requires xlswrite.m  for windows. Get it here: 
% http://www.mathworks.com/matlabcentral/fileexchange/27236-improved-xlswrite-m
% Make sure it is in your home directory
    
% Determine operating system
os_pc = ispc;

% Switch to working folder
cd (pathname);

% If Windows, use xlswrite
if os_pc == 1
    
    %save the header first
    xlswrite(filename,save_header,worksheet2)

    % There should only be one row in the sheet for the header
    start_xlrange = 'A2';      

    % Need to get the number of data columns (if not passed through handle)
    save_size = size (save_data);
    max_row = save_size (1,1) + 1;
    max_col = save_size (1,2); 
    column_number = max_col;
    
    % Make sure you navigate back to the home directory 
    cd (home_dir)
    excelcols 
    
    % Takes you back to the directory where your data is stored
    cd (pathname)
    end_xlrange = strcat (column_letter, num2str(max_row));
    xlRange = strcat(start_xlrange,':',end_xlrange);
    xlswrite(filename,save_data,worksheet2,xlRange)
    
    if summary_data == 1
        % Saves header to summary data in column 1
        worksheet3 = strcat (worksheet, ' Summary');
        xlswrite(filename, save_header2, worksheet3);
        
        % Get dimensions of summary data
        save_size = size (save_summary);
        max_row = save_size (1,1) + 1;
        
        % Save summary data 
        xlRange = strcat('B1:B',num2str(max_row));
        xlswrite(filename, save_summary,worksheet3, xlRange);
        
    end
    
    % Display save path in text_box
    status = strcat ('Analysis complete. Data saved as ', pathname, filename);
    
% If Mac, use xlwrite (java wrap and workaround)
else
    % Importing java files for saving xlsx on Macs
    % Requires jxl.jar, MXl.jar be installed in your PATH
    javaaddpath('/library/java/extensions/jxl.jar');
    javaaddpath('/library/java/extensions/MXL.jar');

    import mymxl.*;
    import jxl.*; 
       
    % Have to save to a different file or else get java error 
    % This is the main difference between the PC and Mac versions
    pathname2 = strcat (pathname, 'Analyzed/');
    filename2 = strcat(pathname2, filename, ' Analyzed');
    
    if ~exist(pathname2, 'dir')
        mkdir(pathname2);
    end
    
    % Need to define a different range for macs because we're combining the
    % save_header and save_data arrays
    save_size = size (save_data);
    mac_max_row = save_size (1,1) + 1;
    mac_max_col = save_size (1,2);
    
    % Setting up current row and column, as well as a new empty array
    % to combine save_header and save_data into a single array
    row = 1;
    col = 1;
    save_mac = {};
    
    % Moving header and data into a single variable
    while col < mac_max_col + 1
        while row < mac_max_row + 1 %%%%%%need to expand to 3 rows
            if row == 1
                save_mac (row, col) = save_header (row, col);
            else
                save_mac (row, col) = num2cell (save_data (row-1, col));
            end
            row = row + 1;
        end
        row = 1;
        col = col + 1;
        
    end 
    cd (home_dir);
    
    % Save to the new worksheet with a java workaround
    xlwrite(filename2, save_mac, worksheet2);
    
    % Save summary data
    row = 1;
    if summary_data == 1
        worksheet3 = strcat (worksheet, ' Summary');
        
        % Get dimensions of summary data
        save_size = size (save_summary);
        mac_max_row = save_size (1,1) + 1;
        
        % Moving header and data into a single variable
        while row < mac_max_row 
            save_mac_summary (row, 1) = save_header2 (row, 1);
            save_mac_summary (row, 2) = num2cell (save_data (row, 1));
            row = row + 1;
        end        
        
        xlwrite (filename2, save_mac_summary, worksheet3);
    end
    
    % Display save path in text_box
    status = strcat ('Analysis complete. Data saved as ', filename2);
    
end

% Update text box with path name
set(handles.txt_status,'String',status);

% Reset Folder
cd (home_dir);