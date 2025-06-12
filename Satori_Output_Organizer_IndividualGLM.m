%% fNIRS Data Organizer for Satori Individual GLM Output
% Originating Author: Andrew Hagen
% Last Revised: 6/12/2025

% This script imports a folder that contains one or multiple raw Satori individual GLM xlsx files
% and organizes each Oxy Results and Deoxy Results sheet to have only beta weights, with
% conditions and participants labeled across each row, and channels across
% each column. Then concatenates all participants together for one
% organized .xlsx file ready for second-level analysis

clc
clear
close all

%% Prompt the user to select a folder with Excel file
folderPath = uigetdir('', 'Select folder containing individual GLM .xlsx files');
fileList = dir(fullfile(folderPath, '*.xlsx'));

if isempty(fileList)
    error('No .xlsx files found in selected folder.');
end

%% Set up containers for HbO and HbR
allData_Oxy = {};
allData_Deoxy = {};
headerWritten = false;
expectedConditions = [];

for i = 1:length(fileList)
    filename = fileList(i).name;
    fullFilepath = fullfile(folderPath, filename);

    %% Loop through each sheet
    for sheetIdx = [1 2]  % 1 = Oxy, 2 = Deoxy

        dataArray = readcell(fullFilepath, "FileType", "spreadsheet", "Sheet", sheetIdx);
            
        % Extract condition names from first row (every 4th column starting at 3)
        currentConditions = dataArray(1, 3:4:end);
        % Delete the last condition (this is just the GLM constants)
        currentConditions = currentConditions(1:(end-1));
    
         % Compare condition names to expected
        if isempty(expectedConditions)
            expectedConditions = currentConditions;
        elseif ~isequal(currentConditions, expectedConditions)
            warning('Condition mismatch in file: %s', filename);
        end
    
        numConditions = length(currentConditions);
    
        %% Delete all columns after column 3 except for every fourth column (This is the beta weight column)
        numCols = size(dataArray, 2);
        columnsToKeep = [1:3, 7:4:numCols]; % Columns 1, 2, 3, and every fourth column after 3
        dataArray = dataArray(:, columnsToKeep);
        
        %% Fill in missing condition names for each file
        colIndex = 3;
        for j = 1:numConditions
            dataArray{1, colIndex} = currentConditions{j};
            colIndex = colIndex + 1;
        end
        
        % Delete extra columns containing GLM onstants - There is group of 4 columns for each participant
        dataArray(:,colIndex:end) = [];
        
        %% Extract participant ID (everything before first underscore)
        underscoreIdx = strfind(filename, '_');
        if isempty(underscoreIdx)
            participantID = filename; % fallback
        else
            participantID = filename(1:underscoreIdx(1)-1);
        end
    
        %% Insert the participant IDs starting from column 3
        % Calculate the total number of columns needed for participant IDs
        totalColumns = 2 + numConditions;
        
        % Create a new row to insert at the top of dataArray
        participantRow = cell(1, totalColumns);
        
        % Insert the participant IDs starting from column 3
        totalColumns = 2 + numConditions;
        participantRow = cell(1, totalColumns);
        participantRow(1, 3:end) = repmat({participantID}, 1, numConditions);
        
        % Insert the participant row at the top of dataArray
        dataArray = [participantRow; dataArray];
        
        %% Transpose dataArray and label
        dataArray = dataArray';
        dataArray{1,1} = 'ParticipantID';
        dataArray{1,2} = 'Condition';
        dataArray{2,2} = [];
        dataArray(:,3) = [];
        
        % Convert all cell contents to strings so we can export
        for k = 1:numel(dataArray)
            if iscell(dataArray{k})
                dataArray{k} = cell2mat(dataArray{k});
            end
        end

       %% Append to master matrix
        if sheetIdx == 1 % Oxy Sheet
            if headerWritten == false
                allData_Oxy = dataArray;
            else
                allData_Oxy = [allData_Oxy; dataArray(3:end,:)]; 
            end
        elseif sheetIdx == 2 % Deoxy Sheet
            if headerWritten == false
                allData_Deoxy = dataArray;
            else
                allData_Deoxy = [allData_Deoxy; dataArray(3:end,:)]; 
            end
        end
    end
    headerWritten = true;  % After first file, we don't want to add the header row to the xlsx
end

%% Export to .xlsx with two sheets
    exportPath = fullfile(folderPath, 'Oxy_Deoxy_Betas_Organized.xlsx');
    writecell(allData_Oxy, exportPath, 'Sheet', 'Oxy_Betas_Organized');
    writecell(allData_Deoxy, exportPath, 'Sheet', 'Deoxy_Betas_Organized');
    disp('Oxy and Deoxy data exported successfully :)')
