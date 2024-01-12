classdef app1_exported < matlab.apps.AppBase

    % Properties that correspond to app components
    properties (Access = public)
        UIFigure                 matlab.ui.Figure
        ChooseCertificateButton  matlab.ui.control.Button
        ChooseDataButton         matlab.ui.control.Button
        Image                    matlab.ui.control.Image
        DataTextAreaLabel        matlab.ui.control.Label
        DataTextArea             matlab.ui.control.TextArea
        CreateButton             matlab.ui.control.Button
        XAxisEditFieldLabel      matlab.ui.control.Label
        XAxisEditField           matlab.ui.control.NumericEditField
        YAxisEditFieldLabel      matlab.ui.control.Label
        YAxisEditField           matlab.ui.control.NumericEditField
    end

    
    properties (Access = private)
        text_names cell
        len
        imgData% Description
    end
    

    % Callbacks that handle component events
    methods (Access = private)

        % Button pushed function: ChooseCertificateButton
        function ChooseCertificateButtonPushed(app, event)
            [file, path] = uigetfile({'*.jpg;*.png;*.jpeg;*.tif', 'Image Files (*.jpg, *.png, *.jpeg, *.tif)'}, 'Select an Image File');
        
        if isequal(file, 0) % Check if the user canceled file selection
            disp('User canceled file selection');
        else
            imagePath = fullfile(path, file);
            
            % Display the selected image in the Image component
            if exist(imagePath, 'file')
                app.imgData = imread(imagePath);
                app.Image.ImageSource = app.imgData;
            else
                disp('File does not exist or is not an image');
            end
        end
        end

        % Button pushed function: ChooseDataButton
        function ChooseDataButtonPushed(app, event)
             [file, path] = uigetfile({'*.csv; *.xlsx; *.xls', 'CSV Files (*.csv; *.xlsx; *.xls)'}, 'Select a CSV File');
             filepath = fullfile(path, file);
             [num,txt] = xlsread(filepath);
             
             % Read Excel sheet containing details. Text is read from the file
% seperately from numbers


            app.len=length(txt);
% obtain length of text in excel or number of certificates to be generated
% This code provides scalability

             app.text_names = cell(app.len, 1); % Preallocate to store names

            % Obtain names from the txt variable which are in the 2nd column
            for i = 1:app.len
                app.text_names{i} = char(txt{i, 2});
            end

            % Display the text_names in the DataTextArea
            set(app.DataTextArea, 'Value', app.text_names);

         
     
            
        end

        % Button pushed function: CreateButton
        function CreateButtonPushed(app, event)
            xValue = app.XAxisEditField.Value;
            yValue = app.YAxisEditField.Value;
            
            for i=2:app.len
%                text_all=[app.text_names(i,2)];
                text_all = app.text_names{i}; % Access the ith element of app.text_names
                % combine names and topics
                
                position = [xValue yValue];         
                % obtain positions to insert on image, MSPaint or any image editor
                
                RGB = insertText(app.imgData,position,text_all,'FontSize',50,'BoxOpacity',0);
                %Provide parameters for function insertText
                %Font size is 22 and opacity of box is 0 means 100% transparent
                
                figure;
                imshow(RGB);        
                %Obtain and display figure in color
                
                y=i-1;
                folderPath = 'E:\Study\University\IT509 DSP\Matlab files\Final project\AUTOMATIC CERTIFICATE GENERATION USING MATLAB\Generated Certificates'; % Replace with your desired folder path
                filename = fullfile(folderPath, ['Certificate_Topic_' num2str(y) '.tif']);
                saveas(gcf, filename);
                % generate and save files with .tif extension
               



            end   
            
        end
    end

    % Component initialization
    methods (Access = private)

        % Create UIFigure and components
        function createComponents(app)

            % Create UIFigure and hide until all components are created
            app.UIFigure = uifigure('Visible', 'off');
            app.UIFigure.Color = [0.651 0.898 0.9922];
            app.UIFigure.Position = [100 100 814 606];
            app.UIFigure.Name = 'UI Figure';

            % Create ChooseCertificateButton
            app.ChooseCertificateButton = uibutton(app.UIFigure, 'push');
            app.ChooseCertificateButton.ButtonPushedFcn = createCallbackFcn(app, @ChooseCertificateButtonPushed, true);
            app.ChooseCertificateButton.BackgroundColor = [0.1294 0.5882 0.9529];
            app.ChooseCertificateButton.FontColor = [0.2235 0.2745 0.2824];
            app.ChooseCertificateButton.Position = [139 204 147 35];
            app.ChooseCertificateButton.Text = 'Choose Certificate';

            % Create ChooseDataButton
            app.ChooseDataButton = uibutton(app.UIFigure, 'push');
            app.ChooseDataButton.ButtonPushedFcn = createCallbackFcn(app, @ChooseDataButtonPushed, true);
            app.ChooseDataButton.BackgroundColor = [0.1294 0.5882 0.9529];
            app.ChooseDataButton.FontColor = [0.2235 0.2745 0.2824];
            app.ChooseDataButton.Position = [587 204 143 35];
            app.ChooseDataButton.Text = 'Choose Data';

            % Create Image
            app.Image = uiimage(app.UIFigure);
            app.Image.Position = [35 259 356 329];

            % Create DataTextAreaLabel
            app.DataTextAreaLabel = uilabel(app.UIFigure);
            app.DataTextAreaLabel.HorizontalAlignment = 'right';
            app.DataTextAreaLabel.FontColor = [0.2235 0.2745 0.2824];
            app.DataTextAreaLabel.Position = [493 564 31 22];
            app.DataTextAreaLabel.Text = 'Data';

            % Create DataTextArea
            app.DataTextArea = uitextarea(app.UIFigure);
            app.DataTextArea.FontColor = [0.2235 0.2745 0.2824];
            app.DataTextArea.BackgroundColor = [0.851 0.9882 0.8314];
            app.DataTextArea.Position = [539 250 254 338];

            % Create CreateButton
            app.CreateButton = uibutton(app.UIFigure, 'push');
            app.CreateButton.ButtonPushedFcn = createCallbackFcn(app, @CreateButtonPushed, true);
            app.CreateButton.BackgroundColor = [0.1294 0.5882 0.9529];
            app.CreateButton.FontColor = [0.2235 0.2745 0.2824];
            app.CreateButton.Position = [563 70 192 54];
            app.CreateButton.Text = 'Create';

            % Create XAxisEditFieldLabel
            app.XAxisEditFieldLabel = uilabel(app.UIFigure);
            app.XAxisEditFieldLabel.HorizontalAlignment = 'right';
            app.XAxisEditFieldLabel.FontColor = [0.2235 0.2745 0.2824];
            app.XAxisEditFieldLabel.Position = [64 123 40 22];
            app.XAxisEditFieldLabel.Text = 'X-Axis';

            % Create XAxisEditField
            app.XAxisEditField = uieditfield(app.UIFigure, 'numeric');
            app.XAxisEditField.FontColor = [0.2235 0.2745 0.2824];
            app.XAxisEditField.BackgroundColor = [0.851 0.9882 0.8314];
            app.XAxisEditField.Position = [119 123 134 22];

            % Create YAxisEditFieldLabel
            app.YAxisEditFieldLabel = uilabel(app.UIFigure);
            app.YAxisEditFieldLabel.HorizontalAlignment = 'right';
            app.YAxisEditFieldLabel.FontColor = [0.2235 0.2745 0.2824];
            app.YAxisEditFieldLabel.Position = [64 86 39 22];
            app.YAxisEditFieldLabel.Text = 'Y-Axis';

            % Create YAxisEditField
            app.YAxisEditField = uieditfield(app.UIFigure, 'numeric');
            app.YAxisEditField.BackgroundColor = [0.851 0.9882 0.8314];
            app.YAxisEditField.Position = [118 86 135 22];

            % Show the figure after all components are created
            app.UIFigure.Visible = 'on';
        end
    end

    % App creation and deletion
    methods (Access = public)

        % Construct app
        function app = app1_exported

            % Create UIFigure and components
            createComponents(app)

            % Register the app with App Designer
            registerApp(app, app.UIFigure)

            if nargout == 0
                clear app
            end
        end

        % Code that executes before app deletion
        function delete(app)

            % Delete UIFigure when app is deleted
            delete(app.UIFigure)
        end
    end
end