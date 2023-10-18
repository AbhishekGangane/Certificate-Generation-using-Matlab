clc 
clear all 
close all 


filename = 'Registration_DetailsA.xls';
[num,txt] = xlsread(filename);
% Read Excel sheet 
% seperately from numbers
len=length(txt);
blankimage = imread('1.png');
% Read blank certificate image

text_names(:, 2) = txt(:, 2);

% Obtain names from the txt variable which are in 2nd column


text_topic(:,3)= txt(:,3);
  
% Obtain topics from the txt variable which are in 3rd column


%Ignore first row which is heading
for i=2:len
    for j=3:6
        text_all=[text_names(i,2) text_topic(i,3)];
        % combine names and topics
        %text_all=[text_names(j,2) text_topic(j,3)];
        
        position = [550 750; 400 1300];         
        % obtain positions to insert on image, MSPaint or any image editor
        %RGB = insertText(blankimage,position,text_all,'FontSize',70,'BoxOpacity',0,'TextColor','black');
        RGB = insertText(blankimage,position,text_all,'FontSize',65,'BoxOpacity',0,'TextColor','black');
        %Provide parameters for function insertText
        %Font size is 22 and opacity of box is 0 means 100% transparent
        
        figure
        imshow(RGB)        
        %Obtain and display figure in color
        
        y=i-1;
        filename=['Certificate_ID' num2str(y)];
        lastf=[filename '.png'];
        saveas(gcf,lastf);
        % generate and save files with .png extension
    end     
end