function PlotOptimizationProgress(CurrentCostFnVal, S11db, Freq)

% Create or activate figure with Tag 'OptProgressFig'
FigTag = 'OptProgressFig';
hFig = findobj('Type','figure', 'Tag',FigTag);  
if isempty(hFig), 
    figure; 
    set(gcf,'Tag',FigTag); 
else
    figure(hFig); 
end
grid on;  hold on;
set(gcf,'Position',[50 200 900 550]); 

% Get last value of the cost fun
LastCostFnVal = getappdata(gcf,'LastCostFnVal');
if isempty(LastCostFnVal)
    LastCostFnVal = CurrentCostFnVal;
    setappdata(gcf,'LastCostFnVal',LastCostFnVal);
end    

% Results for the best cost fun so far
if CurrentCostFnVal<LastCostFnVal
    Tag = 'best';
    delete(findobj(gcf, 'Tag',Tag))
    setappdata(gcf,'LastCostFnVal',CurrentCostFnVal);
    plot(Freq/1e9, S11db, 'k-', 'DisplayName',Tag, 'Tag',Tag, 'LineStyle','-', 'LineWidth',2);
end

% initial or current results
hLines = findobj(gcf, 'Tag','init');
if isempty(hLines)
    Tag = 'init';
    LineStyle = '--';  LineWidth = 1;
else
    delete(findobj(gcf, 'Tag','last'))
    Tag = 'last';
    LineStyle = '-';  LineWidth = 1;
end
plot(Freq/1e9, S11db, 'k-', 'DisplayName',Tag, 'Tag',Tag, 'LineStyle',LineStyle, 'LineWidth',LineWidth);

legend show location west
xlabel('Frequency (GHz)');
ylabel('|S11|(dB)');

ylim([-35 0])



