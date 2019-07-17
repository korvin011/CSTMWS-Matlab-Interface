function Res = CalcCostFn(CSTProject)
% CST calls this function passing the full project file name without
% extension as the argument.

% Create the CST interface object and connect to the CST project
CST = TCSTInterface();
CST.OpenProject([CSTProject '.cst']);

% Get S-params (S11 for the example project, since we have only 1 port) for
% the current run (RunID=0) and Freq=1 GHz
S = CST.GetSParams(0,[],[],[],1); 

% The cost function here is simply |S11| in [dB]
Res = 20*log10(abs(S));

% -------------- It is recommeded to keep the code below ------------------
% Res must be a real scalar finite numeric value. The lines below may help to debug
% potential errors in this function.
assert(isscalar(Res),  'Error in Matlab cost function: Res must be scalar.' );
assert(isreal(Res),    'Error in Matlab cost function: Res must be real.'   );
assert(isnumeric(Res), 'Error in Matlab cost function: Res must be numeric.');
assert(~isinf(Res),    'Error in Matlab cost function: Res must be finite.' );
assert(~isnan(Res),    'Error in Matlab cost function: Res cannot be NaN.'  );
Res = double(Res); % make sure Res is of type double
% -------------------------------------------------------------------------

