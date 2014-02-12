XLSA
====

### Sensitivity analyis for models implemented in Excel


Simple Python script that uses [Python for Windows extensions](http://sourceforge.net/projects/pywin32/) to repeately run spreadsheet models in Excel to perform univariate and probabilistic sensitivity analysis.

#### Mini How-To

Implement your model in Excel, and add a worksheet named "param_distributions". Populate it with  any parameters that should be subjet to the sensitivity analysis. One row per parameter, with the following columns:
* param_name: Name to be assigned to the parameter in the output files
* worksheet: Worksheet of model which contains parameter
* cells_row_col: 1-based row and col index of paramter.
* min, max, mode: min, max, and central value used for univariate sensitivity analysis
* distribution: sampling distribution for probabilistic sensitivity analysis

Add a worksheet named "predictions". List all outcomes of interest, one row per outcome, with the following columns:
* outcome: Name to be assigned to the outcome in the output files
* worksheet: Worksheet of the model which contains outcome
* cell_row_col: 1-based row and col index of outcome

Launch src/XLSA.py (you may need to install [Python for Windows extensions](http://sourceforge.net/projects/pywin32/) first). Results will be written to the output directory. src/analysis is an R-script which generates Tornado plots from the predictions. models/SampleModel.xls provides an example for how to use XLSA. 
