import win32com.client as win32
import csv
import random
import os

n_samples = 500
SA_ids = ['SampleModel']
rootDir = os.path.normpath(os.path.join(os.getcwd(), '..'))


#set all values to mode
def reset(parameters, spreadsheet):
    for param in parameters:
        param_sheet = spreadsheet.Sheets(param['sheet'])
        set_param_values(param, param_sheet, param['mode'])


#one-way sensitivity analysis
def set_param_values(param, param_sheet, value):
    for cell in param['cells']:
        row = int(cell.split(',')[0])
        col = int(cell.split(',')[1])
        param_sheet.Cells(row, col).Value = value


def univariate(parameters, predictions, ss):
    with open(os.path.join(rootDir, 'output', 'univariate_' + SA_id + '_wide.txt'), 'wb') as outfile_wide, \
        open(os.path.join(rootDir, 'output', 'univariate_' + SA_id + '_long.txt'), 'wb') as outfile_long:
        wide_writer = csv.writer(outfile_wide, delimiter='\t')
        wide_writer.writerow(['parname', 'outcome', 'ParLow', 'ParMode', 'ParHigh', 'PredLow', 'PredMode', 'PredHigh'])
        long_writer = csv.writer(outfile_long, delimiter='\t')
        long_writer.writerow(['parname', 'assumption', 'parvalue', 'outcome', 'value'])
        for param in parameters:
            if not (param['mini'] and param['maxi'] and param['mode']):
                continue
            minimum = float(param['mini'])
            maximum = float(param['maxi'])
            mode = float(param['mode'])
            param_sheet = ss.Sheets(param['sheet'])
            for prediction in predictions:
                prediction_sheet = ss.Sheets(prediction['sheet'])
                set_param_values(param, param_sheet, minimum)
                low_value = prediction_sheet.Cells(prediction['row'], prediction['col']).Value
                set_param_values(param, param_sheet, maximum)
                high_value = prediction_sheet.Cells(prediction['row'], prediction['col']).Value
                set_param_values(param, param_sheet, mode)
                mode_value = prediction_sheet.Cells(prediction['row'], prediction['col']).Value
                long_writer.writerow([param['pname'], 'low', minimum, prediction['predname'], low_value])
                long_writer.writerow([param['pname'], 'high', maximum, prediction['predname'], high_value])
                long_writer.writerow([param['pname'], 'mode', mode, prediction['predname'], mode_value])
                wide_writer.writerow([param['pname'], prediction['predname'], param['mini'], param['mode'],
                                      param['maxi'], low_value, mode_value, high_value])


#probabilistic sensitivity analysis
def psa(parameters, parnames, predictions, prednames, ss):
    with open(os.path.join(rootDir, 'output', 'PSA_' + SA_id + '_wide.txt'), 'wb') as outfile_wide,\
        open(os.path.join(rootDir, 'output', 'PSA_' + SA_id + '_long.txt'), 'wb') as outfile_long:
        wide_writer = csv.writer(outfile_wide, delimiter='\t')
        wide_writer.writerow(parnames + prednames)
        long_writer = csv.writer(outfile_long, delimiter='\t')
        long_writer.writerow(['parname', 'parvalue', 'outcome', 'value'])
        for sample in range(n_samples):
            parvalues = []
            for param in parameters:
                if not param['distribution']:
                    continue
                distribution = param['distribution'].split(':')
                if distribution[0] == 'triangular':
                    value = random.triangular(float(distribution[1]), float(distribution[2]), float(distribution[3]))
                elif distribution[0] == 'uniform':
                    value = random.uniform(float(distribution[1]), float(distribution[2]))
                elif distribution[0] == 'integer':
                    value = random.randint(float(distribution[1]), float(distribution[2]))
                elif distribution[0] == 'beta':
                    value = random.betavariate(float(distribution[1]), float(distribution[2]))
                else:
                    # this is a constant, just use mode
                    value = float(param['mode'])
                parvalues.append(value)
                param_sheet = ss.Sheets(param['sheet'])
                set_param_values(param, param_sheet, value)
            predvalues = []
            for pred, predname in zip(predictions, prednames):
                shOut = ss.Sheets(pred['sheet'])
                value = shOut.Cells(pred['row'], pred['col']).Value
                predvalues.append(value)
                for paramvalue, parname in zip(parvalues, parnames):
                    long_writer.writerow([parname, paramvalue, predname, value])
            wide_writer.writerow(parvalues + predvalues)


if __name__ == "__main__":
    for SA_id in SA_ids:
        model_file = os.path.join(rootDir, 'models', SA_id + '.xls')
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        xl.Visible = False
        spreadsheet = xl.Workbooks.Open(model_file)
        param_sheet = spreadsheet.Sheets('param_distributions')
        parameters = []
        param_names = []
        row = 2
        while True:
            param_name = param_sheet.Cells(row, 1).Value
            if not param_name:
                break
            sheet = param_sheet.Cells(row, 2).Value
            inputs = param_sheet.Cells(row, 3).Value.split(':')
            cells = []
            for cell in inputs:
                cells.append(cell)
            mini = param_sheet.Cells(row, 4).Value
            maxi = param_sheet.Cells(row, 5).Value
            mode = param_sheet.Cells(row, 6).Value
            distribution = param_sheet.Cells(row, 7).Value
            parameters.append({'pname': param_name, 'sheet': sheet, 'cells': cells, 'distribution': distribution,
                               'mini': mini, 'maxi': maxi, 'mode': mode})
            param_names.append(param_name)
            row += 1
        pred_sheet = spreadsheet.Sheets('predictions')
        predictions = []
        prediction_names = []
        row = 2
        while True:
            predname = pred_sheet.Cells(row, 1).Value
            if not predname:
                break
            sheet = pred_sheet.Cells(row, 2).Value
            row_col = pred_sheet.Cells(row, 3).Value.split(',')
            row = int(row_col[0])
            col = int(row_col[1])
            predictions.append({'predname': predname, 'sheet': sheet, 'row': row, 'col': col})
            prediction_names.append(predname)
            row += 1
        reset(parameters, spreadsheet)
        univariate(parameters, predictions, spreadsheet)
        spreadsheet.Close(False)
        spreadsheet = xl.Workbooks.Open(model_file)
        reset(parameters, spreadsheet)
        psa(parameters, param_names, predictions, prediction_names, spreadsheet)
        reset(parameters, spreadsheet)
        spreadsheet.Close(True)
        xl.Application.Quit()
        print('Done')
