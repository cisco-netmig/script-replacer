import os
import re
import traceback
import logging
from tempfile import NamedTemporaryFile

from PyQt5 import QtWidgets, QtCore


class FillValuesWorker(QtCore.QThread):
    """
    Worker thread to extract variables from a configuration template file and generate an Excel form.
    The Excel form is pre-filled with variable names to allow the user to provide corresponding values.
    """

    fill_complete = QtCore.pyqtSignal(dict)

    def __init__(self, template_file):
        """
        Initialize the worker with the provided template file path.

        Args:
            template_file (str): Path to the configuration template file.
        """
        super().__init__()
        self.template_file = template_file
        self.config_template = ''
        self.var_list = []

    def run(self):
        """
        Run the worker thread logic. Parses the template, extracts variables,
        creates an Excel form, and emits a signal upon completion.
        """
        if not self.template_file:
            logging.info('Template file cannot be empty')
            return

        logging.info('Building variable form')

        self.config_template, self.var_list = self.get_template_vars(self.template_file)
        variable_file_path = self.create_excel_form(self.var_list)

        logging.info('Opening variable form')
        logging.info('Fill up the values against variables and save')
        logging.info('Once values are filled click "Replace"')

        self.fill_complete.emit({
            'config_template': self.config_template,
            'var_list': self.var_list,
            'variable_file_path': variable_file_path
        })

    def get_template_vars(self, filepath, var_regex=r'\$\w+\$'):
        """
        Extract unique variable names from the configuration template.

        Args:
            filepath (str): Path to the template file.
            var_regex (str): Regular expression pattern to match variable tokens.

        Returns:
            tuple: The raw content of the file and a list of unique variables.
        """
        content = open(filepath).read()
        var_list = []
        for var in re.findall(var_regex, content):
            if var not in var_list:
                var_list.append(var)
        return content, var_list

    def create_excel_form(self, var_list):
        """
        Create an Excel file with a list of variable names for user input.

        Args:
            var_list (list): List of variable strings to include in the form.

        Returns:
            str: File path to the generated Excel file.
        """
        from netcore import XLBW
        tmp = NamedTemporaryFile(suffix='.xlsx', mode='w+', delete=False)
        tmp.close()
        filepath = tmp.name
        workbook = XLBW(filepath)
        worksheet = workbook.add_worksheet('Variables')

        for row_idx in range(500):
            for col_idx in range(50):
                worksheet.write(row_idx, col_idx, '', workbook.ftbody)

        for idx, var in enumerate(var_list):
            worksheet.write(idx, 0, var, workbook.ftbody)

        workbook.close()
        return filepath


class BuildOutputWorker(QtCore.QThread):
    """
    Worker thread to read variable values from an Excel form, replace them in the template,
    and write the result to an output Excel file.
    """

    def __init__(self, template, variable_file_path, output_report):
        """
        Initialize the worker with the configuration template, variable form, and output path.

        Args:
            template (str): The raw configuration template string.
            variable_file_path (str): Path to the filled Excel form with variable values.
            output_report (str): Path to save the output Excel file.
        """
        super().__init__()
        self.template = template
        self.variable_file_path = variable_file_path
        self.output_report = output_report

    def run(self):
        """
        Run the worker thread logic. Replaces variables in the template using values from
        the Excel form and writes the result to an output Excel file.
        """
        if not self.variable_file_path:
            logging.info('Variable form cannot be empty')
            return

        from netcore import XLBW, XLR

        logging.info('Building output')

        workbook_form = XLR(self.variable_file_path).book
        worksheet_form = workbook_form.sheet_by_index(0)

        workbook = XLBW(self.output_report)
        worksheet = workbook.add_worksheet('Output')
        ftb, fthl, fte = workbook.ftbody, workbook.fthighlight, workbook.fterror

        for col_idx in range(1, worksheet_form.ncols):
            cfg = []
            repl_dict = {
                worksheet_form.cell_value(row, 0): worksheet_form.cell_value(row, col_idx)
                for row in range(worksheet_form.nrows)
            }
            cfg.extend(self.sub_get_string(self.template, repl_dict, ftb, fthl, fte))
            self.write_rich_table(worksheet, cfg, ftb, fthl, fte, 0, col_idx - 1, 95)

            if hasattr(logging, 'savings'):
                logging.savings(10)

        workbook.close()

        try:
            os.remove(self.variable_file_path)
            logging.info('Variable form removed successfully')
        except Exception:
            logging.error(traceback.format_exc())

        logging.info('Variables Replaced!')

    def write_rich_table(self, ws, cfg_lines, ftb, fthl, fte, row_index, col_index, col_width):
        """
        Write formatted configuration lines to an Excel worksheet.

        Args:
            ws: The worksheet object.
            cfg_lines (list): List of rich text formatted lines.
            ftb, fthl, fte: Format styles for normal text, highlighted variables, and errors.
            row_index (int): Starting row index.
            col_index (int): Column index to write into.
            col_width (int): Width of the output column.
        """
        ws.set_column(col_index, col_index, col_width)
        for cfg_line in cfg_lines:
            if ftb in cfg_line and (fthl in cfg_line or fte in cfg_line):
                ws.write_rich_string(row_index, col_index, *cfg_line)
            elif fthl in cfg_line:
                ws.write_string(row_index, col_index, cfg_line[-1].strip(), fthl)
            elif fte in cfg_line:
                ws.write_string(row_index, col_index, cfg_line[-1].strip(), fte)
            elif ftb in cfg_line:
                ws.write_string(row_index, col_index, cfg_line[-1], ftb)
            row_index += 1

    def sub_get_string(self, cfg, repl, ftb, fthl, fte, var_regex=r'\$\w+\$'):
        """
        Replace variables in the configuration string using the replacement dictionary.

        Args:
            cfg (str): Configuration template string.
            repl (dict): Dictionary of variable replacements.
            ftb, fthl, fte: Format styles for text rendering.
            var_regex (str): Regular expression pattern for variables.

        Returns:
            list: List of rich text formatted configuration lines.
        """
        cfg_format = []
        for line in cfg.splitlines():
            non_reg = re.split(var_regex, line)
            reg = re.findall(var_regex, line)
            format_line = []
            for i in range(len(reg)):
                if not re.search(r'^\s*$', non_reg[i]):
                    format_line.extend([ftb, non_reg[i]])
                value = repl.get(reg[i], '')
                if not re.search(r'^\s*$', value):
                    format_line.extend([fthl, value])
                else:
                    format_line.extend([fte, reg[i]])
            if not re.search(r'^\s*$', non_reg[-1]):
                format_line.extend([ftb, non_reg[-1]])
            cfg_format.append(format_line)
        return cfg_format
