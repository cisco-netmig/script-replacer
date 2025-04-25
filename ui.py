import logging
import os
from datetime import datetime
from PyQt5 import QtWidgets, QtGui, QtCore
from .workers import FillValuesWorker, BuildOutputWorker


class Ui_Form:
    """
    A PyQt5 UI form class for configuring and executing network diagnostics.

    """

    def setup_ui(self, form):
        """
        Set up the layout and UI elements of the diagnostics form.

        Args:
            form (QWidget): The parent widget to apply the layout and components to.
        """
        self.layout = QtWidgets.QHBoxLayout(form)

        self._create_template_section(form)
        self._create_form_section(form)
        self._create_replace_section(form)
        self._create_output_section(form)

    def _create_template_section(self, parent):
        """
        Create the template selection UI section.

        Args:
            parent (QWidget): The parent widget.
        """
        self.template_widget = QtWidgets.QGroupBox(parent)
        self.template_layout = self._create_section_layout(self.template_widget)

        self.template_header_label = self._create_header_label("Template", self.template_widget)
        self.template_layout.addWidget(self.template_header_label)

        self.template_button = self._create_icon_button("template", self.template_widget)
        self._add_button_to_layout(self.template_layout, self.template_button)

        self.template_description = self._create_description_text(
            'Select a configuration template with variables in "$VARIABLE$" format',
            self.template_widget
        )
        self.template_layout.addWidget(self.template_description)

        self.layout.addWidget(self.template_widget)

    def _create_form_section(self, parent):
        """
        Create the form upload UI section.

        Args:
            parent (QWidget): The parent widget.
        """
        self.form_widget = QtWidgets.QGroupBox(parent)
        self.form_widget.setStyleSheet('QFrame {border: 1px solid grey;}')
        self.form_layout = self._create_section_layout(self.form_widget)

        self.form_header_label = self._create_header_label("Form", self.form_widget)
        self.form_layout.addWidget(self.form_header_label)

        self.form_button = self._create_icon_button("form", self.form_widget)
        self._add_button_to_layout(self.form_layout, self.form_button)

        self.form_description = self._create_description_text(
            "Open excel form and fill values against variables",
            self.form_widget
        )
        self.form_layout.addWidget(self.form_description)

        self.layout.addWidget(self.form_widget)

    def _create_replace_section(self, parent):
        """
        Create the variable replacement section.

        Args:
            parent (QWidget): The parent widget.
        """
        self.replace_widget = QtWidgets.QGroupBox(parent)
        self.replace_widget.setStyleSheet('QFrame {border: 1px solid grey;}')
        self.replace_layout = self._create_section_layout(self.replace_widget)

        self.replace_header_label = self._create_header_label("Find & Replace", self.replace_widget)
        self.replace_layout.addWidget(self.replace_header_label)

        self.replace_button = self._create_icon_button("find-and-replace", self.replace_widget)
        self._add_button_to_layout(self.replace_layout, self.replace_button)

        self.replace_description = self._create_description_text(
            "Find and replace variables with values submitted through the excel form",
            self.replace_widget
        )
        self.replace_layout.addWidget(self.replace_description)

        self.layout.addWidget(self.replace_widget)

    def _create_output_section(self, parent):
        """
        Create the output view section.

        Args:
            parent (QWidget): The parent widget.
        """
        self.output_widget = QtWidgets.QGroupBox(parent)
        self.output_widget.setStyleSheet('QFrame {border: 1px solid grey;}')
        self.output_layout = self._create_section_layout(self.output_widget)

        self.output_header_label = self._create_header_label("Output", self.output_widget)
        self.output_layout.addWidget(self.output_header_label)

        self.output_button = self._create_icon_button("xls", self.output_widget)
        self._add_button_to_layout(self.output_layout, self.output_button)

        self.output_description = self._create_description_text(
            "Open replaced output",
            self.output_widget
        )
        self.output_layout.addWidget(self.output_description)

        self.layout.addWidget(self.output_widget)

    def _create_section_layout(self, parent):
        """
        Create a standardized section layout with spacing.

        Args:
            parent (QWidget): The parent widget.

        Returns:
            QVBoxLayout: Configured layout object.
        """
        layout = QtWidgets.QVBoxLayout(parent)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(40)
        layout.addItem(QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding))
        return layout

    def _create_header_label(self, text, parent):
        """
        Create a styled header label.

        Args:
            text (str): Text to display.
            parent (QWidget): Parent widget.

        Returns:
            QLabel: Configured label widget.
        """
        label = QtWidgets.QLabel(parent)
        label.setText(text)
        label.setStyleSheet("font-weight: 500; font-size: 14pt; border:none;")
        label.setAlignment(QtCore.Qt.AlignCenter)
        return label

    def _create_icon_button(self, icon_name, parent):
        """
        Create a styled button with an icon.

        Args:
            icon_name (str): Name of the icon file (without extension).
            parent (QWidget): Parent widget.

        Returns:
            QPushButton: Configured button widget.
        """
        button = QtWidgets.QPushButton(parent)
        button.setStyleSheet("QPushButton {background: transparent; border-radius: 15px;}")
        button.setIcon(self._get_icon(icon_name))
        button.setMinimumSize(QtCore.QSize(100, 100))
        button.setIconSize(QtCore.QSize(64, 64))
        return button

    def _create_description_text(self, text, parent):
        """
        Create a read-only description text area.

        Args:
            text (str): Description to display.
            parent (QWidget): Parent widget.

        Returns:
            QTextEdit: Configured text edit widget.
        """
        text_edit = QtWidgets.QTextEdit(parent)
        text_edit.setStyleSheet("border:none; background-color: transparent;")
        text_edit.setText(text)
        text_edit.setAlignment(QtCore.Qt.AlignCenter)
        text_edit.setReadOnly(True)
        return text_edit

    def _add_button_to_layout(self, layout, button):
        """
        Add a button centered within a horizontal layout.

        Args:
            layout (QVBoxLayout): The layout to modify.
            button (QPushButton): The button to add.
        """
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addItem(QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Expanding))
        button_layout.addWidget(button)
        button_layout.addItem(QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Expanding))
        layout.addLayout(button_layout)

    def _get_icon(self, filename: str) -> QtGui.QIcon:
        """
        Load an icon from the assets directory.

        Args:
            filename (str): Name of the icon file (without extension).

        Returns:
            QtGui.QIcon: The QIcon object.
        """
        icon_path = os.path.join(os.path.dirname(__file__), "assets", f"{filename}.ico")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(icon_path), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        return icon


class Form(QtWidgets.QWidget, Ui_Form):
    """
    UI Form class.

    """

    def __init__(self, parent=None, **kwargs):
        """
        Initialize the UI form.

        Args:
            parent (QWidget): Parent widget.
            **kwargs: Additional arguments for customization or metadata.
        """
        super().__init__(parent)
        self.kwargs = kwargs
        self.setup_ui(self)
        self.output_dir = str(os.path.join(self.kwargs.get("output_dir"),
                                           os.path.basename(os.path.dirname(__file__)).upper()))
        self.output_report = ''
        self.data = {}

        self.template_button.clicked.connect(self.select_template_event)
        self.form_button.clicked.connect(self.fill_values_event)
        self.replace_button.clicked.connect(self.build_output_event)
        self.output_button.clicked.connect(lambda: self.open_path(self.output_report))

        logging.debug("Form initialized with output_dir: %s", self.output_dir)

    def select_template_event(self):
        """
        Open a file dialog for selecting the configuration template file.
        """
        file_path = QtWidgets.QFileDialog().getOpenFileName(filter='(*.txt *.cfg)')[0]
        if file_path:
            self.data['template_file'] = file_path
            logging.info(f'"{file_path}" selected!')
            logging.info('Now click "Fill Variables" and fill up the values')

    def fill_values_event(self):
        """
        Start the worker thread to extract variables from the selected template.
        """
        if 'template_file' not in self.data:
            logging.info('No template selected')
            return

        self.fill_thread = FillValuesWorker(self.data['template_file'])
        self.fill_thread.fill_complete.connect(self.fill_values_complete)
        self.fill_thread.start()

    def fill_values_complete(self, result):
        """
        Handle the result from the variable extraction process.

        Args:
            result (dict): Dictionary containing parsed template and variables.
        """
        self.data['config_template'] = result['config_template']
        self.data['var_list'] = result['var_list']
        self.data['variable_file_path'] = result['variable_file_path']
        self.open_path(self.data['variable_file_path'])

    def build_output_event(self):
        """
        Start the output generation process by replacing variables in the template.
        """
        if 'variable_file_path' not in self.data or 'config_template' not in self.data:
            logging.info('Missing template or variable file')
            return

        os.makedirs(self.output_dir, exist_ok=True)

        timestamp = datetime.now().strftime('%Y-%m-%d_%H.%M')
        filename = f"{os.path.basename(os.path.dirname(__file__)).title()}_{timestamp}.xlsx"
        self.output_report = os.path.join(self.output_dir, filename)

        self.build_thread = BuildOutputWorker(
            self.data['config_template'],
            self.data['variable_file_path'],
            self.output_report
        )
        self.build_thread.start()
        self.build_thread.finished.connect(self.build_finished)

    def build_finished(self):
        """
        Handle the build completion.
        """
        QtWidgets.QMessageBox.information(self, "Info", "Task completed!!")

    def open_path(self, path: str):
        """
        Open a file or directory using the system's default handler.

        Args:
            path (str): File or directory path to open.
        """
        try:
            if path and os.path.exists(path):
                logging.info(f"Opening path: {path}")
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
            else:
                logging.error(f"Invalid or non-existent path: {path}")
        except Exception as e:
            logging.exception(f"Failed to open path: {e}")
