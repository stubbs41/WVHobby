#!/usr/bin/env python3
import sys
import os
import pandas as pd
import configparser
import re
from PySide6.QtWidgets import (QApplication, QMainWindow, QLabel, QComboBox, 
                             QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, 
                             QWidget, QLineEdit, QMessageBox, QFrame)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon, QPixmap


class POFormatter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'po_formatter.ini')
        self.load_settings()
        self.initUI()
        self.current_file = None
        self.df = None
        
        # Set window icon
        self.setWindowIcon(self.get_app_icon())
        
    def get_app_icon(self):
        """Load the application icon"""
        # Look in several possible locations for the icon file
        icon_paths = [
            os.path.join("Images", "POChange.png"),  # Images folder
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images", "POChange.png"),  # Script dir/images
            os.path.join(sys._MEIPASS, "Images", "POChange.png") if hasattr(sys, '_MEIPASS') else None  # PyInstaller bundle
        ]
        
        for path in icon_paths:
            if path and os.path.exists(path):
                return QIcon(path)
                
        # Return a default icon if the image is not found
        return QIcon()
        
    def load_settings(self):
        """Load saved directory paths from config file"""
        self.last_input_dir = ""
        self.last_output_dir = ""
        
        if os.path.exists(self.config_file):
            config = configparser.ConfigParser()
            config.read(self.config_file)
            
            if 'Directories' in config:
                if 'input_dir' in config['Directories']:
                    self.last_input_dir = config['Directories']['input_dir']
                if 'output_dir' in config['Directories']:
                    self.last_output_dir = config['Directories']['output_dir']
    
    def save_settings(self):
        """Save directory paths to config file"""
        config = configparser.ConfigParser()
        config['Directories'] = {
            'input_dir': self.last_input_dir,
            'output_dir': self.last_output_dir
        }
        
        with open(self.config_file, 'w') as f:
            config.write(f)
        
    def initUI(self):
        self.setWindowTitle('Purchase Order Formatter')
        self.setGeometry(100, 100, 600, 400)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Add logo image at the top
        logo_layout = QHBoxLayout()
        logo_label = QLabel()
        
        # Try to load the logo image
        logo_paths = [
            os.path.join("Images", "POChange.png"),  # Images folder
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images", "POChange.png"),  # Script dir/images
            os.path.join(sys._MEIPASS, "Images", "POChange.png") if hasattr(sys, '_MEIPASS') else None  # PyInstaller bundle
        ]
        
        logo_pixmap = None
        for path in logo_paths:
            if path and os.path.exists(path):
                logo_pixmap = QPixmap(path)
                if not logo_pixmap.isNull():
                    break
        
        if logo_pixmap and not logo_pixmap.isNull():
            # Scale the image to a reasonable size if it's too large
            if logo_pixmap.width() > 200:
                logo_pixmap = logo_pixmap.scaledToWidth(200, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            logo_layout.addStretch()
            logo_layout.addWidget(logo_label)
            logo_layout.addStretch()
            main_layout.addLayout(logo_layout)
        
        # Application title
        title_label = QLabel('Purchase Order Formatter')
        title_label.setFont(QFont('Arial', 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Description
        desc_label = QLabel('Select an Excel file and vendor to format your PO for upload')
        desc_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(desc_label)
        
        # Add separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)
        
        # File selection section
        file_layout = QHBoxLayout()
        file_label = QLabel('PO File:')
        self.file_path_label = QLabel('No file selected')
        self.browse_button = QPushButton('Browse...')
        self.browse_button.clicked.connect(self.browse_file)
        
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_label, 1)
        file_layout.addWidget(self.browse_button)
        main_layout.addLayout(file_layout)
        
        # PO number section
        po_layout = QHBoxLayout()
        po_label = QLabel('PO Number:')
        self.po_input = QLineEdit()
        self.po_input.setEnabled(False)  # Initially disabled until file is selected
        
        po_layout.addWidget(po_label)
        po_layout.addWidget(self.po_input)
        main_layout.addLayout(po_layout)
        
        # Vendor selection section
        vendor_layout = QHBoxLayout()
        vendor_label = QLabel('Vendor:')
        self.vendor_combo = QComboBox()
        self.vendor_combo.addItems(['Select a vendor', 'HorizonHobby/FastServe', 'Stephens', 'HRP', 'AMAIN', 'Traxxas'])
        self.vendor_combo.setEnabled(False)  # Initially disabled until file is selected
        
        vendor_layout.addWidget(vendor_label)
        vendor_layout.addWidget(self.vendor_combo)
        main_layout.addLayout(vendor_layout)
        
        # Action buttons
        button_layout = QHBoxLayout()
        self.process_button = QPushButton('Process')
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.process_file)
        
        self.cancel_button = QPushButton('Cancel')
        self.cancel_button.clicked.connect(self.close)
        
        button_layout.addWidget(self.process_button)
        button_layout.addWidget(self.cancel_button)
        main_layout.addLayout(button_layout)
        
        # Add spacer at the bottom
        main_layout.addStretch()
        
        # Status message
        self.status_label = QLabel('')
        self.status_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_label)
        
    def browse_file(self):
        # Use last input directory if available
        start_dir = self.last_input_dir if self.last_input_dir else ""
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Select PO File', start_dir, 'Excel and CSV Files (*.xlsx *.xls *.csv);; INV Files (*.inv);; All Files (*.*)'
        )
        
        if file_path:
            # Store the directory for future use
            self.last_input_dir = os.path.dirname(file_path)
            self.save_settings()
            
            self.current_file = file_path
            filename = os.path.basename(file_path)
            self.file_path_label.setText(filename)
            
            # Extract PO number from filename (remove extension)
            # For filenames that are just numbers like 17633.xlsx
            po_number = os.path.splitext(filename)[0].strip()
            
            # If the filename starts with "PO", extract the number part
            if po_number.upper().startswith('PO'):
                po_number = po_number[2:].strip()
                
            self.po_input.setText(po_number)
            
            # Enable PO input and vendor selection
            self.po_input.setEnabled(True)
            self.vendor_combo.setEnabled(True)
            self.process_button.setEnabled(True)
            
            try:
                # Load the file based on extension
                if file_path.lower().endswith('.csv'):
                    self.df = pd.read_csv(file_path)
                elif file_path.lower().endswith('.inv'):
                    # INV files are typically for Traxxas
                    self.df = pd.read_csv(file_path)
                    self.vendor_combo.setCurrentText('Traxxas')
                else:
                    self.df = pd.read_excel(file_path)
                
                self.status_label.setText(f"File loaded successfully: {len(self.df)} rows")
                
                # Try to auto-detect vendor format
                # For FastServe CSV format
                if all(col in self.df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                    self.vendor_combo.setCurrentText('HorizonHobby/FastServe')
                
                # For Traxxas format
                if set(['SKU', 'QTY']).issubset(self.df.columns):
                    self.vendor_combo.setCurrentText('Traxxas')
                elif file_path.lower().endswith('.inv'):
                    self.vendor_combo.setCurrentText('Traxxas')
                    
            except Exception as e:
                self.status_label.setText(f"Error loading file: {str(e)}")
                self.df = None
                self.process_button.setEnabled(False)
    
    def process_file(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No valid Excel file loaded")
            return
            
        po_number = self.po_input.text().strip()
        if not po_number:
            QMessageBox.warning(self, "Error", "PO Number is required")
            return
            
        vendor_index = self.vendor_combo.currentIndex()
        if vendor_index == 0:
            QMessageBox.warning(self, "Error", "Please select a vendor")
            return
            
        vendor = self.vendor_combo.currentText()
        
        # Ask user where to save the output file
        default_filename = f"{po_number}_{vendor.replace(' ', '')}"
        
        try:
            # Process based on vendor selection
            if vendor == "HorizonHobby/FastServe":
                output_path = self.format_horizon_fastserve(po_number)
            elif vendor == "Stephens":
                output_path = self.format_stephens(po_number)
            elif vendor == "HRP":
                output_path = self.format_hrp(po_number)
            elif vendor == "AMAIN":
                output_path = self.format_amain(po_number)
            elif vendor == "Traxxas":
                output_path = self.format_traxxas(po_number)
            else:
                QMessageBox.warning(self, "Error", "Invalid vendor selection")
                return
                
            QMessageBox.information(
                self, "Success", 
                f"File successfully processed and saved as:\n{output_path}"
            )
            self.reset_ui()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def reset_ui(self):
        self.current_file = None
        self.df = None
        self.file_path_label.setText('No file selected')
        self.po_input.setText('')
        self.po_input.setEnabled(False)
        self.vendor_combo.setCurrentIndex(0)
        self.vendor_combo.setEnabled(False)
        self.process_button.setEnabled(False)
        self.status_label.setText('')
    
    # Vendor-specific formatting methods
    def format_horizon_fastserve(self, po_number):
        """
        Format for HorizonHobby/FastServe
        Expected format: PO# / SKU1 / Qty1 / SKU2 / Qty2 / ... / END / SkuCount
        Output file: FastServe-[PONumber].txt
        
        Can process either:
        1. Excel files with Sku and Qty columns
        2. CSV files with PO_NUMBER, ITEM_NUMBER, DESCRIPTION, QTY, UNIT_PRICE, TOTAL columns
        """
        try:
            processed_df = self.df.copy()
            
            # Check if this is a CSV format with the expected columns
            if all(col in processed_df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                # Use the CSV format columns
                formatted_df = processed_df[['ITEM_NUMBER', 'QTY']]
                formatted_df.columns = ['Sku', 'Qty']
            else:
                # Check for required columns from Excel format
                required_columns = ['Sku', 'Qty']
                missing_columns = [col for col in required_columns if col not in processed_df.columns]
                
                if missing_columns:
                    raise ValueError(f"Input file missing required columns: {', '.join(missing_columns)}")
                
                # Get only the needed columns
                formatted_df = processed_df[['Sku', 'Qty']]
            
            # Create the formatted string content
            # Start with PO number
            content = po_number
            
            # Add each SKU and quantity pair
            for index, row in formatted_df.iterrows():
                content += f" / {row['Sku']} / {int(row['Qty'])}"
            
            # Add the END marker and SKU count
            sku_count = len(formatted_df)
            content += f" / END / {sku_count}"
            
            # Ask user where to save the file
            # Use last output directory if available, otherwise use input directory
            start_dir = self.last_output_dir if self.last_output_dir else self.last_input_dir
            default_name = f"FastServe-{po_number}.txt"
            start_path = os.path.join(start_dir, default_name) if start_dir else default_name
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 'Save Formatted File', start_path, 'Text Files (*.txt)'
            )
            
            if not file_path:
                raise ValueError("Save operation cancelled by user")
            
            # Store the output directory for future use
            self.last_output_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Write content to file
            with open(file_path, 'w') as f:
                f.write(content)
            
            return file_path
            
        except Exception as e:
            raise Exception(f"Error formatting for HorizonHobby/FastServe: {str(e)}")
    
    def format_stephens(self, po_number):
        """
        Format for Stephens International
        Format:
        - Line 1: PO number
        - Lines 2, 4, 6, etc.: Product SKUs
        - Lines 3, 5, 7, etc.: Quantities for corresponding SKUs
        - Second-to-last line: "END" marker
        - Last line: Count of products ordered
        """
        try:
            processed_df = self.df.copy()
            
            # Check if this is a CSV format with the expected columns
            if all(col in processed_df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                # Use the CSV format columns
                formatted_df = processed_df[['ITEM_NUMBER', 'QTY']]
                formatted_df.columns = ['Sku', 'Qty']
            else:
                # Check for required columns
                required_columns = ['Sku', 'Qty']
                missing_columns = [col for col in required_columns if col not in processed_df.columns]
                
                if missing_columns:
                    # Try to find similar column names
                    sku_columns = [col for col in processed_df.columns if 'sku' in col.lower() or 'item' in col.lower() or 'part' in col.lower()]
                    qty_columns = [col for col in processed_df.columns if 'qty' in col.lower() or 'quantity' in col.lower()]
                    
                    if sku_columns and qty_columns:
                        # Use the first matching columns
                        formatted_df = processed_df[[sku_columns[0], qty_columns[0]]].copy()
                        formatted_df.columns = ['Sku', 'Qty']
                    else:
                        raise ValueError(f"Input file missing required columns: {', '.join(missing_columns)}")
                else:
                    # Get only the needed columns
                    formatted_df = processed_df[['Sku', 'Qty']]
            
            # Create the formatted content
            lines = []
            
            # First line: PO number
            lines.append(po_number)
            
            # Add each product as a pair of lines (SKU line followed by quantity line)
            for index, row in formatted_df.iterrows():
                # SKU line
                lines.append(str(row['Sku']))
                # Quantity line
                lines.append(str(int(row['Qty'])))
            
            # Add the END marker
            lines.append("END")
            
            # Last line: count of products
            product_count = len(formatted_df)
            lines.append(str(product_count))
            
            # Ask user where to save the file
            # Use last output directory if available, otherwise use input directory
            start_dir = self.last_output_dir if self.last_output_dir else self.last_input_dir
            default_name = f"{po_number}_Stephens.txt"
            start_path = os.path.join(start_dir, default_name) if start_dir else default_name
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 'Save Formatted File', start_path, 'Text Files (*.txt)'
            )
            
            if not file_path:
                raise ValueError("Save operation cancelled by user")
            
            # Store the output directory for future use
            self.last_output_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Write content to file
            with open(file_path, 'w') as f:
                f.write('\n'.join(lines))
            
            return file_path
            
        except Exception as e:
            raise Exception(f"Error formatting for Stephens: {str(e)}")
    
    def format_hrp(self, po_number):
        """
        Format for HRP
        Expected format: CSV or TXT with columns:
        - PART # (SKU)
        - QTY
        - WAREHOUSE(Optional)
        
        Output: Tab-delimited text file or CSV file
        """
        try:
            processed_df = self.df.copy()
            
            # Check if this is a CSV format with the expected columns
            if all(col in processed_df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                # Use the CSV format columns
                formatted_df = processed_df[['ITEM_NUMBER', 'QTY']]
                formatted_df.columns = ['Sku', 'Qty']
            else:
                # Check for required columns
                required_columns = ['Sku', 'Qty']
                missing_columns = [col for col in required_columns if col not in processed_df.columns]
                
                if missing_columns:
                    # Try to find similar column names
                    sku_columns = [col for col in processed_df.columns if 'sku' in col.lower() or 'item' in col.lower() or 'part' in col.lower()]
                    qty_columns = [col for col in processed_df.columns if 'qty' in col.lower() or 'quantity' in col.lower()]
                    
                    if sku_columns and qty_columns:
                        # Use the first matching columns
                        formatted_df = processed_df[[sku_columns[0], qty_columns[0]]].copy()
                        formatted_df.columns = ['Sku', 'Qty']
                    else:
                        raise ValueError(f"Input file missing required columns: {', '.join(missing_columns)}")
                else:
                    # Get only the needed columns
                    formatted_df = processed_df[['Sku', 'Qty']]
            
            # Create the output dataframe with the HRP required format
            hrp_df = pd.DataFrame()
            hrp_df['PART #'] = formatted_df['Sku']
            hrp_df['QTY'] = formatted_df['Qty'].astype(int)
            # Add empty WAREHOUSE column - will use user's default warehouse
            hrp_df['WAREHOUSE(Optional)'] = ""
            
            # Ask user for format preference
            format_dialog = QMessageBox()
            format_dialog.setWindowTitle("Choose Output Format")
            format_dialog.setText("Choose the output format for HRP:")
            format_dialog.setIcon(QMessageBox.Question)
            
            csv_button = format_dialog.addButton("CSV", QMessageBox.ActionRole)
            txt_button = format_dialog.addButton("TXT (Tab-delimited)", QMessageBox.ActionRole)
            cancel_button = format_dialog.addButton(QMessageBox.Cancel)
            
            format_dialog.exec()
            
            if format_dialog.clickedButton() == cancel_button:
                raise ValueError("Format selection cancelled by user")
            
            use_txt_format = format_dialog.clickedButton() == txt_button
            
            # Ask user where to save the file
            # Use last output directory if available, otherwise use input directory
            start_dir = self.last_output_dir if self.last_output_dir else self.last_input_dir
            
            if use_txt_format:
                default_name = f"{po_number}_HRP.txt"
                file_filter = 'Text Files (*.txt)'
            else:
                default_name = f"{po_number}_HRP.csv"
                file_filter = 'CSV Files (*.csv)'
                
            start_path = os.path.join(start_dir, default_name) if start_dir else default_name
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 'Save Formatted File', start_path, file_filter
            )
            
            if not file_path:
                raise ValueError("Save operation cancelled by user")
            
            # Store the output directory for future use
            self.last_output_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Save in requested format
            if use_txt_format:
                # Save as tab-delimited text file
                hrp_df.to_csv(file_path, sep='\t', index=False)
            else:
                # Save as CSV
                hrp_df.to_csv(file_path, index=False)
                
            return file_path
            
        except Exception as e:
            raise Exception(f"Error formatting for HRP: {str(e)}")
    
    def format_amain(self, po_number):
        """
        Format for AMAIN
        Supports three formats as shown in the screenshot:
        1. HOST file format - A specialized format with PO number and SKU/quantity pairs
        2. CSV file - Standard CSV format with SKU and quantity
        3. Tab-delimited - Same as CSV but with tab separators
        """
        try:
            processed_df = self.df.copy()
            
            # Check if this is a CSV format with the expected columns
            if all(col in processed_df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                # Use the CSV format columns
                formatted_df = processed_df[['ITEM_NUMBER', 'QTY']]
                formatted_df.columns = ['Sku', 'Qty']
            else:
                # Check for required columns
                required_columns = ['Sku', 'Qty']
                missing_columns = [col for col in required_columns if col not in processed_df.columns]
                
                if missing_columns:
                    # Try to find similar column names
                    sku_columns = [col for col in processed_df.columns if 'sku' in col.lower() or 'item' in col.lower() or 'part' in col.lower()]
                    qty_columns = [col for col in processed_df.columns if 'qty' in col.lower() or 'quantity' in col.lower()]
                    
                    if sku_columns and qty_columns:
                        # Use the first matching columns
                        formatted_df = processed_df[[sku_columns[0], qty_columns[0]]].copy()
                        formatted_df.columns = ['Sku', 'Qty']
                    else:
                        raise ValueError(f"Input file missing required columns: {', '.join(missing_columns)}")
                else:
                    # Get only the needed columns
                    formatted_df = processed_df[['Sku', 'Qty']]
            
            # Ask user for format preference
            format_dialog = QMessageBox()
            format_dialog.setWindowTitle("Choose Output Format")
            format_dialog.setText("Choose the output format for AMAIN:")
            format_dialog.setIcon(QMessageBox.Question)
            
            host_button = format_dialog.addButton("HOST", QMessageBox.ActionRole)
            csv_button = format_dialog.addButton("CSV", QMessageBox.ActionRole)
            tab_button = format_dialog.addButton("Tab-delimited", QMessageBox.ActionRole)
            cancel_button = format_dialog.addButton(QMessageBox.Cancel)
            
            format_dialog.exec()
            
            if format_dialog.clickedButton() == cancel_button:
                raise ValueError("Format selection cancelled by user")
            
            use_host_format = format_dialog.clickedButton() == host_button
            use_tab_format = format_dialog.clickedButton() == tab_button
            
            # Ask user where to save the file
            # Use last output directory if available, otherwise use input directory
            start_dir = self.last_output_dir if self.last_output_dir else self.last_input_dir
            
            if use_host_format:
                default_name = f"{po_number}_AMAIN.txt"
                file_filter = 'HOST Files (*.txt)'
            elif use_tab_format:
                default_name = f"{po_number}_AMAIN.txt"
                file_filter = 'Text Files (*.txt)'
            else:
                default_name = f"{po_number}_AMAIN.csv"
                file_filter = 'CSV Files (*.csv)'
                
            start_path = os.path.join(start_dir, default_name) if start_dir else default_name
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 'Save Formatted File', start_path, file_filter
            )
            
            if not file_path:
                raise ValueError("Save operation cancelled by user")
            
            # Store the output directory for future use
            self.last_output_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Format and save based on selected format
            if use_host_format:
                # Create HOST format file
                # First line: PO number
                lines = [f"{po_number} [This is your PO number]"]
                
                # Item lines: SKU and quantity for each product
                for index, row in formatted_df.iterrows():
                    lines.append(f"{row['Sku']} [Product to Order]")
                    lines.append(f"{int(row['Qty'])} [Qty to order for {row['Sku']}]")
                
                # End line with count
                sku_count = len(formatted_df)
                lines.append(f"END [Added to every HOST file, represents END of list]")
                lines.append(f"{sku_count} [Total number of line items added, not total of any added]")
                
                # Write content to file
                with open(file_path, 'w') as f:
                    f.write('\n'.join(lines))
            elif use_tab_format:
                # Save as tab-delimited text file without headers
                with open(file_path, 'w') as f:
                    # Write each row without headers
                    for _, row in formatted_df.iterrows():
                        f.write(f"{row['Sku']}\t{row['Qty']}\n")
            else:
                # Save as CSV without headers
                with open(file_path, 'w') as f:
                    # Write each row without headers
                    for _, row in formatted_df.iterrows():
                        f.write(f"{row['Sku']},{row['Qty']}\n")
                
            return file_path
            
        except Exception as e:
            raise Exception(f"Error formatting for AMAIN: {str(e)}")
    
    def format_traxxas(self, po_number):
        """
        Format for Traxxas
        Expected format: CSV with SKUs: Must include "SKU" and "QTY" columns
        Note: If SKUs start with "tra", this prefix will be automatically removed
        Output format can be either standard SKU/QTY format or the new template with sku,qty,variant,comment
        """
        try:
            processed_df = self.df.copy()
            
            # Check for color variants in Item/SKU columns (like RED, GRN, BLUE in the SKU)
            has_variants = False
            
            # Try to find relevant columns for processing
            sku_columns = [col for col in processed_df.columns if 'sku' in col.lower() or 'item' in col.lower() or 'part' in col.lower()]
            qty_columns = [col for col in processed_df.columns if 'qty' in col.lower() or 'quantity' in col.lower()]
            
            if not sku_columns or not qty_columns:
                raise ValueError("Could not find required SKU and QTY columns. File must have columns for item SKU and quantity.")
            
            # Select the best columns to use
            sku_col = 'Sku' if 'Sku' in processed_df.columns else sku_columns[0]
            qty_col = 'Qty' if 'Qty' in processed_df.columns else qty_columns[0]
            
            # Check if Item or SKU column has color variants in the format XXXXX-COLOR
            color_pattern = r'-([A-Z]+)$'  # Matches -RED, -GRN, -BLUE at the end of string
            sample_skus = processed_df[sku_col].astype(str).tolist()
            
            for sku in sample_skus:
                if re.search(color_pattern, sku):
                    has_variants = True
                    break
                    
            # Ask user for output format
            format_dialog = QMessageBox()
            format_dialog.setWindowTitle("Choose Output Format")
            format_dialog.setText("Choose the output format for Traxxas:")
            format_dialog.setIcon(QMessageBox.Question)
            
            csv_button = format_dialog.addButton("Standard CSV/INV", QMessageBox.ActionRole)
            template_button = format_dialog.addButton("New Template Format", QMessageBox.ActionRole)
            cancel_button = format_dialog.addButton(QMessageBox.Cancel)
            
            format_dialog.exec()
            
            if format_dialog.clickedButton() == cancel_button:
                raise ValueError("Format selection cancelled by user")
            
            use_template_format = format_dialog.clickedButton() == template_button
            
            # Create the output dataframe based on selected format
            if use_template_format:
                # New template format with sku, qty, variant, comment
                traxxas_df = pd.DataFrame(columns=['sku', 'qty', 'variant', 'comment'])
                
                # Process each row and handle variants
                rows = []
                for _, row in processed_df.iterrows():
                    sku = str(row[sku_col])
                    qty = row[qty_col]
                    variant = ""
                    
                    # Extract variant from SKU if it exists
                    if has_variants:
                        match = re.search(color_pattern, sku)
                        if match:
                            color_code = match.group(1)
                            # Convert color codes to full color names
                            if color_code == "RED":
                                variant = "Red"
                            elif color_code == "GRN":
                                variant = "Green"
                            elif color_code == "BLU" or color_code == "BLUE":
                                variant = "Blue"
                            elif color_code == "YEL":
                                variant = "Yellow"
                            elif color_code == "BLK":
                                variant = "Black"
                            elif color_code == "WHT":
                                variant = "White"
                            elif color_code == "PNK":
                                variant = "Pink"
                            elif color_code == "PUR":
                                variant = "Purple"
                            elif color_code == "ORG":
                                variant = "Orange"
                            else:
                                variant = color_code  # Use code as is if not recognized
                            
                            # Remove the color code from the SKU
                            sku = sku.rsplit('-', 1)[0]
                    
                    # Remove "tra" prefix from SKUs if present
                    if sku.lower().startswith('tra'):
                        sku = sku[3:]
                        
                    rows.append({
                        'sku': sku,
                        'qty': qty,
                        'variant': variant,
                        'comment': ""
                    })
                
                traxxas_df = pd.DataFrame(rows)
            else:
                # Standard format (just SKU and QTY)
                traxxas_df = pd.DataFrame()
                traxxas_df['SKU'] = processed_df[sku_col].astype(str)
                traxxas_df['QTY'] = processed_df[qty_col]
                
                # Remove "tra" prefix from SKUs if present
                traxxas_df['SKU'] = traxxas_df['SKU'].apply(
                    lambda sku: sku[3:] if sku.lower().startswith('tra') else sku
                )
            
            # Ask user where to save the file
            # Use last output directory if available, otherwise use input directory
            start_dir = self.last_output_dir if self.last_output_dir else self.last_input_dir
            
            if use_template_format:
                default_name = f"{po_number}_Traxxas_Template.csv"
            else:
                if format_dialog.clickedButton() == csv_button:
                    default_name = f"{po_number}_Traxxas.csv"
                else:
                    default_name = f"{po_number}_Traxxas.inv"
                    
            file_filter = 'CSV Files (*.csv);;INV Files (*.inv)'
            start_path = os.path.join(start_dir, default_name) if start_dir else default_name
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 'Save Formatted File', start_path, file_filter
            )
            
            if not file_path:
                raise ValueError("Save operation cancelled by user")
            
            # Store the output directory for future use
            self.last_output_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Save in requested format
            traxxas_df.to_csv(file_path, index=False)
            return file_path
            
        except Exception as e:
            raise Exception(f"Error formatting for Traxxas: {str(e)}")

    def process_command_line_file(self, file_path):
        """
        Process a file passed directly from command line
        """
        if not os.path.exists(file_path):
            QMessageBox.critical(self, "Error", f"File not found: {file_path}")
            return
        
        self.current_file = file_path
        filename = os.path.basename(file_path)
        self.file_path_label.setText(filename)
        
        # Extract PO number from filename (remove extension)
        po_number = os.path.splitext(filename)[0].strip()
        
        # If the filename starts with "PO", extract the number part
        if po_number.upper().startswith('PO'):
            po_number = po_number[2:].strip()
        
        self.po_input.setText(po_number)
        
        try:
            # Load the file based on extension
            if file_path.lower().endswith('.csv'):
                self.df = pd.read_csv(file_path)
            elif file_path.lower().endswith('.inv'):
                # INV files are typically for Traxxas
                self.df = pd.read_csv(file_path)
                self.vendor_combo.setCurrentText('Traxxas')
            else:
                self.df = pd.read_excel(file_path)
            
            self.status_label.setText(f"File loaded successfully: {len(self.df)} rows")
            
            # Enable PO input and vendor selection
            self.po_input.setEnabled(True)
            self.vendor_combo.setEnabled(True)
            self.process_button.setEnabled(True)
            
            # Store the directory for future use
            self.last_input_dir = os.path.dirname(file_path)
            self.save_settings()
            
            # Try to auto-detect vendor format
            # For FastServe CSV format
            if all(col in self.df.columns for col in ['PO_NUMBER', 'ITEM_NUMBER', 'QTY']):
                self.vendor_combo.setCurrentText('HorizonHobby/FastServe')
            
            # For Traxxas format
            if set(['SKU', 'QTY']).issubset(self.df.columns):
                self.vendor_combo.setCurrentText('Traxxas')
            elif file_path.lower().endswith('.inv'):
                self.vendor_combo.setCurrentText('Traxxas')
                
        except Exception as e:
            self.status_label.setText(f"Error loading file: {str(e)}")
            self.df = None
            self.process_button.setEnabled(False)


def main():
    # Set application info
    app = QApplication(sys.argv)
    app.setApplicationName("PO Formatter")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("WVHobby")
    app.setStyle('Fusion')  # Modern looking style
    
    # Ensure app icon is set for the entire application
    icon_paths = [
        os.path.join("Images", "POChange.png"),  # Images folder
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images", "POChange.png"),  # Script dir/images
        os.path.join(sys._MEIPASS, "Images", "POChange.png") if hasattr(sys, '_MEIPASS') else None  # PyInstaller bundle
    ]
    
    for path in icon_paths:
        if path and os.path.exists(path):
            app.setWindowIcon(QIcon(path))
            break
    
    # Create and show the main window
    window = POFormatter()
    
    # Handle command line arguments - if a file path is provided, open it
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        window.process_command_line_file(sys.argv[1])
    
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main() 