"""
Dr. Alfredo Pio De Roda ES - QR Code Generator (Toga Version)
Version: TOGA MIGRATION - Cross-platform with Briefcase
Date: January 30, 2026

FIXED: Cleaner UI matching Tkinter design
"""

import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
from openpyxl import load_workbook
import qrcode
import os
from datetime import datetime
from pathlib import Path


class QRGenerator(toga.App):
    def startup(self):
        """Setup the application"""
        # Setup folders
        self.home_dir = Path.home()
        self.base_folder = self.home_dir / "SF2_Files"
        self.qr_folder = self.base_folder / "QR_Codes"
        self.active_folder = self.base_folder / "Active"
        
        for folder in [self.qr_folder, self.active_folder]:
            folder.mkdir(parents=True, exist_ok=True)
        
        self.sf2_file = None
        self.student_names = []
        
        # Build UI
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = self.create_ui()
        self.main_window.show()
    
    def is_valid_student_name(self, name):
        """Validate student name"""
        if not name or not isinstance(name, str):
            return False
        
        name = name.strip()
        if len(name) < 2:
            return False
        
        name_upper = name.upper()
        
        excluded_patterns = [
            "SUMIF", "COUNTIF", "AVERAGE", "SCHOOL FORM", "SF2", 
            "TOTAL", "MALE", "FEMALE", "PERCENTAGE", "ENROLMENT",
            "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY",
            "NAN", "NONE", "N/A", "NULL", "BLANK", "EMPTY",
        ]
        
        for pattern in excluded_patterns:
            if pattern in name_upper:
                return False
        
        if not any(c.isalpha() for c in name):
            return False
        
        return True
    
    def create_ui(self):
        """Create the UI"""
        main_box = toga.Box(style=Pack(direction=COLUMN, padding=15))
        
        # Header
        header_box = toga.Box(style=Pack(direction=COLUMN, padding=10))
        title = toga.Label("🎯 QR Code Generator", style=Pack(padding=5, font_size=18, font_weight='bold'))
        subtitle = toga.Label("Generate QR codes for all students in your SF2 file", style=Pack(padding=5))
        header_box.add(title)
        header_box.add(subtitle)
        main_box.add(header_box)
        
        # File selection
        file_box = toga.Box(style=Pack(direction=COLUMN, padding=10))
        file_label = toga.Label("📄 Select SF2 File", style=Pack(padding=5, font_weight='bold'))
        file_box.add(file_label)
        
        file_button_box = toga.Box(style=Pack(direction=ROW, padding=5))
        browse_btn = toga.Button("📁 Browse & Load SF2 File", on_press=self.browse_file, style=Pack(padding=5))
        file_button_box.add(browse_btn)
        
        self.file_status_label = toga.Label("No file selected", style=Pack(padding=5))
        file_button_box.add(self.file_status_label)
        file_box.add(file_button_box)
        main_box.add(file_box)
        
        # Info section
        info_box = toga.Box(style=Pack(direction=COLUMN, padding=10))
        info_title = toga.Label("📊 File Information", style=Pack(padding=5, font_weight='bold'))
        info_box.add(info_title)
        
        self.info_label = toga.Label("Students: 0\nQR Codes: 0\nStatus: Ready to load", style=Pack(padding=5))
        info_box.add(self.info_label)
        main_box.add(info_box)
        
        # Progress section
        progress_box = toga.Box(style=Pack(direction=COLUMN, padding=10))
        progress_title = toga.Label("⏳ Generation Progress", style=Pack(padding=5, font_weight='bold'))
        progress_box.add(progress_title)
        
        self.progress_bar = toga.ProgressBar(max=100, style=Pack(padding=5, width=500))
        progress_box.add(self.progress_bar)
        
        self.status_label = toga.Label("Status: Idle", style=Pack(padding=5))
        progress_box.add(self.status_label)
        main_box.add(progress_box)
        
        # Generate buttons
        button_box = toga.Box(style=Pack(direction=ROW, padding=10))
        
        self.generate_btn = toga.Button("🎯 Generate All QR Codes", on_press=self.generate_qr_codes,
                                       enabled=False, style=Pack(flex=1, padding=5))
        qr_folder_btn = toga.Button("📂 QR FOLDER", on_press=self.open_qr_folder, style=Pack(flex=1, padding=5))
        
        button_box.add(self.generate_btn)
        button_box.add(qr_folder_btn)
        main_box.add(button_box)
        
        # Info text
        help_box = toga.Box(style=Pack(direction=COLUMN, padding=10))
        help_title = toga.Label("ℹ️ How It Works", style=Pack(padding=5, font_weight='bold'))
        help_box.add(help_title)
        
        help_text = toga.MultilineTextInput(
            value=(
                "1. Click 'Browse & Load SF2 File' to select your Excel file\n"
                "2. System loads all valid student names from Column B\n"
                "3. Click 'Generate All QR Codes' to create QR codes\n"
                f"4. QR codes are saved to: {self.qr_folder}\n"
                "5. Print the QR codes and attach to student IDs\n"
                "6. Use the Attendance System to scan QR codes\n\n"
                "WHAT IT GENERATES:\n"
                "• Individual QR code images for each student\n"
                "• QR code contains student name\n"
                "• Organized in QR_Codes folder\n"
                "• Easy to print and laminate\n\n"
                "TIPS:\n"
                "• Print on sticker sheets for easy ID attachment\n"
                "• Use 2\"x2\" or 1.5\"x1.5\" size for best scanning\n"
                "• Laminate for durability"
            ),
            readonly=True,
            style=Pack(flex=1, padding=5)
        )
        help_box.add(help_text)
        main_box.add(help_box)
        
        return main_box
    
    def browse_file(self, widget):
        """Browse for SF2 file"""
        try:
            self.main_window.open_file_dialog(
                title="Select SF2 Excel File",
                initial_directory=self.active_folder,
                file_types=['xlsx', 'xls'],
                on_result=self.load_file
            )
        except Exception as e:
            print(f"Browse error: {e}")
            self.main_window.error_dialog("Error", f"Failed to browse: {e}")
    
    def load_file(self, widget, file_path):
        """Load SF2 file"""
        if file_path is None:
            return
        
        try:
            print(f"\nLOADING FILE: {file_path.name}")
            
            workbook = load_workbook(file_path)
            sheet = workbook.active
            
            self.sf2_file = file_path
            self.student_names = []
            
            print(f"👥 Extracting students from Column B...")
            
            for row in range(13, sheet.max_row + 1):
                name_cell = sheet.cell(row, 2).value
                
                if not name_cell:
                    continue
                
                if self.is_valid_student_name(name_cell):
                    name = name_cell.strip()
                    self.student_names.append(name)
                    print(f"  ✅ {name}")
                else:
                    print(f"  ⊘  FILTERED - '{name_cell}'")
            
            print(f"✅ Loaded {len(self.student_names)} valid students\n")
            
            # Update UI
            self.file_status_label.text = f"✅ Loaded: {file_path.name}"
            self.info_label.text = f"Students: {len(self.student_names)}\nQR Codes: Ready to generate\nStatus: File loaded successfully"
            self.generate_btn.enabled = True
        
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            self.main_window.error_dialog("Error", f"Failed to load file: {e}")
    
    def generate_qr_codes(self, widget):
        """Generate QR codes for all students"""
        if not self.student_names:
            self.main_window.info_dialog("Warning", "No students loaded!")
            return
        
        try:
            print(f"\nGENERATING QR CODES")
            print(f"Total students: {len(self.student_names)}")
            print(f"Output folder: {self.qr_folder}\n")
            
            self.generate_btn.enabled = False
            self.progress_bar.max = len(self.student_names)
            
            for index, student_name in enumerate(self.student_names):
                # Generate QR code
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=4,
                )
                qr.add_data(student_name)
                qr.make(fit=True)
                
                img = qr.make_image(fill_color="black", back_color="white")
                
                # Save QR code
                filename = self.qr_folder / f"{student_name.replace(' ', '_')}.png"
                img.save(filename)
                
                print(f"  ✅ {index+1}/{len(self.student_names)}: {student_name}")
                
                # Update progress
                self.progress_bar.value = index + 1
                self.status_label.text = f"Status: Generated {index+1}/{len(self.student_names)}"
            
            print(f"\n✅ All {len(self.student_names)} QR codes generated!")
            print(f"📁 Location: {self.qr_folder}\n")
            
            self.status_label.text = f"Status: ✅ Completed! Generated {len(self.student_names)} QR codes"
            self.main_window.info_dialog("Success", 
                f"✅ Generated {len(self.student_names)} QR codes successfully!\n\n"
                f"Location: {self.qr_folder}")
            
            self.generate_btn.enabled = True
        
        except Exception as e:
            print(f"❌ Error generating QR codes: {e}")
            self.main_window.error_dialog("Error", f"Failed to generate QR codes: {e}")
            self.generate_btn.enabled = True
    
    def open_qr_folder(self, widget):
        """Open QR folder"""
        try:
            import subprocess
            if os.name == 'nt':
                subprocess.Popen(f'explorer "{self.qr_folder}"')
            else:
                subprocess.Popen(['open', str(self.qr_folder)])
        except Exception as e:
            self.main_window.error_dialog("Error", f"Cannot open folder: {e}")


def main():
    return QRGenerator('QR Code Generator', 'org.dralfredroda.qrgenerator')


if __name__ == '__main__':
    main().main_loop()
