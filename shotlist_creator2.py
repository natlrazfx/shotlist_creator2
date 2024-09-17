# Natalia Raz
# Shotlist Creator for DaVinci Resolve Studio

import os
import platform
import subprocess
import time
import DaVinciResolveScript as dvr_script
import xlsxwriter
from PySide6 import QtWidgets, QtCore, QtGui
from PIL import Image
from pynput.keyboard import Controller

# Connect to the currently running instance of Resolve
resolve = dvr_script.scriptapp("Resolve")

# Create an instance of the keyboard controller
keyboard = Controller()

# Function to get the save file name from the user
def get_save_file_name(project_name):
    app = QtWidgets.QApplication.instance()
    if not app:
        app = QtWidgets.QApplication([])
    options = QtWidgets.QFileDialog.Options()
    default_filename = f"{project_name}_shotlist_v001.xlsx" if project_name else ""
    file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
        None,
        "Save As",
        default_filename,
        "Excel Files (*.xlsx);;All Files (*)",
        options=options,
    )
    return file_name

# Function to handle folder/file replacement or renaming
def ask_replace_or_rename(file_or_folder):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Question)
    msgBox.setText(f"'{file_or_folder}' already exists. What would you like to do?")
    msgBox.setWindowTitle("File/Folder Exists")
    replace_button = msgBox.addButton("Replace", QtWidgets.QMessageBox.AcceptRole)
    rename_button = msgBox.addButton("Rename", QtWidgets.QMessageBox.NoRole)
    cancel_button = msgBox.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
    msgBox.setDefaultButton(replace_button)

    msgBox.exec()

    if msgBox.clickedButton() == replace_button:
        return "replace"
    elif msgBox.clickedButton() == rename_button:
        return "rename"
    else:
        return "cancel"

# Function to create a subfolder, handling conflicts with existing folders
def ask_create_subfolder(output_path, file_name):
    subfolder_name = os.path.splitext(file_name)[0]
    subfolder_path = os.path.join(output_path, subfolder_name)

    while True:
        if os.path.exists(subfolder_path):
            action = ask_replace_or_rename(subfolder_name)
            if action == "replace":
                # Delete the existing folder if replacing
                for root, dirs, files in os.walk(subfolder_path, topdown=False):
                    for file in files:
                        os.remove(os.path.join(root, file))
                    for dir in dirs:
                        os.rmdir(os.path.join(root, dir))
                break
            elif action == "rename":
                app = QtWidgets.QApplication.instance()
                if not app:
                    app = QtWidgets.QApplication([])
                new_name, ok = QtWidgets.QInputDialog.getText(
                    None,
                    "Rename",
                    "Enter new name for the folder and file:",
                    text=subfolder_name,
                )
                if ok and new_name:
                    subfolder_path = os.path.join(output_path, new_name)
                    file_name = f"{new_name}.xlsx"
                    subfolder_name = new_name  # Update subfolder_name for accurate display
                else:
                    return None, None  # Cancel operation
            else:
                return None, None  # Cancel operation
        else:
            os.makedirs(subfolder_path)
            break

    return subfolder_path, file_name

# Function to get the appropriate color format for Excel cells
def get_color_format(workbook, color_name):
    color_map = {
        "Rose": "#FF007F",
        "Pink": "#FFC0CB",
        "Lavender": "#E6E6FA",
        "Cyan": "#00FFFF",
        "Fuchsia": "#FF00FF",
        "Mint": "#98FF98",
        "Sand": "#C2B280",
        "Yellow": "#FFFF00",
        "Green": "#00FF00",
        "Blue": "#0000FF",
        "Purple": "#800080",
        "Red": "#FF0000",
        "Cocoa": "#D2691E",
        "Sky": "#87CEEB",
        "Lemon": "#FFF44F",
        "Cream": "#FFFDD0",
    }
    hex_color = color_map.get(color_name, "#FFFFFF")  # Default to white if color is not found
    return workbook.add_format(
        {"bg_color": hex_color, "valign": "vcenter", "align": "center"}
    )

# Function to open the output folder in Explorer (Windows) or Finder (macOS)
def open_folder_in_explorer(output_path):
    if platform.system() == "Windows":
        subprocess.Popen(["explorer", os.path.normpath(output_path)])
    elif platform.system() == "Darwin":  # macOS
        subprocess.Popen(["open", output_path])
    else:
        print("Unsupported OS for auto-opening folder.")

# Function to focus back on the timeline (edit page) before keyboard presses
def focus_on_timeline():
    app = QtWidgets.QApplication.instance()
    if not app:
        app = QtWidgets.QApplication([])
    # Ensure Resolve is brought to the front
    if platform.system() == "Darwin":  # macOS
        subprocess.run(
            ["osascript", "-e", 'tell application "DaVinci Resolve" to activate']
        )
    elif platform.system() == "Windows":
        resolve_window_name = "DaVinci Resolve"  # Replace with the actual window name if necessary
        subprocess.run(f'cmdow "{resolve_window_name}" /ACT', shell=True)

# Function to export markers and create an Excel file
def export_markers(timeline, output_path, timecodes, excel_filename, metadata_list, selected_fields, image_size):
    markers = timeline.GetMarkers()
    workbook = xlsxwriter.Workbook(os.path.join(output_path, excel_filename))
    worksheet = workbook.add_worksheet()

    # Set up cell formats for text
    text_format = workbook.add_format({"valign": "vcenter", "align": "left"})

    # Set image size based on user selection
    max_size = image_size

    # Prepare headers
    headers = selected_fields + ["Still"]

    # Write header row
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, text_format)

    row = 1

    # Get the current gallery and still album
    gallery = currentProject.GetGallery()
    currentStillAlbum = gallery.GetCurrentStillAlbum()

    # Get all the stills in the current still album
    stills = currentStillAlbum.GetStills()

    for idx, (frame, marker) in enumerate(markers.items()):
        col = 0
        # Prepare data row
        data_row = {}
        # Add standard fields
        data_row["Frame"] = frame
        data_row["Name"] = marker["name"]
        data_row["Note"] = marker["note"]
        data_row["Duration"] = marker["duration"]
        data_row["Color"] = marker["color"]
        data_row["Timecode"] = timecodes[idx]

        # Add metadata
        clip_metadata = metadata_list[idx]
        data_row.update(clip_metadata)

        # Write selected fields to worksheet
        for field in selected_fields:
            if field == "Color":
                worksheet.write(
                    row, col, "", get_color_format(workbook, data_row.get(field, ""))
                )
            else:
                worksheet.write(row, col, data_row.get(field, ""), text_format)
            col += 1
        row += 1

    # Adjust column index for the image based on added columns
    image_col_index = len(headers) - 1  # Since headers include 'Still' at the end

    row = 1

    # Loop through the stills and export them using a unique temporary file name for each still
    for i, still in enumerate(stills):
        # Generate a unique suffix for the temporary file name
        suffix = f"{i + 1:03}"
        tmp_name = f"tmp_{suffix}"
        # Export the still to the output folder using PNG format and the unique temporary file name
        currentStillAlbum.ExportStills([still], output_path, tmp_name, "png")

    # Create a list of file names in the folder, sorted by name
    files = sorted(os.listdir(output_path))

    for file in files:
        if file.startswith("tmp") and file.endswith(".png"):
            # Extract the file number from the file name
            file_number = int(file.split("_")[1])
            new_name = f"thumb{file_number:03d}.png"

            while os.path.exists(os.path.join(output_path, new_name)):
                action = ask_replace_or_rename(new_name)
                if action == "replace":
                    os.remove(os.path.join(output_path, new_name))
                elif action == "rename":
                    new_name, ok = QtWidgets.QInputDialog.getText(
                        None, "Rename", "Enter new name for the image:", text=new_name
                    )
                    if not ok or not new_name:
                        return
                else:
                    return

            # Rename the file
            os.rename(
                os.path.join(output_path, file), os.path.join(output_path, new_name)
            )

            # Embed the still in the worksheet
            image_file_path = os.path.normpath(os.path.join(output_path, new_name))
            print(image_file_path)

            # Open the image and get its size
            image = Image.open(image_file_path)
            width, height = image.size

            # Calculate the new size (maintaining aspect ratio)
            if width > height:
                new_width = int(max_size)
                new_height = int((max_size / width) * height)
            else:
                new_height = int(max_size)
                new_width = int((max_size / height) * width)

            # Resize the image
            resized_image = image.resize((new_width, new_height))

            # Save the resized image back to the file
            resized_image.save(image_file_path)

            # Insert the resized image into the worksheet
            worksheet.insert_image(
                row,
                image_col_index,
                image_file_path,
                {"x_scale": 1, "y_scale": 1, "object_position": 1},
            )

            # Set the width and height of the cell to match the image's aspect ratio
            worksheet.set_column(
                image_col_index, image_col_index, new_width / 6
            )
            worksheet.set_row(
                row, new_height / 1.33
            )

            row += 1

    # Autofit the worksheet columns
    worksheet.autofit()

    workbook.close()

# Function to set a dark theme for the application
def set_dark_theme(app):
    # Set the Fusion style
    app.setStyle('Fusion')
    # Now set a dark palette
    dark_palette = QtGui.QPalette()

    # Set the background color
    dark_color = QtGui.QColor(45, 45, 45)
    disabled_color = QtGui.QColor(127, 127, 127)

    dark_palette.setColor(QtGui.QPalette.Window, dark_color)
    dark_palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(18, 18, 18))
    dark_palette.setColor(QtGui.QPalette.AlternateBase, dark_color)
    dark_palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, disabled_color)
    dark_palette.setColor(QtGui.QPalette.Button, dark_color)
    dark_palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, disabled_color)
    dark_palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    dark_palette.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Highlight, QtGui.QColor(80, 80, 80))
    dark_palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.HighlightedText, disabled_color)

    app.setPalette(dark_palette)

# Class for the combined user input dialog
class UserInputDialog(QtWidgets.QDialog):
    def __init__(self, all_fields, parent=None):
        super(UserInputDialog, self).__init__(parent)
        self.setWindowTitle("Shotlist Creator Options")

        layout = QtWidgets.QVBoxLayout(self)

        instructions = QtWidgets.QLabel("Please set the options below:")
        layout.addWidget(instructions)

        # Timecode input
        timecode_label = QtWidgets.QLabel("Enter custom timecode (default is 01:00:00:00):")
        layout.addWidget(timecode_label)

        self.timecode_input = QtWidgets.QLineEdit()
        self.timecode_input.setText("01:00:00:00")
        layout.addWidget(self.timecode_input)

        # Checkbox for deleting stills
        self.delete_stills_checkbox = QtWidgets.QCheckBox("Delete all stills from the gallery album")
        layout.addWidget(self.delete_stills_checkbox)

        # Metadata selection
        metadata_label = QtWidgets.QLabel("Select the metadata fields to include:")
        layout.addWidget(metadata_label)

        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QtWidgets.QWidget()
        scroll_layout = QtWidgets.QVBoxLayout(scroll_content)

        # Default selected fields and order
        default_selected_fields = [
            "Frame", "Timecode", "Name", "Note", "Duration", "Color",
            "Clip Name", "FPS", "File Path", "Video Codec",
            "Resolution", "Start TC", "End TC"
        ]

        # Create checkboxes for all metadata fields
        self.checkboxes = []
        for key in all_fields:
            checkbox = QtWidgets.QCheckBox(key)
            if key in default_selected_fields:
                checkbox.setChecked(True)
            scroll_layout.addWidget(checkbox)
            self.checkboxes.append(checkbox)

        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        # Select All and Deselect All buttons
        button_layout = QtWidgets.QHBoxLayout()
        select_all_button = QtWidgets.QPushButton("Select All")
        deselect_all_button = QtWidgets.QPushButton("Deselect All")
        button_layout.addWidget(select_all_button)
        button_layout.addWidget(deselect_all_button)
        layout.addLayout(button_layout)

        # Dropdown for selecting image size
        size_label = QtWidgets.QLabel("Choose the size for the still images:")
        layout.addWidget(size_label)

        size_layout = QtWidgets.QHBoxLayout()
        self.size_combo = QtWidgets.QComboBox()
        self.size_combo.addItems(["SMALL", "LARGE", "CUSTOM"])
        size_layout.addWidget(self.size_combo)

        # Input for custom size (hidden by default)
        self.custom_size_input = QtWidgets.QDoubleSpinBox()
        self.custom_size_input.setRange(0.1, 10)
        self.custom_size_input.setSingleStep(0.1)
        self.custom_size_input.setValue(1)
        self.custom_size_input.setVisible(False)
        size_layout.addWidget(self.custom_size_input)

        layout.addLayout(size_layout)

        # Show custom size input when "CUSTOM" is selected
        def on_size_change():
            self.custom_size_input.setVisible(self.size_combo.currentText() == "CUSTOM")

        self.size_combo.currentIndexChanged.connect(on_size_change)

        # OK and Cancel buttons
        ok_cancel_layout = QtWidgets.QHBoxLayout()
        ok_button = QtWidgets.QPushButton("OK")
        cancel_button = QtWidgets.QPushButton("Cancel")
        ok_cancel_layout.addWidget(ok_button)
        ok_cancel_layout.addWidget(cancel_button)
        layout.addLayout(ok_cancel_layout)

        # Functions to select/deselect all
        def select_all():
            for checkbox in self.checkboxes:
                checkbox.setChecked(True)

        def deselect_all():
            for checkbox in self.checkboxes:
                checkbox.setChecked(False)

        select_all_button.clicked.connect(select_all)
        deselect_all_button.clicked.connect(deselect_all)

        # Function to handle OK and Cancel
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

    def get_values(self):
        selected_fields = [checkbox.text() for checkbox in self.checkboxes if checkbox.isChecked()]
        size = self.size_combo.currentText()
        if size == "CUSTOM":
            multiplier = self.custom_size_input.value()
            image_size = 260 * multiplier
        elif size == "LARGE":
            image_size = 520
        else:
            image_size = 260  # SMALL
        timecode = self.timecode_input.text() if self.timecode_input.text() else "01:00:00:00"
        delete_stills = self.delete_stills_checkbox.isChecked()
        return selected_fields, image_size, timecode, delete_stills

if __name__ == "__main__":
    # Create a Qt Application instance
    app = QtWidgets.QApplication.instance()
    if not app:
        app = QtWidgets.QApplication([])
    # Set dark theme
    set_dark_theme(app)

    # Get the current project and timeline
    projectManager = resolve.GetProjectManager()
    currentProject = projectManager.GetCurrentProject()
    currentTimeline = currentProject.GetCurrentTimeline()

    # Get the project name
    project_name = currentProject.GetName()

    # Get all metadata keys from the first clip (assuming all clips have the same keys)
    markers = currentTimeline.GetMarkers()

    # Initialize a list to store metadata
    metadata_list = []

    # Get the clip at the current playhead position to extract metadata keys
    current_clip = currentTimeline.GetCurrentVideoItem()
    if current_clip:
        media_pool_item = current_clip.GetMediaPoolItem()
        if media_pool_item:
            clip_properties = media_pool_item.GetClipProperty()
            # Include all properties
            all_metadata_keys = list(clip_properties.keys())
        else:
            all_metadata_keys = []
    else:
        all_metadata_keys = []

    # Add standard fields to metadata keys
    standard_fields = ["Frame", "Timecode", "Name", "Note", "Duration", "Color"]
    all_fields = standard_fields + all_metadata_keys

    # Remove duplicates while preserving order
    seen = set()
    all_fields = [x for x in all_fields if not (x in seen or seen.add(x))]

    # Show the combined user input dialog
    dialog = UserInputDialog(all_fields)
    if dialog.exec() == QtWidgets.QDialog.Accepted:
        selected_fields, image_size, timecode_to_set, delete_stills = dialog.get_values()

        if not selected_fields:
            print("No metadata fields selected.")
        else:
            # Set the timecode
            currentTimeline.SetCurrentTimecode(timecode_to_set)

            # Check if the user wants to delete the stills
            if delete_stills:
                # Switch to the color page for stills deletion
                resolve.OpenPage("color")

                # Get the current gallery and still album
                gallery = currentProject.GetGallery()
                currentStillAlbum = gallery.GetCurrentStillAlbum()

                # Get all the stills in the album
                stills = currentStillAlbum.GetStills()

                if stills:
                    success = currentStillAlbum.DeleteStills(stills)
                    if success:
                        print("All stills have been successfully deleted.")
                    else:
                        print("Failed to delete stills.")
                else:
                    print("No stills found in the album.")

            # Proceed with other actions, such as setting metadata and exporting stills

            # Initialize timecodes list to store the timecodes
            timecodes = []

            # Initialize a list to store metadata
            metadata_list = []

            # Ensure focus is back on the timeline/edit page before keyboard pressing
            focus_on_timeline()

            # Get the markers from the timeline
            markers = currentTimeline.GetMarkers()

            # Loop through each marker on the timeline
            for i, (frame_id, marker) in enumerate(markers.items()):
                # Calculate the number of markers until the end of the timeline
                numMarkersToEnd = len(markers) - (i + 1)

                # Print the number of markers until the end of the timeline
                print("Number of markers until the end of the timeline:", numMarkersToEnd)

                # Send a "0" key press event to move to the next marker
                keyboard.press("0")
                keyboard.release("0")

                # Wait for the playhead to move to the next marker
                time.sleep(0.2)

                # Get the current timecode
                currentTimecode = currentTimeline.GetCurrentTimecode()

                # Append the current timecode to the list of timecodes
                timecodes.append(currentTimecode)

                # Grab a still from the current video clip
                galleryStill = currentTimeline.GrabStill()

                # Get the clip at the current playhead position
                current_clip = currentTimeline.GetCurrentVideoItem()

                # Extract metadata from the clip
                clip_metadata = {}
                if current_clip:
                    clip_metadata["Clip Name"] = current_clip.GetName()
                    media_pool_item = current_clip.GetMediaPoolItem()
                    if media_pool_item:
                        clip_properties = media_pool_item.GetClipProperty()
                        # Include all properties
                        for key, value in clip_properties.items():
                            clip_metadata[key] = value
                    else:
                        # If media_pool_item is None, set properties to "N/A"
                        clip_metadata["Clip Name"] = current_clip.GetName()
                else:
                    # If current_clip is None, set all metadata to "N/A"
                    clip_metadata["Clip Name"] = "N/A"

                # Append the metadata to the list
                metadata_list.append(clip_metadata)

                # Check if this is the last marker on the timeline
                if numMarkersToEnd == 0:
                    break

            # Use QFileDialog.getSaveFileName to select both filename and location
            full_path = get_save_file_name(project_name)

            if not full_path:
                print("No output folder and filename selected.")
            else:
                output_path, excel_filename = os.path.split(full_path)
                if not excel_filename.endswith(".xlsx"):
                    excel_filename += ".xlsx"

                # Ask if a subfolder should be created
                output_path, excel_filename = ask_create_subfolder(output_path, excel_filename)

                if output_path:
                    # Export markers and embed stills in the specified Excel file
                    export_markers(
                        currentTimeline, output_path, timecodes, excel_filename, metadata_list, selected_fields, image_size
                    )
                    print("DONE")
                    # Open the output directory in Finder or Explorer
                    open_folder_in_explorer(output_path)
    else:
        print("Operation cancelled.")

    # End of the script
