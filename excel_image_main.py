import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.table import Table
import os
import sys
from matplotlib.patches import Rectangle
import re

def clean_sheet_name(sheet_name):
    """Clean the sheet name by removing digits and special symbols after the first string."""
    # Find the first word (string) and remove any digits or special symbols that follow
    match = re.match(r"^[^\d\W]+", sheet_name)
    if match:
        return match.group(0)  # Return the cleaned string
    return sheet_name
def format_time(value):
    """Convert numeric value to time format HH:MM AM/PM."""
    try:
        if pd.notna(value):
            value = str(int(value))  # Ensure it's a string for processing
            if len(value) == 4 and value.isdigit():
                hours = int(value[:2])
                minutes = value[2:]
                period = "AM" if hours < 12 else "PM"
                if hours == 0:
                    hours = 12
                elif hours > 12:
                    hours -= 12
                return f"{hours:02}:{minutes} {period}"
            elif len(value) == 2 and value.isdigit():  # Handle cases like '09' for '9 AM'
                hours = int(value)
                period = "AM" if hours < 12 else "PM"
                if hours == 0:
                    hours = 12
                elif hours > 12:
                    hours -= 12
                return f"{hours:02}:00 {period}"
    except ValueError:
        print(f"Error formatting value: {value}")
    return value

def get_current_time_formatted():
    """Get the current time formatted as HH:MM AM/PM."""
    return time.strftime("%I:%M %p")

def add_watermark(ax, watermark_text, fontsize=0, opacity=0.3):
    """
    Add a vertical watermark to the given axes.
    
    :param ax: The matplotlib axes to add the watermark to.
    :param watermark_text: The text for the watermark.
    :param fontsize: The font size of the watermark text.
    :param opacity: The opacity of the watermark text (0.0 to 1.0).
    """
    fig = ax.figure
    fig.text(0.5, 0.5, watermark_text, fontsize=fontsize, color='gray', 
             ha='center', va='center', rotation=45, alpha=opacity, 
             transform=fig.transFigure, zorder=10, fontweight='bold')


def excel_to_table_image(excel_file, output_dir):
    start_time = time.time()  # Start timing the processing
    image_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(excel_file))[0]}.png")
 
 
    # Remove the previous output image or any other relevant file before starting the process
    if os.path.exists(image_file):
        os.remove(image_file)
        
    # Remove the existing image file if it exists
    if os.path.exists(image_file):
        os.remove(image_file)

    try:
        # Retry mechanism
        max_retries = 5
        for attempt in range(max_retries):
            try:
                df = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
                break  # If successful, break out of the retry loop
            except PermissionError as e:
                if attempt < max_retries - 1:
                    print(f"PermissionError: {e}. Retrying in 5 seconds...")
                    time.sleep(5)  # Wait for 5 seconds before retrying
                else:
                    print(f"Failed to read {excel_file} after {max_retries} attempts.")
                    return
            except OSError as e:
                if "used by another process" in str(e):
                    print(f"File is being used by another process. Retrying in 5 seconds...")
                    time.sleep(5)  # Wait for 5 seconds before retrying
                else:
                    print(f"Error reading {excel_file}: {e}")
                    os.remove(excel_file)
                    return
            except Exception as e:
                print(f"Error reading {excel_file}: {e}")
                os.remove(excel_file)
                return
            
    
        # Check if the number of sheets is less than the minimum required
        if len(df) <= 3:
            os.remove(excel_file)
            print(f"{excel_file} has fewer than 3 sheets. Skipping conversion.")
            return
        if not df:
            print(f"{excel_file} is empty.")
            return


        # Check "Time" column in sheets 2, 3, 4
        current_time = get_current_time_formatted()
        time_match_found = False

        for sheet_index in [1, 2, 3]:
            if sheet_index >= len(df):
                continue
            
            sheet_name, data = list(df.items())[sheet_index]
            
            if "Time" not in data.columns:
                continue
            
            print(f"Processing sheet: {sheet_name}")

             # Convert rows 3 through 8 in the "Time" column to formatted time
            time_column_index = data.columns.get_loc('Time')
            for i in range(2, 20):  # Rows 3 to 8 (index 2 to 7)
                if i < len(data):
                    data.iloc[i, time_column_index] = format_time(data.iloc[i, time_column_index])


            # Check if first row is empty and only process the second row if so
            if pd.isna(data.iloc[0]).all():
                if len(data) > 1:
                    second_row_time = data['Time'].iloc[1]  # Get the time value from the second row
                    formatted_time = format_time(second_row_time)
                    
                    print(f"Second row time value: {second_row_time}, Formatted time: {formatted_time}")
                    
                    if formatted_time == current_time:
                        time_match_found = True
                        break
            else:
                print(f"Sheet {sheet_name} has data in the first row. Skipping...")
                continue
        
        if not time_match_found:
            os.remove(excel_file)
            print(f"No matching time found in {excel_file}. Removed file.")
            return


        num_sheets = len(df)
        max_cols = max(len(data.columns) for data in df.values())
        max_rows = min(max(len(data) for data in df.values()), 20)

        fig_width = max(12, 1 + 0.2 * max_cols)
        fig_height = 4 * num_sheets + 0.3 * max_rows+2

        fig, axs = plt.subplots(num_sheets, 1, figsize=(fig_width, fig_height))
        if num_sheets == 1:
            axs = [axs]

        fig.patch.set_facecolor('black')

        # Adjust figure layout to add space at the top
        fig.subplots_adjust(top=1.00)  # Increase the top margin

        for sheet_index, (ax, (sheet_name, data)) in enumerate(zip(axs, df.items())):
            ax.axis('off')

            # Clean the first sheet name
            if sheet_index == 0:
                sheet_name = clean_sheet_name(sheet_name)

            # Check if the sheet contains any data
            if data.empty:
                print(f"Sheet {sheet_name} has no data. Skipping...")
                continue

            # Determine the background color for the title based on the sheet name
            if "3 Min" in sheet_name or "5 Min" in sheet_name or "15 Min" in sheet_name:
                title_bg_color = '#FFA500' if "3 Min" in sheet_name else '#008000' if "5 Min" in sheet_name else '#0000FF'
                title_font_color = 'black'
                title_font_size = 25
            else:
                title_bg_color = 'black'
                title_font_color = '#ffffff'
                title_font_size = 14

            ax.add_patch(Rectangle((0, 1.02), 1, 0.1, transform=ax.transAxes, color=title_bg_color, zorder=1, clip_on=False))

            if sheet_index == 0:
                ax.text(0.05, 1.05, sheet_name, fontsize=title_font_size, fontweight='bold', color=title_font_color, ha='left', va='bottom',
                        bbox=dict(facecolor=title_bg_color, edgecolor='none', boxstyle='round,pad=0'))
                ax.text(0.5, 1.00, "USE DATA AFTER 10:30 AM", fontsize=12, fontweight='bold', color='yellow', ha='center', va='bottom',
                        bbox=dict(facecolor='black', edgecolor='none', boxstyle='round,pad=0'))
            else:
                ax.text(0.5, 1.02, sheet_name, fontsize=title_font_size, fontweight='bold', color=title_font_color, ha='center', va='bottom',
                        bbox=dict(facecolor=title_bg_color, edgecolor='none', boxstyle='round,pad=0'))

            data = data.head(20)
            


            if sheet_index in [1, 2, 3] and 'Time' in data.columns:
                if len(data) > 1:
                    time_column_index = data.columns.get_loc('Time')
                    # Convert only the second row's time value to formatted time
                    data.iloc[1, time_column_index] = format_time(data.iloc[1, time_column_index])

            table = Table(ax, bbox=[0, 0, 1, 1])
            nrows, ncols = data.shape
            if ncols == 0 or nrows == 0:
                print(f"Sheet {sheet_name} in {excel_file} has no data.")
                continue

            width = 1.0 / ncols
            height = 1.0 / (nrows + 1)

            header_color = '#333333'
            header_font_color = '#ffffff'

            for i, column in enumerate(data.columns):
                table.add_cell(0, i, width, height, text=column, facecolor=header_color,
                            edgecolor='#444444', loc='center')
                cell = table[(0, i)]
                cell._text.set_fontsize(10)
                cell._text.set_color(header_font_color)
                cell._text.set_weight('bold')

            # Initialize median_row_index
            median_row_index = None

            # Clean and process the 5th column for median calculation
            if sheet_index == 0 and ncols > 5:
                # Remove percentage signs and convert to numeric
                data.iloc[:, 5] = data.iloc[:, 5].astype(str).str.replace('%', '', regex=False).replace('-', '', regex=False)
                data.iloc[:, 5] = pd.to_numeric(data.iloc[:, 5], errors='coerce')
                median_value = data.iloc[:, 5].median()
                median_row = data[data.iloc[:, 5] == median_value].index.tolist()

                if median_row:
                    median_row_index = median_row[0]

            # Find maximum values for specified columns
            max_values = {}
            if sheet_index == 0 and ncols > 1:
                if len(data.columns) > 7:
                    max_values[1] = data.iloc[:, 1].max()
                    max_values[7] = data.iloc[:, 7].max()



            # Add Data Rows
            for i, row in enumerate(data.values):
                for j, cell_value in enumerate(row):
                    font_color = '#cccccc'  # Default light gray font color

                    # Alternate background colors for rows
                    if i % 2 == 0:
                        bg_color = 'black'  # Black background color for even rows
                    else:
                        bg_color = '#333333'  # Gray background color for odd rows

                    if 'Total' in str(cell_value) and sheet_index == 0:
                        if j in [5, 6, 7, 8]:
                            font_color = '#FFFF00'  # Yellow font color
                        bg_color = '#FFFF00'  # Yellow background color for the "Total" row
                        font_color = '#000000'  # Black font color for the "Total" row
                        # Set font color for the entire row to yellow
                        for k in range(len(row)):
                            try:
                                table[(i + 1, k)]._text.set_color('#FFFF00')
                            except KeyError:
                                pass  # Skip if the cell does not exist

                    # Determine font color based on specific conditions
                    if sheet_index == 0:
                        # First sheet: Green font color for columns 2 and 6, red for negative numbers
                        if j == 2 or j == 6:
                            if isinstance(cell_value, (int, float)) and cell_value < 0:
                                font_color = '#ff0000'  # Red font color for negative numbers
                            else:
                                font_color = '#00ff00'  # Green font color
                    else:
                        # Other sheets: Green font color for columns 3, 4, 5, and 8, red for negative numbers
                        if j in [3, 4, 5, 8]:
                            if isinstance(cell_value, (int, float)) and cell_value < 0:
                                font_color = '#ff0000'  # Red font color for negative numbers
                            else:
                                font_color = '#00ff00'  # Green font color

                    # Handle empty cells
                    if pd.isna(cell_value):
                        cell_value = ''  # Keep the cell empty

                    # Highlight "sell" keyword in non-first sheets
                    if sheet_index != 0 and (j == 5 or j == 8) and isinstance(cell_value, str) and 'sell' in cell_value.lower() and 'SELL' in cell_value.upper():
                        font_color = '#ff0000'  # Red font color

                    # Highlight the corresponding value in the 4th column of the row with the median percentage value
                    if sheet_index == 0 and median_row_index is not None and i == median_row_index and j == 4:
                        bg_color = '#FFFF00'  # Yellow background color
                        font_color = '#000000'  # Black font color

                    # Highlight the maximum values in specified columns
                    if sheet_index == 0 and j in max_values:
                        if cell_value == max_values[j]:
                            font_color = '#FFFF00'  # Yellow background color

                    table.add_cell(i + 1, j, width, height, text=str(cell_value),
                                facecolor=bg_color, edgecolor='#444444', loc='center')

                    cell = table[(i + 1, j)]
                    cell._text.set_fontsize(10)
                    cell._text.set_color(font_color)

            ax.add_table(table)

        add_watermark(ax, "Sample Data", fontsize=90, opacity=0.5)

    # Save the figure only if there are any valid sheets
        if any(ax.get_children() for ax in axs):
            plt.savefig(image_file, bbox_inches='tight', pad_inches=0.1, dpi=300, transparent=True)
            print(f"Saved image for {excel_file} to {image_file}")
            os.remove(excel_file)
        else:
            print(f"No valid sheets with data to save for {excel_file}. Removing empty image file if exists.")
            if os.path.exists(image_file):
                os.remove(image_file)
    except Exception as e:
        print(f"Error processing {excel_file}: {e}")

    plt.close()
    

    end_time = time.time()  #                End timing the processing
    processing_time = end_time - start_time

    # Save execution time to the second sheet of the Excel file
    if len(df) > 1:
        sheet_2 = list(df.items())[1][1]  # Second sheet data
        sheet_2_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(excel_file))[0]}_sheet_2.xlsx")
        with pd.ExcelWriter(sheet_2_file, engine='openpyxl') as writer:
            sheet_2.to_excel(writer, index=False, sheet_name='Sheet2')
            # Write execution time
            exec_time_df = pd.DataFrame({'Execution Time': [f"{processing_time:.2f} seconds"]})
            exec_time_df.to_excel(writer, index=False, startrow=len(sheet_2) + 2, sheet_name='Sheet2')

    print(f"Processed {excel_file} in {processing_time:.2f} seconds")

    return image_file


import pandas as pd
import os

def convert_to_time_format(cell_value):
    try:
        time_value = pd.to_datetime(cell_value, format='%H:%M:%S').time()
        return time_value
    except Exception as e:
        return cell_value


import time

def process_directory_continuously(input_dir, output_dir, run_duration= 5 * 3600, check_interval=10):
    """
    Process the directory continuously for a specified duration and convert any new Excel files to images.
    
    :param input_dir: Directory to monitor for new Excel files.
    :param output_dir: Directory to save the converted images.
    :param run_duration: Duration to run the process (in seconds). Default is 1 hour.
    :param check_interval: Interval (in seconds) to check for new files. Default is 10 seconds.
    """
    start_time = time.time()
    processed_files = set()

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    while time.time() - start_time < run_duration:
        for excel_file in os.listdir(input_dir):
            if excel_file.endswith('.xlsx') and excel_file not in processed_files:
                excel_path = os.path.join(input_dir, excel_file)
                image_file = excel_to_table_image(excel_path, output_dir)
                if image_file:
                    print(f"Processed and saved image: {image_file}")
                    processed_files.add(excel_file)

        time.sleep(check_interval)  # Wait before checking for new files again

    print("Processing complete.")

if __name__ == "__main__":
    if len(sys.argv) == 3:
        # Command-line mode for directory processing
        input_dir = sys.argv[1]
        output_dir = sys.argv[2]
        process_directory_continuously(input_dir, output_dir)
    else:
        print("Usage:")
        print("For directory: python excel_image.py <input_dir> <output_dir>")
        sys.exit(1)