from flask import Flask, request, render_template, send_file, session, redirect, url_for
import pandas as pd
import io
import os
import openpyxl
import time
import numpy as np

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Change this to a random secret key

# Columns to compare between uploads
columns_to_compare = ["Sec Faculty Info", "Sec All Faculty Last Names", "Total FTE", "FTE Count"]

# Alert thresholds
FTE_CHANGE_THRESHOLD = 0.10  # 10% change threshold
FACULTY_COLUMNS = ["Sec Faculty Info", "Sec All Faculty Last Names"]
CRITICAL_COLUMNS = ["Sec Faculty Info", "Total FTE", "FTE Count"]

def check_fte_change_alert(old_val, new_val, identifier):
    """Check if FTE change exceeds threshold."""
    try:
        old_val = float(old_val) if pd.notna(old_val) else 0
        new_val = float(new_val) if pd.notna(new_val) else 0
        
        if old_val == 0:
            return None
        
        change_pct = abs(new_val - old_val) / old_val
        if change_pct > FTE_CHANGE_THRESHOLD:
            return {
                'severity': 'warning',
                'type': 'Large FTE Change',
                'identifier': identifier,
                'description': f'FTE changed by {change_pct*100:.1f}%',
                'old_value': old_val,
                'new_value': new_val
            }
    except (ValueError, TypeError):
        pass
    return None

def check_capacity_alert(row, identifier):
    """Check if FTE Count exceeds Capacity."""
    try:
        if 'FTE Count' in row and 'Capacity' in row:
            fte_count = float(row['FTE Count']) if pd.notna(row['FTE Count']) else 0
            capacity = float(row['Capacity']) if pd.notna(row['Capacity']) else 0
            
            if capacity > 0 and fte_count > capacity:
                return {
                    'severity': 'critical',
                    'type': 'Capacity Exceeded',
                    'identifier': identifier,
                    'description': f'FTE Count ({fte_count}) exceeds Capacity ({capacity})',
                    'old_value': capacity,
                    'new_value': fte_count
                }
    except (ValueError, TypeError):
        pass
    return None

def check_missing_data_alert(row, identifier):
    """Check for missing critical data in new rows."""
    missing_fields = []
    for col in CRITICAL_COLUMNS:
        if col in row and (pd.isna(row[col]) or str(row[col]).strip() == ''):
            missing_fields.append(col)
    
    if missing_fields:
        return {
            'severity': 'critical',
            'type': 'Missing Critical Data',
            'identifier': identifier,
            'description': f'Missing data in: {", ".join(missing_fields)}',
            'old_value': '',
            'new_value': ''
        }
    return None

def check_faculty_change_alert(old_val, new_val, identifier, col_name):
    """Check if faculty information changed."""
    if str(old_val).strip() != str(new_val).strip():
        return {
            'severity': 'warning',
            'type': 'Faculty Change',
            'identifier': identifier,
            'description': f'{col_name} changed',
            'old_value': str(old_val).strip(),
            'new_value': str(new_val).strip()
        }
    return None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filename = file.filename.lower()
            
            # Read the new file (CSV or Excel)
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                df = pd.read_excel(file)
            else:
                df = pd.read_csv(file)
            
            df.columns = df.columns.str.strip()

            # Save the new data temporarily
            df.to_csv("temp_new_data.csv", index=False)
            
            # Store in session that we have new data
            session['has_new_data'] = True
            session['new_data_preview'] = df.head(10).to_html(classes='data')
            
            return redirect(url_for('select_master'))

    return render_template('index.html')

@app.route('/select-master', methods=['GET', 'POST'])
def select_master():
    if not session.get('has_new_data'):
        return redirect(url_for('upload_file'))
    
    if request.method == 'POST':
        # Start timing
        start_time = time.time()
        
        master_file = request.files['master_file']
        if master_file:
            master_filename = master_file.filename.lower()
            
            # Load the master file (CSV or Excel)
            if master_filename.endswith('.xlsx') or master_filename.endswith('.xls'):
                is_excel = True
                excel_file_path = os.path.join(os.path.dirname(__file__), "temp_master.xlsx")
                master_file.save(excel_file_path)
                master_df = pd.read_excel(excel_file_path)
            else:
                is_excel = False
                excel_file_path = None
                master_df = pd.read_csv(master_file)
            
            master_df.columns = master_df.columns.str.strip()
            
            # Load the new data
            new_df = pd.read_csv("temp_new_data.csv")
            new_df.columns = new_df.columns.str.strip()
            
            # Check if files have a common key column (try common column names)
            key_column = None
            possible_keys = ['Sec Name', 'Section Name', 'ID', 'Section', 'Name']
            
            for key in possible_keys:
                if key in new_df.columns and key in master_df.columns:
                    key_column = key
                    break
            
            if not key_column:
                # If no common key found, use first column
                if len(new_df.columns) > 0 and len(master_df.columns) > 0:
                    if new_df.columns[0] == master_df.columns[0]:
                        key_column = new_df.columns[0]
            
            if not key_column:
                return "Error: Could not find a matching column between the files to compare."
            
            # VECTORIZED APPROACH: Use merge to identify new, modified, and deleted rows
            # Add indicators to track row sources
            master_df['_master_flag'] = True
            new_df['_new_flag'] = True
            
            # Merge on key column to identify relationships
            merged = pd.merge(
                master_df, 
                new_df, 
                on=key_column, 
                how='outer', 
                suffixes=('_master', '_new'),
                indicator=True
            )
            
            # Initialize tracking lists
            changes_list = []
            alerts_list = []
            
            # Track new rows (in new file but not in master)
            new_rows_mask = merged['_merge'] == 'right_only'
            new_rows = merged[new_rows_mask]
            
            for idx, row in new_rows.iterrows():
                identifier = row[key_column]
                change_info = {
                    'identifier': identifier,
                    'change_type': 'new',
                    'changes': {}
                }
                changes_list.append(change_info)
                
                # Check for alerts on new rows
                # Capacity alert
                row_data = {}
                for col in new_df.columns:
                    if col != '_new_flag':
                        col_new = col + '_new' if col + '_new' in row else col
                        if col_new in row:
                            row_data[col] = row[col_new]
                
                alert = check_capacity_alert(row_data, identifier)
                if alert:
                    alerts_list.append(alert)
                
                alert = check_missing_data_alert(row_data, identifier)
                if alert:
                    alerts_list.append(alert)
                
                # Info alert for new row
                alerts_list.append({
                    'severity': 'info',
                    'type': 'New Row',
                    'identifier': identifier,
                    'description': 'New row added to dataset',
                    'old_value': '',
                    'new_value': ''
                })
            
            # Track deleted rows (in master but not in new file)
            deleted_rows_mask = merged['_merge'] == 'left_only'
            deleted_rows = merged[deleted_rows_mask]
            
            for idx, row in deleted_rows.iterrows():
                identifier = row[key_column]
                change_info = {
                    'identifier': identifier,
                    'change_type': 'deleted',
                    'changes': {}
                }
                changes_list.append(change_info)
            
            # Track modified rows (in both files)
            both_rows_mask = merged['_merge'] == 'both'
            both_rows = merged[both_rows_mask]
            
            for idx, row in both_rows.iterrows():
                identifier = row[key_column]
                has_changes = False
                change_details = {}
                
                # Check each column for changes
                for col in columns_to_compare:
                    col_master = col + '_master'
                    col_new = col + '_new'
                    
                    if col_master in row and col_new in row:
                        old_val = row[col_master]
                        new_val = row[col_new]
                        
                        # Compare values (handling NaN)
                        if pd.isna(old_val) and pd.isna(new_val):
                            continue
                        elif pd.isna(old_val) or pd.isna(new_val) or old_val != new_val:
                            has_changes = True
                            change_details[col] = {
                                'old': old_val if not pd.isna(old_val) else '',
                                'new': new_val if not pd.isna(new_val) else ''
                            }
                            
                            # Check for alerts
                            if 'FTE' in col:
                                alert = check_fte_change_alert(old_val, new_val, identifier)
                                if alert:
                                    alerts_list.append(alert)
                            
                            if col in FACULTY_COLUMNS:
                                alert = check_faculty_change_alert(old_val, new_val, identifier, col)
                                if alert:
                                    alerts_list.append(alert)
                
                if has_changes:
                    change_info = {
                        'identifier': identifier,
                        'change_type': 'modified',
                        'changes': change_details
                    }
                    changes_list.append(change_info)
                    
                    # Check capacity for modified rows
                    row_data = {}
                    for col in new_df.columns:
                        if col != '_new_flag':
                            col_new = col + '_new'
                            if col_new in row:
                                row_data[col] = row[col_new]
                    
                    alert = check_capacity_alert(row_data, identifier)
                    if alert:
                        alerts_list.append(alert)
            
            # Build updated master using vectorized operations
            # Start with all rows from master, then update with new values
            updated_master = master_df.drop('_master_flag', axis=1, errors='ignore').copy()
            
            # Update existing rows with new values (needed for Excel formatting preservation)
            # Note: This loop is required to maintain Excel cell formatting when is_excel=True
            for idx, new_row in new_df.iterrows():
                sec_name = new_row[key_column]
                master_idx = master_df[master_df[key_column] == sec_name].index
                
                if len(master_idx) > 0:
                    master_idx = master_idx[0]
                    # Update ALL columns from new data
                    for col in new_df.columns:
                        if col in updated_master.columns and col != '_new_flag':
                            updated_master.loc[master_idx, col] = new_row[col]
            
            # Add new rows that don't exist in master
            new_rows_df = new_df[new_df[key_column].isin(new_rows[key_column])].copy()
            if '_new_flag' in new_rows_df.columns:
                new_rows_df = new_rows_df.drop('_new_flag', axis=1)
            
            if not new_rows_df.empty:
                updated_master = pd.concat([updated_master, new_rows_df], ignore_index=True)
            
            # Calculate processing time
            processing_time = time.time() - start_time
            
            # Count statistics
            num_new = len(new_rows)
            num_modified = sum(1 for c in changes_list if c['change_type'] == 'modified')
            num_deleted = len(deleted_rows)
            total_changes = len(changes_list)
            
            # Save updated master (Excel handling)
            if is_excel:
                # Load workbook to preserve formatting
                from openpyxl import load_workbook
                wb = load_workbook(excel_file_path)
                ws = wb.active
                
                # Create a mapping of column names to Excel column indices in master
                col_name_to_idx = {}
                header_row = [cell.value for cell in ws[1]]
                for col_idx, col_name in enumerate(header_row):
                    if col_name:
                        col_name_to_idx[str(col_name).strip()] = col_idx + 1  # Excel is 1-indexed
                
                # Update cells with new data
                for idx, new_row in new_df.iterrows():
                    sec_name = new_row[key_column]
                    master_idx = master_df[master_df[key_column] == sec_name].index
                    
                    if len(master_idx) > 0:
                        row_num = master_idx[0] + 2  # +2 because Excel is 1-indexed and has header
                        
                        # Update each column that exists in both files
                        for col_name in new_df.columns:
                            col_name_str = str(col_name).strip()
                            if col_name_str in col_name_to_idx and col_name != '_new_flag':
                                excel_col = col_name_to_idx[col_name_str]
                                cell_value = new_row[col_name]
                                # Handle NaN values
                                if pd.isna(cell_value):
                                    cell_value = None
                                ws.cell(row=row_num, column=excel_col).value = cell_value
                
                output_path = os.path.join(os.path.dirname(__file__), "updated_master.xlsx")
                wb.save(output_path)
                output_ext = '.xlsx'
            else:
                output_path = os.path.join(os.path.dirname(__file__), "updated_master.csv")
                updated_master.to_csv(output_path, index=False)
                output_ext = '.csv'
            
            # Store in session
            session['has_new_data'] = False
            session['output_file'] = output_path
            session['output_ext'] = output_ext
            session['changes_list'] = changes_list
            session['alerts_list'] = alerts_list
            session['processing_time'] = processing_time
            session['num_rows_processed'] = len(master_df) + len(new_df)
            session['num_new'] = num_new
            session['num_modified'] = num_modified
            session['num_deleted'] = num_deleted
            
            return render_template('results.html', 
                                 num_changes=total_changes,
                                 num_new=num_new,
                                 num_modified=num_modified,
                                 num_deleted=num_deleted,
                                 changes_list=changes_list,
                                 alerts_list=alerts_list,
                                 processing_time=processing_time,
                                 num_rows_processed=len(master_df) + len(new_df))
    
    return render_template('select_master.html', 
                         preview=session.get('new_data_preview'))
@app.route('/download')
def download():
    output_path = session.get('output_file')
    output_ext = session.get('output_ext', '.csv')
    
    if output_path and os.path.exists(output_path):
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if output_ext == '.xlsx' else 'text/csv'
        download_name = 'updated_master' + output_ext
        return send_file(output_path, mimetype=mimetype, as_attachment=True, download_name=download_name)
    return "No updated file found."

if __name__ == '__main__':
    app.run(debug=True)