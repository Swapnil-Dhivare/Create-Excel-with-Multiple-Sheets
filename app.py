from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
import io
import xlsxwriter

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # For flash messages

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_excel():
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Loop through submitted sheets
            sheet_count = 1
            while f'sheet_title_{sheet_count}' in request.form:
                # Get data for each sheet
                sheet_title = request.form[f'sheet_title_{sheet_count}']
                columns = request.form[f'columns_{sheet_count}'].split(',')
                data = request.form[f'data_{sheet_count}'].split(';')

                # Create a dictionary for the current sheet's DataFrame
                data_dict = {}
                for i in range(len(columns)):
                    column_data = []
                    for row in data:
                        row_values = row.split(',')
                        if i < len(row_values):
                            value = row_values[i].strip()
                            # Convert to numeric if possible, otherwise keep as string
                            try:
                                column_data.append(float(value) if '.' in value else int(value))
                            except ValueError:
                                column_data.append(value)  # Keep as string if not numeric
                        else:
                            column_data.append('')  # Fill missing data with empty strings
                    data_dict[columns[i].strip()] = column_data

                # Convert to DataFrame
                df = pd.DataFrame(data_dict)

                # Write the DataFrame to the corresponding sheet
                df.to_excel(writer, sheet_name=sheet_title.strip(), index=False)

                # Auto-adjust column width
                worksheet = writer.sheets[sheet_title.strip()]
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2  # Add extra padding
                    worksheet.set_column(idx, idx, max_len)

                sheet_count += 1

        output.seek(0)
        return send_file(output, download_name="multi_sheet_course_file.xlsx", as_attachment=True)

    except Exception as e:
        flash(f"Error generating file: {str(e)}", 'error')
        return redirect(url_for('index'))

if __name__ == "__main__":
    app.run(debug=True)
