import os
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from flask import Flask, render_template, request, redirect, url_for, session, make_response
from werkzeug.utils import secure_filename
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

app = Flask(__name__)
app.secret_key = 'super secret key' # Change this to a random secure key
ALLOWED_EXTENSIONS = {'csv', 'xls', 'xlsx'}
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Create the upload folder if it doesn't exist

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to generate visualizations and analysis
def analyze_data(file_path):
    """
    Analyzes the data from the given file path and generates various visualizations.

    Args:
        file_path (str): Path to the CSV or Excel file.

    Returns:
        dict: A dictionary containing the analysis results, including:
            - 'summary': Descriptive statistics of the dataset.
            - 'plots': A dictionary of base64 encoded plot images (histograms, boxplots, etc.).
            - 'column_details': Details about each column (name, type, unique values).
            - 'error': Error message if any error occurs, None otherwise.
    """
    analysis_results = {'summary': '', 'plots': {}, 'column_details': [], 'error': None}
    try:
        # Read the file using pandas
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)

        # Get summary statistics
        analysis_results['summary'] = df.describe().to_html()

        # Get column details
        for col in df.columns:
            unique_values = df[col].unique()
            analysis_results['column_details'].append({
                'name': col,
                'type': str(df[col].dtype),
                'unique_values': len(unique_values),
                'sample_values': list(unique_values[:5])  # Show only the first 5 unique values
            })

        # Generate plots (store as base64 strings)
        for col in df.columns:
            plt.figure(figsize=(8, 6))  # Create a new figure for each plot
            if pd.api.types.is_numeric_dtype(df[col]):
                # Histogram for numeric columns
                plt.hist(df[col], bins=20)
                plt.title(f'Histogram of {col}')
                plt.xlabel(col)
                plt.ylabel('Frequency')
            elif pd.api.types.is_object_dtype(df[col]):
                # Bar chart for categorical columns
                value_counts = df[col].value_counts()
                plt.bar(value_counts.index, value_counts.values)
                plt.title(f'Bar Chart of {col}')
                plt.xlabel(col)
                plt.ylabel('Count')
            else:
                analysis_results['plots'][col] = "Skipped: Unsupported data type"
                continue # Skip if the data type is not supported

            # Save the plot to a BytesIO object
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png')
            img_buf.seek(0)
            img_base64 = base64.b64encode(img_buf.getvalue()).decode('utf-8')
            analysis_results['plots'][col] = img_base64
            plt.close()  # Close the figure to free memory

        return analysis_results

    except Exception as e:
        analysis_results['error'] = str(e)
        return analysis_results
    except:
        analysis_results['error'] = "Unknown Error"
        return analysis_results

# Flask routes
@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Handles the main page, allowing users to upload a file.
    """
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            return render_template('index.html', error='No selected file')
        if file and allowed_file(file.filename):
            # Secure the filename and save the file
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            session['file_path'] = file_path  # Store the file path in session
            return redirect(url_for('analyze'))  # Redirect to the analyze route
        else:
            return render_template('index.html', error='Invalid file type. Please upload a CSV or Excel file.')
    return render_template('index.html')

@app.route('/analyze')
def analyze():
    """
    Analyzes the uploaded file and displays the results.
    """
    file_path = session.get('file_path')  # Retrieve the file path from the session
    if not file_path:
        return redirect(url_for('index'))  # Redirect to home if no file

    analysis_results = analyze_data(file_path) # Analyze the data

    if analysis_results['error']:
        return render_template('error.html', error=analysis_results['error'])

    return render_template('analyze.html',
                           summary=analysis_results['summary'],
                           plots=analysis_results['plots'],
                           column_details=analysis_results['column_details'])

@app.route('/download_results')  # New route for downloading results
def download_results():
    """
    Generates an Excel file containing the analysis results and allows the user to download it.
    """
    file_path = session.get('file_path')
    if not file_path:
        return redirect(url_for('index'))

    analysis_results = analyze_data(file_path)

    if analysis_results['error']:
        return render_template('error.html', error=analysis_results['error'])

    # Create a new Excel file in memory
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # 1. Write Summary Statistics to Excel
    try:
        summary_df = pd.read_html(analysis_results['summary'])[0]  # Convert HTML table to DataFrame
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    except ValueError:
        # Handle the case where the summary is empty
        empty_df = pd.DataFrame({'No Summary Available': []})
        empty_df.to_excel(writer, sheet_name='Summary', index=False)
    # 2. Write Column Details to Excel
    column_details_df = pd.DataFrame(analysis_results['column_details'])
    column_details_df.to_excel(writer, sheet_name='Column Details', index=False)

     # 3. Create sheets for the first 5 plots.  Write the plot to the excel file.
    plot_count = 0
    for col, img_base64 in analysis_results['plots'].items():
        if plot_count >= 5:  # Limit to a maximum of 5 plots
            break
        if "Skipped" in img_base64:
            continue
        try:
            # Decode the base64 string to get the image data
            img_data = base64.b64decode(img_base64)
            # Read the image data using BytesIO
            img_io = io.BytesIO(img_data)

            # Add the plot to a new sheet in the Excel file.
            worksheet = writer.book.add_worksheet(f'Plot_{col}')
            worksheet.insert_image('A1', '', {'image_data': img_io})
            plot_count += 1
        except Exception as e:
            print(f"Error adding plot for {col} to Excel: {e}")

    # Save the Excel file
    writer.close()
    output.seek(0)

    # Create a response to download the file
    response = make_response(output.read())
    response.headers.set('Content-Disposition', 'attachment', filename='analysis_results.xlsx')
    response.headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return response

if __name__ == '__main__':
    app.run(debug=True) # Set debug to false in production
