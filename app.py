from flask import Flask, render_template, request, jsonify, send_file,redirect,url_for
from flask_cors import CORS
from docx import Document
from openpyxl import Workbook
import os
import tempfile
import shutil
import sys
import zipfile
from werkzeug.utils import secure_filename
from Styles import StyleExtractor
from Styles import process_documents_and_update_xml
from Styles import WordToJLDConverter
from rational import web_bp  # Import the blueprint
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from flask_cors import CORS
from docx import Document
from openpyxl import Workbook
import os
import tempfile
import shutil
from werkzeug.utils import secure_filename
from rational import web_bp  # Import the blueprint


app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or 'dev-secret-key'

app.register_blueprint(web_bp)

CORS(app)  # Enable cross-origin requests

# Configuration

UPLOAD_FOLDER = 'uploads'

OUTPUT_FOLDER = 'output'

REPORT_FOLDER= 'reports'

ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

app.config['REPORT_FOLDER'] = OUTPUT_FOLDER


# Ensure upload and output directories exist

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

os.makedirs(REPORT_FOLDER, exist_ok=True)



def allowed_file(filename):

    """Check if the file has an allowed extension."""

    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")

def homepage():

    with app.app_context():



        return render_template("LatestHomePage.html")


@app.route("/page2")

def page2():

    with app.app_context():



        return render_template("FeaturesInfo.html")


@app.route("/page3")

def page3():

    with app.app_context():



        return render_template("StylesExtraction.html")

   
@app.route("/page4")

def page4():

    with app.app_context():



        return render_template("JLD.html")



@app.route("/back/<current_page>")
def back(current_page):
    # Page4 specific back navigation
    if current_page == "page4":
        return redirect(url_for("page2"))
    
    # Page3 specific back navigation
    if current_page == "page3":
        return redirect(url_for("page2"))
    
    # Page2 specific back navigation
    if current_page == "page2":
        return redirect(url_for("LatestHomePage"))
    
    # Default navigation for any other case
    return redirect(url_for("LatestHomePage"))


    
# Style extraction routes

@app.route("/api/extract", methods=["POST"])

def extract_files():

    temp_dir = None

    try:

        # Check if files were uploaded

        if "files" not in request.files:

            return jsonify({"error": "No files uploaded"}), 400


        files = request.files.getlist("files")

        # Create temporary directory to store uploaded files

        temp_dir = tempfile.mkdtemp()

        file_paths = []

        # Save uploaded files temporarily, preserving folder structure

        for file in files:

            if file.filename.endswith(".docx"):

                # Get the relative path

                rel_path = file.filename

                # Create necessary subdirectories

                file_dir = os.path.dirname(os.path.join(temp_dir, rel_path))

                if file_dir and not os.path.exists(file_dir):

                    os.makedirs(file_dir, exist_ok=True)

                # Save the file

                temp_path = os.path.join(temp_dir, rel_path)

                file.save(temp_path)

                file_paths.append(temp_path)


        if not file_paths:

            return jsonify({"error": "No valid Word documents (.docx) uploaded"}), 400


        # Create an instance of StyleExtractor

        extractor = StyleExtractor()

        # Create output folder within temp directory

        output_folder = os.path.join(temp_dir, "Excel_Output")

        os.makedirs(output_folder, exist_ok=True)

        output_excel = os.path.join(output_folder, "StylesData.xlsx")

        # Process the files

        success = extractor.process_documents(file_paths, output_excel)

        if not success:

            return jsonify({"error": "Failed to process documents"}), 500



        # Return the generated Excel file

        return send_file(

            output_excel,

            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            as_attachment=True,

            download_name="StylesData.xlsx",

        )

       

    except Exception as e:

        app.logger.error(f"Error processing files: {str(e)}", exc_info=True)



        return jsonify({"error": str(e)}), 500



    finally:

        # Clean up temporary directory

        if temp_dir and os.path.exists(temp_dir):

            try:

                shutil.rmtree(temp_dir)

            except Exception as e:

                app.logger.warning(f"Failed to clean up temp directory: {str(e)}")





@app.route("/api/extract-path", methods=["POST"])

def extract_from_path():

    try:

        data = request.json

        if not data or "path" not in data:

            return jsonify({"error": "No path provided"}), 400



        input_path = data["path"]

        # Create an instance of StyleExtractor

        extractor = StyleExtractor()

        # Process the path using extract_data method

        success, message = extractor.extract_data(input_path)

        if not success:

            return jsonify({"error": message}), 500

        # Get the expected output Excel path

        if os.path.isdir(input_path):

            output_excel = os.path.join(input_path, "Excel_Output", "StylesData.xlsx")

        else:

            folder_path = os.path.dirname(input_path)

            output_excel = os.path.join(folder_path, "Excel_Output", "StylesData.xlsx")



        if not os.path.exists(output_excel):

            return jsonify({"error": "Output file not generated"}), 500

        # Return the generated Excel file

        return send_file(

            output_excel,

            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            as_attachment=True,

            download_name="StylesData.xlsx",

        )



    except Exception as e:

        app.logger.error(f"Error processing path: {str(e)}", exc_info=True)



        return jsonify({"error": str(e)}), 500



# XML processing routes



@app.route('/upload', methods=['POST'])

def upload_files():

    """Handle file upload and process DOCX files for XML updating."""

    if 'files' not in request.files:

        return jsonify({"error": "No files provided"}), 400



    files = request.files.getlist('files')

    uploaded_files = []

    # Save uploaded files

    for file in files:

        if file and allowed_file(file.filename):

            filename = secure_filename(file.filename)

            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

            file.save(file_path)

            uploaded_files.append(file_path)



    if not uploaded_files:

        return jsonify({"error": "No valid DOCX files uploaded"}), 400



    # Define paths for XML files



    # Assuming word2jld.xml is in the same directory as your_backend_functions.py

    xml_input_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WORD2JLD.xml')  # Reference XML file

    xml_output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'OUTPUT_JLD.xml')  # Output XML file

    # Process the uploaded DOCX files

    try:

        process_documents_and_update_xml(app.config['UPLOAD_FOLDER'], xml_input_path, xml_output_path)



        return jsonify({"message": "XML file generated successfully!", "output_file": xml_output_path})



    except Exception as e:

        return jsonify({"error": str(e)}), 500

   

# Add the new route that matches what your frontend is calling\

@app.route("/api/generate-xml", methods=["POST"])

def generate_xml():

    """Handle file upload and process DOCX files for XML updating."""

    if 'files' not in request.files:

        return jsonify({"error": "No files provided"}), 400



    files = request.files.getlist('files')

    uploaded_files = []

    # Save uploaded files

    for file in files:

        if file and allowed_file(file.filename):

            filename = secure_filename(file.filename)

            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            uploaded_files.append(file_path)

    if not uploaded_files:
        return jsonify({"error": "No valid DOCX files uploaded"}), 400

    # Define paths for XML files
    xml_input_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WORD2JLD.xml')
    xml_output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'updatedWord2jld.xml')
    # Process the uploaded DOCX files
    try:
        process_documents_and_update_xml(app.config['UPLOAD_FOLDER'], xml_input_path, xml_output_path)
        # Verify the file exists before attempting to send it
        if not os.path.exists(xml_output_path):
            return jsonify({"error": "Output file not generated"}), 500

        # Return the XML file directly
        return send_file(
            xml_output_path,
            mimetype="application/xml",
            as_attachment=True,
            download_name="updatedWord2jld.xml"
        )

    except Exception as e:
        app.logger.error(f"Error in generate_xml: {str(e)}", exc_info=True)

        return jsonify({"error": str(e)}), 500


@app.route('/download')
def download_file():
    """Allow users to download the generated XML file."""
    output_file = request.args.get('file')
    if not output_file or not os.path.exists(output_file):
        return jsonify({"error": "File not found"}), 404

    return send_file(output_file, as_attachment=True)





# Configure upload and output folders
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Ensure upload and output folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# @app.route('/')
# def index():
#     return render_template('JLD.html')  # Serve the front-end HTML

@app.route('/process', methods=['POST'])
def process_files():
    try:
        # Save uploaded files
        excel_files = request.files.getlist('excelFiles')
        base_template = request.files['baseTemplate']
        variables_file = request.files['variablesFile']
        output_path = request.form['outputPath']

        # Save files to upload folder
        excel_paths = []
        for file in excel_files:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(file_path)
            excel_paths.append(file_path)

        base_template_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(base_template.filename))
        base_template.save(base_template_path)

        variables_file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(variables_file.filename))
        variables_file.save(variables_file_path)

        # Process files using your backend class
        converter = WordToJLDConverter()
        converter.baseblock_path = base_template_path
        converter.variables_file_path = variables_file_path
        converter.output_folder = output_path

        for excel_path in excel_paths:
            converter.process_excel_file(excel_path)

        return jsonify({"message": "JLD files generated successfully!"}), 200

    except Exception as e:
        return jsonify({"message": f"Error: {str(e)}"}), 500



@app.route("/api/convert-docx", methods=["POST"])
def convert_docx_to_jld():
    """Handle direct conversion from DOCX to JLD without exposing Excel files"""
    try:
        # Check if files were uploaded
        if "files" not in request.files:
            return jsonify({"error": "No files uploaded"}), 400

        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "No files selected"}), 400

        # Get other required files
        if "baseTemplate" not in request.files:
            return jsonify({"error": "Base template file is required"}), 400
        
        if "variablesFile" not in request.files:
            return jsonify({"error": "Variables file is required"}), 400

        base_template = request.files["baseTemplate"]
        variables_file = request.files["variablesFile"]
        output_path = request.form.get("outputPath")

        if not output_path:
            return jsonify({"error": "Output path is required"}), 400

        # Save files temporarily
        temp_dir = tempfile.mkdtemp()
        docx_paths = []
        
        try:
            # Save DOCX files
            for file in files:
                if file.filename.endswith('.docx'):
                    temp_path = os.path.join(temp_dir, secure_filename(file.filename))
                    file.save(temp_path)
                    docx_paths.append(temp_path)

            # Save other required files
            base_template_path = os.path.join(temp_dir, secure_filename(base_template.filename))
            base_template.save(base_template_path)
            
            variables_file_path = os.path.join(temp_dir, secure_filename(variables_file.filename))
            variables_file.save(variables_file_path)

            # Create StyleExtractor instance for DOCX processing
            extractor = StyleExtractor()
            converter = WordToJLDConverter()
            converter.baseblock_path = base_template_path
            converter.variables_file_path = variables_file_path
            converter.output_folder = output_path
            
            # Process each DOCX file
            generated_files = []
            for docx_path in docx_paths:
                # First extract styles from DOCX and save to Excel
                doc = Document(docx_path)
                styled_content = extractor.extract_text_with_styles(doc)
                
                # Save to temporary Excel file
                filename = os.path.splitext(os.path.basename(docx_path))[0]
                temp_excel_path = os.path.join(temp_dir, f"{filename}_temp.xlsx")
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(["Document", "Paragraph", "Style", "Text"])  # Add headers
                for para_num, para_style, text_style, text in styled_content:
                    sheet.append([filename, para_num, para_style, text_style, text])
                workbook.save(temp_excel_path)

                # Create document data dictionary
                document_data = {filename: []}
                for para_num, para_style, text_style, text in styled_content:
                    document_data[filename].append({
                        "paragraph_number": para_num,
                        "paragraph_style": para_style,
                        "text_style": text_style,
                        "text": text
                    })

                # Generate JLD using create_jld_files
                converter.create_jld_files(document_data, base_template_path, output_path, variables_file_path)
                generated_files.append(os.path.join(output_path, f"{filename}.jld"))

            return jsonify({
                "message": "JLD files generated successfully",
                "output_path": output_path,
                "generated_files": generated_files
            }), 200

        finally:
            # Clean up temporary files
            shutil.rmtree(temp_dir)

    except Exception as e:
        app.logger.error(f"Error in convert_docx_to_jld: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":

    app.run(debug=True, port=5000)