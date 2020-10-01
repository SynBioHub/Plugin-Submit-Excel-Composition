from flask import Flask, request, abort, send_file, jsonify
import os, shutil, tempfile, io
import pandas as pd
from excel2sbol.functions import quality_check_metadata, load_libraries
from excel2sbol.functions import get_data, get_parts, check_name, write_sbol_comp, fix_msec_sbol
from sbol2 import *

app = Flask(__name__)

@app.route("/status")
def status():
    return("The Submit Excel Plugin Flask Server is up and running")



@app.route("/evaluate", methods=["POST"])
def evaluate():
    #uses MIME types
    #https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
    
    eval_manifest = request.get_json(force=True)
    files = eval_manifest['manifest']['files']
    
    #temp
    cwd = os.getcwd()
    data = str(eval_manifest)
    
    eval_response_manifest = {"manifest":[]}
    
    for file in files:
        file_name = file['filename']
        file_type = file['type']
        file_url = file['url']
        
        ########## REPLACE THIS SECTION WITH OWN RUN CODE #################
        acceptable_types = {'application/vnd.ms-excel',
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}

        #could change what appears in the useful_types based on the file content
        useful_types = {}
        
        file_type_acceptable = file_type in acceptable_types
        file_type_useable = file_type in useful_types
        ################## END SECTION ####################################
        
        if file_type_acceptable:
            useableness = 2
        elif file_type_useable:
            useableness = 1
        else:
            useableness = 0
        
        eval_response_manifest["manifest"].append({
            "filename": file_name,
            "requirement": useableness})
        
    return jsonify(eval_response_manifest)


@app.route("/run", methods=["POST"])
def run():

    cwd = os.getcwd()
    
    #create a temporary directory
    temp_dir = tempfile.TemporaryDirectory()
    zip_in_dir_name = temp_dir.name
    
    #take in run manifest
    run_manifest = request.get_json(force=True)
    files = run_manifest['manifest']['files']
    
    #Read in template to compare to
    template_path = os.path.join(cwd, "templates", "darpa_template_blank.xlsx")
    
    #initiate response manifest
    run_response_manifest = {"results":[]}
    
    for a_file in files:
        try:
            file_name = a_file['filename']
            file_type = a_file['type']
            file_url = a_file['url']
            data = str(a_file)
           
            converted_file_name = f"{file_name}.converted"
            file_path_out = os.path.join(zip_in_dir_name, converted_file_name)
        
            ########## REPLACE THIS SECTION WITH OWN RUN CODE #################
            #Load Data
            startrow_composition = 9
            sheet_name = "Composite Parts"
            nrows = 8
            use_cols = [0,1]
            
            #read in whole composite sheet below metadata
            table = pd.read_excel (file_url, sheet_name = sheet_name, 
                                   header = None, skiprows = startrow_composition)
            
            #Load Metadata
            filled_composition_metadata = pd.read_excel (file_url, sheet_name = sheet_name,
                                          header= None, nrows = nrows, usecols = use_cols)
            blank_composition_metadata = pd.read_excel (template_path, sheet_name = sheet_name,
                                          header= None, nrows = nrows, usecols = use_cols)
            
            #Compare the metadata to the template
            quality_check_metadata(filled_composition_metadata, blank_composition_metadata)
            
            #Load Libraries required for Parts
            libraries = load_libraries(table)
            
            #Loop over all rows and find those where each block begins
            compositions, list_of_rows = get_data(table)
                        
            #Extract parts from table
            compositions, all_parts = get_parts(list_of_rows, table, compositions)
            
            #Check if Collection names are alphanumeric and separated by underscore
            compositions = check_name(compositions)
            
            
            #Create sbol
            doc = write_sbol_comp(libraries, compositions, all_parts)
            doc.write(file_path_out)
            
            
            #fix millisecond bug in pysbol/sbol
            fix_msec_sbol(file_path_out)
            ################## END SECTION ####################################
        
            # add name of converted file to manifest
            run_response_manifest["results"].append({"filename":converted_file_name,
                                    "sources":[file_name]})

        except Exception as e:
            print(e)
            abort(415)
            
    #create manifest file
    file_path_out = os.path.join(zip_in_dir_name, "manifest.json")
    with open(file_path_out, 'w') as manifest_file:
            manifest_file.write(str(run_response_manifest)) 
      
    
    with tempfile.NamedTemporaryFile() as temp_file:
        #create zip file of converted files and manifest
        shutil.make_archive(temp_file.name, 'zip', zip_in_dir_name)
        
        #delete zip in directory
        shutil.rmtree(zip_in_dir_name)
        
        #return zip file
        return send_file(f"{temp_file.name}.zip")
            

