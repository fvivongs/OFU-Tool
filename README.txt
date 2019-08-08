Welcome to the OFU Tool Documentation!

This tool was created to review outage follow-up data and make correction suggestions based on a series
of tests that are created by the user. 

A flow of the application can be found in the file (OFU_Tool_Diagram.pdf).

An explanation of the folder and files can be viewed below:

Folders:
--> static stores "non-dynamic" files that are specifically used for the front end of this applications. 
    Files include css, javascript, and bootstrap references. 
--> templates contains the html files or pages of the application. index.html is the first page that is
    displayed when accessing the application. 
--> uploads is a folder that contains the files that are unploaded to the tool. Uploaded files are renamed 
    and sent to this folder before being processed. 

Files:
--> flask_app.py is the main python application that references index.html (front end) and contains all of the functions
    for data structuring and tests/scans. 
--> code_type.xlsx contains the information that the tool uses to restructure the data into high level and
    low level codes. 
--> sample_outage_data.xlsx (this is a test document that can be used to upload to the tool)