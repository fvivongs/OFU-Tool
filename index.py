from flask import Flask, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename
from tkinter import filedialog
from pandas import DataFrame, merge, ExcelFile

import numpy as np
import pandas as pd
import os

df = 0

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__, static_url_path="/static")
DIR_PATH = os.path.dirname(os.path.realpath(__file__))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# limit upload size upto 8mb
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            process_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    return render_template('./index.html')

def process_file(path):
    global df
    df = pd.read_excel(open(path, 'rb'))
    remove_watermark()
    return df

def remove_watermark():
    global df
    Doc_Prep()
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    df.to_excel(export_file_path, index=None, header=True)

#========================================================================================================================
# Insert '---' for all null cells & Constructing codes for low level filtering
#========================================================================================================================
    # Category Focus | outageID, Clearing Device, Responsible System, Cause(IEEE), Failure Mode,
    # AT/D/M, Equipment Code, Manufacturer/Species, Planned Outages, Weather

def Doc_Prep():
    global df
    # ========================================================================================================================
    # Separate pages on "code_type.xlsx" into dataframes
    # ========================================================================================================================
    rules = pd.ExcelFile("code_type.xlsx")
    df_cd = pd.read_excel(rules, 'CD Type')
    df_rs = pd.read_excel(rules, 'RS Type')
    df_cause = pd.read_excel(rules, 'Cause Type')
    df_fm = pd.read_excel(rules, 'FM Type')
    df_atdm = pd.read_excel(rules, 'ATDM Type')
    df_ec = pd.read_excel(rules, 'EC Type')
    df_ms = pd.read_excel(rules, 'MS Type')
    df_po = pd.read_excel(rules, 'PO Type')
    df_we = pd.read_excel(rules, 'WE Type')
    # ========================================================================================================================
    # END
    # ========================================================================================================================
    cdNullCondition = [df['Clearing Device'].isnull()]
    rsNullCondition = [df['Resp. System'].isnull()]
    causeNullCondition = [df['Cause (IEEE)'].isnull()]
    fmNullCondition = [df['Failure Mode'].isnull()]
    atdmNullCondition = [df['AT/D/M'].isnull()]
    ecNullCondition = [df['Eq. Code'].isnull()]
    msNullCondition = [df['Manuf./ Species'].isnull()]
    poNullCondition = [df['Planned Outages'].isnull()]
    weNullCondition = [df['WE'].isnull()]

    general_error_choices = ['---']

    # ========================================================================================================================
    # Constructing codes for low level filtering
    # ========================================================================================================================

    df['LOW_CD'] = np.select(cdNullCondition, general_error_choices, default= df['Clearing Device'].astype(str).str[0:2])
    df['LOW_RS'] = np.select(rsNullCondition, general_error_choices, default= df['Resp. System'].astype(str).str[0:2])
    df['LOW_Cause'] = np.select(causeNullCondition , general_error_choices, default= df['Cause (IEEE)'].astype(str).str[0:2])
    df['LOW_FM'] = np.select(fmNullCondition, general_error_choices, default= df['Failure Mode'].astype(str).str[0:2])
    df['LOW_ATDM'] = np.select(atdmNullCondition, general_error_choices, default= df['AT/D/M'].astype(str).str[0:2])
    df['LOW_EC'] = np.select(ecNullCondition, general_error_choices, default= df['Eq. Code'].astype(str).str[0:2])
    df['LOW_MS'] = np.select(msNullCondition, general_error_choices, default= df['Manuf./ Species'].astype(str).str[0:2])
    df['LOW_PO'] = np.select(poNullCondition, general_error_choices, default= df['Planned Outages'].astype(str).str[0:2])
    df['LOW_WE'] = np.select(weNullCondition, general_error_choices, default= df['WE'].astype(str).str[0:2])

    # ========================================================================================================================
    # Constructing codes for high level filtering
    # ========================================================================================================================

    df['trimmed_CD'] = np.select(cdNullCondition, general_error_choices, default= 'CD_' + df['Clearing Device'].astype(str).str[0:2])
    df['trimmed_RS'] = np.select(rsNullCondition, general_error_choices, default= 'RS_' + df['Resp. System'].astype(str).str[0:2])
    df['trimmed_Cause'] = np.select(causeNullCondition , general_error_choices, default= 'Cause_' + df['Cause (IEEE)'].astype(str).str[0:2])
    df['trimmed_FM'] = np.select(fmNullCondition, general_error_choices, default='FM_' + df['Failure Mode'].astype(str).str[0:2])
    df['trimmed_ATDM'] = np.select(atdmNullCondition, general_error_choices, default='AT_' + df['AT/D/M'].astype(str).str[0:2])
    df['trimmed_EC'] = np.select(ecNullCondition, general_error_choices, default='EC_' + df['Eq. Code'].astype(str).str[0:2])
    df['trimmed_MS'] = np.select(msNullCondition, general_error_choices, default='MS_' + df['Manuf./ Species'].astype(str).str[0:2])
    df['trimmed_PO'] = np.select(poNullCondition, general_error_choices, default='PO_' + df['Planned Outages'].astype(str).str[0:2])
    df['trimmed_WE'] = np.select(weNullCondition, general_error_choices, default= 'WE_' + df['WE'].astype(str).str[0:2])

    # ========================================================================================================================
    # high level codes continued (references "code_type.xlsx")
    # ========================================================================================================================

    df = df.merge(df_cd, on=['trimmed_CD'], how='left')
    df = df.merge(df_rs, on=['trimmed_RS'], how='left')
    df = df.merge(df_cause, on=['trimmed_Cause'], how='left')
    df = df.merge(df_fm, on=['trimmed_FM'], how='left')
    df = df.merge(df_atdm, on=['trimmed_ATDM'], how='left')
    df = df.merge(df_ec, on=['trimmed_EC'], how='left')
    df = df.merge(df_ms, on=['trimmed_MS'], how='left')
    df = df.merge(df_po, on=['trimmed_PO'], how='left')
    df = df.merge(df_we, on=['trimmed_WE'], how='left')

    df.fillna('---', inplace = True)

    df['Comments'] = ''
    df['Notification'] = 'Valid'
    df['# Corrections'] = 0

    # ================================
    # NULL CLEARING DEVICE
    # ================================

    def findNull_CD(df):
        if (df['Clearing Device'] == '---'):
            return False
        else:
            return True

    df['findNull_CD'] = df.apply(findNull_CD, axis=1)

    df.loc[df['findNull_CD'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_CD'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_CD'] == False, 'Comments'] = df['Comments'] + 'Missing Responsible System. '

    # ================================
    # NULL RESPONSIBLE SYSTEM
    # ================================

    def findNull_RS(df):
        if (df['Resp. System'] == '---'):
            return False
        else:
            return True

    df['findNull_RS'] = df.apply(findNull_RS, axis=1)

    df.loc[df['findNull_RS'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_RS'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_RS'] == False, 'Comments'] = df['Comments'] + 'Missing Responsible System. '

    # ================================
    # NULL CAUSE
    # ================================

    def findNull_Cause(df):
        if (df['Cause (IEEE)'] == '---'):
            return False
        else:
            return True

    df['findNull_Cause'] = df.apply(findNull_Cause, axis=1)

    df.loc[df['findNull_Cause'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_Cause'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_Cause'] == False, 'Comments'] = df['Comments'] + 'Missing Cause (IEEE). '

    # ================================
    # NULL FAILUREMODE
    # ================================

    def findNull_FM(df):
        if (df['Failure Mode'] == '---'):
            return False
        else:
            return True

    df['findNull_FM'] = df.apply(findNull_FM, axis=1)

    df.loc[df['findNull_FM'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_FM'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_FM'] == False, 'Comments'] = df['Comments'] + 'Missing Cause (IEEE). '

    # ================================
    # NULL AT/D/M
    # ================================

    def findNull_ATDM(df):
        if (df['AT/D/M'] == '---'):
            return False
        else:
            return True

    df['findNull_ATDM'] = df.apply(findNull_ATDM, axis=1)

    df.loc[df['findNull_ATDM'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_ATDM'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_ATDM'] == False, 'Comments'] = df['Comments'] + 'Missing AT/D/M. '

    # ================================
    # NULL AT/D/M
    # ================================

    def findNull_EC(df):
        if (df['Eq. Code'] == '---'):
            return False
        else:
            return True

    df['findNull_EC'] = df.apply(findNull_EC, axis=1)

    df.loc[df['findNull_EC'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_EC'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_EC'] == False, 'Comments'] = df['Comments'] + 'Eq. Code. '

    # ================================
    # NULL MANUFACTURER/SPECIES
    # ================================

    def findNull_MS(df):
        if (df['Manuf./ Species'] == '---'):
            return False
        else:
            return True

    df['findNull_MS'] = df.apply(findNull_MS, axis=1)

    df.loc[df['findNull_MS'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_MS'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_MS'] == False, 'Comments'] = df['Comments'] + 'Missing Manuf./ Species. '

    # ================================
    # NULL MANUFACTURER/SPECIES
    # ================================

    def findNull_PO(df):
        if (df['Planned Outages'] == '---'):
            return False
        else:
            return True

    df['findNull_PO'] = df.apply(findNull_PO, axis=1)

    df.loc[df['findNull_PO'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_PO'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_PO'] == False, 'Comments'] = df['Comments'] + 'Missing Planned Outage. '

    # ================================
    # NULL MANUFACTURER/SPECIES
    # ================================

    def findNull_WE(df):
        if (df['WE'] == '---'):
            return False
        else:
            return True

    df['findNull_WE'] = df.apply(findNull_WE, axis=1)

    df.loc[df['findNull_WE'] == False, 'Notification'] = 'Correction'
    df.loc[df['findNull_WE'] == False, '# Corrections'] = df['# Corrections'] + 1
    df.loc[df['findNull_WE'] == False, 'Comments'] = df['Comments'] + 'Missing Weather. '

    # ========================================================================================================================
    # HIGH LEVEL TEST: Compares Cause (IEEE) to High Level Failure Mode
    # ========================================================================================================================

    df['test1'] = (
            (df['LOW_Cause'] == '03') & (df['highFM'] == 'tree codes') |
            # Vegetation
            (df['LOW_Cause'] == '20') & ((df['highFM'] == 'deterioration') | (df['highFM'] == 'design issues')) |
            # Equipment Failure
            (df['LOW_Cause'] == '09') & ((df['highFM'] == 'human intervention') | (df['highFM'] == 'tree codes')) |
            # Public Accident/Damage
            (df['LOW_Cause'] == '04') & ((df['highFM'] == 'environment') | (df['highFM'] == 'design issues')) |
            # WildLife
            (df['LOW_Cause'] == '19') & ((df['highFM'] == 'environment') | (df['highFM'] == 'design issues')) |
            # Lightning Strike
            (df['LOW_Cause'] == 'EA') & (df['highFM'] == 'environment') |
            # Weather
            (df['LOW_Cause'] == '05') & ((df['highFM'] == 'work request') | (df['highFM'] == 'design issues')) |

            # ================================
            # The following may need revision:
            # ================================

            (df['LOW_Cause'] == '41') |
            # Loss of Transmission/Generation
            (df['LOW_Cause'] == '11') |
            # Unknown Cause
            (df['LOW_Cause'] == '28')
            # Other Cause
    )

    # ========================================================================================================================
    # TEST OUTPUTS
    # ========================================================================================================================

    df.loc[df['test1'] == False, 'Notification'] = 'Mismatch'
    df.loc[df['test1'] == False, 'Comments'] = df['Comments'] + 'Cause and FailureMode do not match. '

    del df['Category']
    del df['Op Center']
    del df['Circuit']
    del df['Time Off']
    del df['Time On']
    del df['Device & Ph']
    del df['# Cust']
    del df['Ckt Cust']
    del df['Dur']
    del df['Fault Location']
    del df['trimmed_FM']
    del df['trimmed_ATDM']
    del df['trimmed_EC']
    del df['trimmed_MS']
    del df['trimmed_PO']
    del df['LOW_CD']
    del df['LOW_RS']
    del df['LOW_Cause']
    del df['LOW_FM']
    del df['LOW_ATDM']
    del df['LOW_EC']
    del df['LOW_MS']
    del df['LOW_PO']
    del df['LOW_WE']
    del df['trimmed_CD']
    del df['trimmed_RS']
    del df['trimmed_Cause']
    del df['highEC']
    del df['highMS']
    del df['highPO']
    del df['highWE']
    del df['trimmed_WE']
    del df['highCD']
    del df['highRS']
    del df['highCause']
    del df['highFM']
    del df['highATDM']

    del df['test1']
    del df['findNull_CD']
    del df['findNull_RS']
    del df['findNull_Cause']
    del df['findNull_FM']
    del df['findNull_ATDM']
    del df['findNull_EC']
    del df['findNull_MS']
    del df['findNull_PO']
    del df['findNull_WE']
#========================================================================================================================
# END
#========================================================================================================================
# HIGH LEVEL TEST: Null Values
#========================================================================================================================

# ================================
# NOTE: Does not include cnt
# ================================

# ================================
# NULL CLEARING DEVICE
# ================================


if __name__ == "__main__":
    app.run(debug=True)
