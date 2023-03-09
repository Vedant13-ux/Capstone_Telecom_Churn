import pickle 
import openpyxl
import pandas as pd
from sklearn.preprocessing import StandardScaler, LabelEncoder
import os 

# Set wd
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Load the model from the file
classifier_from_pickle = pickle.load(open('classifier.pkl', 'rb'))

# Load the Scaler and Label Encoder from the file
label_encoder = pickle.load(open('label_encoder.pkl', 'rb'))
scaler = pickle.load(open('scaler.pkl', 'rb'))


# Get Inputs using openpyxl
wb = openpyxl.load_workbook('Telecom_Churn.xlsm')
sheet = wb.active

# Get the inputs from the excel sheet and apply the transformations
gender = label_encoder['Gender'].transform([sheet['B2'].value])[0]
senior_citizen = label_encoder['SeniorCitizen'].transform([sheet['B3'].value])[0]
partner = label_encoder['Partner'].transform([sheet['B4'].value])[0]
dependents = label_encoder['Dependents'].transform([sheet['B5'].value])[0]
tenure = sheet['B6'].value
phone_service = label_encoder['PhoneService'].transform([sheet['B7'].value])[0]
multiple_lines = 1 if sheet['B8'].value=='Yes' else 0
internet_service = sheet['B9'].value
online_security = 1 if sheet['B10'].value=='Yes' else 0
online_backup = 1 if sheet['B11'].value=='Yes' else 0
device_protection = 1 if sheet['B12'].value=='Yes' else 0
tech_support = 1 if sheet['B13'].value=='Yes' else 0
streaming_tv = 1 if sheet['B14'].value=='Yes' else 0
streaming_movies = 1 if sheet['B15'].value=='Yes' else 0
contract = sheet['B16'].value
paperless_billing = label_encoder['PaperlessBilling'].transform([sheet['B17'].value])[0]
payment_method = sheet['B18'].value
monthly_charges = sheet['B19'].value

# Scaling tenure and monthly charges
[[monthly_charges, tenure]] = scaler.transform([[monthly_charges, tenure]])


# One hot encoding for contract and payment method
contract_mm = 0
contract_1 = 0
contract_2 = 0
if contract == 'One year':
    contract_1 = 1
elif contract == 'Two year':
    contract_2 = 1
else:
    contract_mm = 1


payment_method_bank = 0
payment_method_cc = 0
payment_method_ee = 0
payment_method_mm = 0
if payment_method == 'Bank transfer (automatic)':
    payment_method_bank = 1
elif payment_method == 'Credit card (automatic)':
    payment_method_cc = 1
elif payment_method == 'Electronic check':
    payment_method_ee = 1
else:
    payment_method_mm = 1


internet_service_dsl = 0
internet_service_fiber = 0
internet_service_no = 0

if internet_service == 'DSL':
    internet_service_dsl = 1
elif internet_service == 'Fiber optic':
    internet_service_fiber = 1
else:
    internet_service_no = 1


# Create a dataframe with the inputs
X = [gender, senior_citizen, partner, dependents, tenure, phone_service, multiple_lines, online_security, online_backup, device_protection, tech_support, streaming_tv, streaming_movies,paperless_billing, monthly_charges, contract_mm, contract_1, contract_2, payment_method_bank, payment_method_cc, payment_method_ee, payment_method_mm, internet_service_dsl, internet_service_fiber, internet_service_no]

# Predict the output
y_pred = classifier_from_pickle.predict([X])

# Write the output to the excel sheet
print(y_pred[0])
sheet['F10'] = y_pred[0]



